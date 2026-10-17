import {
  classifyBuilding,
  classifyPriority,
  directApplicationUrl,
  normalizeStage,
  transformFeatures
} from "./transform.mjs";
import {
  documentUrl,
  extractBudget,
  extractBuilder,
  extractEnergyEvidence,
  reviewPdfDocument
} from "./document-intelligence.mjs";

const ARC_GIS_FIELDS = [
  "DA_NUMBER", "AMENDMENT_VERSION", "LODGEMENT_DATE", "DATE_START", "DATE_END", "DA_STAGE",
  "APPLICATION_TYPE", "APPLICATION_SUB_TYPE", "DISTRICT", "SUBURB", "STREET_ADDRESS", "BLOCK",
  "SECTION", "BLOCK_KEY", "PROPOSAL_TEXT", "COMPANYORG_NAME", "APPLICANT_NAME", "DECISION", "DECISION_DATE",
  "DATECREATED", "PUBLICLY_NOTIFIED", "APPLICATION_AMENDED", "APPLICATION_CORRECTED",
  "REPRESENTATION_MADE", "NOT_NOTIFIED_REASON", "REASON_EXTENSION_PERIOD", "RECONSIDERATION",
  "ACAT_REVIEW", "REPRESENTATION_COUNT", "PRE_DECISION_ADVICE_DATE", "PREDECISION_DUE_DATE",
  "COMMENTS", "ENTITY_ADVICE", "CENTROID_LAT", "CENTROID_LONG", "OBJECTID"
];
const ARC_GIS_URL = `https://services1.arcgis.com/E5n4f1VY84i0xSjy/arcgis/rest/services/ACTGOV_DAFINDER_LIST_VIEW/FeatureServer/0/query?${new URLSearchParams({
  f: "json",
  where: "DA_STAGE <> 'Decided'",
  outFields: ARC_GIS_FIELDS.join(","),
  returnGeometry: "false",
  resultRecordCount: "2000",
  orderByFields: "LODGEMENT_DATE DESC,DA_NUMBER DESC"
})}`;
const ARC_GIS_DOCUMENT_URL = "https://services1.arcgis.com/E5n4f1VY84i0xSjy/arcgis/rest/services/ACTGOV_DAFINDER_DETAILED_VIEW/FeatureServer/0/query";
const GRAPH = "https://graph.microsoft.com/v1.0";
const WORKBOOK_PATH = "/ACT_DA_Opportunity_Tracker_Interactive.xlsx";
const SHEET_NAME = "ACT Live Feed";
const COLUMN_COUNT = 28;

// System-populated fields stay on the left; user-maintained workflow fields
// stay on the right. Lower-value technical fields remain available but hidden.
const REGISTER_AUTO_VISIBLE_HEADERS = [
  "DA number", "Amendment", "All documents", "Stage", "Priority", "Address", "Suburb", "District",
  "Building type", "Proposal", "Lodged", "Notification closes", "Days to notification close",
  "Organisation / developer lead", "Applicant", "Builder", "Project budget (AUD)", "Solar allowance",
  "Battery allowance", "Cost plan document", "Document review", "Decision", "Decision date"
];
const REGISTER_AUTO_HIDDEN_HEADERS = [
  "Record Key", "Application type", "Application subtype", "Block", "Section", "Block key",
  "Representation starts", "Application amended", "Publicly notified", "Document count", "Document pack",
  "Document intelligence checked", "Builder source", "Budget source", "Solar / battery evidence",
  "Source last updated", "Last checked", "Latitude", "Longitude"
];
const REGISTER_MANUAL_HEADERS = [
  "Next action", "Action owner", "Follow-up due", "Follow-up status", "Tender status", "Proposal status",
  "Estimated fee (AUD)", "Relationship / contact", "Last contact", "Notes"
];
const REGISTER_HEADERS = [...REGISTER_AUTO_VISIBLE_HEADERS, ...REGISTER_AUTO_HIDDEN_HEADERS, ...REGISTER_MANUAL_HEADERS];

const required = ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_REFRESH_TOKEN"];
for (const key of required) {
  if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);
}

const accessToken = await refreshAccessToken();
let sessionId = "";
const arcResponse = await fetchJson(ARC_GIS_URL);
if (arcResponse.error) throw new Error(`ArcGIS error: ${arcResponse.error.message ?? "Unknown error"}`);
const rows = transformFeatures(arcResponse.features);
const documentsByDa = await fetchRelevantDocuments(arcResponse.features);
if (rows.some((row) => row.length !== COLUMN_COUNT)) throw new Error("Generated row width does not match the workbook table.");

const root = `/me/drive/root:${encodePath(WORKBOOK_PATH)}:/workbook`;
const session = await graph(`${root}/createSession`, {
  method: "POST",
  body: { persistChanges: true },
  useSession: false
});
sessionId = session.id;

try {
  const documentIndex = await synchroniseDocumentsSheet(root, documentsByDa);
  const tables = await graph(`${root}/worksheets/${encodeURIComponent(SHEET_NAME)}/tables`);
  if (!tables.value?.length) throw new Error(`No table found on worksheet ${SHEET_NAME}.`);
  const tableId = encodeURIComponent(tables.value[0].id);
  const tableBase = `${root}/tables/${tableId}`;
  let tableRange = await graph(`${tableBase}/range`);
  if (tableRange.columnCount !== COLUMN_COUNT) {
    throw new Error(`Expected ${COLUMN_COUNT} columns, found ${tableRange.columnCount}.`);
  }

  const totalRows = tableRange.rowCount - (tables.value[0].showHeaders === false ? 0 : 1) - (tables.value[0].showTotals ? 1 : 0);
  const difference = rows.length - totalRows;

  if (difference > 0) {
    const blanks = Array.from({ length: difference }, () => Array(COLUMN_COUNT).fill(""));
    await graph(`${tableBase}/rows/add`, { method: "POST", body: { index: null, values: blanks } });
  } else if (difference < 0) {
    for (let index = totalRows - 1; index >= rows.length; index -= 1) {
      await graph(`${tableBase}/rows/${index}`, { method: "DELETE" });
    }
  }

  tableRange = await graph(`${tableBase}/range`);
  const { firstColumn, headerRow } = parseRangeAddress(tableRange.address);
  const lastColumn = columnName(columnNumber(firstColumn) + COLUMN_COUNT - 1);
  const firstDataRow = headerRow + 1;
  const lastDataRow = firstDataRow + rows.length - 1;
  const address = `${firstColumn}${firstDataRow}:${lastColumn}${lastDataRow}`;

  await graph(`${root}/worksheets/${encodeURIComponent(SHEET_NAME)}/range(address='${address}')`, {
    method: "PATCH",
    body: { values: rows }
  });

  const registerStats = await synchroniseRegister(root, arcResponse.features, documentsByDa, documentIndex);
  await repairReports(root, rows.length);
  await graph(`${root}/application/calculate`, {
    method: "POST",
    body: { calculationType: "Full" }
  });

  console.log(JSON.stringify({ updatedRows: rows.length, registerRows: rows.length, documentRows: documentIndex.documentRows, ...registerStats, updatedAtUtc: new Date().toISOString() }));
} finally {
  await graph(`${root}/closeSession`, { method: "POST", body: {}, tolerate404: true });
}

async function synchroniseRegister(root, features, documentsByDa, documentIndex) {
  const sheetName = "DA Register";
  const tables = await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/tables`);
  const table = tables.value?.find((item) => item.name === "DARegisterTable");
  if (!table) throw new Error("DARegisterTable was not found.");
  const tableBase = `${root}/tables/${encodeURIComponent(table.id)}`;
  let tableRange = await graph(`${tableBase}/range`);
  const originalHeaders = tableRange.values?.[0] ?? [];

  const existingRows = (tableRange.values ?? []).slice(1);
  const existingByKey = new Map();
  const keyIndex = originalHeaders.indexOf("Record Key");
  const daIndex = originalHeaders.indexOf("DA number");
  const amendmentIndex = originalHeaders.indexOf("Amendment");
  for (const row of existingRows) {
    const daNumber = textValue(row[daIndex]);
    const amendment = textValue(row[amendmentIndex]);
    const key = textValue(row[keyIndex]) || recordKey(daNumber, amendment);
    if (key) existingByKey.set(key, row);
  }

  const intelligenceByKey = await extractMissingDocumentIntelligence(features, originalHeaders, existingByKey, documentsByDa);

  const dataRows = features.map(({ attributes: a = {} }) => {
    const daNumber = textValue(a.DA_NUMBER);
    const amendment = textValue(a.AMENDMENT_VERSION);
    const key = recordKey(daNumber, amendment);
    const existing = existingByKey.get(key) ?? [];
    const old = (header) => existing[originalHeaders.indexOf(header)] ?? "";
    const proposal = textValue(a.PROPOSAL_TEXT);
    const classifiedBuilding = classifyBuilding(`${textValue(a.APPLICATION_TYPE)} ${textValue(a.APPLICATION_SUB_TYPE)} ${proposal}`);
    const buildingType = old("Building type") || registerBuildingType(classifiedBuilding);
    const priority = old("Priority") || classifyPriority(classifiedBuilding, proposal);
    const organisation = textValue(a.COMPANYORG_NAME);
    const documents = documentsByDa.get(daNumber) ?? [];
    const costDocument = preferredDocument(documents, amendment, isCostDocument);
    const applicationDocument = preferredDocument(documents, amendment, isApplicationDocument);
    const zipDocument = preferredDocument(documents, amendment, isZipDocument);
    const intelligence = intelligenceByKey.get(key) ?? {};
    const projectBudget = firstNonBlank(old("Project budget (AUD)"), old("Project Budget"));
    const previousBuilder = /^Not publicly identified/i.test(textValue(old("Builder"))) ? "" : old("Builder");
    const builder = firstNonBlank(previousBuilder, intelligence.builder, "Not publicly identified (DA stage)");
    const builderSource = firstNonBlank(
      intelligence.builder ? `DA document: ${applicationDocument?.name ?? "public application documents"}` : "",
      old("Builder source"),
      previousBuilder ? "Previously recorded" : "Planning ACT application form does not disclose a builder"
    );
    const sourceValues = {
      "DA number": daNumber,
      Amendment: amendment,
      Lodged: excelDate(a.LODGEMENT_DATE),
      "Notification closes": excelDate(a.DATE_END),
      Stage: normalizeStage(a.DA_STAGE),
      Priority: priority,
      "Building type": buildingType,
      District: textValue(a.DISTRICT),
      Suburb: textValue(a.SUBURB),
      Address: textValue(a.STREET_ADDRESS),
      Block: numberOrBlank(a.BLOCK),
      Section: numberOrBlank(a.SECTION),
      Proposal: proposal,
      "Organisation / developer lead": organisation,
      Applicant: textValue(a.APPLICANT_NAME),
      Builder: builder,
      "Project budget (AUD)": firstNonBlank(old("Project budget (AUD)"), projectBudget, intelligence.budget),
      "Budget source": firstNonBlank(old("Budget source") === "Not found in source feed" ? "" : old("Budget source"), intelligence.budget ? `DA document: ${costDocument?.name ?? "cost plan"}` : "Not found in source feed"),
      "Solar allowance": firstNonBlank(old("Solar allowance") === "Not stated - review documents" ? "" : old("Solar allowance"), intelligence.solar ? "Mentioned in DA documents" : "", proposal.toUpperCase().includes("SOLAR") ? "Mentioned in proposal" : "Not stated - review documents"),
      "Battery allowance": firstNonBlank(old("Battery allowance") === "Not stated - review documents" ? "" : old("Battery allowance"), intelligence.battery ? "Mentioned in DA documents" : "", proposal.toUpperCase().includes("BATTERY") ? "Mentioned in proposal" : "Not stated - review documents"),
      "Solar / battery evidence": firstNonBlank(old("Solar / battery evidence"), old("Evidence"), intelligence.energyEvidence, /(SOLAR|BATTERY)/i.test(proposal) ? "Proposal text" : ""),
      "Document review": firstNonBlank(intelligence.reviewStatus, old("Document review"), priority === "High" ? "Queued" : "Not started"),
      "Next action": old("Next action"),
      "Action owner": old("Action owner"),
      "Follow-up status": old("Follow-up status"),
      "Source last updated": excelDate(a.DATECREATED, true),
      "Last checked": excelDate(Date.now()),
      Latitude: numberOrBlank(a.CENTROID_LAT),
      Longitude: numberOrBlank(a.CENTROID_LONG),
      "Days to notification close": "",
      "Review due": old("Review due"),
      "Follow-up due": firstNonBlank(old("Follow-up due"), old("Follow-up Date")),
      "Tender status": old("Tender status"),
      "Proposal status": old("Proposal status"),
      "Estimated fee (AUD)": firstNonBlank(old("Estimated fee (AUD)"), old("Estimated Fee")),
      "Relationship / contact": old("Relationship / contact"),
      "Last contact": old("Last contact"),
      Notes: old("Notes"),
      "Record Key": key,
      "Application type": textValue(a.APPLICATION_TYPE),
      "Application subtype": textValue(a.APPLICATION_SUB_TYPE),
      "Block key": numberOrBlank(a.BLOCK_KEY),
      "Representation starts": excelDate(a.DATE_START),
      Decision: textValue(a.DECISION),
      "Decision date": excelDate(a.DECISION_DATE),
      "Application amended": textValue(a.APPLICATION_AMENDED),
      "Publicly notified": textValue(a.PUBLICLY_NOTIFIED),
      "Document pack": zipDocument ? documentUrl(zipDocument.name) : "",
      "Cost plan document": costDocument ? documentUrl(costDocument.name) : "",
      "Document intelligence checked": firstNonBlank(intelligence.checked, old("Document intelligence checked")),
      "Builder source": builderSource,
      "Document count": documents.length,
      "All documents": documentIndex.firstRowByDa.get(daNumber) ? `#'DA Documents'!A${documentIndex.firstRowByDa.get(daNumber)}` : ""
    };
    return REGISTER_HEADERS.map((header) => safeCellValue(sourceValues[header] ?? old(header)));
  });

  const writableRowCount = await resizeTableRows(tableBase, tableRange, dataRows.length, originalHeaders.length);
  const rowsToWrite = dataRows.concat(
    Array.from({ length: writableRowCount - dataRows.length }, () => Array(REGISTER_HEADERS.length).fill(""))
  );
  if (!sameHeaders(originalHeaders, REGISTER_HEADERS)) {
    await rebuildTableColumns(tableBase, rowsToWrite, REGISTER_HEADERS);
  }
  tableRange = await graph(`${tableBase}/range`);
  const headers = tableRange.values?.[0] ?? [];
  const { firstColumn, headerRow } = parseRangeAddress(tableRange.address);
  const firstDataRow = headerRow + 1;
  const lastDataRow = firstDataRow + rowsToWrite.length - 1;
  const lastColumn = columnName(columnNumber(firstColumn) + headers.length - 1);
  await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='${firstColumn}${firstDataRow}:${lastColumn}${lastDataRow}')`, {
    method: "PATCH",
    body: { values: rowsToWrite }
  });

  await writeColumnFormulas(root, sheetName, headers, firstColumn, firstDataRow, dataRows, "DA number", ({ row, feature }) => {
    const a = feature.attributes ?? {};
    return `=HYPERLINK("${formulaText(directApplicationUrl(a.DA_NUMBER, a.AMENDMENT_VERSION))}","${formulaText(a.DA_NUMBER)}")`;
  }, features);
  await writeOptionalLinkColumn(root, sheetName, headers, firstColumn, firstDataRow, dataRows, "Document pack", "Open documents");
  await writeOptionalLinkColumn(root, sheetName, headers, firstColumn, firstDataRow, dataRows, "Cost plan document", "Open cost plan");
  await writeColumnFormulas(root, sheetName, headers, firstColumn, firstDataRow, dataRows, "All documents", ({ feature }) => {
    const daNumber = textValue(feature.attributes?.DA_NUMBER);
    const row = documentIndex.firstRowByDa.get(daNumber);
    const count = documentIndex.countByDa.get(daNumber) ?? 0;
    return row ? `=HYPERLINK("#'DA Documents'!A${row}","View ${count} documents")` : "";
  }, features);
  const notificationColumn = columnName(columnNumber(firstColumn) + headers.indexOf("Notification closes"));
  await writeColumnFormulas(root, sheetName, headers, firstColumn, firstDataRow, dataRows, "Days to notification close", ({ row }) =>
    `=IF(${notificationColumn}${row}="","",${notificationColumn}${row}-TODAY())`, features);

  for (const header of ["Lodged", "Notification closes", "Last checked", "Last contact", "Follow-up due", "Representation starts", "Decision date"]) {
    await setColumnNumberFormat(root, sheetName, headers, firstColumn, firstDataRow, lastDataRow, header, "dd-mmm-yyyy");
  }
  for (const header of ["Project budget (AUD)", "Estimated fee (AUD)"]) {
    await setColumnNumberFormat(root, sheetName, headers, firstColumn, firstDataRow, lastDataRow, header, '"$"#,##0');
  }

  // Column A is an orphaned legacy Record Key copy. The actual DARegisterTable starts in column B.
  if (columnNumber(firstColumn) > 1) {
    const legacyHeader = await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='A1:A1')`);
    if (String(legacyHeader.values?.[0]?.[0] ?? "").trim().toLowerCase() === "record key") {
      await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='A1:A${Math.max(lastDataRow, 1000)}')/clear`, {
        method: "POST",
        body: { applyTo: "Contents" }
      });
    }
  }
  const legacyCell = await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='A1:A1')`);
  await formatRegisterZones(root, sheetName, headers, firstColumn, headerRow, lastDataRow);

  const budgetIndex = headers.indexOf("Project budget (AUD)");
  const documentPackIndex = headers.indexOf("Document pack");
  const costPlanIndex = headers.indexOf("Cost plan document");
  return {
    legacyRecordKeyCleared: String(legacyCell.values?.[0]?.[0] ?? "").trim() === "",
    registerColumns: headers.length,
    visibleAutofillColumns: REGISTER_AUTO_VISIBLE_HEADERS.length,
    hiddenTechnicalColumns: REGISTER_AUTO_HIDDEN_HEADERS.length,
    manualInputColumns: REGISTER_MANUAL_HEADERS.length,
    budgetsRecorded: dataRows.filter((row) => budgetIndex >= 0 && row[budgetIndex] !== "").length,
    documentPacksLinked: dataRows.filter((row) => documentPackIndex >= 0 && row[documentPackIndex] !== "").length,
    costPlansLinked: dataRows.filter((row) => costPlanIndex >= 0 && row[costPlanIndex] !== "").length
  };
}

async function synchroniseDocumentsSheet(root, documentsByDa) {
  const sheetName = "DA Documents";
  const worksheets = await graph(`${root}/worksheets`);
  if (!(worksheets.value ?? []).some((sheet) => sheet.name === sheetName)) {
    await graph(`${root}/worksheets/add`, { method: "POST", body: { name: sheetName } });
  }

  const headers = ["Record Key", "DA number", "Amendment", "Document type", "Document name", "Size (KB)", "Document URL"];
  const rows = [];
  const firstRowByDa = new Map();
  const countByDa = new Map();
  const sortedEntries = [...documentsByDa.entries()].sort(([a], [b]) => Number(b) - Number(a));
  for (const [daNumber, documents] of sortedEntries) {
    firstRowByDa.set(daNumber, rows.length + 2);
    countByDa.set(daNumber, documents.length);
    for (const document of documents) {
      rows.push([
        recordKey(daNumber, document.amendment),
        daNumber,
        document.amendment,
        document.type,
        document.name,
        document.sizeKb,
        documentUrl(document.name)
      ].map(safeCellValue));
    }
  }

  const tables = await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/tables`);
  const table = (tables.value ?? []).find((item) => item.name === "DADocumentsTable") ?? tables.value?.[0];
  const targetLastRow = Math.max(rows.length + 1, 2);
  let previousLastRow = 2;
  if (table) {
    const tableBase = `${root}/tables/${encodeURIComponent(table.id)}`;
    const currentRange = await graph(`${tableBase}/range`);
    previousLastRow = currentRange.rowCount;
    await graph(`${tableBase}/resize`, {
      method: "POST",
      body: { newRange: `'${sheetName}'!A1:G${targetLastRow}` }
    });
  }
  await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='A1:G${Math.max(previousLastRow, targetLastRow)}')/clear`, {
    method: "POST",
    body: { applyTo: "Contents" }
  });
  await patchRange(root, sheetName, "A1:G1", { values: [headers] });
  await patchRowsInChunks(root, sheetName, rows, 2, 1000);

  for (let offset = 0; offset < rows.length; offset += 1000) {
    const chunk = rows.slice(offset, offset + 1000);
    const startRow = offset + 2;
    const endRow = startRow + chunk.length - 1;
    const daFormulas = chunk.map((row) => [
      `=HYPERLINK("${formulaText(directApplicationUrl(row[1], row[2]))}","${formulaText(row[1])}")`
    ]);
    const documentFormulas = chunk.map((row) => [
      `=HYPERLINK("${formulaText(row[6])}","${formulaText(row[4])}")`
    ]);
    await patchRange(root, sheetName, `B${startRow}:B${endRow}`, { formulas: daFormulas });
    await patchRange(root, sheetName, `E${startRow}:E${endRow}`, { formulas: documentFormulas });
  }

  if (!table) {
    await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/tables/add`, {
      method: "POST",
      body: { address: `A1:G${targetLastRow}`, hasHeaders: true }
    });
    const newTables = await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/tables`);
    const created = newTables.value?.[0];
    if (created && created.name !== "DADocumentsTable") {
      await graph(`${root}/tables/${encodeURIComponent(created.id)}`, {
        method: "PATCH",
        body: { name: "DADocumentsTable" }
      });
    }
  }
  return { firstRowByDa, countByDa, documentRows: rows.length };
}

async function patchRowsInChunks(root, sheetName, rows, firstRow, chunkSize) {
  for (let offset = 0; offset < rows.length; offset += chunkSize) {
    const chunk = rows.slice(offset, offset + chunkSize);
    const startRow = firstRow + offset;
    const endRow = startRow + chunk.length - 1;
    await patchRange(root, sheetName, `A${startRow}:G${endRow}`, { values: chunk });
  }
}

async function fetchRelevantDocuments(features) {
  const daNumbers = [...new Set(features.map(({ attributes = {} }) => textValue(attributes.DA_NUMBER)).filter(Boolean))];
  const byDa = new Map();
  for (let start = 0; start < daNumbers.length; start += 35) {
    const numbers = daNumbers.slice(start, start + 35).map((value) => Number(value)).filter(Number.isFinite);
    if (!numbers.length) continue;
    const where = `DA_NUMBER IN (${numbers.join(",")})`;
    let offset = 0;
    while (true) {
      const url = `${ARC_GIS_DOCUMENT_URL}?${new URLSearchParams({
        f: "json",
        where,
        outFields: "DA_NUMBER,AMENDMENT_VERSION,DOCUMENT_NAME,DOCUMENT_TYPE,DOCUMENT_SIZE_KB,OBJECTID",
        returnGeometry: "false",
        resultOffset: String(offset),
        resultRecordCount: "2000",
        orderByFields: "DA_NUMBER,OBJECTID"
      })}`;
      const response = await fetchJson(url);
      if (response.error) throw new Error(`ArcGIS document error: ${response.error.message ?? "Unknown error"}`);
      const page = response.features ?? [];
      for (const { attributes: a = {} } of page) {
        const da = textValue(a.DA_NUMBER);
        const item = {
          amendment: textValue(a.AMENDMENT_VERSION),
          name: textValue(a.DOCUMENT_NAME),
          type: textValue(a.DOCUMENT_TYPE),
          sizeKb: numberOrBlank(a.DOCUMENT_SIZE_KB)
        };
        if (!byDa.has(da)) byDa.set(da, []);
        byDa.get(da).push(item);
      }
      if (!response.exceededTransferLimit && page.length < 2000) break;
      offset += page.length;
      if (!page.length) break;
    }
  }
  return byDa;
}

async function extractMissingDocumentIntelligence(features, headers, existingByKey, documentsByDa) {
  const results = new Map();
  const candidates = [];
  for (const feature of features) {
    const a = feature.attributes ?? {};
    const daNumber = textValue(a.DA_NUMBER);
    const amendment = textValue(a.AMENDMENT_VERSION);
    const key = recordKey(daNumber, amendment);
    const existing = existingByKey.get(key) ?? [];
    const old = (header) => existing[headers.indexOf(header)] ?? "";
    const documents = documentsByDa.get(daNumber) ?? [];
    const cost = preferredDocument(documents, amendment, isCostDocument);
    const application = preferredDocument(documents, amendment, isApplicationDocument);
    const developmentReport = preferredDocument(documents, amendment, isDevelopmentReport);
    const unchanged = Number(old("Document count")) === documents.length && /^Auto-reviewed v2/.test(textValue(old("Document review")));
    if (unchanged || (!cost && !application && !developmentReport)) continue;
    candidates.push({
      key,
      cost,
      application,
      developmentReport,
      priority: cost ? 0 : normalizeStage(a.DA_STAGE) === "On notification" ? 1 : 2
    });
  }
  candidates.sort((a, b) => a.priority - b.priority);
  for (const candidate of candidates.slice(0, 10)) {
    const documents = [...new Map([candidate.cost, candidate.application, candidate.developmentReport].filter(Boolean).map((doc) => [doc.name, doc])).values()];
    let combinedText = "";
    const reviewed = [];
    for (const document of documents) {
      try {
        const parsed = await reviewPdfDocument(document);
        combinedText += `\n${parsed.text}`;
        reviewed.push(document.name);
      } catch (error) {
        console.warn(`Could not review ${document.name}: ${error.message}`);
      }
    }
    const budget = extractBudget(combinedText);
    const builder = extractBuilder(combinedText);
    const energy = extractEnergyEvidence(combinedText);
    results.set(candidate.key, {
      budget,
      builder,
      solar: energy.some((value) => value.startsWith("Solar")),
      battery: energy.some((value) => value.startsWith("Battery")),
      energyEvidence: energy.join("; "),
      evidence: reviewed.length ? `Automatically reviewed: ${reviewed.join(", ")}` : "",
      reviewStatus: reviewed.length ? `Auto-reviewed v2 ${new Date().toISOString().slice(0, 10)}` : "Review queued",
      checked: reviewed.length ? excelDate(Date.now()) : ""
    });
  }
  return results;
}

function documentsForVersion(documents, amendment) {
  const exact = documents.filter((document) => textValue(document.amendment) === textValue(amendment));
  return exact.length ? exact : documents.filter((document) => !textValue(document.amendment));
}

function preferredDocument(documents, amendment, predicate) {
  const exact = documents.filter((document) => textValue(document.amendment) === textValue(amendment) && predicate(document));
  const base = documents.filter((document) => !textValue(document.amendment) && predicate(document));
  return exact.at(-1) ?? base.at(-1) ?? documents.filter(predicate).at(-1);
}

function isCostDocument(document) {
  return /bill of quantities|cost/i.test(`${document.type} ${document.name}`);
}

function isApplicationDocument(document) {
  return /application form/i.test(document.type);
}

function isDevelopmentReport(document) {
  return /development outcomes report/i.test(document.type);
}

function isZipDocument(document) {
  return /zipped file/i.test(document.type) || /^ZIPPED-/i.test(document.name);
}

async function repairReports(root, registerRowCount) {
  const criteria = '(DARegisterTable[Record Key]<>"")*(DARegisterTable[Priority]="High")';
  const indexed = (header, row) => `INDEX(FILTER(DARegisterTable[${header}],${criteria}),ROWS($A$12:A${row}))`;
  const guarded = (header, row) => `IF(${indexed(header, row)}="","",${indexed(header, row)})`;
  const dailyRows = Array.from({ length: 5 }, (_, offset) => {
    const row = 12 + offset;
    return [
      `=IFERROR(${indexed("DA number", row)},"")`,
      `=IFERROR(${indexed("Stage", row)},"")`,
      `=IFERROR(${indexed("Address", row)},"")`,
      `=IFERROR(${guarded("Proposal", row)},"")`,
      `=IFERROR(${guarded("Building type", row)},"")`,
      `=IFERROR(${guarded("Organisation / developer lead", row)},"")`,
      `=IFERROR(${guarded("Builder", row)},"")`,
      `=IFERROR(${guarded("Project budget (AUD)", row)},"")`,
      `=IFERROR(${guarded("Follow-up due", row)},"")`
    ];
  });
  await patchRange(root, "Daily Report", "A11:I11", {
    values: [["DA Number", "Stage", "Project / Address", "Proposal", "Building type", "Developer / Owner", "Builder", "Project Budget", "Follow-up Date"]]
  });
  await patchRange(root, "Daily Report", "A12:I16", { formulas: dailyRows });
  await patchRange(root, "Daily Report", "I12:I16", { numberFormat: Array.from({ length: 5 }, () => ["dd-mmm-yyyy"]) });
  await patchRange(root, "Daily Report", "H12:H16", { numberFormat: Array.from({ length: 5 }, () => ['"$"#,##0']) });

  await patchRange(root, "Daily Report", "J4:J7", { values: [["Under assessment"], ["Lodged"], ["On notification"], [""]] });
  await patchRange(root, "Daily Report", "K4:K7", {
    formulas: [["=COUNTIFS(DARegisterTable[Stage],J4)"], ["=COUNTIFS(DARegisterTable[Stage],J5)"], ["=COUNTIFS(DARegisterTable[Stage],J6)"], [""]]
  });

  await patchRange(root, "Daily Report", "A19:A27", { values: [
    ["Bilal's daily market questions"],
    ["What is happening around the ACT?"],
    ["What buildings are coming up?"],
    ["Which owners or developers are identified?"],
    ["Which builders are identified?"],
    ["Which project budgets are known?"],
    ["Is solar allowed for?"],
    ["Is battery storage allowed for?"],
    ["What is happening with tenders and proposals?"]
  ] });
  await patchRange(root, "Daily Report", "D20:D27", { formulas: [
    ['="Tracking "&A5&" current applications across "&COUNTIFS(DARegisterTable[District],"<>")&" populated district records."'],
    ['="See the building-type chart; "&COUNTIFS(DARegisterTable[Building type],"<>")&" applications have a classified building type."'],
    ['=COUNTIFS(DARegisterTable[Organisation / developer lead],"<>")&" applications have an owner or developer identified."'],
    ['=COUNTIFS(DARegisterTable[Builder],"<>",DARegisterTable[Builder],"<>Not publicly identified*")&" applications have a builder publicly identified; the others require tender/network follow-up."'],
    ['=COUNTIFS(DARegisterTable[Project budget (AUD)],"<>")&" budgets are recorded, totalling "&TEXT(SUM(DARegisterTable[Project budget (AUD)]),"$#,##0")&"."'],
    ['=COUNTIFS(DARegisterTable[Solar allowance],"Mentioned*")&" applications mention solar; the remaining documents require review."'],
    ['=COUNTIFS(DARegisterTable[Battery allowance],"Mentioned*")&" applications mention battery storage; the remaining documents require review."'],
    ['=COUNTIFS(DARegisterTable[Tender status],"<>Unknown")&" tender statuses and "&COUNTIFS(DARegisterTable[Proposal status],"<>Not started")&" active proposal statuses are recorded."']
  ] });

  const summaryCriteria = (row) => `((($B$16="All")+(DARegisterTable[Priority]=$B$16))>0)*((($D$16="All")+(DARegisterTable[Stage]=$D$16))>0)*(DARegisterTable[Record Key]<>"")`;
  const summaryIndexed = (header, row) => `INDEX(FILTER(DARegisterTable[${header}],${summaryCriteria(row)}),ROWS($A$20:A${row}))`;
  const count = Math.max(registerRowCount, 1);
  const ownerFormulas = [];
  const actionOwnerFormulas = [];
  const followUpFormulas = [];
  for (let offset = 0; offset < count; offset += 1) {
    const row = 20 + offset;
    ownerFormulas.push([`=IFERROR(IF(${summaryIndexed("Organisation / developer lead", row)}="","",${summaryIndexed("Organisation / developer lead", row)}),"")`]);
    actionOwnerFormulas.push([`=IFERROR(IF(${summaryIndexed("Action owner", row)}="","",${summaryIndexed("Action owner", row)}),"")`]);
    followUpFormulas.push([`=IFERROR(IF(${summaryIndexed("Follow-up due", row)}="","",${summaryIndexed("Follow-up due", row)}),"")`]);
  }
  const lastSummaryRow = 19 + count;
  await patchRange(root, "Summary", `F20:F${lastSummaryRow}`, { formulas: ownerFormulas });
  await patchRange(root, "Summary", `H20:H${lastSummaryRow}`, { formulas: actionOwnerFormulas });
  await patchRange(root, "Summary", `I20:I${lastSummaryRow}`, { formulas: followUpFormulas });
  await patchRange(root, "Summary", `I20:I${lastSummaryRow}`, { numberFormat: Array.from({ length: count }, () => ["dd-mmm-yyyy"]) });
}

async function resizeTableRows(tableBase, tableRange, targetRows, columnCount) {
  const currentRows = tableRange.rowCount - 1;
  const difference = targetRows - currentRows;
  if (difference > 0) {
    await graph(`${tableBase}/rows/add`, {
      method: "POST",
      body: { index: null, values: Array.from({ length: difference }, () => Array(columnCount).fill("")) }
    });
  }
  return Math.max(currentRows, targetRows);
}

function sameHeaders(left, right) {
  return left.length === right.length && left.every((header, index) => header === right[index]);
}

async function rebuildTableColumns(tableBase, rows, headers) {
  const originalColumns = (await graph(`${tableBase}/columns`)).value ?? [];
  const temporaryNames = headers.map((_, index) => `__NEW_${String(index + 1).padStart(2, "0")}__`);

  // Insert in reverse at index zero so the final physical order matches headers.
  for (let index = headers.length - 1; index >= 0; index -= 1) {
    await graph(`${tableBase}/columns/add`, {
      method: "POST",
      body: {
        index: 0,
        values: [[temporaryNames[index]], ...rows.map((row) => [row[index]])]
      }
    });
  }

  for (const column of [...originalColumns].reverse()) {
    await graph(`${tableBase}/columns/${encodeURIComponent(column.id)}`, { method: "DELETE" });
  }

  const rebuiltColumns = (await graph(`${tableBase}/columns`)).value ?? [];
  for (let index = 0; index < headers.length; index += 1) {
    const column = rebuiltColumns.find((item) => item.name === temporaryNames[index]);
    if (!column) throw new Error(`Could not find rebuilt register column ${temporaryNames[index]}.`);
    await graph(`${tableBase}/columns/${encodeURIComponent(column.id)}`, {
      method: "PATCH",
      body: { name: headers[index] }
    });
  }
}

async function formatRegisterZones(root, sheetName, headers, firstColumn, headerRow, lastDataRow) {
  const firstNumber = columnNumber(firstColumn);
  const autoVisibleLast = columnName(firstNumber + REGISTER_AUTO_VISIBLE_HEADERS.length - 1);
  const autoHiddenFirst = columnName(firstNumber + REGISTER_AUTO_VISIBLE_HEADERS.length);
  const autoHiddenLast = columnName(firstNumber + REGISTER_AUTO_VISIBLE_HEADERS.length + REGISTER_AUTO_HIDDEN_HEADERS.length - 1);
  const manualFirst = columnName(firstNumber + REGISTER_AUTO_VISIBLE_HEADERS.length + REGISTER_AUTO_HIDDEN_HEADERS.length);
  const manualLast = columnName(firstNumber + headers.length - 1);
  const dataFirstRow = headerRow + 1;

  await patchRange(root, sheetName, `${firstColumn}${headerRow}:${manualLast}${lastDataRow}`, { columnHidden: false });
  await patchRange(root, sheetName, `${autoHiddenFirst}${headerRow}:${autoHiddenLast}${lastDataRow}`, { columnHidden: true });

  await patchRangeFill(root, sheetName, `${firstColumn}${headerRow}:${autoHiddenLast}${headerRow}`, "#1F4E78");
  await patchRangeFont(root, sheetName, `${firstColumn}${headerRow}:${autoHiddenLast}${headerRow}`, { bold: true, color: "#FFFFFF" });
  await patchRangeFill(root, sheetName, `${manualFirst}${headerRow}:${manualLast}${headerRow}`, "#7F6000");
  await patchRangeFont(root, sheetName, `${manualFirst}${headerRow}:${manualLast}${headerRow}`, { bold: true, color: "#FFFFFF" });

  if (lastDataRow >= dataFirstRow) {
    await patchRangeFill(root, sheetName, `${firstColumn}${dataFirstRow}:${autoVisibleLast}${lastDataRow}`, "#DDEBF7");
    await patchRangeFill(root, sheetName, `${manualFirst}${dataFirstRow}:${manualLast}${lastDataRow}`, "#FFF2CC");
  }
}

async function patchRangeFill(root, sheetName, address, color) {
  return graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='${address}')/format/fill`, {
    method: "PATCH",
    body: { color }
  });
}

async function patchRangeFont(root, sheetName, address, body) {
  return graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='${address}')/format/font`, {
    method: "PATCH",
    body
  });
}

async function writeColumnFormulas(root, sheetName, headers, firstColumn, firstDataRow, rows, header, factory, features) {
  const index = headers.indexOf(header);
  if (index < 0 || rows.length === 0) return;
  const column = columnName(columnNumber(firstColumn) + index);
  const formulas = rows.map((_, offset) => [factory({ row: firstDataRow + offset, feature: features[offset] })]);
  await patchRange(root, sheetName, `${column}${firstDataRow}:${column}${firstDataRow + rows.length - 1}`, { formulas });
}

async function writeOptionalLinkColumn(root, sheetName, headers, firstColumn, firstDataRow, rows, header, label) {
  const index = headers.indexOf(header);
  if (index < 0 || rows.length === 0) return;
  const column = columnName(columnNumber(firstColumn) + index);
  const formulas = rows.map((row) => {
    const value = textValue(row[index]);
    return [value ? `=HYPERLINK("${formulaText(value)}","${formulaText(label)}")` : ""];
  });
  await patchRange(root, sheetName, `${column}${firstDataRow}:${column}${firstDataRow + rows.length - 1}`, { formulas });
}

async function setColumnNumberFormat(root, sheetName, headers, firstColumn, firstDataRow, lastDataRow, header, format) {
  const index = headers.indexOf(header);
  if (index < 0) return;
  const column = columnName(columnNumber(firstColumn) + index);
  await patchRange(root, sheetName, `${column}${firstDataRow}:${column}${lastDataRow}`, {
    numberFormat: Array.from({ length: lastDataRow - firstDataRow + 1 }, () => [format])
  });
}

async function patchRange(root, sheetName, address, body) {
  return graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='${address}')`, {
    method: "PATCH",
    body
  });
}

function recordKey(daNumber, amendment) {
  return daNumber ? `${daNumber}|${amendment || "BASE"}` : "";
}

function registerBuildingType(value) {
  return ({
    Warehouse: "Warehouse / industrial",
    Apartments: "Apartments / multi-unit",
    Commercial: "Commercial / office",
    Retail: "Retail / hospitality",
    Education: "Education / childcare",
    Health: "Health / aged care",
    Community: "Community / sport",
    Infrastructure: "Infrastructure / utilities"
  })[value] ?? value;
}

function excelDate(value, includeTime = false) {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  const serial = value / 86400000 + 25569;
  return includeTime ? serial : Math.floor(serial);
}

function numberOrBlank(value) {
  return typeof value === "number" && Number.isFinite(value) ? value : "";
}

function textValue(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function firstNonBlank(...values) {
  return values.find((value) => value !== "" && value !== null && value !== undefined) ?? "";
}

function formulaText(value) {
  return textValue(value).replaceAll('"', '""');
}

function safeCellValue(value) {
  return typeof value === "string" && /^[=+\-@]/.test(value) ? `'${value}` : value;
}

async function refreshAccessToken() {
  const body = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: process.env.MS_REFRESH_TOKEN,
    scope: "offline_access Files.ReadWrite User.Read"
  });
  const response = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });
  if (!response.ok) throw new Error(`Microsoft token refresh failed (${response.status}): ${await response.text()}`);
  return (await response.json()).access_token;
}

async function graph(path, options = {}) {
  const response = await fetch(`${GRAPH}${path}`, {
    method: options.method ?? "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(sessionId && options.useSession !== false ? { "Workbook-Session-Id": sessionId } : {})
    },
    body: options.body === undefined ? undefined : JSON.stringify(options.body)
  });
  if (options.tolerate404 && response.status === 404) return {};
  if (!response.ok) throw new Error(`Graph request failed (${response.status}) ${path}: ${await response.text()}`);
  return response.status === 204 ? {} : response.json();
}

async function fetchJson(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Request failed (${response.status}) for ${url}`);
  return response.json();
}

function encodePath(path) {
  return path.split("/").map(encodeURIComponent).join("/");
}

function parseRangeAddress(address) {
  const match = address.match(/!\$?([A-Z]+)\$?(\d+):/i);
  if (!match) throw new Error(`Could not parse table address: ${address}`);
  return { firstColumn: match[1].toUpperCase(), headerRow: Number(match[2]) };
}

function columnNumber(name) {
  return [...name].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function columnName(number) {
  let result = "";
  for (let value = number; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(((value - 1) % 26) + 65) + result;
  }
  return result;
}
