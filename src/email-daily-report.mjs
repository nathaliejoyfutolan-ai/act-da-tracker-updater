const GRAPH = "https://graph.microsoft.com/v1.0";
const WORKBOOK_PATH = "/ACT_DA_Opportunity_Tracker_Interactive.xlsx";
const REPORT_SHEET = "Daily Report";
const REPORT_TIME_ZONE = "Asia/Manila";
const refreshStatus = process.env.REPORT_REFRESH_STATUS ?? "fresh";
const ACT_POSTCODES = postcodeMap({
  "2600": "BARTON|CANBERRA|CAPITAL HILL|DEAKIN|HARMAN|HMAS HARMAN|PARKES|RUSSELL|YARRALUMLA",
  "2601": "ACTON|CANBERRA|CITY",
  "2602": "AINSLIE|DICKSON|DOWNER|HACKETT|LYNEHAM|WATSON",
  "2603": "FORREST|GRIFFITH|MANUKA|RED HILL",
  "2604": "CAUSEWAY|KINGSTON|NARRABUNDAH",
  "2605": "CURTIN|GARRAN|HUGHES",
  "2606": "CHIFLEY|LYONS|PHILLIP|WODEN",
  "2607": "FARRER|ISAACS|MAWSON|PEARCE|TORRENS",
  "2609": "CANBERRA AIRPORT|FYSHWICK|MAJURA|PIALLIGO|SYMONSTON",
  "2611": "CHAPMAN|COOMBS|COREE|DENMAN PROSPECT|DUFFY|FISHER|HOLDER|MOLONGLO|RIVETT|STIRLING|STROMLO|URIARRA VILLAGE|WARAMANGA|WESTON|WESTON CREEK|WHITLAM|WRIGHT",
  "2612": "BRADDON|CAMPBELL|REID|TURNER",
  "2614": "ARANDA|COOK|HAWKER|JAMISON CENTRE|MACQUARIE|PAGE|SCULLIN|WEETANGERA",
  "2615": "CHARNWOOD|DUNLOP|FLOREY|FLYNN|FRASER|HIGGINS|HOLT|KIPPAX|LATHAM|MACGREGOR|MACNAMARA|MELBA|SPENCE|STRATHNAIRN",
  "2617": "BELCONNEN|BRUCE|EVATT|GIRALANG|KALEEN|LAWSON|MCKELLAR|UNIVERSITY OF CANBERRA",
  "2618": "HALL",
  "2620": "BEARD|BOOTH|HUME|KOWEN|OAKS ESTATE|PADDYS RIVER|TENNENT|THARWA|WILLIAMSDALE",
  "2900": "GREENWAY",
  "2902": "KAMBAH|KAMBAH VILLAGE",
  "2903": "ERINDALE CENTRE|OXLEY|WANNIASSA",
  "2904": "FADDEN|GOWRIE|MACARTHUR|MONASH",
  "2905": "BONYTHON|CALWELL|CHISHOLM|GILMORE|ISABELLA PLAINS|RICHARDSON|THEODORE",
  "2906": "BANKS|CONDER|GORDON",
  "2911": "CRACE|KENNY|MITCHELL",
  "2912": "GUNGAHLIN",
  "2913": "CASEY|FRANKLIN|NGUNNAWAL|NICHOLLS|PALMERSTON|TAYLOR",
  "2914": "AMAROO|BONNER|FORDE|HARRISON|JACKA|MONCRIEFF|THROSBY"
});

const required = ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_REFRESH_TOKEN", "REPORT_RECIPIENT"];
for (const key of required) {
  if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);
}

const accessToken = await refreshAccessToken();
const encodedPath = WORKBOOK_PATH.split("/").map(encodeURIComponent).join("/");
const itemRoot = `/me/drive/root:${encodedPath}`;
const workbookRoot = `${itemRoot}:/workbook`;
let sessionId = "";

const session = await graph(`${workbookRoot}/createSession`, {
  method: "POST",
  body: { persistChanges: false },
  useSession: false
});
sessionId = session.id;

try {
  await graph(`${workbookRoot}/application/calculate`, {
    method: "POST",
    body: { calculationType: "Full" }
  });

  const [cards, questions, workbookItem, opportunityData] = await Promise.all([
    range("A4:G5"),
    range("A19:D27"),
    graph(itemRoot, { useSession: false }),
    loadOpportunityData()
  ]);

  const reportDate = new Intl.DateTimeFormat("en-AU", {
    timeZone: REPORT_TIME_ZONE,
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric"
  }).format(new Date());

  const html = buildReportHtml({
    reportDate,
    cards: cards.values ?? [],
    opportunities: opportunityData.opportunities,
    questions: questions.values ?? [],
    workbookUrl: workbookItem.webUrl,
    refreshStatus,
    documentLinks: opportunityData.documentLinks
  });

  await graph("/me/sendMail", {
    method: "POST",
    body: {
      message: {
        subject: `ACT DA Daily Report – ${reportDate}`,
        body: { contentType: "HTML", content: html },
        toRecipients: [{ emailAddress: { address: process.env.REPORT_RECIPIENT } }]
      },
      saveToSentItems: true
    },
    useSession: false,
    expectedStatus: 202
  });

  console.log(JSON.stringify({ recipient: process.env.REPORT_RECIPIENT, reportDate, sentAtUtc: new Date().toISOString() }));
} finally {
  await graph(`${workbookRoot}/closeSession`, { method: "POST", body: {}, tolerate404: true });
}

async function range(address) {
  return graph(`${workbookRoot}/worksheets/${encodeURIComponent(REPORT_SHEET)}/range(address='${address}')`);
}

async function loadOpportunityData() {
  const tableRange = await graph(`${workbookRoot}/tables/${encodeURIComponent("DARegisterTable")}/range`);
  const headers = tableRange.values?.[0] ?? [];
  const values = tableRange.values?.slice(1) ?? [];
  const formulas = tableRange.formulas?.slice(1) ?? [];
  const daIndex = headers.indexOf("DA number");
  const countIndex = headers.indexOf("Document count");
  const allDocumentsIndex = headers.indexOf("All documents");
  const costPlanIndex = headers.indexOf("Cost plan document");
  const solarIndex = headers.indexOf("Solar allowance");
  const batteryIndex = headers.indexOf("Battery allowance");
  const evidenceIndex = headers.indexOf("Solar / battery evidence");
  const priorityIndex = headers.indexOf("Priority");
  const stageIndex = headers.indexOf("Stage");
  const lodgedIndex = headers.indexOf("Lodged");
  const addressIndex = headers.indexOf("Address");
  const suburbIndex = headers.indexOf("Suburb");
  const proposalIndex = headers.indexOf("Proposal");
  const buildingTypeIndex = headers.indexOf("Building type");
  const developerIndex = headers.indexOf("Organisation / developer lead");
  const builderIndex = headers.indexOf("Builder");
  const budgetIndex = headers.indexOf("Project budget (AUD)");
  const requiredIndexes = [
    daIndex, countIndex, allDocumentsIndex, costPlanIndex, solarIndex, batteryIndex, evidenceIndex,
    priorityIndex, stageIndex, lodgedIndex, addressIndex, suburbIndex, proposalIndex, buildingTypeIndex,
    developerIndex, builderIndex, budgetIndex
  ];
  if (requiredIndexes.some((index) => index < 0)) {
    throw new Error("DARegisterTable is missing required document-link or document-insight columns");
  }

  const documentLinks = new Map();
  const candidates = [];
  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex] ?? [];
    const daNumber = String(values[rowIndex]?.[daIndex] ?? "").trim();
    if (!daNumber) continue;
    const documents = {
      count: row[countIndex] ?? 0,
      target: hyperlinkUrl(formulas[rowIndex]?.[allDocumentsIndex]),
      costPlan: Boolean(hyperlinkUrl(formulas[rowIndex]?.[costPlanIndex]) || row[costPlanIndex]),
      solar: String(row[solarIndex] ?? ""),
      battery: String(row[batteryIndex] ?? ""),
      evidence: String(row[evidenceIndex] ?? "")
    };
    documentLinks.set(daNumber, documents);
    const stage = String(row[stageIndex] ?? "");
    if (String(row[priorityIndex] ?? "") !== "High" || /^Decided$/i.test(stage)) continue;
    candidates.push({
      score: opportunityScore(documents, stage),
      lodged: sortableDate(row[lodgedIndex]),
      row: [
        daNumber,
        stage,
        projectAddress(row[addressIndex], row[suburbIndex]),
        proposalHighlight(row[proposalIndex]),
        row[buildingTypeIndex] ?? "",
        row[developerIndex] ?? "",
        row[builderIndex] ?? "",
        row[budgetIndex] ?? "",
        `View ${Number(documents.count) || 0} documents`
      ]
    });
  }
  candidates.sort((a, b) => b.score - a.score || b.lodged - a.lodged);
  return {
    documentLinks,
    opportunities: [
      ["DA Number", "Stage", "Project / Address", "Decarby-relevant highlights", "Building type", "Developer / Owner", "Builder", "Project Budget", "Documents"],
      ...candidates.slice(0, 20).map((candidate) => candidate.row)
    ]
  };
}

function opportunityScore(documents, stage) {
  const evidence = `${documents.solar}; ${documents.battery}; ${documents.evidence}`;
  return (documents.costPlan ? 10 : 0)
    + (/solar|photovoltaic|\bpv\b/i.test(evidence) ? 20 : 0)
    + (/batter(?:y|ies)|\bbess\b/i.test(evidence) ? 20 : 0)
    + (/inverter/i.test(evidence) ? 20 : 0)
    + (/^On notification$/i.test(stage) ? 6 : 0)
    + (/^Under assessment$/i.test(stage) ? 5 : 0)
    + (/^Lodged$/i.test(stage) ? 4 : 0);
}

function sortableDate(value) {
  if (typeof value === "number") return value;
  const parsed = Date.parse(String(value ?? ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function postcodeMap(groups) {
  return new Map(Object.entries(groups).flatMap(([postcode, suburbs]) =>
    suburbs.split("|").map((suburb) => [suburb, postcode])
  ));
}

function projectAddress(address, suburb) {
  const street = String(address ?? "").trim();
  const locality = String(suburb ?? "").trim().toUpperCase();
  const postcode = ACT_POSTCODES.get(locality);
  const parts = [street];
  if (locality && !street.toUpperCase().includes(locality)) parts.push(locality);
  parts.push(postcode ? `ACT ${postcode}` : "ACT");
  return parts.filter(Boolean).join(", ");
}

function proposalHighlight(value) {
  const source = String(value ?? "")
    .replace(/\s+/g, " ")
    .replace(/^PUBLIC NOTIFICATION(?: EXTENDED)?(?: UNTIL [^-]+)?\s*-\s*/i, "")
    .replace(/PLANNING ACT 2023\s*-\s*/gi, "")
    .replace(/^(?:AMENDMENT TO (?:DA )?\d+(?:-\w+)?\s*-\s*)?PROPOSAL FOR\s*/i, "")
    .trim();
  if (!source) return "—";

  const segments = source.split(/(?<=[.;])\s+|\s+-\s+/)
    .map((segment) => segment.replace(/^[.;,:\-\s]+|[.;,:\-\s]+$/g, "").trim())
    .filter(Boolean);
  const ranked = segments.map((text, index) => {
    const energy = (text.match(/\b(?:solar|photovoltaic|pv|battery|bess|inverter|electrical|substation|switchboard|ev charg\w*)\b/gi) ?? []).length;
    const asset = (text.match(/\b(?:warehouse|industrial|commercial|apartment|multi-unit|office|retail|hospital|school|education|childcare|building|storey|roof|carport|mechanical|services)\b/gi) ?? []).length;
    const work = (text.match(/\b(?:construction|development|alteration|addition|extension|fitout|demolition|refurbishment|new)\b/gi) ?? []).length;
    return { text, index, score: energy * 10 + asset * 4 + work * 2 };
  });
  const selected = ranked
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, 2)
    .sort((a, b) => a.index - b.index)
    .map((item) => truncateAtWord(item.text, 125));
  return truncateAtWord((selected.length ? selected : [segments[0] ?? source]).join("; "), 230);
}

function truncateAtWord(value, limit) {
  const text = String(value ?? "").trim();
  if (text.length <= limit) return text;
  const shortened = text.slice(0, limit - 1).replace(/\s+\S*$/, "").replace(/[,:;\-\s]+$/, "");
  return `${shortened || text.slice(0, limit - 1)}…`;
}

function buildReportHtml({ reportDate, cards, opportunities, questions, workbookUrl, refreshStatus, documentLinks }) {
  const cardColumns = [0, 2, 4, 6];
  const cardPalettes = [
    ["#eaf2fa", "#214b76", "#cbdced"],
    ["#f4ecff", "#7137a8", "#ddc7f2"],
    ["#fff0f5", "#bd2f72", "#f3c7d9"],
    ["#fff0ee", "#cf3d57", "#f3c8cc"]
  ];
  const cardHtml = cardColumns.map((column, index) => `
    <td style="padding:16px 12px;border:6px solid #ffffff;background:${cardPalettes[index][0]};text-align:center;border-radius:12px;">
      <div style="font-size:11px;line-height:16px;letter-spacing:.5px;text-transform:uppercase;font-weight:700;color:${cardPalettes[index][1]};">${escapeHtml(cards[0]?.[column] ?? "")}</div>
      <div style="margin-top:3px;font-size:30px;line-height:34px;font-weight:800;color:${cardPalettes[index][1]};">${escapeHtml(cards[1]?.[column] ?? "0")}</div>
      <div style="width:34px;height:3px;margin:8px auto 0;background:${cardPalettes[index][2]};border-radius:3px;"></div>
    </td>`).join("");

  const headers = [...(opportunities[0] ?? []), "Document insights"];
  const opportunityRows = opportunities.slice(1).filter((row) => row.some(nonBlank));
  const opportunityTable = opportunityRows.length
    ? `<table role="presentation" style="width:100%;border-collapse:separate;border-spacing:0;font-size:11px;line-height:15px;border:1px solid #d9dfeb;border-radius:10px;overflow:hidden;">
        <thead><tr>${headers.map((header) => `<th style="padding:10px 8px;border-right:1px solid #567391;background:#234b75;color:#fff;text-align:left;vertical-align:bottom;font-size:10px;line-height:13px;letter-spacing:.3px;text-transform:uppercase;">${escapeHtml(header)}</th>`).join("")}</tr></thead>
        <tbody>${opportunityRows.map((row, rowIndex) => {
          const daNumber = String(row[0] ?? "").trim();
          const documents = documentLinks.get(daNumber);
          const background = rowIndex % 2 === 0 ? "#ffffff" : "#f8f7fc";
          const cells = row.map((value, index) => {
          const displayValue = index === 8 && documents
            ? `View ${Number(documents.count) || 0} documents`
            : formatCell(value, index);
          const display = escapeHtml(displayValue);
          if (index === 0 && nonBlank(value)) {
            return `<td style="padding:9px 8px;border-right:1px solid #e4e3eb;border-bottom:1px solid #e4e3eb;background:${background};vertical-align:top;"><a href="${escapeAttribute(applicationUrl(value))}" style="color:#7a35b2;font-weight:800;text-decoration:none;border-bottom:1px solid #c9a9df;">${display}</a></td>`;
          }
          if (index === 1) {
            return `<td style="padding:9px 8px;border-right:1px solid #e4e3eb;border-bottom:1px solid #e4e3eb;background:${background};vertical-align:top;">${stageBadge(displayValue)}</td>`;
          }
          if (index === 3) {
            return `<td style="padding:9px 8px;border-right:1px solid #e4e3eb;border-bottom:1px solid #e4e3eb;background:${background};vertical-align:top;color:#303746;">
              <div>${display}</div>
              <a href="${escapeAttribute(applicationUrl(daNumber))}" style="display:inline-block;margin-top:8px;padding:5px 9px;border-radius:14px;background:#fff0f5;color:#bd2f72;font-weight:800;text-decoration:none;white-space:nowrap;">View project proposal&nbsp;→</a>
            </td>`;
          }
          if (index === 8 && documents) {
            const destination = workbookLocationUrl(workbookUrl, documents.target) || workbookUrl;
            return `<td style="padding:9px 8px;border-right:1px solid #e4e3eb;border-bottom:1px solid #e4e3eb;background:${background};vertical-align:top;"><a href="${escapeAttribute(destination)}" style="display:inline-block;padding:6px 9px;border-radius:14px;background:#f2e8fb;color:#7137a8;font-weight:800;text-decoration:none;white-space:nowrap;">${display} →</a></td>`;
          }
          return `<td style="padding:9px 8px;border-right:1px solid #e4e3eb;border-bottom:1px solid #e4e3eb;background:${background};vertical-align:top;color:#303746;">${display}</td>`;
        });
          cells.push(`<td style="padding:9px 8px;border-bottom:1px solid #e4e3eb;background:${background};vertical-align:top;">${insightBadges(documents)}</td>`);
          return `<tr>${cells.join("")}</tr>`;
        }).join("")}</tbody>
      </table>`
    : "<p>No high-priority opportunities are currently listed.</p>";

  const questionRows = questions.slice(1).filter((row) => nonBlank(row[0]));
  const questionHtml = questionRows.map((row, index) => `
    <tr>
      <td style="width:34%;padding:12px 14px;border-bottom:1px solid #e4e3eb;background:${index % 2 === 0 ? "#f8f7fc" : "#ffffff"};font-weight:700;color:#234b75;vertical-align:top;">${escapeHtml(row[0])}</td>
      <td style="padding:12px 14px;border-bottom:1px solid #e4e3eb;background:${index % 2 === 0 ? "#f8f7fc" : "#ffffff"};color:#444a59;vertical-align:top;">${escapeHtml(row[3] ?? "")}</td>
    </tr>`).join("");
  const refreshNotice = refreshStatus === "fresh" ? "" : `
    <p style="margin:0 0 18px;padding:12px 14px;background:#fff0f5;color:#9e285f;border:1px solid #edbfd2;border-radius:8px;">
      The latest refresh was temporarily unavailable. This report uses the most recent saved workbook data.
    </p>`;

  return `<!doctype html>
  <html><body style="margin:0;background:#f2f2f7;font-family:Arial,Helvetica,sans-serif;color:#242a37;">
    <div style="max-width:1180px;margin:0 auto;padding:28px 14px;">
      <div style="background:#234b75;background-image:linear-gradient(125deg,#234b75 0%,#7137a8 58%,#d53083 82%,#ef4660 100%);color:#fff;padding:30px 30px 27px;border-radius:14px 14px 0 0;">
        <div style="font-size:11px;line-height:16px;letter-spacing:2px;text-transform:uppercase;color:#ffd7e4;font-weight:800;">Decarby Market Intelligence</div>
        <h1 style="margin:7px 0 0;font-size:28px;line-height:34px;font-weight:800;">ACT DA Opportunity Tracker</h1>
        <div style="margin-top:7px;color:#dcebe4;font-size:14px;">Daily opportunity brief · ${escapeHtml(reportDate)}</div>
      </div>
      <div style="background:#fff;padding:24px 26px 30px;border-radius:0 0 14px 14px;box-shadow:0 10px 28px rgba(35,75,117,.12);">
        ${refreshNotice}
        <table role="presentation" style="width:100%;border-collapse:collapse;table-layout:fixed;margin:0 0 26px;"><tr>${cardHtml}</tr></table>
        <div style="margin:6px 0 14px;border-left:4px solid #e43f74;padding-left:12px;">
          <h2 style="margin:0;color:#234b75;font-size:20px;line-height:26px;">Top 20 highest-priority opportunities</h2>
          <div style="margin-top:3px;color:#6d7180;font-size:12px;">Ranked for Decarby relevance, live stage and available document evidence.</div>
        </div>
        ${opportunityTable}
        <div style="margin:30px 0 14px;border-left:4px solid #e43f74;padding-left:12px;">
          <h2 style="margin:0;color:#234b75;font-size:20px;line-height:26px;">Bilal's market questions</h2>
          <div style="margin-top:3px;color:#6d7180;font-size:12px;">A concise readout from the latest public application data.</div>
        </div>
        <table role="presentation" style="width:100%;border-collapse:separate;border-spacing:0;font-size:13px;border:1px solid #d9dfeb;border-radius:10px;overflow:hidden;">${questionHtml}</table>
        <p style="margin:28px 0 14px;text-align:center;"><a href="${escapeAttribute(workbookUrl)}" style="display:inline-block;background:#7137a8;background-image:linear-gradient(90deg,#7137a8,#d53083,#ef4660);color:#fff;text-decoration:none;padding:13px 22px;border-radius:24px;font-weight:800;box-shadow:0 5px 14px rgba(113,55,168,.24);">Open the live workbook&nbsp;&nbsp;→</a></p>
        <p style="margin:0;text-align:center;font-size:10px;line-height:15px;color:#7b7f8b;">Generated automatically after the ACT DA tracker refresh. Publicly unavailable builder, budget, solar, battery and inverter information is never inferred.</p>
      </div>
    </div>
  </body></html>`;
}

function stageBadge(value) {
  const stage = String(value ?? "");
  const palette = /^On notification$/i.test(stage)
    ? ["#fff0f5", "#bd2f72"]
    : /^Under assessment$/i.test(stage)
      ? ["#eee8fa", "#7137a8"]
      : ["#eaf2fa", "#214b76"];
  return `<span style="display:inline-block;padding:5px 8px;border-radius:13px;background:${palette[0]};color:${palette[1]};font-weight:800;white-space:nowrap;">${escapeHtml(stage || "—")}</span>`;
}

function insightBadges(documents) {
  const insight = documentInsight(documents);
  if (/^No /i.test(insight)) return `<span style="color:#809087;">${escapeHtml(insight)}</span>`;
  const palettes = {
    "Cost plan available": ["#f4ecff", "#7137a8"],
    "Solar mentioned": ["#fff0ee", "#cf3d57"],
    "Battery mentioned": ["#eaf2fa", "#214b76"],
    "Inverter mentioned": ["#fff0f5", "#bd2f72"]
  };
  return insight.split("; ").map((item) => {
    const palette = palettes[item] ?? ["#eef2f0", "#4d5d55"];
    return `<span style="display:inline-block;margin:0 4px 4px 0;padding:4px 7px;border-radius:12px;background:${palette[0]};color:${palette[1]};font-weight:700;white-space:nowrap;">${escapeHtml(item)}</span>`;
  }).join("");
}

function formatCell(value, columnIndex) {
  if (!nonBlank(value)) return "—";
  if (columnIndex === 7 && typeof value === "number") {
    return new Intl.NumberFormat("en-AU", { style: "currency", currency: "AUD", maximumFractionDigits: 0 }).format(value);
  }
  return String(value);
}

function hyperlinkUrl(formula) {
  const match = String(formula ?? "").match(/^=HYPERLINK\("((?:[^"]|"")*)"/i);
  return match ? match[1].replaceAll('""', '"') : "";
}

function workbookLocationUrl(workbookUrl, internalTarget) {
  if (!workbookUrl || !internalTarget) return "";
  const destination = new URL(workbookUrl);
  destination.searchParams.set("ActiveCell", String(internalTarget).replace(/^#/, ""));
  return destination.toString();
}

function documentInsight(documents) {
  if (!documents) return "No document insight available yet";
  const evidence = `${documents.solar}; ${documents.battery}; ${documents.evidence}`;
  const insights = [];
  if (documents.costPlan) insights.push("Cost plan available");
  if (/solar|photovoltaic|\bpv\b/i.test(evidence)) insights.push("Solar mentioned");
  if (/batter(?:y|ies)|\bbess\b/i.test(evidence)) insights.push("Battery mentioned");
  if (/inverter/i.test(evidence)) insights.push("Inverter mentioned");
  return insights.length ? [...new Set(insights)].join("; ") : "No relevant mention identified yet";
}

function applicationUrl(daNumber) {
  const value = encodeURIComponent(String(daNumber).trim());
  return `https://www.planning.act.gov.au/applications-and-assessments/development-applications/browse-das/development-application-details?da-number=${value}`;
}

function nonBlank(value) {
  return value !== "" && value !== null && value !== undefined;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value ?? "#");
}

async function refreshAccessToken() {
  const body = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: process.env.MS_REFRESH_TOKEN,
    scope: "offline_access Files.ReadWrite User.Read Mail.Send"
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
  const requestBody = options.body === undefined ? undefined : JSON.stringify(options.body);
  for (let attempt = 1; attempt <= 4; attempt += 1) {
    const response = await fetch(`${GRAPH}${path}`, {
      method: options.method ?? "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...(options.useSession === false || !sessionId ? {} : { "workbook-session-id": sessionId })
      },
      body: requestBody
    });
    if (options.tolerate404 && response.status === 404) return null;
    const succeeded = options.expectedStatus ? response.status === options.expectedStatus : response.ok;
    if (succeeded) return response.status === 204 || response.status === 202 ? null : response.json();

    const details = await response.text();
    const retryable = response.status === 429 || response.status >= 500;
    if (!retryable || attempt === 4) {
      throw new Error(`Graph request failed (${response.status}) for ${path}: ${details}`);
    }
    const retryAfter = Number(response.headers.get("retry-after"));
    const delayMs = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 1000 * 2 ** (attempt - 1);
    await delay(delayMs);
  }
  throw new Error(`Graph request failed after retries: ${path}`);
}

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}
