import { transformFeatures } from "./transform.mjs";

const ARC_GIS_URL = "https://services1.arcgis.com/E5n4f1VY84i0xSjy/arcgis/rest/services/ACTGOV_DAFINDER_LIST_VIEW/FeatureServer/0/query?f=json&where=DA_STAGE%20%3C%3E%20%27Decided%27&outFields=DA_NUMBER%2CAMENDMENT_VERSION%2CLODGEMENT_DATE%2CDATE_END%2CDA_STAGE%2CAPPLICATION_TYPE%2CAPPLICATION_SUB_TYPE%2CDISTRICT%2CSUBURB%2CSTREET_ADDRESS%2CBLOCK%2CSECTION%2CPROPOSAL_TEXT%2CCOMPANYORG_NAME%2CAPPLICANT_NAME%2CDATECREATED%2CCENTROID_LAT%2CCENTROID_LONG%2COBJECTID&returnGeometry=false&resultRecordCount=2000&orderByFields=LODGEMENT_DATE%20DESC%2CDA_NUMBER%20DESC";
const GRAPH = "https://graph.microsoft.com/v1.0";
const WORKBOOK_PATH = "/ACT_DA_Opportunity_Tracker_Interactive.xlsx";
const SHEET_NAME = "ACT Live Feed";
const COLUMN_COUNT = 28;

const required = ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_REFRESH_TOKEN"];
for (const key of required) {
  if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);
}

const accessToken = await refreshAccessToken();
let sessionId = "";
const arcResponse = await fetchJson(ARC_GIS_URL);
if (arcResponse.error) throw new Error(`ArcGIS error: ${arcResponse.error.message ?? "Unknown error"}`);
const rows = transformFeatures(arcResponse.features);
if (rows.some((row) => row.length !== COLUMN_COUNT)) throw new Error("Generated row width does not match the workbook table.");

const root = `/me/drive/root:${encodePath(WORKBOOK_PATH)}:/workbook`;
const session = await graph(`${root}/createSession`, {
  method: "POST",
  body: { persistChanges: true },
  useSession: false
});
sessionId = session.id;

try {
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
  await graph(`${root}/application/calculate`, {
    method: "POST",
    body: { calculationType: "Full" }
  });

  console.log(JSON.stringify({ updatedRows: rows.length, updatedAtUtc: new Date().toISOString() }));
} finally {
  await graph(`${root}/closeSession`, { method: "POST", body: {}, tolerate404: true });
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
