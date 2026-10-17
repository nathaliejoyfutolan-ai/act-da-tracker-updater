import crypto from "node:crypto";
import fs from "node:fs";

const GRAPH = "https://graph.microsoft.com/v1.0";
const WORKBOOK_PATH = "/ACT_DA_Opportunity_Tracker_Interactive.xlsx";
const required = ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_REFRESH_TOKEN", "WORKBOOK_INSPECT_KEY"];

for (const key of required) {
  if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);
}

const accessToken = await refreshAccessToken();
const root = `/me/drive/root:${encodePath(WORKBOOK_PATH)}:/workbook`;
let sessionId = "";
const session = await graph(`${root}/createSession`, {
  method: "POST",
  body: { persistChanges: false },
  useSession: false
});
sessionId = session.id;

try {
  const worksheets = await graph(`${root}/worksheets`);
  const sheetNames = worksheets.value.map((sheet) => sheet.name);
  const requestedRanges = {
    Summary: "A1:Z100",
    "Daily Report": "A1:Z120",
    "DA Register": "A1:AB400",
    "Summary Previous": "A1:Z100",
    Guide: "A1:Z100"
  };
  const ranges = {};
  const charts = {};

  for (const [sheetName, address] of Object.entries(requestedRanges)) {
    if (!sheetNames.includes(sheetName)) continue;
    ranges[sheetName] = await graph(
      `${root}/worksheets/${encodeURIComponent(sheetName)}/range(address='${address}')`
    );
    charts[sheetName] = await graph(`${root}/worksheets/${encodeURIComponent(sheetName)}/charts`);
  }

  const payload = JSON.stringify({ sheetNames, ranges, charts });
  await fs.promises.writeFile("/tmp/workbook-inspection.enc", encrypt(payload, process.env.WORKBOOK_INSPECT_KEY));
  console.log("Encrypted workbook inspection created.");
} finally {
  await graph(`${root}/closeSession`, { method: "POST", body: {}, tolerate404: true });
}

function encrypt(plaintext, keyHex) {
  const key = Buffer.from(keyHex, "hex");
  if (key.length !== 32) throw new Error("WORKBOOK_INSPECT_KEY must be a 32-byte hex key.");
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, ciphertext]).toString("base64");
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
  if (!response.ok) throw new Error(`Microsoft token refresh failed (${response.status}).`);
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

function encodePath(path) {
  return path.split("/").map(encodeURIComponent).join("/");
}
