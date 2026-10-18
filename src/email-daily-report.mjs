const GRAPH = "https://graph.microsoft.com/v1.0";
const WORKBOOK_PATH = "/ACT_DA_Opportunity_Tracker_Interactive.xlsx";
const REPORT_SHEET = "Daily Report";
const REPORT_TIME_ZONE = "Asia/Manila";

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

  const [cards, opportunities, questions, workbookItem] = await Promise.all([
    range("A4:G5"),
    range("A11:I16"),
    range("A19:D27"),
    graph(itemRoot, { useSession: false })
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
    opportunities: opportunities.values ?? [],
    questions: questions.values ?? [],
    workbookUrl: workbookItem.webUrl
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

function buildReportHtml({ reportDate, cards, opportunities, questions, workbookUrl }) {
  const cardColumns = [0, 2, 4, 6];
  const cardHtml = cardColumns.map((column) => `
    <td style="padding:10px;border:1px solid #d8e3dc;background:#edf5ea;text-align:center;">
      <div style="font-size:12px;color:#49615a;">${escapeHtml(cards[0]?.[column] ?? "")}</div>
      <div style="font-size:26px;font-weight:700;color:#174f3b;">${escapeHtml(cards[1]?.[column] ?? "0")}</div>
    </td>`).join("");

  const headers = opportunities[0] ?? [];
  const opportunityRows = opportunities.slice(1).filter((row) => row.some(nonBlank));
  const opportunityTable = opportunityRows.length
    ? `<table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr>${headers.map((header) => `<th style="padding:7px;border:1px solid #d6ddd9;background:#2f6d4f;color:#fff;text-align:left;">${escapeHtml(header)}</th>`).join("")}</tr></thead>
        <tbody>${opportunityRows.map((row) => `<tr>${row.map((value, index) => {
          const display = escapeHtml(formatCell(value, index));
          if (index === 0 && nonBlank(value)) {
            return `<td style="padding:7px;border:1px solid #d6ddd9;"><a href="${applicationUrl(value)}">${display}</a></td>`;
          }
          return `<td style="padding:7px;border:1px solid #d6ddd9;">${display}</td>`;
        }).join("")}</tr>`).join("")}</tbody>
      </table>`
    : "<p>No high-priority opportunities are currently listed.</p>";

  const questionRows = questions.slice(1).filter((row) => nonBlank(row[0]));
  const questionHtml = questionRows.map((row) => `
    <tr>
      <td style="padding:8px;border-bottom:1px solid #d6ddd9;font-weight:600;vertical-align:top;">${escapeHtml(row[0])}</td>
      <td style="padding:8px;border-bottom:1px solid #d6ddd9;vertical-align:top;">${escapeHtml(row[3] ?? "")}</td>
    </tr>`).join("");

  return `<!doctype html>
  <html><body style="margin:0;background:#f4f6f5;font-family:Arial,sans-serif;color:#1f2d28;">
    <div style="max-width:1000px;margin:0 auto;padding:24px;">
      <div style="background:#174f3b;color:#fff;padding:18px 22px;">
        <h1 style="margin:0;font-size:24px;">ACT DA Opportunity Tracker</h1>
        <div style="margin-top:5px;color:#d8e9df;">Daily report · ${escapeHtml(reportDate)}</div>
      </div>
      <div style="background:#fff;padding:22px;">
        <table style="width:100%;border-collapse:collapse;table-layout:fixed;"><tr>${cardHtml}</tr></table>
        <h2 style="margin-top:28px;color:#174f3b;">Highest-priority opportunities</h2>
        ${opportunityTable}
        <h2 style="margin-top:28px;color:#174f3b;">Bilal's market questions</h2>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">${questionHtml}</table>
        <p style="margin-top:24px;"><a href="${escapeAttribute(workbookUrl)}" style="display:inline-block;background:#2f6d4f;color:#fff;text-decoration:none;padding:11px 17px;border-radius:4px;">Open the live workbook</a></p>
        <p style="font-size:11px;color:#68756f;">Generated automatically after the ACT DA tracker refresh. Publicly unavailable builder, budget, solar and battery information is not inferred.</p>
      </div>
    </div>
  </body></html>`;
}

function formatCell(value, columnIndex) {
  if (!nonBlank(value)) return "—";
  if (columnIndex === 7 && typeof value === "number") {
    return new Intl.NumberFormat("en-AU", { style: "currency", currency: "AUD", maximumFractionDigits: 0 }).format(value);
  }
  if (columnIndex === 8 && typeof value === "number") return excelDate(value);
  return String(value);
}

function excelDate(serial) {
  const date = new Date(Date.UTC(1899, 11, 30) + Number(serial) * 86400000);
  return new Intl.DateTimeFormat("en-AU", { timeZone: "UTC", day: "2-digit", month: "short", year: "numeric" }).format(date);
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
  const response = await fetch(`${GRAPH}${path}`, {
    method: options.method ?? "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.useSession === false || !sessionId ? {} : { "workbook-session-id": sessionId })
    },
    body: options.body === undefined ? undefined : JSON.stringify(options.body)
  });
  if (options.tolerate404 && response.status === 404) return null;
  if (options.expectedStatus && response.status !== options.expectedStatus) {
    throw new Error(`Graph request failed (${response.status}) for ${path}: ${await response.text()}`);
  }
  if (!options.expectedStatus && !response.ok) {
    throw new Error(`Graph request failed (${response.status}) for ${path}: ${await response.text()}`);
  }
  if (response.status === 204 || response.status === 202) return null;
  return response.json();
}
