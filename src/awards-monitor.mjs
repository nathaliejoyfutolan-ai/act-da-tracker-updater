import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { awards } from "./awards-catalog.mjs";

const GRAPH = "https://graph.microsoft.com/v1.0";
const STATE_PATH = "data/awards-scan-state.json";
const JSON_REPORT = "reports/awards-latest.json";
const HTML_REPORT = "reports/awards-latest.html";
const TIME_ZONE = "Asia/Manila";
const shouldEmail = String(process.env.SEND_EMAIL || "").toLowerCase() === "true";

const previous = await readJson(STATE_PATH, { pages: {} });
const scans = [];
for (const award of awards) scans.push(await scanAward(award, previous.pages?.[award.source]));

const generatedAt = new Date();
const changed = scans.filter((scan) => scan.changed);
const unavailable = scans.filter((scan) => !scan.ok);
const report = {
  generatedAtUtc: generatedAt.toISOString(),
  generatedAtPht: formatDateTime(generatedAt),
  awardsChecked: scans.length,
  pagesChanged: changed.length,
  pagesUnavailable: unavailable.length,
  scans
};

const html = buildEmail(report);
await mkdir("data", { recursive: true });
await mkdir("reports", { recursive: true });
await writeFile(STATE_PATH, `${JSON.stringify({ generatedAtUtc: report.generatedAtUtc, pages: Object.fromEntries(scans.map((scan) => [scan.source, { hash: scan.hash, status: scan.status, checkedAtUtc: scan.checkedAtUtc }])) }, null, 2)}\n`);
await writeFile(JSON_REPORT, `${JSON.stringify(report, null, 2)}\n`);
await writeFile(HTML_REPORT, html);

if (shouldEmail) await sendEmail(html, report);
console.log(JSON.stringify({ checked: scans.length, changed: changed.length, unavailable: unavailable.length, emailed: shouldEmail, generatedAtPht: report.generatedAtPht }));

async function scanAward(award, oldPage) {
  const checkedAtUtc = new Date().toISOString();
  try {
    const response = await fetch(award.source, { headers: { "User-Agent": "DecarbyAwardsMonitor/1.0 (+https://www.decarbysolar.com.au/)" }, redirect: "follow", signal: AbortSignal.timeout(25000) });
    const body = await response.text();
    const text = cleanHtml(body).slice(0, 250000);
    const hash = createHash("sha256").update(text).digest("hex");
    return { ...award, ok: response.ok, status: response.status, changed: Boolean(oldPage?.hash && oldPage.hash !== hash), firstScan: !oldPage?.hash, hash, checkedAtUtc, pageTitle: extractTitle(body), dateSignals: extractDates(text) };
  } catch (error) {
    return { ...award, ok: false, status: 0, changed: false, firstScan: !oldPage?.hash, hash: oldPage?.hash || "", checkedAtUtc, pageTitle: "", dateSignals: [], error: String(error.message || error) };
  }
}

function buildEmail(report) {
  const priority = report.scans.filter((item) => item.recommendation === "Target 2027");
  const changedRows = report.scans.filter((item) => item.changed);
  const changeSection = changedRows.length ? changedRows.map(card).join("") : `<div style="padding:16px;border:1px solid #d9e4e2;border-radius:10px;background:#f7fbfa;color:#46635f;">No monitored official page changed since the previous scan.</div>`;
  const priorityRows = priority.map((item) => `<tr>
    <td style="padding:12px;border-bottom:1px solid #e8eeed;"><a href="${escAttr(item.source)}" style="color:#0f6b68;font-weight:700;text-decoration:none;">${esc(item.name)}</a><div style="font-size:12px;color:#667773;margin-top:3px;">${esc(item.candidate)} · ${esc(item.expectedWindow)}</div></td>
    <td style="padding:12px;border-bottom:1px solid #e8eeed;text-align:center;font-weight:700;color:#17324d;">${item.probability}%</td>
    <td style="padding:12px;border-bottom:1px solid #e8eeed;"><span style="display:inline-block;padding:4px 8px;border-radius:999px;background:${item.eligibilityStatus === "Likely eligible" ? "#e2f0d9" : "#fff2cc"};color:#3d5b45;font-size:12px;">${esc(item.eligibilityStatus)}</span></td>
    <td style="padding:12px;border-bottom:1px solid #e8eeed;color:#43524f;">${esc(item.requirements)}</td>
  </tr>`).join("");
  return `<!doctype html><html><body style="margin:0;background:#eef3f2;font-family:Arial,sans-serif;color:#1f2937;">
  <div style="max-width:920px;margin:0 auto;padding:24px 12px;">
    <div style="background:linear-gradient(135deg,#17324d,#0f6b68);color:#fff;padding:28px;border-radius:16px 16px 0 0;">
      <div style="font-size:12px;letter-spacing:1.5px;text-transform:uppercase;color:#cfe7e4;">Decarby · Awards intelligence</div>
      <h1 style="margin:8px 0 4px;font-size:28px;">Monday awards briefing</h1>
      <div style="color:#dcebea;">${esc(report.generatedAtPht)} · ${report.awardsChecked} official sources checked</div>
    </div>
    <div style="background:#fff;padding:24px;border-radius:0 0 16px 16px;box-shadow:0 8px 24px rgba(23,50,77,.08);">
      <table role="presentation" style="width:100%;border-collapse:separate;border-spacing:8px;margin:-8px;"><tr>
        ${metric("Sources checked", report.awardsChecked)}${metric("Pages changed", report.pagesChanged)}${metric("Needs recheck", report.pagesUnavailable)}
      </tr></table>
      <h2 style="margin:28px 0 12px;color:#17324d;font-size:19px;">What changed</h2>${changeSection}
      <h2 style="margin:28px 0 12px;color:#17324d;font-size:19px;">Five priority applications</h2>
      <table style="width:100%;border-collapse:collapse;font-size:13px;"><thead><tr style="background:#17324d;color:#fff;"><th style="padding:10px;text-align:left;">Award</th><th style="padding:10px;">Win probability</th><th style="padding:10px;text-align:left;">Eligibility</th><th style="padding:10px;text-align:left;">Requirements</th></tr></thead><tbody>${priorityRows}</tbody></table>
      <div style="margin-top:24px;padding:16px;background:#f5f8fa;border-left:4px solid #0f6b68;color:#43524f;"><strong>How to use this:</strong> open any award name to review its official source. Probabilities are internal planning estimates, not organiser-published odds. Friday scans refresh the evidence; Monday scans send this briefing.</div>
      <p style="margin:22px 0 0;font-size:11px;color:#75827f;">Automated by Decarby's GitHub workflow. Eligibility and requirements must be reconfirmed against the live rules before submission.</p>
    </div>
  </div></body></html>`;
}

function card(item) {
  const dates = item.dateSignals.length ? ` Date signals: ${item.dateSignals.slice(0, 4).join(", ")}.` : "";
  return `<div style="padding:16px;margin-bottom:10px;border:1px solid #d9e4e2;border-radius:10px;"><a href="${escAttr(item.source)}" style="color:#0f6b68;font-weight:700;text-decoration:none;">${esc(item.name)}</a><div style="margin-top:6px;font-size:13px;color:#4f615d;">Official page content changed.${esc(dates)} Review the live eligibility and closing date before acting.</div></div>`;
}

function metric(label, value) { return `<td style="width:33%;padding:14px;background:#eaf4f2;border-radius:10px;text-align:center;"><div style="font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:#5f7470;">${esc(label)}</div><div style="font-size:26px;font-weight:700;color:#0f6b68;margin-top:3px;">${esc(value)}</div></td>`; }

async function sendEmail(html, report) {
  for (const key of ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_REFRESH_TOKEN", "REPORT_RECIPIENT"]) if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);
  const accessToken = await refreshAccessToken();
  const workbook = await readFile("outputs/decarby_awards_2027/Decarby_Awards_Pipeline_2027.xlsx");
  const subjectDate = new Intl.DateTimeFormat("en-AU", { timeZone: TIME_ZONE, day: "2-digit", month: "short", year: "numeric" }).format(new Date());
  const response = await fetch(`${GRAPH}/me/sendMail`, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({ message: { subject: `Decarby awards briefing – ${subjectDate}`, body: { contentType: "HTML", content: html }, toRecipients: [{ emailAddress: { address: process.env.REPORT_RECIPIENT } }], attachments: [{ "@odata.type": "#microsoft.graph.fileAttachment", name: "Decarby_Awards_Pipeline_2027.xlsx", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", contentBytes: workbook.toString("base64") }] }, saveToSentItems: true })
  });
  if (response.status !== 202) throw new Error(`Microsoft Graph sendMail failed (${response.status}): ${await response.text()}`);
  console.log(JSON.stringify({ recipient: process.env.REPORT_RECIPIENT, subjectDate, pagesChanged: report.pagesChanged }));
}

async function refreshAccessToken() {
  const body = new URLSearchParams({ client_id: process.env.MS_CLIENT_ID, grant_type: "refresh_token", refresh_token: process.env.MS_REFRESH_TOKEN, scope: "offline_access Files.ReadWrite User.Read Mail.Send" });
  const response = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body });
  if (!response.ok) throw new Error(`Microsoft token refresh failed (${response.status}): ${await response.text()}`);
  return (await response.json()).access_token;
}

async function readJson(path, fallback) { try { return JSON.parse(await readFile(path, "utf8")); } catch { return fallback; } }
function cleanHtml(html) { return html.replace(/<script[\s\S]*?<\/script>/gi, " ").replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/\s+/g, " ").trim(); }
function extractTitle(html) { return cleanHtml(html.match(/<title[^>]*>([\s\S]*?)<\/title>/i)?.[1] || "").slice(0, 200); }
function extractDates(text) { return [...new Set(text.match(/\b(?:[0-3]?\d\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)[,\s]+20\d{2}|20\d{2}-\d{2}-\d{2})\b/gi) || [])].slice(0, 12); }
function formatDateTime(date) { return new Intl.DateTimeFormat("en-AU", { timeZone: TIME_ZONE, weekday: "long", day: "2-digit", month: "long", year: "numeric", hour: "2-digit", minute: "2-digit", timeZoneName: "short" }).format(date); }
function esc(value) { return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;"); }
function escAttr(value) { return esc(value || "#"); }
