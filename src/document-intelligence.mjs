import pdfParse from "pdf-parse/lib/pdf-parse.js";

const DOCUMENT_BASE = "https://dafinder.blob.core.windows.net/dadocuments/DOCs";

export function documentUrl(name) {
  const clean = String(name ?? "").trim();
  return clean ? `${DOCUMENT_BASE}/${encodeURIComponent(clean)}` : "";
}

export function extractBudget(text) {
  const source = normaliseText(text);
  const labels = [
    "construction works total",
    "total development cost",
    "total estimated cost",
    "estimated development cost",
    "grand total"
  ];
  const candidates = [];
  for (const label of labels) {
    let offset = 0;
    while ((offset = source.indexOf(label, offset)) >= 0) {
      const window = source.slice(offset + label.length, offset + label.length + 260);
      const values = [...window.matchAll(/(?:\$\s*)?((?:\d{1,3}(?:,\d{3})+)|(?:\d{6,}))(?:\.\d{1,2})?/g)]
        .map((match) => Number(match[1].replaceAll(",", "")))
        .filter((value) => Number.isFinite(value) && value >= 100000 && value <= 100000000000);
      if (values.length) candidates.push(Math.max(...values));
      offset += label.length;
    }
  }
  return candidates.length ? Math.max(...candidates) : "";
}

export function extractBuilder(text) {
  const source = String(text ?? "").replace(/\r/g, "");
  const patterns = [
    /(?:builder|principal contractor|head contractor)\s*(?:name)?\s*[:\-]\s*([^\n]{2,100})/gi,
    /(?:builder|principal contractor|head contractor)\s*\n\s*([^\n]{2,100})/gi
  ];
  for (const pattern of patterns) {
    const match = pattern.exec(source);
    if (!match) continue;
    const value = match[1].replace(/\s{2,}/g, " ").trim();
    if (!/^(n\/?a|not known|not appointed|tbc|to be confirmed|unknown)$/i.test(value)) return value;
  }
  return "";
}

export function extractEnergyEvidence(text) {
  const source = normaliseText(text);
  const evidence = [];
  if (/\bsolar\b|photovoltaic|\bpv\b/.test(source)) evidence.push("Solar mentioned in DA documents");
  if (/\bbatter(?:y|ies)\b|battery energy storage|\bbess\b/.test(source)) evidence.push("Battery mentioned in DA documents");
  return evidence;
}

export async function reviewPdfDocument(document, fetchImpl = fetch) {
  const url = documentUrl(document.name);
  if (!url) return { text: "", url };
  const response = await fetchImpl(url);
  if (!response.ok) throw new Error(`Document download failed (${response.status}) for ${url}`);
  const buffer = Buffer.from(await response.arrayBuffer());
  const parsed = await pdfParse(buffer);
  return { text: parsed.text ?? "", url };
}

function normaliseText(value) {
  return String(value ?? "").toLowerCase().replace(/\s+/g, " ");
}
