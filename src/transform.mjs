const DETAIL_URL = "https://www.planning.act.gov.au/applications-and-assessments/development-applications/browse-das";

export function transformFeatures(features) {
  if (!Array.isArray(features)) {
    throw new TypeError("ArcGIS response did not contain a features array.");
  }

  return features.map(({ attributes: a = {} }) => {
    const daNumber = text(a.DA_NUMBER);
    const amendment = text(a.AMENDMENT_VERSION);
    const proposal = text(a.PROPOSAL_TEXT);
    const buildingType = classifyBuilding(`${text(a.APPLICATION_TYPE)} ${text(a.APPLICATION_SUB_TYPE)} ${proposal}`);
    const priority = classifyPriority(buildingType, proposal);

    return [
      amendment ? `${daNumber}|${amendment}` : daNumber,
      daNumber,
      amendment,
      DETAIL_URL,
      dateOnly(a.LODGEMENT_DATE),
      dateOnly(a.DATE_END),
      text(a.DA_STAGE),
      priority,
      buildingType,
      text(a.DISTRICT),
      text(a.SUBURB),
      text(a.STREET_ADDRESS),
      numberOrBlank(a.BLOCK),
      numberOrBlank(a.SECTION),
      proposal,
      text(a.COMPANYORG_NAME),
      text(a.APPLICANT_NAME),
      text(a.APPLICATION_TYPE),
      text(a.APPLICATION_SUB_TYPE),
      "",
      dateTimeMinute(a.DATECREATED),
      numberOrBlank(a.CENTROID_LAT),
      DETAIL_URL,
      "",
      numberOrBlank(a.CENTROID_LONG),
      numberOrBlank(a.OBJECTID),
      "",
      ""
    ];
  });
}

function text(value) {
  return value === null || value === undefined ? "" : String(value);
}

function numberOrBlank(value) {
  return typeof value === "number" && Number.isFinite(value) ? value : "";
}

function dateOnly(value) {
  return typeof value === "number" && Number.isFinite(value)
    ? new Date(value).toISOString().slice(0, 10)
    : "";
}

function dateTimeMinute(value) {
  return typeof value === "number" && Number.isFinite(value)
    ? new Date(value).toISOString().slice(0, 16).replace("T", " ")
    : "";
}

export function classifyBuilding(source) {
  const value = source.toUpperCase();
  if (value.includes("WAREHOUSE") || value.includes("INDUSTRIAL")) return "Warehouse";
  if (value.includes("APARTMENT") || value.includes("MULTI UNIT") || value.includes("MULTI-UNIT")) return "Apartments";
  if (value.includes("COMMERCIAL") || value.includes("OFFICE")) return "Commercial";
  if (value.includes("RETAIL") || value.includes("SHOP")) return "Retail";
  if (value.includes("SCHOOL") || value.includes("EDUCATION")) return "Education";
  if (value.includes("HOSPITAL") || value.includes("HEALTH") || value.includes("MEDICAL")) return "Health";
  if (value.includes("COMMUNITY")) return "Community";
  if (value.includes("SUBSTATION") || value.includes("INFRASTRUCTURE")) return "Infrastructure";
  if (value.includes("LEASE") || value.includes("LAND")) return "Land / lease";
  if (value.includes("DWELLING") || value.includes("RESIDENTIAL") || value.includes("HOUSE")) return "Residential";
  return "Other";
}

export function classifyPriority(buildingType, proposal) {
  const value = proposal.toUpperCase();
  if (
    ["Warehouse", "Apartments", "Commercial", "Retail", "Infrastructure"].includes(buildingType) ||
    value.includes("SOLAR") ||
    value.includes("BATTERY") ||
    value.includes("SUBSTATION") ||
    value.includes("ELECTRIC VEHICLE")
  ) {
    return "High";
  }
  if (["Education", "Health", "Community"].includes(buildingType)) return "Medium";
  return "Low";
}
