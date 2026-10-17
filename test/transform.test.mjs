import test from "node:test";
import assert from "node:assert/strict";
import { classifyBuilding, classifyPriority, directApplicationUrl, normalizeStage, transformFeatures } from "../src/transform.mjs";

test("classifies priority opportunities", () => {
  assert.equal(classifyBuilding("new industrial warehouse"), "Warehouse");
  assert.equal(classifyPriority("Warehouse", "ordinary project"), "High");
  assert.equal(classifyPriority("Other", "battery installation"), "High");
});

test("maps ArcGIS attributes to the workbook's 28 columns", () => {
  const rows = transformFeatures([{ attributes: {
    DA_NUMBER: "202600001",
    AMENDMENT_VERSION: "A",
    LODGEMENT_DATE: Date.UTC(2026, 0, 2),
    DATE_END: null,
    DA_STAGE: "Lodged",
    APPLICATION_TYPE: "Commercial",
    APPLICATION_SUB_TYPE: "Office",
    DISTRICT: "CITY",
    SUBURB: "CITY",
    STREET_ADDRESS: "1 Example St",
    BLOCK: 1,
    SECTION: 2,
    PROPOSAL_TEXT: "New office",
    COMPANYORG_NAME: "Builder",
    APPLICANT_NAME: "Applicant",
    DATECREATED: Date.UTC(2026, 0, 2, 3, 4),
    CENTROID_LAT: -35.2,
    CENTROID_LONG: 149.1,
    OBJECTID: 10
  } }]);

  assert.equal(rows.length, 1);
  assert.equal(rows[0].length, 28);
  assert.equal(rows[0][0], "202600001|A");
  assert.equal(new URL(rows[0][3]).hostname, "www.planning.act.gov.au");
  assert.equal(rows[0][7], "High");
  assert.equal(rows[0][8], "Commercial");
  assert.equal(rows[0][20], "2026-01-02 03:04");
});

test("normalizes stage labels and creates absolute ACT application links", () => {
  assert.equal(normalizeStage("Under Assessment"), "Under assessment");
  const url = directApplicationUrl("202645643", "S168B");
  assert.equal(new URL(url).searchParams.get("da-number"), "202645643");
});
