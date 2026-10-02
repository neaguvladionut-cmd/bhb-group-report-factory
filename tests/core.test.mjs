import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { buildPayload, createAuditWorkbook, EVAL_SHEET_HEADERS, mergeSelectedFiles } from "../src/rebuild-core.js";
import { methodologyColumns, reportPlan } from "../src/rebuild-report-plan.js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const vendor = await readFile(resolve(root, "src/assets/vendor/xlsx.full.min.js"), "utf8");
const sandbox = { exports: {}, module: { exports: {} }, Buffer, process };
vm.runInNewContext(vendor, sandbox);
const XLSX = sandbox.exports;
const workbook = (rows) => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  return XLSX.write(book, { type: "buffer", bookType: "xlsx" });
};

const summary = workbook([
  ["CODE", "name", "cod cp", "Leadership", "Colaborare"],
  ["North", "Synthetic Ana", "A-1", 4, 3],
  ["South", "Synthetic Bogdan", "A-2", 2, 5]
]);
const detailed = workbook([
  ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", "Leadership", "Colaborare"],
  ["", "", "", "", "Subcompetente", "L", "C"],
  ["", "", "", "", "behavior", "Behavior one", "Behavior two"],
  ["North", "Synthetic Ana", "Nord", "A-1", "", 2, 1],
  ["South", "Synthetic Bogdan", "Sud", "A-2", "", 0, 2]
]);

test("summary and detailed exports normalize to one payload", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project" });
  assert.equal(payload.blockers.length, 0);
  assert.equal(payload.records.length, 2);
  assert.equal(payload.calculations[0].mean, 3);
  assert.equal(payload.codeReadiness.groups.join("|"), "North|South");
  assert.equal(payload.regionReadiness.available, 2);
  assert.equal(payload.behaviorAggregates.length, 2);
});

test("CODE groups and regions are available without changing participant identity", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project" });
  assert.equal(payload.records.map((record) => record.identity).join("|"), "synthetic ana|a-1|synthetic bogdan|a-2");
  assert.equal(payload.groups.length, 2);
  assert.equal([...new Set(payload.zoneCalculations.map((item) => item.region))].join("|"), "Nord|Sud");
});

test("Dev Plan template uses the exact seven required headers", () => {
  assert.deepEqual(EVAL_SHEET_HEADERS, ["competency", "subcompetency", "behavior", "objective_text_score_0", "objective_text_score_-1", "objective_text_score_1", "objective_text_score_2"]);
});

test("duplicate and detailed-only identities block readiness", () => {
  const duplicate = workbook([
    ["CODE", "name", "cod cp", "Leadership"],
    ["", "Synthetic Ana", "A-1", 4],
    ["", "Synthetic Ana", "A-1", 3]
  ]);
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: duplicate }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project" });
  assert.equal(payload.readiness, false);
  assert(payload.blockers.some((item) => item.code === "duplicate-identity"));
});

test("audit workbook exposes source, validation, normalized data, calculation and grouping sheets", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project" });
  const audit = createAuditWorkbook(XLSX, payload);
  for (const sheet of ["Rezumat", "Surse", "Validare", "Date normalizate", "Calcule", "Comportamente", "Grupuri", "Zone", "Clasament"]) assert(audit.SheetNames.includes(sheet), sheet);
});

test("selected file merge replaces same-name selections and preserves other files", () => {
  const first = { name: "summary.xlsx", bytes: new Uint8Array([1]) };
  const second = { name: "detail.xlsx", bytes: new Uint8Array([2]) };
  const replacement = { name: "SUMMARY.xlsx", bytes: new Uint8Array([3]) };
  assert.deepEqual(mergeSelectedFiles(mergeSelectedFiles([], [first]), [second]), [first, second]);
  assert.deepEqual(mergeSelectedFiles([first, second], [replacement]), [replacement, second]);
});

test("GRF-UX methodology derives real dates, teams, roles and location from detailed export", () => {
  const detailedWithMethodology = workbook([
    ["CODE", "name the person evaluated", "job", "regiune", "cod ac", "Competente", "date", "invited at", "certification location", "principal evaluator", "secondary evaluator", "evaluator 3", "Evaluatori", "Leadership"],
    ["", "", "", "", "", "Subcompetente", "", "", "", "", "", "", "", "L"],
    ["", "", "", "", "", "behavior", "", "", "", "", "", "", "", "Behavior one"],
    ["North", "Synthetic Ana", "Manager", "Nord", "A-1", "", "2026-10-01", "2026-10-01", "București", "Ana", "Mihai", "Sistem AC", "", 2],
    ["South", "Synthetic Bogdan", "Specialist", "Sud", "A-2", "", "2026-10-02", "2026-10-02", "București", "Ana", "Mihai", "System user", "", 1]
  ]);
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailedWithMethodology }], { projectName: "Proiect sintetic", program: "Centru de Dezvoltare", reportDate: "2026-10-01" });
  assert.deepEqual(payload.methodology.evaluatorNames, ["Ana", "Mihai"]);
  assert.deepEqual(payload.methodology.dates, ["01.10.2026", "02.10.2026"]);
  assert.equal(payload.methodology.commonTeamSize, 2);
  assert.deepEqual(payload.methodology.populationByRole, [{ role: "Manager", count: 1 }, { role: "Specialist", count: 1 }]);
  assert.deepEqual(payload.methodology.locations, ["București"]);
  const method = methodologyColumns(payload);
  assert(method.left.some((line) => /Fiecare participant a fost observat de o echipă formată din 2 consultanți/u.test(line)));
  assert(method.left.some((line) => /2 participanți \(1 Manager, 1 Specialist\)/u.test(line)));
  assert(method.left.some((line) => /01\.10\.2026 – 02\.10\.2026/u.test(line)));
  const cover = reportPlan(payload)[0];
  assert.equal(cover.title, "Proiect sintetic – Centru de Dezvoltare");
  assert.equal(cover.reportDate, "01.10.2026");
});
