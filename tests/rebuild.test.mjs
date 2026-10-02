import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { buildPayload, createAuditWorkbook, EVAL_SHEET_HEADERS } from "../src/rebuild-core.js";
import { rankBehaviors, reportPlan } from "../src/rebuild-report-plan.js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const vendor = await readFile(resolve(root, "src/assets/vendor/xlsx.full.min.js"), "utf8");
const sandbox = { exports: {}, module: { exports: {} }, Buffer, process };
vm.runInNewContext(vendor, sandbox);
const XLSX = sandbox.exports;
const workbook = (rows) => { const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1"); return XLSX.write(book, { type: "buffer", bookType: "xlsx" }); };

const summary = workbook([
  ["CODE", "name", "cod cp", "Leadership", "Colaborare"],
  ["North", "Synthetic Ana", "A-1", 4, 3],
  ["South", "Synthetic Bogdan", "A-2", 2, 5],
  ["", "Synthetic Carmen", "A-3", 3, 3]
]);
const detailed = workbook([
  ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", "Leadership", "Colaborare"],
  ["", "", "", "", "Subcompetente", "L", "C"],
  ["", "", "", "", "behavior", "Behaviou­r one", "Behaviour two"],
  ["North", "Synthetic Ana", "Nord", "A-1", "", 2, 1],
  ["South", "Synthetic Bogdan", "Sud", "A-2", "", 0, 2],
  ["", "Synthetic Carmen", "Centru", "A-3", "", 1, ""]
]);
const descriptors = new TextEncoder().encode(`${EVAL_SHEET_HEADERS.join(",")}\nLeadership,,Behaviou­r one,Să exersezi,,,Ai exersat\nColaborare,,Behaviour two,Să cooperezi,,,Ai cooperat\n`);

test("GRF-R payload keeps CODE groups, regions, missing scores and exact descriptor wording", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }, { name: "evaluation-sheet-template.csv", bytes: descriptors }], { projectName: "Synthetic project", clientName: "Synthetic client" }, {}, { acknowledgedWarningIds: [] });
  assert.equal(payload.blockers.length, 0);
  assert.equal(payload.codeReadiness.groups.join("|"), "North|South");
  assert.equal(payload.regionReadiness.available, 3);
  assert(payload.warnings.some((item) => item.code === "detailed-blank"));
  assert.equal(payload.behaviorAggregates.find((item) => item.behavior === "Behaviou­r one").score2, "Ai exersat");
  assert.equal(payload.behaviorAggregates.find((item) => item.behavior === "Behaviour two").score0, "Să cooperezi");
  const reviewed = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }, { name: "evaluation-sheet-template.csv", bytes: descriptors }], { projectName: "Synthetic project" }, {}, { acknowledgedWarningIds: payload.warnings.map((item) => item.id) });
  assert.equal(reviewed.readiness, true);
});

test("behaviour ranking uses present-score sums, fixed cuts and import-order ties", () => {
  const rows = Array.from({ length: 10 }, (_, index) => ({ competency: "Leadership", behavior: `B${index + 1}`, sum: index < 5 ? 4 : 1, sourceIndex: index }));
  const ranking = rankBehaviors(rows)[0];
  assert.deepEqual(ranking.key.map((row) => row.behavior), ["B1", "B2", "B3", "B4", "B5"]);
  assert.deepEqual(ranking.development.map((row) => row.behavior), ["B6", "B7", "B8", "B9", "B10"]);
  assert.equal(ranking.median, null);
});

test("split report plan places whole project before CODE groups and keeps named slides in the annex", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project", splitGroups: true, annex: "separate" }, {}, { acknowledgedWarningIds: [] });
  const plan = reportPlan(payload);
  const firstGroup = plan.findIndex((item) => item.groupKey === "North");
  const firstAnnex = plan.findIndex((item) => item.deliverable === "appendix");
  assert(firstGroup > 0);
  assert(firstAnnex > firstGroup);
  assert(plan.filter((item) => item.deliverable === "appendix").every((item) => ["appendix-divider", "participant-mean", "participant-comparison", "competency-participants"].includes(item.family)));
  assert(plan.filter((item) => item.deliverable === "appendix").every((item) => !item.groupKey));
});

test("audit workbook carries groups, zones and ranking sheets", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project" }, {}, { acknowledgedWarningIds: [] });
  const audit = createAuditWorkbook(XLSX, payload);
  assert(audit.SheetNames.includes("Grupuri"));
  assert(audit.SheetNames.includes("Zone"));
  assert(audit.SheetNames.includes("Clasament"));
});

test("summary competency scores accept decimal averages on the 1–5 scale and reject values outside it", () => {
  const decimalSummary = workbook([
    ["CODE", "name", "cod cp", "Leadership", "Colaborare"],
    ["", "Synthetic Dana", "D-1", 1.5, 3.25],
    ["", "Synthetic Emil", "D-2", 4.75, 2],
    ["", "Synthetic Flavia", "D-3", 5.5, 0.75]
  ]);
  const payload = buildPayload(XLSX, [{ name: "decimal-summary.xlsx", bytes: decimalSummary }], { projectName: "Decimal synthetic" });
  const invalid = payload.blockers.filter((item) => item.code === "summary-score");
  assert.equal(invalid.length, 2, "only 5.5 and 0.75 are outside 1–5");
  assert.ok(invalid.every((item) => /Synthetic Flavia|5\.5|0\.75/u.test(JSON.stringify(item)) || item.row === 4 || item.rowNumber === 4));
  const dana = payload.records.find((record) => record.name === "Synthetic Dana");
  assert.deepEqual(dana.scores, { Leadership: 1.5, Colaborare: 3.25 });
  const emil = payload.records.find((record) => record.name === "Synthetic Emil");
  assert.equal(emil.scores.Leadership, 4.75);
});

test("participants without CODE raise one grouped warning, not one per participant", () => {
  const rows = [["CODE", "name", "cod cp", "Leadership"]];
  for (let index = 1; index <= 12; index += 1) rows.push([index <= 2 ? "North" : "", `Synthetic Nocode ${String(index).padStart(2, "0")}`, `N-${index}`, 3]);
  const payload = buildPayload(XLSX, [{ name: "nocode-summary.xlsx", bytes: workbook(rows) }], { projectName: "Grouped warning synthetic" });
  const codeWarnings = payload.warnings.filter((item) => item.code === "code-missing");
  assert.equal(codeWarnings.length, 1);
  assert.match(codeWarnings[0].message, /^10 participanți nu au CODE/u);
  assert.match(codeWarnings[0].message, /și încă 2/u);
  assert.equal(codeWarnings[0].identities.length, 10);
});
