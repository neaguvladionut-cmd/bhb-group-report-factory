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
  assert.deepEqual(ranking.development.map((row) => row.behavior), ["B6", "B7", "B8", "B9", "B10"], "equal means and shares fall back to column order");
  assert.equal(ranking.median, null);
});

test("split report plan places whole project before CODE groups and keeps named slides in the annex", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project", splitGroups: true, annex: "separate" }, {}, { acknowledgedWarningIds: [] });
  const plan = reportPlan(payload);
  const firstGroup = plan.findIndex((item) => item.groupKey === "North");
  const firstAnnex = plan.findIndex((item) => item.deliverable === "appendix");
  assert(firstGroup > 0);
  assert(firstAnnex > firstGroup);
  assert(plan.filter((item) => item.deliverable === "appendix").every((item) => ["appendix-divider", "participant-mean", "participant-comparison", "divider-observations", "key-findings", "competency-participants", "divider-behaviors", "behavior"].includes(item.family)));
  assert(plan.filter((item) => ["participant-mean", "participant-comparison", "competency-participants"].includes(item.family)).every((item) => !item.groupKey));
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

test("GRF-UX scale fixture keeps 155 identical blockers on one bounded, searchable kind card", async () => {
  const rows = [["CODE", "name", "cod cp", "Leadership"]];
  for (let index = 1; index <= 200; index += 1) rows.push(["North", `Synthetic participant ${String(index).padStart(3, "0")}`, `S-${index}`, index <= 155 ? 9 : 3]);
  const payload = buildPayload(XLSX, [{ name: "summary-export-synthetic.xlsx", bytes: workbook(rows) }], { projectName: "Synthetic scale fixture" });
  assert.equal(payload.blockers.filter((item) => item.code === "summary-score").length, 155);
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  const issueRenderer = await readFile(resolve(root, "src/rebuild-issues.js"), "utf8");
  const css = await readFile(resolve(root, "src/styles.css"), "utf8");
  assert.match(app, /groupedIssues\(blockers, "blocker"\)/u);
  assert.match(app, /renderIssueGroup/u);
  assert.match(issueRenderer, /data-issue-search/u);
  assert.match(css, /\.issue-panel\{/u);
  assert.match(css, /\.issue-toggle/u);
  assert.doesNotMatch(css, /\.issue-items\{[^}]*overflow/u);
});

test("R5 by mean (Vlad 2026-10-02): complete data keeps the sum order; missing scores rank by mean; spread then column order break ties", async () => {
  const { topOrder } = await import("../src/rebuild-report-plan.js");
  const row = (behavior, scores, sourceIndex) => { const present = scores.filter((value) => value !== null); return { competency: "C", behavior, sourceIndex, n: present.length, sum: present.reduce((a, b) => a + b, 0), mean: present.reduce((a, b) => a + b, 0) / present.length, pct2: present.filter((v) => v === 2).length / present.length, pct0: present.filter((v) => v === 0).length / present.length }; };
  // Complete data, no ties: mean order equals the old sum order.
  const complete = [row("A", [2, 2, 1, 1], 0), row("B", [2, 1, 1, 0], 1), row("C", [2, 2, 2, 1], 2), row("D", [0, 0, 1, 0], 3), row("E", [1, 1, 1, 2], 4), row("F", [0, 1, 0, 0], 5)];
  const bySum = complete.slice().sort((a, b) => b.sum - a.sum || a.sourceIndex - b.sourceIndex).map((r) => r.behavior);
  assert.deepEqual(complete.slice().sort(topOrder).map((r) => r.behavior), bySum);
  // Missing scores: A has the larger sum but the lower mean.
  const missing = [row("A", [1, 1, 1, 1, 1, 2], 0), row("B", [2, 2, 2, null, null, null], 1), row("C", [0, 0, 1, 0, 0, 0], 2), row("D", [0, 1, 0, 1, 0, 0], 3), row("E", [1, 0, 1, 0, 1, 0], 4), row("F", [0, 0, 0, 0, 0, 0], 5)];
  assert(missing[0].sum > missing[1].sum && missing[1].mean > missing[0].mean);
  const ranked = rankBehaviors(missing)[0];
  assert.deepEqual(ranked.key.map((r) => r.behavior), ["B", "A", "E"]);
  assert.equal(ranked.development[0].behavior, "F");
  // Equal means: the top list prefers the larger share of 2; the bottom list the larger share of 0; then column order.
  const ties = [row("P", [2, 0, 1, 1], 0), row("Q", [1, 1, 1, 1], 1), row("R", [2, 0, 2, 0], 2), row("S", [1, 1, 1, 1], 3), row("T", [0, 0, 0, 0], 4), row("U", [0, 0, 0, 1], 5)];
  const tied = rankBehaviors(ties)[0];
  assert.deepEqual(tied.key.map((r) => r.behavior), ["R", "P", "Q"]);
  assert.deepEqual(tied.development.map((r) => r.behavior), ["T", "U", "S"]);
});

test("R5 by mean in a group view ranks the group's own scores", () => {
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project", splitGroups: true }, {}, { acknowledgedWarningIds: [] });
  const plan = reportPlan(payload);
  const groupFindings = plan.filter((item) => item.family === "key-findings" && item.groupKey);
  for (const item of groupFindings) for (const row of [...item.insight.key, ...item.insight.development]) assert(row.n <= payload.groups.find((group) => group.code === item.groupKey).records.length);
  const audit = createAuditWorkbook(XLSX, payload);
  assert.equal(JSON.stringify(XLSX.utils.sheet_to_json(audit.Sheets.Clasament, { header: 1 })[0].slice(0, 5)), JSON.stringify(["Competență", "Comportament", "Medie 0–2", "% scor 2", "% scor 0"]));
});

test("assessment days count „invited at” dates when the export has no „date” column", () => {
  const rows = [["CODE", "name", "invited at", "cod cp", "Leadership"], ["", "Synthetic Ana", "11-05-2026", "A-1", 3], ["", "Synthetic Bob", "12-05-2026", "A-2", 3], ["", "Synthetic Cia", "12-05-2026", "A-3", 3]];
  const payload = buildPayload(XLSX, [{ name: "invited-summary.xlsx", bytes: workbook(rows) }], { projectName: "Invited synthetic" });
  const dates = new Set((payload.schemas || []).flatMap((schema) => schema.methodology?.dates || []));
  assert.equal(dates.size, 2);
});
