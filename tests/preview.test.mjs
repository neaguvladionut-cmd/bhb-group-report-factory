import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import test from "node:test";
import { behaviorInsights, methodologyColumns, reportPlan, rankBehaviors } from "../src/rebuild-report-plan.js";

const root = resolve(new URL("..", import.meta.url).pathname);

test("report plan includes the Trend sections and an explicit whole-project-first group order", () => {
  const payload = { metadata: { projectName: "Synthetic project", clientName: "Synthetic client", annex: "separate", splitGroups: true, groupNames: { North: "Nord" } }, participantCounts: { included: 2 }, bands: { low: 2.75, high: 3.5, below: 0, typical: 2, above: 0, n: 2 }, calculations: [{ competency: "Leadership", mean: 3.5, median: 3.5, n: 2 }], behaviorAggregates: [], behaviorRecords: [], records: [], groups: [{ code: "North", records: [] }, { code: "South", records: [] }], regionReadiness: { available: 0 }, zoneCalculations: [] };
  const plan = reportPlan(payload);
  assert(plan.some((slide) => slide.family === "how-to-read"));
  assert(plan.some((slide) => slide.family === "executive-summary"));
  assert(plan.some((slide) => slide.family === "key-findings"));
  assert(plan.some((slide) => slide.groupKey === "North"));
  assert(plan.findIndex((slide) => slide.groupKey === "North") > 0);
  assert(plan.some((slide) => slide.deliverable === "appendix"));
});

test("methodology has editable unknown fields and no internal score language", () => {
  const payload = { metadata: { projectName: "Synthetic project", clientName: "Synthetic client" }, participantCounts: { included: 2 }, behaviorAggregates: [], schemas: [{ methodology: { evaluators: ["Consultant A", "Consultant B"], dates: ["2026-06-01", "2026-06-02"], teamSizes: [2, 2] } }] };
  const columns = methodologyColumns(payload);
  const copy = [...columns.left, ...columns.right].join(" ");
  assert.match(copy, /consultanți TREND implicați/u);
  assert.equal(columns.missing.exercises, true);
  assert.doesNotMatch(copy, /0\s*[–,-]\s*1\s*[–,-și ]+2|\(0-2\)/u);
  const edited = methodologyColumns({ ...payload, metadata: { ...payload.metadata, evaluators: "4", days: "3", exercises: "un exercițiu de grup, un studiu de caz" } });
  assert.deepEqual(edited.facts.slice(1, 3).map((fact) => fact.number), ["4", "3"]);
  assert.equal(edited.facts.at(-1).number, "2");
  assert.match(edited.facts.at(-1).text, /: un exercițiu de grup, un studiu de caz$/u);
});

test("behaviour insights follow the fixed summed-score cuts", () => {
  const rows = Array.from({ length: 6 }, (_, index) => ({ competency: "Leadership", behavior: `B${index + 1}`, sum: 6 - index, sourceIndex: index }));
  const insight = behaviorInsights(rows)[0];
  assert.deepEqual(insight.key.map((row) => row.behavior), ["B1", "B2", "B3"]);
  assert.deepEqual(insight.development.map((row) => row.behavior), ["B6", "B5", "B4"], "the bottom list starts with the lowest");
  assert.equal(rankBehaviors(rows)[0].median, null);
});

test("preview source keeps accessible thumbnails, keyboard controls and escaped text sinks", async () => {
  const source = await readFile(resolve(root, "src/preview.js"), "utf8");
  assert.match(source, /data-slide/u);
  assert.match(source, /ArrowLeft/u);
  assert.match(source, /ArrowRight/u);
  assert.match(source, /replace\(/u);
});
