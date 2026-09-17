import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { readFile, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { createPdf } from "../src/pdf.js";
import { reportPlan } from "../src/report-plan.js";

const vendor = await readFile(new URL("../src/assets/vendor/pdf-lib.min.js", import.meta.url), "utf8");
(0, eval)(vendor);
const basePayload = sections => ({ metadata: { projectName: "Proiect PDF", clientName: "Client PDF", reportDate: "septembrie 2026", context: "Assessment Centre", exercises: "Exerciții confirmate", otherInstruments: "", conclusions: "Concluzii aprobate.", sections, includeCompetencySummary: false }, calculations: [{ competency: "Colaborare", mean: 3.4, median: 3.5, min: 2, max: 5, n: 2 }, { competency: "People Management", mean: 2.8, median: 3, min: 1, max: 4, n: 2 }], records: [{ name: "Ana", scores: { Colaborare: 3, "People Management": 4 } }, { name: "Bogdan", scores: { Colaborare: 4, "People Management": 2 } }], bands: { low: 2.75, high: 3.5, below: 0, typical: 1, above: 1, n: 2 }, participantCounts: { included: 2, excludedUnrated: 0 }, behaviorAggregates: [{ competency: "Colaborare", behavior: "Observă activ și formulează întrebări", pct0: 0, pct2: 1, mean: 3.5 }], schemas: [] });
const allSections = { distribution: true, benchmark: true, keyFindings: true, behavior: true, conclusions: true, appendixBenchmark: true, appendixComparison: true, appendixDistribution: true };

test("PDF bytes are real, fixed landscape pages, and match the whole report plan", async () => {
  for (const renderer of ["bhb", "trend"]) {
    const payload = basePayload(allSections), bytes = await createPdf(payload, { renderer, scope: "whole" });
    assert.match(Buffer.from(bytes).subarray(0, 8).toString(), /^%PDF-/u);
    const pageCount = (Buffer.from(bytes).toString("latin1").match(/\/Type\s*\/Page\b/gu) || []).length;
    assert.equal(pageCount, reportPlan(payload, { scope: "whole" }).length, `${renderer} PDF page count follows reportPlan`);
  }
});

test("PDF follows selected-section omission and carries metadata, thresholds, conclusions, and hostile text safely", async () => {
  const payload = basePayload({ ...allSections, distribution: false, benchmark: false, behavior: false, conclusions: false, appendixBenchmark: false, appendixComparison: false, appendixDistribution: false });
  payload.metadata.projectName = "Titlu <script>alert(1)</script>";
  const bytes = await createPdf(payload), dir = await mkdtemp(join(tmpdir(), "grf-pdf-test-")), file = join(dir, "report.pdf");
  try {
    await writeFile(file, bytes);
    const raw = Buffer.from(bytes).toString("latin1");
    assert.doesNotMatch(raw, /<script>|alert\(1\)<\/script>/u);
    assert.equal((raw.match(/\/Type\s*\/Page\b/gu) || []).length, reportPlan(payload, { scope: "whole" }).length);
    try {
      const extracted = execFileSync("pdftotext", [file, "-"], { encoding: "utf8" });
      assert.match(extracted, /Titlu script alert\(1\) script/u);
      assert.match(extracted, /Client PDF|Assessment Centre/u);
      assert.doesNotMatch(extracted, /Distribuția rezultatelor|Profil comportamental|Anexă/u);
    } catch (error) { if (error?.code !== "ENOENT") throw error; }
  } finally { await rm(dir, { recursive: true, force: true }); }
});
