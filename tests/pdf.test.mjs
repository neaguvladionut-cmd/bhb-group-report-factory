import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { readFile, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import vm from "node:vm";
import test from "node:test";
import { createPdf } from "../src/pdf.js";
import { reportPlan } from "../src/report-plan.js";

const vendor = await readFile(new URL("../src/assets/vendor/pdf-lib.min.js", import.meta.url), "utf8");
const fontkitVendor = await readFile(new URL("../src/assets/vendor/fontkit.umd.min.js", import.meta.url), "utf8");
const fontDataVendor = await readFile(new URL("../src/assets/fonts/noto-sans-data.js", import.meta.url), "utf8");
const pdfjsVendor = await readFile(new URL("../src/assets/vendor/pdf.min.js", import.meta.url), "utf8");
const pdfjsWorkerVendor = await readFile(new URL("../src/assets/vendor/pdf.worker.min.js", import.meta.url), "utf8");
(0, eval)(vendor);
(0, eval)(fontkitVendor);
(0, eval)(fontDataVendor);
async function extractPdfText(bytes) {
  const context = { console, DOMMatrix: class DOMMatrix {}, Path2D: class Path2D {}, URL, TextEncoder, TextDecoder, Uint8Array, ArrayBuffer, Promise, AbortController, AbortSignal, structuredClone, ReadableStream, TransformStream, atob, setTimeout, clearTimeout, navigator: { userAgent: "node", platform: "MacIntel" }, location: { href: "file:///tmp/index.html" } };
  context.globalThis = context; context.self = context; context.window = context;
  vm.createContext(context);
  vm.runInContext(pdfjsWorkerVendor, context, { filename: "pdf.worker.min.js" });
  vm.runInContext(pdfjsVendor, context, { filename: "pdf.min.js" });
  const doc = await context.pdfjsLib.getDocument({ data: new Uint8Array(bytes), disableWorker: true }).promise;
  let output = "";
  for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
    const page = await doc.getPage(pageNumber), content = await page.getTextContent();
    output += content.items.map(item => item.str).join(" ") + "\n";
  }
  return output;
}
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

test("PDF preserves exact Romanian diacritics in embedded text", async () => {
  const payload = basePayload(allSections);
  payload.metadata.projectName = "raportul schimbării";
  payload.metadata.context = "gândire";
  payload.metadata.conclusions = "priorităților și concluziilor aprobate";
  payload.calculations[0].competency = "schimbării";
  const bytes = await createPdf(payload), raw = Buffer.from(bytes).toString("latin1");
  assert.match(raw, /ToUnicode/u, "the PDF must embed Unicode mappings for text extraction");
  const dir = await mkdtemp(join(tmpdir(), "grf-pdf-diacritics-")), file = join(dir, "report.pdf");
  try {
    await writeFile(file, bytes);
    const extracted = await extractPdfText(bytes);
    for (const value of ["schimbării", "gândire", "priorităților", "și"]) assert.match(extracted, new RegExp(value, "u"));
  } finally { await rm(dir, { recursive: true, force: true }); }
});
