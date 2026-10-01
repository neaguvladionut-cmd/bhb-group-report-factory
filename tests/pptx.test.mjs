import assert from "node:assert/strict";
import { access, readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { resolve, join } from "node:path";
import { tmpdir } from "node:os";
import test from "node:test";
import { buildPayload, createEvaluationSheetTemplate } from "../src/rebuild-core.js";
import { HOW_TO_READ, reportPlan } from "../src/rebuild-report-plan.js";
import { FIXED_TEMPLATE_LABELS, buildBundleArtifacts, fillFixedTemplateLabels, generateTrendPptx, selfCheckPptx } from "../src/template-pptx.js";

const run = promisify(execFile);
const root = resolve(new URL("..", import.meta.url).pathname);
const asset = resolve(root, "src/assets/trend/template-raport-de-grup-RO.pptx");

const loadUmd = async (path) => { const module = { exports: {} }; const source = await readFile(path, "utf8"); new Function("module", "exports", "require", source)(module, module.exports, undefined); return module.exports; };
const XLSX = await loadUmd(resolve(root, "src/assets/vendor/xlsx.full.min.js"));
const JSZip = await loadUmd(resolve(root, "src/assets/vendor/jszip.min.js"));
const opentype = await loadUmd(resolve(root, "src/assets/vendor/opentype.min.js"));
const workbook = (rows) => { const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1"); return XLSX.write(book, { type: "buffer", bookType: "xlsx" }); };
const summary = workbook([["CODE", "name", "cod cp", "Leadership", "Colaborare"], ["NORD", "Synthetic Ana", "A-1", 4, 3], ["SUD", "Synthetic Bogdan", "A-2", 2, 5]]);
const detailed = workbook([["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", "Leadership", "Colaborare"], ["", "", "", "", "Subcompetente", "L", "C"], ["", "", "", "", "behavior", "B one", "B two"], ["NORD", "Synthetic Ana", "Nord", "A-1", "", 2, 1], ["SUD", "Synthetic Bogdan", "Sud", "A-2", "", 0, 2]]);
const makePayload = (metadata = {}) => buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "Synthetic project", ...metadata }, {}, { acknowledgedWarningIds: [] });
const decode = (value) => String(value).replace(/&amp;/gu, "&").replace(/&lt;/gu, "<").replace(/&gt;/gu, ">").replace(/&quot;/gu, '"').replace(/&apos;/gu, "'");
const slideText = (xml) => [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/gu)].map((match) => decode(match[1])).filter(Boolean);
const relationshipTarget = (source, target) => { const base = source.split("/").slice(0, -1).join("/"); const parts = `${base}/${target}`.split("/"); const result = []; for (const part of parts) { if (part === "..") result.pop(); else if (part && part !== ".") result.push(part); } return result.join("/"); };
const generatedDeck = async (payload, options = {}) => { global.window = { JSZip, XLSX, __GRF_TEMPLATE_BASE64__: (await readFile(asset)).toString("base64") }; const blob = await generateTrendPptx(payload, options); const bytes = Buffer.from(await blob.arrayBuffer()); const path = join(tmpdir(), `grf-r-${Date.now()}-${Math.random().toString(16).slice(2)}.pptx`); await writeFile(path, bytes); return { path, bytes, zip: await JSZip.loadAsync(bytes) }; };

test("Trend route uses the cleaned native template, not a generated picture deck", async () => {
  await access(asset);
  const source = await readFile(resolve(root, "src/template-pptx.js"), "utf8");
  assert.doesNotMatch(source, /PptxGenJS/u);
  assert.match(source, /ppt\/charts/u);
  assert.match(source, /ppt\/embeddings/u);
  const { stdout } = await run("unzip", ["-l", asset]);
  assert.match(stdout, /ppt\/slideMasters\/slideMaster1\.xml/u);
  assert.match(stdout, /ppt\/slideLayouts\/slideLayout1\.xml/u);
  assert.match(stdout, /ppt\/charts\/chart1\.xml/u);
  assert.match(stdout, /ppt\/embeddings\/Microsoft_Excel_Worksheet\.xlsx/u);
  assert.doesNotMatch(stdout, /ppt\/notesSlides\//u);
  assert.doesNotMatch(stdout, /customXml\//u);
  assert.doesNotMatch(stdout, /ppt\/changesInfos\//u);
  assert.doesNotMatch(stdout, /ppt\/revisionInfo\.xml/u);
  assert.doesNotMatch(stdout, /ppt\/fonts\//u);
  const presentation = (await run("unzip", ["-p", asset, "ppt/presentation.xml"])).stdout;
  assert.doesNotMatch(presentation, /notesMasterIdLst/u);
});

test("G11 retires the old renderer, report plan and BHB PowerPoint route", async () => {
  for (const retired of ["src/app.js", "src/core.js", "src/index.html", "src/pptx.js", "src/report-plan.js", "deploy/core.js", "deploy/pptx.js", "deploy/report-plan.js", "core.js", "pptx.js", "report-plan.js"]) {
    await assert.rejects(access(resolve(root, retired)), new RegExp("ENOENT", "u"));
  }
});

test("clean-template declares the public-package scrub and local-only source boundary", async () => {
  const script = await readFile(resolve(root, "tools/clean-template.mjs"), "utf8");
  assert.match(script, /notesSlides/u);
  assert.match(script, /customXml/u);
  assert.match(script, /changesInfos/u);
  assert.match(script, /dc:title/u);
  assert.match(script, /selfCheck/u);
  assert.match(script, /embeddedFontLst/u);
  assert.match(script, /ppt\/fonts/u);
  assert.match(script, /git-ignored/iu);
});

test("Trend generator re-inserts every fixed template label into generated slide XML", () => {
  const generated = fillFixedTemplateLabels("<p:sld><p:spTree></p:spTree></p:sld>", FIXED_TEMPLATE_LABELS);
  for (const label of FIXED_TEMPLATE_LABELS) assert.match(generated, new RegExp(label.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"), "u"));
});

test("generated Trend deck has one cloned output slide per plan item across configurations", async () => {
  for (const metadata of [{ annex: "end" }, { annex: "separate" }, { annex: "none" }, { annex: "end", splitGroups: true, slideToggles: { range: false, competencyMean: false, benchmark: false, competencyDistribution: false, zone: false, observation: false, behavior: false, conclusions: false } }]) {
    const payload = makePayload(metadata);
    const plan = reportPlan(payload);
    const generated = await generatedDeck(payload);
    const slideNames = Object.keys(generated.zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).sort((a, b) => Number(a.match(/slide(\d+)/u)[1]) - Number(b.match(/slide(\d+)/u)[1]));
    assert.equal(slideNames.length, plan.length);
    for (const [index, item] of plan.entries()) {
      const xml = await generated.zip.file(slideNames[index]).async("string");
      const titleShape = xml.match(/name="GRF-R role:title"[\s\S]*?<a:t>([\s\S]*?)<\/a:t>/u);
      assert(titleShape, `slide ${index + 1} has no generated title role`);
      assert.equal(decode(titleShape[1]), item.title);
      assert.match(xml, /name="GRF-R role:body"/u, `slide ${index + 1} has no generated body role`);
    }
    await selfCheckPptx(generated.zip);
  }
});

test("generated group slides and charts use NORD/SUD participant-only values", async () => {
  const payload = makePayload({ splitGroups: true });
  const plan = reportPlan(payload);
  const generated = await generatedDeck(payload);
  for (const [group, expected] of [["NORD", "4.00"], ["SUD", "2.00"]]) {
    const groupItems = plan.filter((item) => item.groupKey === group && ["range", "ranking", "competency-distribution"].includes(item.family));
    assert(groupItems.length);
    for (const item of groupItems) {
      const index = plan.indexOf(item);
      const xml = await generated.zip.file(`ppt/slides/slide${index + 1}.xml`).async("string");
      assert.match(xml, new RegExp(`${group}[\\s\\S]*Leadership: ${expected}`, "u"));
    }
  }
  for (const slide of Object.keys(generated.zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name))) {
    const rels = await generated.zip.file(slide.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels").async("string");
    for (const rel of [...rels.matchAll(/Target="([^"]*charts\/chart\d+\.xml)"/gu)]) {
      const chart = relationshipTarget(slide, rel[1]);
      const chartXml = await generated.zip.file(chart).async("string");
      assert.match(chartXml, /<c:f>Sheet1!\$A\$2:/u);
      assert.match(chartXml, /<c:numCache>/u);
    }
  }
});

test("every generated chart cache, formula and relationship workbook agree", async () => {
  const generated = await generatedDeck(makePayload());
  for (const slide of Object.keys(generated.zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name))) {
    const relsPath = slide.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    const rels = await generated.zip.file(relsPath).async("string");
    for (const rel of [...rels.matchAll(/Target="([^"]*charts\/chart\d+\.xml)"/gu)]) {
      const chartPath = relationshipTarget(slide, rel[1]); const chart = await generated.zip.file(chartPath).async("string");
      const chartRelsPath = chartPath.replace("ppt/charts/", "ppt/charts/_rels/") + ".rels"; const chartRels = await generated.zip.file(chartRelsPath).async("string");
      const workbookRel = chartRels.match(/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/package" Target="([^"]+)"/u);
      assert(workbookRel);
      const workbookPath = relationshipTarget(chartPath, workbookRel[1]); const book = XLSX.read(await generated.zip.file(workbookPath).async("array"), { type: "array" }); const rows = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { header: 1, defval: "" });
      const categoryCaches = [...chart.matchAll(/<c:cat>[\s\S]*?<c:strCache>[\s\S]*?<\/c:strCache>[\s\S]*?<\/c:cat>/gu)].map((match) => [...match[0].matchAll(/<c:v>([\s\S]*?)<\/c:v>/gu)].map((value) => decode(value[1])));
      const cells = rows.slice(1).map((row) => String(row[0] ?? ""));
      for (const categories of categoryCaches) assert.deepEqual(categories, cells);
      assert.match(chart, /<c:f>Sheet1!\$A\$2:\$A\$/u);
      assert.match(chart, /<c:f>Sheet1!\$B\$2:\$B\$/u);
    }
  }
});

test("bundle output is outlined, font-embedded and free of undefined/null/NaN", async () => {
  global.window = { opentype, __GRF_FONT_BASE64__: (await readFile(resolve(root, "src/assets/vendor/Poppins-Regular.ttf"))).toString("base64") };
  const artifacts = await buildBundleArtifacts(XLSX, makePayload());
  for (const item of artifacts.svg) { assert.doesNotMatch(item.content, /<text|foreignObject|font-family/iu); assert.match(item.content, /<path\b/u); assert.doesNotMatch(item.content, /undefined|null|NaN/iu); }
  const manifestText = JSON.stringify(artifacts.manifest); assert.doesNotMatch(manifestText, /undefined|null|NaN/iu);
  const dataBook = XLSX.read(artifacts.workbookBytes, { type: "array" });
  for (const name of dataBook.SheetNames) for (const row of XLSX.utils.sheet_to_json(dataBook.Sheets[name], { header: 1, defval: "" })) for (const cell of row) assert.doesNotMatch(String(cell), /undefined|null|NaN/iu);
  const pdf = new TextDecoder("latin1").decode(artifacts.pdfBytes); assert.match(pdf, /\/FontName \/Poppins/u); assert.doesNotMatch(pdf, /\/Helvetica/u); const pdfStreams = [...pdf.matchAll(/stream\n([\s\S]*?)\nendstream/gu)].slice(1).map((match) => match[1]).join("\n"); assert.doesNotMatch(pdfStreams, /\?/u);
});

test("generated CSV template retains all imported behaviors, including all-missing rows", () => {
  const behaviorHeaders = ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", ...Array.from({ length: 30 }, (_, index) => `B${index + 1}`)];
  const detailedRows = [behaviorHeaders, ["", "", "", "", "Subcompetente", ...Array(30).fill("S")], ["", "", "", "", "behavior", ...Array.from({ length: 30 }, (_, index) => `Behavior ${index + 1}`)], ["NORD", "Synthetic Ana", "Nord", "A-1", "", ...Array(30).fill("")]];
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: workbook(detailedRows) }], { projectName: "Synthetic project" }, {}, { acknowledgedWarningIds: [] });
  const output = createEvaluationSheetTemplate(XLSX, payload);
  const rows = XLSX.utils.sheet_to_json(output.Sheets["Evaluation sheet"], { header: 1, defval: "" });
  assert.equal(rows.length - 1, 30);
  assert.deepEqual(rows.slice(1).map((row) => row[2]), Array.from({ length: 30 }, (_, index) => `Behavior ${index + 1}`));
});

test("generated annex output applies the how-to-read wording for all settings", async () => {
  for (const annex of ["end", "separate", "none"]) {
    const payload = makePayload({ annex });
    const generated = await generatedDeck(payload);
    const text = [];
    for (const name of Object.keys(generated.zip.files).filter((entry) => /^ppt\/slides\/slide\d+\.xml$/u.test(entry))) text.push(...slideText(await generated.zip.file(name).async("string")));
    const output = text.join(" ");
    if (annex === "none") assert.doesNotMatch(output, /Rezultatele individuale se regăsesc/u);
    if (annex === "end") assert.match(output, /Rezultatele individuale se regăsesc în anexă\./u);
    if (annex === "separate") assert.match(output, /Rezultatele individuale se regăsesc în anexa transmisă separat\./u);
    await selfCheckPptx(generated.zip);
  }
});

test("README describes the rebuilt Trend tool and its actual dependencies", async () => {
  const readme = await readFile(resolve(root, "README.md"), "utf8");
  assert.doesNotMatch(readme, /alege(?:rea)?\s+(?:BHB|TREND)|BHB\/TREND|PptxGenJS/iu);
  for (const dependency of ["SheetJS", "JSZip", "opentype.js", "Poppins"]) assert.match(readme, new RegExp(dependency.replace(".", "\\."), "u"));
});
