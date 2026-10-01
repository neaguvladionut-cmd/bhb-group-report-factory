import assert from "node:assert/strict";
import { access, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { join, resolve } from "node:path";
import { tmpdir } from "node:os";
import test from "node:test";
import { buildPayload } from "../src/rebuild-core.js";
import { reportPlan } from "../src/rebuild-report-plan.js";
import { buildBundleArtifacts, generateTrendPptx, selfCheckPptx } from "../src/template-pptx.js";
import { behaviorRows, createFixture } from "./fixtures/grf-r-acceptance-fixture.mjs";

const run = promisify(execFile);
const root = resolve(new URL("..", import.meta.url).pathname);
const asset = resolve(root, "src/assets/trend/template-raport-de-grup-RO.pptx");
const font = resolve(root, "src/assets/vendor/Poppins-Regular.ttf");
const loadUmd = async (path) => { const module = { exports: {} }; const source = await readFile(path, "utf8"); new Function("module", "exports", "require", source)(module, module.exports, undefined); return module.exports; };
const XLSX = await loadUmd(resolve(root, "src/assets/vendor/xlsx.full.min.js"));
const JSZip = await loadUmd(resolve(root, "src/assets/vendor/jszip.min.js"));
const fixture = createFixture(XLSX);
const decode = (value) => String(value).replace(/&amp;/gu, "&").replace(/&lt;/gu, "<").replace(/&gt;/gu, ">").replace(/&quot;/gu, '"').replace(/&apos;/gu, "'");
const slideFiles = (zip) => Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).sort((a, b) => Number(a.match(/slide(\d+)/u)[1]) - Number(b.match(/slide(\d+)/u)[1]));
const relationshipTarget = (source, target) => { const base = source.split("/").slice(0, -1).join("/"); const parts = `${base}/${target}`.split("/"); const result = []; for (const part of parts) { if (part === "..") result.pop(); else if (part && part !== ".") result.push(part); } return result.join("/"); };
const relationships = (xml) => [...xml.matchAll(/<Relationship\b([^>]*)\/>/gu)].map((match) => Object.fromEntries([...match[1].matchAll(/(Id|Type|Target|TargetMode)="([^"]*)"/gu)].map((part) => [part[1], part[2]])));
const acceptancePayload = (splitGroups) => buildPayload(XLSX, [{ name: "summary.xlsx", bytes: fixture.summary }, { name: "detail.xlsx", bytes: fixture.detailed }, { name: "evaluation-sheet-template.csv", bytes: fixture.csv }], { ...fixture.metadata, splitGroups }, {}, { acknowledgedWarningIds: [] });
const generatedDeck = async (payload, scope = "whole") => { global.window = { JSZip, XLSX, __GRF_TEMPLATE_BASE64__: (await readFile(asset)).toString("base64") }; const blob = await generateTrendPptx(payload, { scope }); const bytes = Buffer.from(await blob.arrayBuffer()); return { bytes, zip: await JSZip.loadAsync(bytes) }; };
const textFromSlide = (xml) => [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/gu)].map((match) => decode(match[1])).filter(Boolean).join(" ");
const binaryCommandAvailable = async (command) => { try { await run("which", [command]); return true; } catch { return false; } };

test("permanent GRF-R acceptance fixture has the §13 shape", () => {
  const payload = acceptancePayload(true);
  assert.equal(payload.records.length, 20);
  assert.deepEqual(payload.groups.map((group) => [group.code, group.records.length]), [["NORD", 12], ["SUD", 7]]);
  assert.equal(payload.records.filter((record) => !record.code).length, 1);
  assert.equal(payload.regionReadiness.available, 3);
  assert.equal(payload.regionReadiness.blank, 2);
  assert.deepEqual(payload.competencies.map((name) => behaviorRows.filter((row) => row.competency === name).length), [10, 9, 6, 5]);
  assert.equal(behaviorRows.length, 30);
  assert.equal(new Set(behaviorRows.map((row) => row.behavior)).size, 29);
  assert.deepEqual([...fixture.csv.slice(0, 3)], [0xef, 0xbb, 0xbf]);
});

test("F11 generated whole and split decks have valid cloned-slide identity and package structure", async () => {
  for (const splitGroups of [false, true]) {
    const payload = acceptancePayload(splitGroups);
    const generated = await generatedDeck(payload);
    await selfCheckPptx(generated.zip);
    const slides = slideFiles(generated.zip);
    const chartUniqueIds = new Set();
    const chartSupportTargets = new Set();
    assert.equal(slides.length, reportPlan(payload).length);
    const presentation = await generated.zip.file("ppt/presentation.xml").async("string");
    const slideIds = [...presentation.matchAll(/<p:sldId\b[^>]*id="(\d+)"[^>]*r:id="([^"]+)"/gu)].map((match) => match[1]);
    assert.equal(new Set(slideIds).size, slideIds.length);
    assert.equal(slideIds.length, slides.length);
    for (const slide of slides) {
      const xml = await generated.zip.file(slide).async("string");
      assert.doesNotMatch(xml, /<a:ext\b[^>]*cy="[^" ]*\/>/u);
      assert.match(xml, /name="GRF-R role:title"/u);
      assert.match(xml, /name="GRF-R role:body"/u);
      const rels = relationships(await generated.zip.file(`ppt/slides/_rels/${slide.split("/").at(-1)}.rels`).async("string"));
      for (const rel of rels.filter((entry) => entry.Type.endsWith("/chart"))) {
        const chartPath = relationshipTarget(slide, rel.Target); const chart = await generated.zip.file(chartPath).async("string");
        for (const unique of chart.matchAll(/<c16:uniqueId\b[^>]*val="([^"]+)"/gu)) { assert(!chartUniqueIds.has(unique[1]), `duplicate chart uniqueId ${unique[1]}`); chartUniqueIds.add(unique[1]); }
        const chartRels = relationships(await generated.zip.file(`ppt/charts/_rels/${chartPath.split("/").at(-1)}.rels`).async("string"));
        for (const support of chartRels.filter((entry) => /\/(?:chartStyle|chartColorStyle|themeOverride)$/u.test(entry.Type))) { assert(!chartSupportTargets.has(support.Target), `shared chart support part ${support.Target}`); chartSupportTargets.add(support.Target); }
        [...chart.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/gu)].forEach((series, index) => { assert.match(series[1], new RegExp(`<c:idx val="${index}"\\/>`, "u")); assert.match(series[1], new RegExp(`<c:order val="${index}"\\/>`, "u")); });
      }
    }
  }
});

test("F3 every generated slide renders its inserted role text and long labels inside the text extents", async (t) => {
  if (!(await binaryCommandAvailable("soffice")) || !(await binaryCommandAvailable("pdftoppm")) || !(await binaryCommandAvailable("python3"))) { t.skip("LibreOffice, Poppler or Python is unavailable"); return; }
  const renderRoot = await mkdtemp(join(tmpdir(), "grf-r-render-"));
  try {
    for (const splitGroups of [false, true]) {
      const payload = acceptancePayload(splitGroups);
      const generated = await generatedDeck(payload);
      const expectedBody = [];
      for (const slide of slideFiles(generated.zip)) { const xml = await generated.zip.file(slide).async("string"); const body = xml.match(/<p:cNvPr\b[^>]*name="GRF-R role:body"[\s\S]*?<p:txBody>([\s\S]*?)<\/p:txBody>/u); expectedBody.push(Boolean(body && /<a:t>[^<\s][\s\S]*?<\/a:t>/u.test(body[1]))); }
      const deckPath = join(renderRoot, `${splitGroups ? "split" : "whole"}.pptx`);
      await writeFile(deckPath, generated.bytes);
      await run("soffice", ["--headless", "--convert-to", "pdf", "--outdir", renderRoot, deckPath]);
      const pdfPath = join(renderRoot, `${splitGroups ? "split" : "whole"}.pdf`);
      const pngDir = join(renderRoot, splitGroups ? "split-png" : "whole-png");
      await run("mkdir", ["-p", pngDir]);
      await run("pdftoppm", ["-png", "-r", "96", pdfPath, join(pngDir, "slide")]);
      const analysis = `import json,sys
from PIL import Image
from pathlib import Path
results=[]
for path in sorted(Path(sys.argv[1]).glob('slide-*.png')):
  image=Image.open(path).convert('RGB'); w,h=image.size
  def ink(box):
    crop=image.crop(box); pixels=crop.load(); xs=[]
    for y in range(crop.height):
      for x in range(crop.width):
        r,g,b=pixels[x,y]
        if r < 120 and g < 120 and b < 120: xs.append(x)
    return len(xs), (max(xs) if xs else -1)
  title=(*ink((int(w*.06),int(h*.04),int(w*.96),int(h*.24))),w)
  body=(*ink((int(w*.06),int(h*.16),int(w*.96),int(h*.84))),w)
  results.append((path.name,title,body))
print(json.dumps(results))`;
      const { stdout } = await run("python3", ["-c", analysis, pngDir]);
      const raster = JSON.parse(stdout);
      assert.equal(raster.length, reportPlan(payload).length);
      for (const [index, [name, title, body]] of raster.entries()) {
        assert(title[0] > 0, `${name} has no rendered title ink`);
        assert(title[1] < title[2] * .96, `${name} title ink reaches the text-box edge`);
        if (expectedBody[index]) { assert(body[0] > 0, `${name} has no rendered body ink`); assert(body[1] < body[2] * .96, `${name} body ink reaches the text-box edge`); }
      }
    }
  } finally { await rm(renderRoot, { recursive: true, force: true }); }
});

test("F2 group slides and benchmark charts use only NORD/SUD participant values", async () => {
  const payload = acceptancePayload(true);
  const plan = reportPlan(payload);
  const generated = await generatedDeck(payload);
  for (const [group, expected] of [["NORD", [0, 0, 12]], ["SUD", [7, 0, 0]]]) {
    const benchmark = plan.findIndex((item) => item.groupKey === group && item.family === "benchmark");
    assert(benchmark >= 0);
    const slideText = textFromSlide(await generated.zip.file(`ppt/slides/slide${benchmark + 1}.xml`).async("string"));
    assert.match(slideText, new RegExp(`Sub 2\\.75: ${expected[0]}`, "u"));
    assert.match(slideText, new RegExp(`În interval: ${expected[1]}`, "u"));
    assert.match(slideText, new RegExp(`Peste 3\\.5: ${expected[2]}`, "u"));
    assert.doesNotMatch(slideText, new RegExp(`Sub 2\\.75: ${expected[0] === 0 ? 7 : 0}`, "u"));
  }
  for (const item of plan.filter((candidate) => candidate.groupKey)) {
    const index = plan.indexOf(item);
    const xml = await generated.zip.file(`ppt/slides/slide${index + 1}.xml`).async("string");
    assert.match(xml, new RegExp(` · ${item.groupKey}`, "u"));
  }
});

test("F4 every chart series has its own non-empty workbook column, cache and range", async () => {
  const generated = await generatedDeck(acceptancePayload(true));
  for (const slide of slideFiles(generated.zip)) {
    const rels = relationships(await generated.zip.file(`ppt/slides/_rels/${slide.split("/").at(-1)}.rels`).async("string"));
    for (const rel of rels.filter((entry) => entry.Type.endsWith("/chart"))) {
      const chartPath = relationshipTarget(slide, rel.Target); const chart = await generated.zip.file(chartPath).async("string");
      const chartRels = relationships(await generated.zip.file(`ppt/charts/_rels/${chartPath.split("/").at(-1)}.rels`).async("string"));
      const workbookRel = chartRels.find((entry) => entry.Type.endsWith("/package"));
      assert(workbookRel);
      const workbookPath = relationshipTarget(chartPath, workbookRel.Target); const book = XLSX.read(await generated.zip.file(workbookPath).async("array"), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { header: 1, defval: "" });
      const usedColumns = new Set();
      for (const series of chart.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/gu)) {
        const body = series[1]; const cat = body.match(/<c:cat>[\s\S]*?<c:f>([^<]+)<\/c:f>[\s\S]*?<c:strCache>([\s\S]*?)<\/c:strCache>[\s\S]*?<\/c:cat>/u); const val = body.match(/<c:val>[\s\S]*?<c:f>([^<]+)<\/c:f>[\s\S]*?<c:numCache>([\s\S]*?)<\/c:numCache>[\s\S]*?<\/c:val>/u);
        assert(cat && val, `chart ${chartPath} has a series without category/value cache`);
        const categoryColumn = cat[1].match(/\$([A-Z]+)\$/u)?.[1]; const valueColumn = val[1].match(/\$([A-Z]+)\$/u)?.[1];
        assert.equal(categoryColumn, "A"); assert(valueColumn); assert(!usedColumns.has(valueColumn), `chart ${chartPath} reuses ${valueColumn}`); usedColumns.add(valueColumn);
        const cachedCategories = [...cat[2].matchAll(/<c:v>([^<]*)<\/c:v>/gu)].map((match) => decode(match[1]));
        const cachedValues = [...val[2].matchAll(/<c:v>([^<]*)<\/c:v>/gu)].map((match) => Number(match[1]));
        assert.deepEqual(cachedCategories, rows.slice(1).map((row) => String(row[0] ?? "")));
        const columnIndex = valueColumn.split("").reduce((sum, letter) => sum * 26 + letter.charCodeAt(0) - 64, 0) - 1;
        assert.deepEqual(cachedValues, rows.slice(1).map((row) => Number(row[columnIndex] ?? 0)));
      }
    }
  }
});

test("F13/F6/F12 bundle output is one cropped, nonblank, outlined item per manifest entry", async (t) => {
  const havePdfTools = await binaryCommandAvailable("pdfinfo") && await binaryCommandAvailable("pdftoppm") && await binaryCommandAvailable("python3");
  if (!havePdfTools || !(await binaryCommandAvailable("qlmanage"))) { t.skip("PDF or Quick Look raster tools are unavailable"); return; }
  global.window = { opentype: await loadUmd(resolve(root, "src/assets/vendor/opentype.min.js")), __GRF_FONT_BASE64__: (await readFile(font)).toString("base64") };
  const artifacts = await buildBundleArtifacts(XLSX, acceptancePayload(true));
  assert.equal(artifacts.manifest.items.length, 75);
  for (const item of artifacts.svg) { assert.doesNotMatch(item.content, /<text|foreignObject|font-family/iu); assert.match(item.content, /<path\b/u); }
  const temp = await mkdtemp(join(tmpdir(), "grf-r-bundle-"));
  try {
    const pdfPath = join(temp, "items.pdf"); await writeFile(pdfPath, artifacts.pdfBytes);
    const { stdout: info } = await run("pdfinfo", [pdfPath]);
    assert.match(info, /Pages:\s+75/u); assert.match(info, /Page size:\s+1600 x 900/u);
    const pdfPng = join(temp, "pdf"); await run("mkdir", ["-p", pdfPng]); await run("pdftoppm", ["-png", "-r", "72", pdfPath, join(pdfPng, "item")]);
    const svgDir = join(temp, "svg"); await run("mkdir", ["-p", svgDir]);
    const svgSamples = artifacts.svg.map((item, index) => ({ item, index, score: artifacts.manifest.items[index].sourceValues.join(" ").length })).sort((a, b) => b.score - a.score).slice(0, 4);
    for (const { item, index } of svgSamples) { const path = join(temp, `${String(index + 1).padStart(3, "0")}.svg`); await writeFile(path, item.content); await run("qlmanage", ["-t", "-s", "1600", "-o", svgDir, path]); }
    const check = `from PIL import Image
from pathlib import Path
import sys,json
def ink(path):
 im=Image.open(path).convert('RGB'); w,h=im.size; px=im.load(); coords=[]
 for y in range(int(h*.12),int(h*.88)):
  for x in range(int(w*.06),int(w*.96)):
   r,g,b=px[x,y]
   if r<120 and g<120 and b<120: coords.append((x,y))
 return len(coords),max((x for x,_ in coords),default=-1),w
print(json.dumps([(str(p),ink(p)) for p in sorted(Path(sys.argv[1]).glob('*.png'))]))`;
    const { stdout: pdfCheck } = await run("python3", ["-c", check, pdfPng]);
    for (const [path, count] of JSON.parse(pdfCheck)) { assert(count[0] > 0, `${path} is blank`); assert(count[1] < count[2] * .96, `${path} clips ink at the right edge`); }
    const { stdout: svgCheck } = await run("python3", ["-c", check, svgDir]);
    for (const [path, count] of JSON.parse(svgCheck)) { assert(count[0] > 0, `${path} is blank`); assert(count[1] < count[2] * .96, `${path} clips ink at the right edge`); }
  } finally { await rm(temp, { recursive: true, force: true }); }
});

test("DoD 12 split bundle items equal the generated deck items one for one", async () => {
  const payload = acceptancePayload(true); const plan = reportPlan(payload).filter((item) => !["cover", "close"].includes(item.family));
  global.window = { opentype: await loadUmd(resolve(root, "src/assets/vendor/opentype.min.js")), __GRF_FONT_BASE64__: (await readFile(font)).toString("base64") };
  const artifacts = await buildBundleArtifacts(XLSX, payload);
  assert.deepEqual(artifacts.manifest.items.map((item) => [item.title, item.group]), plan.map((item) => [item.title, item.groupKey || "whole-project"]));
});
