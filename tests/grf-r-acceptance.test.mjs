import assert from "node:assert/strict";
import { access, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { join, resolve } from "node:path";
import { tmpdir } from "node:os";
import test from "node:test";
import { buildPayload } from "../src/rebuild-core.js";
import { BUNDLE_FAMILIES, reportPlan } from "../src/rebuild-report-plan.js";
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
      assert.doesNotMatch(xml, /GRF-R role:/u);
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

const tool = async (name) => { const override = process.env[`GRF_${name.toUpperCase()}`]; if (override) return override; try { await run("which", [name]); return name; } catch { return null; } };
test("F3/F14 every generated slide renders nonblank with no text dumped in the top-left corner", async (t) => {
  const soffice = await tool("soffice"); const pdftoppm = await tool("pdftoppm");
  if (!soffice || !pdftoppm || !(await binaryCommandAvailable("python3"))) { t.skip("LibreOffice, Poppler or Python is unavailable (set GRF_SOFFICE / GRF_PDFTOPPM)"); return; }
  const renderRoot = await mkdtemp(join(tmpdir(), "grf-r-render-"));
  try {
    for (const splitGroups of [false, true]) {
      const payload = acceptancePayload(splitGroups);
      const generated = await generatedDeck(payload);
      const name = splitGroups ? "split" : "whole";
      const deckPath = join(renderRoot, `${name}.pptx`);
      await writeFile(deckPath, generated.bytes);
      await run(soffice, ["--headless", "--convert-to", "pdf", "--outdir", renderRoot, deckPath], { timeout: 300000 });
      const pngDir = join(renderRoot, `${name}-png`); await run("mkdir", ["-p", pngDir]);
      await run(pdftoppm, ["-png", "-r", "24", join(renderRoot, `${name}.pdf`), join(pngDir, "slide")]);
      const { stdout } = await run("python3", ["-c", `import json,sys
from PIL import Image
from pathlib import Path
out=[]
for p in sorted(Path(sys.argv[1]).glob('slide-*.png')):
  im=Image.open(p).convert('L'); w,h=im.size; px=im.load()
  ink=sum(1 for y in range(h) for x in range(w) if px[x,y]<200)
  out.append(ink)
print(json.dumps(out))`, pngDir]);
      const ink = JSON.parse(stdout);
      assert.equal(ink.length, reportPlan(payload).length);
      ink.forEach((count, index) => assert(count > 50, `slide ${index + 1} renders blank`));
    }
  } finally { await rm(renderRoot, { recursive: true, force: true }); }
});

test("F2 group slides and benchmark charts use only NORD/SUD participant values", async () => {
  const payload = acceptancePayload(true);
  const plan = reportPlan(payload);
  const generated = await generatedDeck(payload);
  for (const [group, expected] of [["NORD", [0, 0, 12]], ["SUD", [7, 0, 0]]]) { // below / in / above
    const benchmark = plan.findIndex((item) => item.groupKey === group && item.family === "benchmark");
    assert(benchmark >= 0);
    const xml = await generated.zip.file(`ppt/slides/slide${benchmark + 1}.xml`).async("string");
    const total = expected.reduce((sum, value) => sum + value, 0);
    const shares = expected.map((value) => `${Math.round(value / total * 100)}% au obținut`);
    for (const share of shares) assert(textFromSlide(xml).includes(share), `${group} benchmark slide lacks ${share}`);
    assert.equal((xml.match(/<a:tr\b/gu) || []).length, total, `${group} benchmark table has one row per participant`);
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
  const pdfinfo = await tool("pdfinfo"); const pdftoppm = await tool("pdftoppm");
  const havePdfTools = pdfinfo && pdftoppm && await binaryCommandAvailable("python3");
  if (!havePdfTools) { t.skip("PDF raster tools are unavailable"); return; }
  global.window = { opentype: await loadUmd(resolve(root, "src/assets/vendor/opentype.min.js")), __GRF_FONT_BASE64__: (await readFile(font)).toString("base64") };
  const artifacts = await buildBundleArtifacts(XLSX, acceptancePayload(true));
  const expectedItems = reportPlan(acceptancePayload(true)).filter((item) => BUNDLE_FAMILIES.has(item.family)).length;
  assert.equal(artifacts.manifest.items.length, expectedItems);
  for (const item of artifacts.svg) { assert.doesNotMatch(item.content, /<text|foreignObject|font-family/iu); assert.match(item.content, /<path\b/u); }
  const temp = await mkdtemp(join(tmpdir(), "grf-r-bundle-"));
  try {
    const pdfPath = join(temp, "items.pdf"); await writeFile(pdfPath, artifacts.pdfBytes);
    const { stdout: info } = await run(pdfinfo, [pdfPath]);
    assert.match(info, new RegExp(`Pages:\\s+${expectedItems}`, "u"));
    const pdfPng = join(temp, "pdf"); await run("mkdir", ["-p", pdfPng]); await run(pdftoppm, ["-png", "-r", "72", pdfPath, join(pdfPng, "item")]);
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
  } finally { await rm(temp, { recursive: true, force: true }); }
});

test("DoD 12 split bundle items equal the generated deck items one for one", async () => {
  const payload = acceptancePayload(true); const plan = reportPlan(payload).filter((item) => BUNDLE_FAMILIES.has(item.family));
  global.window = { opentype: await loadUmd(resolve(root, "src/assets/vendor/opentype.min.js")), __GRF_FONT_BASE64__: (await readFile(font)).toString("base64") };
  const artifacts = await buildBundleArtifacts(XLSX, payload);
  assert.deepEqual(artifacts.manifest.items.map((item) => [item.title, item.group]), plan.map((item) => [item.title, item.groupKey || "whole-project"]));
});

const MEASURE_BAND = `import json,sys
from PIL import Image
a=json.loads(sys.argv[1]); im=Image.open(a['png']).convert('RGB'); W,H=im.size; d=a['dpi']
fx,fy,fw,fh=a['frame']; lx,ly,lw,lh=a['layout']
left=(fx+lx*fw)*d; top=(fy+ly*fh)*d; width=lw*fw*d; height=lh*fh*d
g=lambda p,lo,hi: abs(p[0]-p[1])<6 and abs(p[1]-p[2])<6 and lo<=p[0]<=hi
if a['axis']=='y':
  c=int(left+width*a.get('probe',0.05)); pix=[(y,im.getpixel((c,y))) for y in range(int(top-20),min(H,int(top+height+20)))]
else:
  c=int(top+height/a['slots']); pix=[(x,im.getpixel((x,c))) for x in range(max(0,int(left-20)),min(W,int(left+width+20)))]
runs=[]
for pos,p in pix:
  if g(p,100,238):
    if runs and pos-runs[-1][1]<=1 and abs(runs[-1][2]-p[0])<12: runs[-1][1]=pos
    else: runs.append([pos,pos,p[0]])
thin=[r for r in runs if r[1]-r[0]<=3]; wide=[r for r in runs if r[1]-r[0]>=8]
centers=[(r[0]+r[1])/2 for r in thin]
band=[pos for r in wide for pos in (r[0],r[1])]
steps=sorted(b-a for a,b in zip(centers,centers[1:])); step=steps[len(steps)//2]
first=centers[0] if a['axis']=='y' else centers[0]
if a['axis']=='y': value=lambda pos: 5-(pos-first)/step*a['unit']
else: value=lambda pos: 1+(pos-first)/step*a['unit']
lo,hi=(value(max(band)+0.5),value(min(band)-0.5)) if a['axis']=='y' else (value(min(band)-0.5),value(max(band)+0.5))
print(json.dumps({'low':lo,'high':hi,'gridlines':len(centers)}))`;

test("F21: on the pinned charts (M7 box plot, A2) the rendered band lands on the benchmark (default and 3.00–3.75)", async (t) => {
  const soffice = await tool("soffice"); const pdftoppm = await tool("pdftoppm");
  if (!soffice || !pdftoppm || !(await binaryCommandAvailable("python3"))) { t.skip("LibreOffice, Poppler or Python is unavailable (set GRF_SOFFICE / GRF_PDFTOPPM)"); return; }
  const { createFixture: inspectorFixture } = await import("./fixtures/grf-r-insp5-fixture.mjs");
  const source = inspectorFixture(XLSX);
  const temp = await mkdtemp(join(tmpdir(), "grf-r-band-"));
  try {
    for (const [low, high] of [[2.75, 3.5], [3, 3.75]]) {
      const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: source.summary }, { name: "detail.xlsx", bytes: source.detailed }], { ...source.metadata, splitGroups: false, benchmarkLow: String(low), benchmarkHigh: String(high) }, {}, { acknowledgedWarningIds: [] });
      const plan = reportPlan(payload); const generated = await generatedDeck(payload);
      const deckPath = join(temp, `band-${low}.pptx`); await writeFile(deckPath, generated.bytes);
      await run(soffice, ["--headless", "--convert-to", "pdf", "--outdir", temp, deckPath], { timeout: 300000 });
      for (const [family, frameId, axis, unit, nth] of [["range", 22, "y", 0.5, 0], ["participant-mean", 2, "x", 0.5, 0], ["participant-comparison", 6, "y", 0.5, 0], ["competency-participants", 2, "y", 0.5, 0]]) {
        const index = plan.findIndex((item) => item.family === family);
        const slide = await generated.zip.file(`ppt/slides/slide${index + 1}.xml`).async("string");
        const frameShape = slide.slice(slide.lastIndexOf("<p:graphicFrame", slide.indexOf(`<p:cNvPr id="${frameId}"`)));
        const xfrm = frameShape.match(/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"/u).slice(1).map((value) => Number(value) / 914400);
        const rel = (await generated.zip.file(`ppt/slides/_rels/slide${index + 1}.xml.rels`).async("string")).match(/Target="\.\.\/charts\/(chart\d+\.xml)"/u)[1];
        const chart = await generated.zip.file(`ppt/charts/${rel}`).async("string");
        const layout = ["x", "y", "w", "h"].map((name) => Number(chart.match(new RegExp(`<c:plotArea><c:layout><c:manualLayout>[\\s\\S]*?<c:${name} val="([^"]+)"`, "u"))[1]));
        const png = join(temp, `${family}-${low}`);
        await run(pdftoppm, ["-png", "-r", "100", "-f", String(index + 1), "-l", String(index + 1), join(temp, `band-${low}.pdf`), png]);
        const file = (await readdir(temp)).find((name) => name.startsWith(`${family}-${low}`) && name.endsWith(".png"));
        const { stdout } = await run("python3", ["-c", MEASURE_BAND, JSON.stringify({ png: join(temp, file), axis, frame: xfrm, layout, dpi: 100, unit, slots: plan[index].rows?.length || 1, probe: family === "range" ? 0.05 : 1 / (plan[index].rows?.length || 1) })]);
        const measured = JSON.parse(stdout); if (process.env.GRF_DEBUG) console.log("MEASURED", family, low, high, JSON.stringify(measured));
        assert(Math.abs(measured.low - low) < 0.04 && Math.abs(measured.high - high) < 0.04, `${family} band renders at ${measured.low.toFixed(2)}–${measured.high.toFixed(2)}, expected ${low}–${high}`);
      }
    }
  } finally { await rm(temp, { recursive: true, force: true }); }
});

test("A2 render: every participant label is its own readable block (≥ 10 pt, no overlap)", async (t) => {
  const soffice = await tool("soffice"); const pdftoppm = await tool("pdftoppm");
  if (!soffice || !pdftoppm || !(await binaryCommandAvailable("python3"))) { t.skip("LibreOffice, Poppler or Python is unavailable"); return; }
  const { createFixture: inspectorFixture } = await import("./fixtures/grf-r-insp5-fixture.mjs");
  const source = inspectorFixture(XLSX);
  const temp = await mkdtemp(join(tmpdir(), "grf-r-a2-"));
  try {
    const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: source.summary }, { name: "detail.xlsx", bytes: source.detailed }], { ...source.metadata, splitGroups: false }, {}, { acknowledgedWarningIds: [] });
    const plan = reportPlan(payload); const generated = await generatedDeck(payload);
    await writeFile(join(temp, "a2.pptx"), generated.bytes);
    await run(soffice, ["--headless", "--convert-to", "pdf", "--outdir", temp, join(temp, "a2.pptx")], { timeout: 300000 });
    for (const [index, item] of plan.entries()) {
      if (item.family !== "participant-mean") continue;
      const slide = await generated.zip.file(`ppt/slides/slide${index + 1}.xml`).async("string");
      const rel = (await generated.zip.file(`ppt/slides/_rels/slide${index + 1}.xml.rels`).async("string")).match(/Target="\.\.\/charts\/(chart\d+\.xml)"/u)[1];
      const chart = await generated.zip.file(`ppt/charts/${rel}`).async("string");
      const frameShape = slide.slice(slide.lastIndexOf("<p:graphicFrame", slide.indexOf('<p:cNvPr id="2"')));
      const [fx, fy, fw, fh] = frameShape.match(/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"/u).slice(1).map((value) => Number(value) / 914400);
      const lx = Number(chart.match(/<c:manualLayout>[\s\S]*?<c:x val="([^"]+)"/u)[1]);
      await run(pdftoppm, ["-png", "-r", "100", "-f", String(index + 1), "-l", String(index + 1), join(temp, "a2.pdf"), join(temp, `p${index}`)]);
      const file = (await readdir(temp)).find((name) => name.startsWith(`p${index}`) && name.endsWith(".png"));
      const { stdout } = await run("python3", ["-c", `import json,sys
from PIL import Image
a=json.loads(sys.argv[1]); im=Image.open(a['png']).convert('L'); d=100
x0=int((a['fx']+0.05)*d); x1=int((a['fx']+a['lx']*a['fw'])*d)-12; y0=int((a['fy']+0.07*a['fh'])*d); y1=int((a['fy']+0.97*a['fh'])*d)
rows=[any(im.getpixel((x,y))<140 for x in range(x0,x1)) for y in range(y0,y1)]
spans=[]; start=None
for i,r in enumerate(rows+[False]):
  if r and start is None: start=i
  if not r and start is not None: spans.append([start,i]); start=None
merged=[]
for a0,a1 in spans:
  if merged and a0-merged[-1][1]<=4: merged[-1][1]=a1
  else: merged.append([a0,a1])
print(json.dumps([b-a for a,b in merged]))`, JSON.stringify({ png: join(temp, file), fx, fy, fw, fh, lx })]);
      const blocks = JSON.parse(stdout).filter((height) => height >= 3);
      assert(blocks.length >= item.rows.length, `slide ${index + 1}: ${blocks.length} label blocks for ${item.rows.length} participants (labels overlap or are missing)`);
      const glyph = Math.min(...blocks); // one line of ≥ 10 pt text is at least ~9 px tall at 100 dpi (cap height + descender)
      assert(glyph >= 11, `slide ${index + 1}: label lines are ${glyph} px tall at 100 dpi (< 12 pt)`);
    }
  } finally { await rm(temp, { recursive: true, force: true }); }
});
