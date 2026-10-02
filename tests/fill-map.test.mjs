// BP-GRF-R fill map (binding): assertions on the GENERATED files, not on the code that makes them.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import test from "node:test";
import { buildPayload } from "../src/rebuild-core.js";
import { BUNDLE_FAMILIES, f2, participantsPerSlide, reportPlan, splitEqual } from "../src/rebuild-report-plan.js";
import { buildBundleArtifacts, generateTrendPptx } from "../src/template-pptx.js";
import { getShape, shapeText, tableRows, visualBox, xfrmOf, NEW_SHAPE_PREFIX } from "../src/trend-fill.js";
import { createFixture as acceptanceFixture, fixtureRows as acceptanceRows } from "./fixtures/grf-r-acceptance-fixture.mjs";
import { createFixture as variedFixture, participants as variedParticipants } from "./fixtures/grf-r-varied-fixture.mjs";

const root = resolve(new URL("..", import.meta.url).pathname);
const asset = resolve(root, "src/assets/trend/template-raport-de-grup-RO.pptx");
const loadUmd = async (path) => { const module = { exports: {} }; new Function("module", "exports", "require", await readFile(path, "utf8"))(module, module.exports, undefined); return module.exports; };
const XLSX = await loadUmd(resolve(root, "src/assets/vendor/xlsx.full.min.js"));
const JSZip = await loadUmd(resolve(root, "src/assets/vendor/jszip.min.js"));
const opentype = await loadUmd(resolve(root, "src/assets/vendor/opentype.min.js"));
globalThis.window = { JSZip, XLSX, opentype, __GRF_TEMPLATE_BASE64__: (await readFile(asset)).toString("base64"), __GRF_FONT_BASE64__: (await readFile(resolve(root, "src/assets/vendor/Poppins-Regular.ttf"))).toString("base64") };
const decode = (value) => String(value).replace(/&amp;/gu, "&").replace(/&lt;/gu, "<").replace(/&gt;/gu, ">").replace(/&quot;/gu, "\"").replace(/&apos;/gu, "'");
const text = (shapeXml) => decode(shapeText(shapeXml));
const EMU = 914400;

const payloadOf = (fixture, metadata = {}) => buildPayload(XLSX, [{ name: "summary.xlsx", bytes: fixture.summary }, { name: "detail.xlsx", bytes: fixture.detailed }, { name: "evaluation-sheet-template.csv", bytes: fixture.csv }], { ...fixture.metadata, reportDate: "2026-10-01", ...metadata }, {}, { acknowledgedWarningIds: [] });
const acceptance = acceptanceFixture(XLSX); const varied = variedFixture(XLSX);
const template = await JSZip.loadAsync(await readFile(asset));
const templateSlides = new Map(); for (let index = 1; index <= 22; index += 1) templateSlides.set(index, await template.file(`ppt/slides/slide${index}.xml`).async("string"));
const shapeIds = (xml) => [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"[^>]*\bname="([^"]*)"/gu)].map((match) => ({ id: match[1], name: decode(match[2]) }));
const relationships = (xml) => [...xml.matchAll(/<Relationship\b([^>]*)\/>/gu)].map((match) => Object.fromEntries([...match[1].matchAll(/(Id|Type|Target)="([^"]*)"/gu)].map((part) => [part[1], part[2]])));
const resolveTarget = (source, target) => { const parts = `${source.split("/").slice(0, -1).join("/")}/${target}`.split("/"); const output = []; for (const part of parts) { if (part === "..") output.pop(); else if (part && part !== ".") output.push(part); } return output.join("/"); };

async function deck(payload, scope = "whole") {
  const zip = await JSZip.loadAsync(Buffer.from(await (await generateTrendPptx(payload, { scope })).arrayBuffer()));
  const plan = reportPlan(payload, { scope });
  const slides = []; for (let index = 1; index <= plan.length; index += 1) slides.push(await zip.file(`ppt/slides/slide${index}.xml`).async("string"));
  const chartsOf = async (index) => { const rels = relationships(await zip.file(`ppt/slides/_rels/slide${index + 1}.xml.rels`).async("string")); const output = []; for (const rel of rels.filter((entry) => entry.Type.endsWith("/chart"))) { const path = resolveTarget(`ppt/slides/slide${index + 1}.xml`, rel.Target); const chart = await zip.file(path).async("string"); const chartRels = relationships(await zip.file(path.replace("charts/", "charts/_rels/") + ".rels").async("string")); const book = XLSX.read(await zip.file(resolveTarget(path, chartRels.find((entry) => entry.Type.endsWith("/package")).Target)).async("array"), { type: "array" }); output.push({ path, chart, book }); } return output; };
  return { zip, plan, slides, chartsOf };
}
const whole = await deck(payloadOf(acceptance, { splitGroups: false, annex: "end" }));
const split = await deck(payloadOf(acceptance, { splitGroups: true, annex: "end" }));
const variedDeck = await deck(payloadOf(varied, { splitGroups: true, annex: "end" }));

test("rule 1: every generated slide holds only its template slide's shapes plus the map's New shapes", () => {
  const allowedNew = { cover: 1, "key-findings": 1, "executive-summary": 9 };
  for (const generated of [whole, split, variedDeck]) {
    generated.plan.forEach((item, index) => {
      const templateIds = new Set(shapeIds(templateSlides.get(item.templateIndex)).map((shape) => shape.id));
      const shapes = shapeIds(generated.slides[index]);
      const extra = shapes.filter((shape) => !templateIds.has(shape.id));
      assert(extra.every((shape) => shape.name.startsWith(NEW_SHAPE_PREFIX)), `slide ${index + 1} (${item.family}) has an undeclared shape: ${extra.map((shape) => shape.name).join(", ")}`);
      assert.equal(extra.length, allowedNew[item.family] || 0, `slide ${index + 1} (${item.family}) New-shape count`);
      const removed = item.noRanking ? new Set(["17", "18", "19", "20"]) : new Set(); // F41: R5 boxes removed when nothing is ranked
      for (const id of templateIds) if (!removed.has(id)) assert(shapes.some((shape) => shape.id === id), `slide ${index + 1} lost template shape ${id}`);
    });
  }
});

test("report order follows map §1: no duplicate key findings, t11 only as the annex divider, closing last", () => {
  const payload = payloadOf(acceptance, { splitGroups: true });
  const families = split.plan.map((item) => item.family);
  assert.deepEqual(families.slice(0, 4), ["cover", "how-to-read", "methodology", "executive-summary"]);
  assert.equal(split.plan.filter((item) => item.family === "key-findings").length, payload.competencies.length);
  assert.equal(split.plan.filter((item) => item.templateIndex === 12).length, payload.competencies.length);
  assert(!families.includes("observation"));
  assert.deepEqual(split.plan.filter((item) => item.templateIndex === 11).map((item) => item.family), ["appendix-divider"]);
  assert.equal(families.at(-1), "close");
  assert.deepEqual(split.slides.filter((xml) => /Constatări cheie|Observații ·/u.test(xml)), []);
  const groupStart = split.plan.findIndex((item) => item.groupKey);
  assert(split.plan.slice(0, groupStart).every((item) => !item.groupKey));
  assert(split.plan.filter((item) => item.groupKey).every((item) => ["divider-results", "range", "ranking", "benchmark", "population", "zone", "divider-behaviors", "behavior", "divider-conclusions", "conclusions"].includes(item.family)));
});

test("M1–M4: cover, how-to-read, methodology and executive summary fill their named shapes", () => {
  const [cover, howTo, method, summary] = whole.slides;
  assert.equal(text(getShape(cover, 15)), "Client sintetic cu etichete românești lungi – Centru de Dezvoltare");
  assert.match(text(getShape(cover, 23)), /^2026 © www\.trendconsult\.eu/u);
  assert.equal(text(getShape(cover, 12)), "Raport de grup");
  assert.match(cover, /GRF-R new:confidential[\s\S]*?CONFIDENȚIAL/u);
  assert(xfrmOf(cover.slice(cover.indexOf("GRF-R new:confidential") - 200)).y > xfrmOf(getShape(cover, 20)).y + xfrmOf(getShape(cover, 20)).cy);
  assert.equal(text(getShape(howTo, 93)), "CUM CITIM ACEST RAPORT");
  assert.equal(text(getShape(howTo, 5)).split("\n").length, 2);
  assert.equal(text(getShape(howTo, 6)).split("\n").length, 2);
  assert.match(text(getShape(howTo, 6)), /Rezultatele individuale se regăsesc în anexă\.$/u);
  assert.match(text(getShape(method, 93)), /METODOLOGIE/u);
  assert.match(text(getShape(method, 5)), /^20 participanți Client sintetic/u);
  assert.match(text(getShape(method, 6)), /scală de la 1 la 5/u);
  for (const slide of [howTo, method]) assert.doesNotMatch(slide, /\(0, 1|0–1–2|0-1-2|\(0-2\)/u);
  assert.equal(text(getShape(summary, 5)), "Executive Summary");
  assert.equal(text(getShape(summary, 8)), "Imaginea de ansamblu");
  assert.match(text(getShape(summary, 4)), /^Evaluarea celor 20 participanți: \d+% dintre participanți/u);
  assert.equal(text(getShape(summary, 9)), "Competențe");
  assert.match(text(getShape(summary, 7)), /^Cel mai bine reprezentată: .+ \(\d\.\d\d\)\nPrincipala oportunitate: .+ \(\d\.\d\d\)$/u);
  assert.equal(text(getShape(summary, 10)), "Concluzii comportamentale");
  const box = xfrmOf(getShape(summary, 4));
  for (const shape of summary.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?GRF-R new:infographic[\s\S]*?<\/p:sp>/gu)) { const at = xfrmOf(shape); assert(at.x >= box.x && at.x + at.cx <= box.x + box.cx + 1 && at.y >= box.y && at.y + at.cy <= box.y + box.cy + 1, "infographic inside id 4's box"); }
});

test("M5 key findings: table rows per scored participant, braces span their bands, R5 lists with shares", () => {
  const payload = payloadOf(acceptance);
  for (const [index, item] of whole.plan.entries()) {
    if (item.family !== "key-findings") continue;
    const xml = whole.slides[index];
    assert.equal(text(getShape(xml, 21)), `Distribuția pe competențe – ${item.competency}`);
    const frame = getShape(xml, 6); const rows = tableRows(frame); const box = xfrmOf(frame);
    assert.equal(rows.length, payload.records.filter((record) => Number.isFinite(record.scores[item.competency])).length);
    const cells = rows.map((row) => (row.match(/<a:tc\b[\s\S]*?<\/a:tc>/gu) || []).map((cell) => decode([...cell.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].map((match) => match[1]).join(""))));
    assert.deepEqual(cells, item.ladder.map((entry) => [entry.name, f2(entry.value)]), "name beside its score, in sorted order");
    // Hand-checked pairs from the fixture: participants 01–12 score 4.25, 13–19 score 2.25, 20 scores 3.25.
    for (const [name, score] of [["Participant sintetic 01", "4.25"], ["Participant sintetic 13", "2.25"], ["Participant sintetic 20", "3.25"]]) assert(cells.some(([cellName, cellScore]) => cellName === name && cellScore === score), `${name} → ${score}`);
    let y = box.y; const spans = {};
    rows.forEach((row, rowIndex) => { const h = Number(row.match(/\bh="(\d+)"/u)[1]); const value = item.scores[rowIndex]; const key = value > item.high ? "above" : value < item.low ? "below" : "in"; spans[key] = spans[key] ? { ...spans[key], bottom: y + h } : { top: y, bottom: y + h }; y += h; });
    for (const [key, braceId, labelId] of [["above", 8, 11], ["in", 14, 16], ["below", 12, 13]]) {
      const brace = getShape(xml, braceId);
      if (!spans[key]) { assert.match(brace, /hidden="1"/u); continue; }
      const at = xfrmOf(brace); assert(Math.abs(at.y - spans[key].top) < 2 && Math.abs(at.y + at.cy - spans[key].bottom) < 2, `brace ${braceId} spans its band`);
      assert.equal(text(getShape(xml, labelId)), String(item.counts[key]));
    }
    assert.equal(text(getShape(xml, 19)), "Abilități cheie – Puncte forte");
    assert.equal(text(getShape(xml, 20)), "Arii de dezvoltare");
    const strengths = text(getShape(xml, 17)).split("\n").filter(Boolean); assert.equal(strengths.length, item.insight.key.length);
    strengths.forEach((line) => assert.match(line, / \(\d+%\)$/u));
    assert.match(xml, new RegExp(`GRF-R new:subtitle[\\s\\S]*?medie ${f2(item.mean)} · mediană ${f2(item.median)}`, "u"));
  }
});

test("M7–M15 and annex charts are native clones with cache = workbook = <c:f>, bands on a fixed 1–5 axis", async () => {
  for (const generated of [split, variedDeck]) {
    for (const [index, item] of generated.plan.entries()) {
      const charts = await generated.chartsOf(index);
      if (["range", "ranking", "population", "zone", "participant-mean", "participant-comparison", "competency-participants"].includes(item.family)) assert.equal(charts.length, 1, `slide ${index + 1} ${item.family} keeps its native chart`);
      for (const { chart, book } of charts) {
        assert.deepEqual(book.SheetNames, ["Sheet1"]);
        const rows = XLSX.utils.sheet_to_json(book.Sheets.Sheet1, { header: 1, defval: null });
        const series = chart.match(/<c:ser>[\s\S]*?<\/c:ser>/gu);
        series.forEach((body, seriesIndex) => {
          const column = String.fromCharCode(66 + seriesIndex);
          assert.match(body, new RegExp(`<c:tx><c:strRef><c:f>Sheet1!\\$${column}\\$1</c:f>`, "u"));
          assert.match(body, new RegExp(`<c:val><c:numRef><c:f>Sheet1!\\$${column}\\$2:\\$${column}\\$${rows.length}</c:f>`, "u"));
          assert.equal(decode(body.match(/<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/u)[1]), rows[0][seriesIndex + 1]);
          const categories = [...body.match(/<c:cat>[\s\S]*?<\/c:cat>/u)[0].matchAll(/<c:v>([^<]*)<\/c:v>/gu)].map((match) => decode(match[1]));
          assert.deepEqual(categories, rows.slice(1).map((row) => String(row[0])));
          const cached = new Map([...body.match(/<c:val>[\s\S]*?<\/c:val>/u)[0].matchAll(/<c:pt idx="(\d+)"><c:v>([^<]*)<\/c:v>/gu)].map((match) => [Number(match[1]), Number(match[2])]));
          rows.slice(1).forEach((row, rowIndex) => assert.equal(cached.get(rowIndex) ?? null, row[seriesIndex + 1]));
        });
        assert.doesNotMatch(chart, /<c:v>(?:undefined|null|NaN)<\/c:v>|Maximizarea|<c:v>C1<\/c:v>|<c:v>Cluj<|<c:v>Iasi</u);
        if (item.family !== "population") assert.match(chart, /<c:valAx>[\s\S]*?<c:scaling><c:orientation val="minMax"\/><c:max val="5"\/><c:min val="1"\/><\/c:scaling>/u);
      }
    }
  }
});

test("rule 4: bands move with an edited benchmark (t13 band maps 1–5 from the template rectangle)", async () => {
  const edited = await deck(payloadOf(acceptance, { benchmarkLow: "3", benchmarkHigh: "4" }));
  const index = edited.plan.findIndex((item) => item.family === "competency-participants");
  // The band is placed from the chart's pinned plot area (F21): y(v) = plotTop + (5 − v) / 4 × plotHeight.
  const moved = visualBox(getShape(edited.slides[index], 3)); const frame = xfrmOf(getShape(edited.slides[index], 2));
  const [{ chart }] = await edited.chartsOf(index);
  const [y, h] = ["y", "h"].map((name) => Number(chart.match(new RegExp(`<c:plotArea><c:layout><c:manualLayout>[\\s\\S]*?<c:${name} val="([^"]+)"`, "u"))[1]));
  const plotTop = frame.y + y * frame.cy; const plotHeight = h * frame.cy;
  assert(Math.abs(moved.y - (plotTop + 1 / 4 * plotHeight)) < 3000 && Math.abs(moved.cy - plotHeight / 4) < 3000, JSON.stringify({ moved, plotTop, plotHeight }));
  const benchmark = edited.slides[edited.plan.findIndex((item) => item.family === "benchmark")];
  assert.equal(text(getShape(benchmark, 19)), "Rezultate raportate la benchmark (3.00-4.00)");
  const range = edited.slides[edited.plan.findIndex((item) => item.family === "range")];
  assert.match(text(getShape(range, 26)), /între 3\.00 și 4\.00/u);
  for (const xml of edited.slides) assert.doesNotMatch(text(xml.replace(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?hidden="1"[\s\S]*?<\/p:sp>/gu, "")), /2\.75/u);
});

test("rules 5/6/9: tables resize by rows, braces follow rows, group suffix on titles and dividers", () => {
  for (const [index, item] of split.plan.entries()) {
    const xml = split.slides[index];
    if (item.family === "benchmark") { const rows = tableRows(getShape(xml, 2)); assert.equal(rows.length, item.table.rows.length); const ids = rows.map((row) => row.match(/rowId[^>]*val="(\d+)"/u)?.[1]); assert.equal(new Set(ids).size, ids.length); }
    if (item.family === "behavior") assert.equal(tableRows(getShape(xml, 7)).length, 1 + Math.max(item.key.length, item.development.length));
    if (item.groupKey) assert(decode(xml).includes(` · ${item.groupLabel}`), `slide ${index + 1} lacks the group suffix`);
    const frame = item.family === "benchmark" ? getShape(xml, 2) : item.family === "key-findings" ? getShape(xml, 6) : null;
    if (frame) { const total = tableRows(frame).reduce((sum, row) => sum + Number(row.match(/\bh="(\d+)"/u)[1]), 0); assert.equal(xfrmOf(frame).cy, total); assert(total <= 10.0 * EMU + 2); }
  }
});

test("rule 10 / A2–A4: annex charts are real charts, participants split equally and readably", () => {
  assert.deepEqual(splitEqual(Array.from({ length: 21 }, (_, index) => index), 20).map((page) => page.length), [11, 10]);
  assert.deepEqual(splitEqual(Array.from({ length: 23 }, (_, index) => index), 5).map((page) => page.length), [5, 5, 5, 4, 4]);
  assert.equal(participantsPerSlide("participant-comparison", { seriesCount: 6 }), 5);
  const plan = variedDeck.plan;
  const pages = (family, competency) => plan.filter((item) => item.family === family && (!competency || item.competency === competency)).map((item) => item.rows.length);
  assert.deepEqual(pages("participant-mean"), [23]);
  assert.deepEqual(pages("participant-comparison"), [5, 5, 5, 4, 4]);
  for (const name of plan.filter((item) => item.family === "competency-participants").map((item) => item.competency)) assert.deepEqual(pages("competency-participants", name), [12, 11]);
  const sixth = plan.findIndex((item) => item.family === "competency-participants" && item.competencyIndex === 5);
  assert.equal(plan[sixth].templateIndex, 13);
  assert.equal(variedParticipants.length, 23);
  assert.match(variedDeck.slides[sixth], /<p:graphicFrame>/u);
});

test("sixth competency colour FF9D75 on the A4 clone, the M8 series and the zone points", async () => {
  const sixth = variedDeck.plan.findIndex((item) => item.family === "competency-participants" && item.competencyIndex === 5);
  const [{ chart }] = await variedDeck.chartsOf(sixth);
  assert.match(chart.match(/<c:ser>[\s\S]*?<\/c:ser>/u)[0], /FF9D75/u);
  assert.doesNotMatch(chart.match(/<c:ser>[\s\S]*?<\/c:ser>/u)[0], /003057/u);
  const [{ chart: ranking }] = await variedDeck.chartsOf(variedDeck.plan.findIndex((item) => item.family === "ranking"));
  assert.equal(ranking.match(/<c:ser>/gu).length, 6); assert.match(ranking.match(/<c:ser>[\s\S]*?<\/c:ser>/gu)[5], /FF9D75/u);
  const [{ chart: zone }] = await variedDeck.chartsOf(variedDeck.plan.findIndex((item) => item.family === "zone"));
  assert.equal(zone.match(/<c:ser>/gu).length, 4); assert.match(zone, /<c:dPt><c:idx val="5"\/>[\s\S]*?FF9D75/u);
});

test("no undefined/null/NaN or template example data in any generated slide", () => {
  for (const generated of [whole, split, variedDeck]) for (const xml of generated.slides) {
    const content = decode([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].map((match) => match[1]).join(" "));
    assert.doesNotMatch(content, /undefined|\bnull\b|NaN|People Management|Colaborare și asertivitate|Gestionarea schimbării|Workshop de Change|2024|\(0, 1/u);
  }
});

test("separate annex file: cover with „Anexă”, A1–A4, closing", async () => {
  const annex = await deck(payloadOf(acceptance, { annex: "separate" }), "appendix");
  assert.equal(annex.plan[0].family, "cover"); assert.match(text(getShape(annex.slides[0], 12)), /Raport de grup – Anexă/u);
  assert.equal(annex.plan.at(-1).family, "close");
  assert(annex.plan.slice(1, -1).every((item) => item.deliverable === "appendix"));
  const main = await deck(payloadOf(acceptance, { annex: "separate" }), "main");
  assert(!main.plan.some((item) => item.deliverable === "appendix"));
  assert.match(text(getShape(main.slides[1], 6)), /în anexa transmisă separat\.$/u);
});

test("bundle (F15): one cropped item per chart and table, values equal to the deck", async () => {
  const payload = payloadOf(acceptance, { splitGroups: true });
  const artifacts = await buildBundleArtifacts(XLSX, payload);
  const items = split.plan.filter((item) => BUNDLE_FAMILIES.has(item.family));
  assert.equal(artifacts.manifest.items.length, items.length);
  const pdf = new TextDecoder("latin1").decode(artifacts.pdfBytes);
  const boxes = [...pdf.matchAll(/\/MediaBox \[0 0 (\d+) (\d+)\]/gu)].map((match) => [Number(match[1]), Number(match[2])]);
  for (const [index, entry] of artifacts.manifest.items.entries()) {
    const svg = artifacts.svg[index].content;
    assert.match(svg, new RegExp(`width="${entry.width}" height="${entry.height}"`, "u"));
    assert.deepEqual(boxes[index], [entry.pdfPage.width, entry.pdfPage.height]);
    assert(entry.pdfPage.height > entry.height, "the PDF page adds the caption strip outside the item (F27)");
    assert(!(entry.width === 1600 && entry.height === 900));
    assert.doesNotMatch(svg, /<text|<image|href=|font-family|undefined|NaN/u);
    const slideIndex = entry.deckSlide - 1; assert.equal(split.plan[slideIndex].title, entry.title);
    const charts = await split.chartsOf(slideIndex);
    const haystack = decode(split.slides[slideIndex] + charts.map((chart) => chart.chart).join("")).replace(/[\u200B-\u200D\u2060\uFEFF]/gu, "");
    for (const value of entry.sourceValues) assert(haystack.includes(value), `bundle item ${entry.id} value ${value} is not in deck slide ${entry.deckSlide}`);
  }
});

test("E: a key-finding share equals a hand count from the raw fixture rows", () => {
  const { detailed } = acceptanceRows();
  const column = 5; // first behaviour of the first competency (Leadership …)
  const scores = detailed.slice(3).map((row) => row[column]).filter((value) => value !== "");
  const expected = Math.round(scores.filter((value) => value === 2).length / scores.length * 100);
  const index = whole.plan.findIndex((item) => item.family === "key-findings" && item.competency === detailed[0][column]);
  const behaviour = detailed[2][column];
  const lines = text(getShape(whole.slides[index], 17)).split("\n").concat(text(getShape(whole.slides[index], 18)).split("\n"));
  const line = lines.find((entry) => entry.includes(behaviour));
  assert(line, "the behaviour is ranked into a key-findings list");
  const share = /Să exersezi/u.test(line) ? Math.round(scores.filter((value) => value === 0).length / scores.length * 100) : expected;
  assert(line.endsWith(`(${share}%)`), `${line} ≠ (${share}%)`);
  const shares = new Set(whole.slides.filter((_, slideIndex) => whole.plan[slideIndex].family === "key-findings").flatMap((xml) => text(getShape(xml, 17)).match(/\(\d+%\)/gu) || []));
  assert(shares.size > 1, "shares differ between behaviours");
});

test("D: box plot keeps the template's median diamond and the varied fixture has a median strictly inside the range", async () => {
  const index = variedDeck.plan.findIndex((item) => item.family === "range");
  const item = variedDeck.plan[index];
  assert(item.items.some((row) => row.median > row.min && row.median < row.max));
  const [{ chart }] = await variedDeck.chartsOf(index);
  const median = chart.match(/<c:ser>[\s\S]*?<\/c:ser>/gu).find((series) => /<c:v>MEDIAN<\/c:v>/u.test(series));
  assert.match(median, /<c:marker><c:symbol val="diamond"\/>/u);
});

test("ruling 8: R5 below six behaviours splits evenly or skips the middle one, on the generated tables", async () => {
  const workbook = (rows) => { const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1"); return XLSX.write(book, { type: "buffer", bookType: "xlsx" }); };
  const counts = [1, 2, 3, 4, 5]; const names = counts.map((count) => `Competență cu ${count} comportamente`);
  const columns = counts.flatMap((count, competencyIndex) => Array.from({ length: count }, (_, index) => ({ competency: names[competencyIndex], behavior: `C${count}-B${index + 1}`, index })));
  const people = Array.from({ length: 4 }, (_, index) => ({ name: `Test ${index + 1}`, id: `T-${index + 1}` }));
  const summary = workbook([["CODE", "name", "cod cp", ...names], ...people.map((person, index) => ["", person.name, person.id, ...names.map(() => 2 + (index % 3))])]);
  const detailed = workbook([["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", ...columns.map((column) => column.competency)], ["", "", "", "", "Subcompetente", ...columns.map(() => "S")], ["", "", "", "", "behavior", ...columns.map((column) => column.behavior)], ...people.map((person) => ["", person.name, "R", person.id, "", ...columns.map((column) => Math.max(0, 2 - Math.floor(column.index / 2)))])]);
  const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: summary }, { name: "detail.xlsx", bytes: detailed }], { projectName: "R5 boundary", clientName: "Test" }, {}, { acknowledgedWarningIds: [] });
  const generated = await deck(payload);
  const expected = { 1: 0, 2: 1, 3: 1, 4: 2, 5: 2 };
  for (const count of counts) {
    const name = names[count - 1];
    const behaviorIndex = generated.plan.findIndex((item) => item.family === "behavior" && item.competency === name);
    if (!expected[count]) { assert.equal(behaviorIndex, -1, "one behaviour: nothing to rank, no table"); continue; }
    const plan = generated.plan[behaviorIndex];
    assert.equal(plan.key.length, expected[count]); assert.equal(plan.development.length, expected[count]);
    assert.equal(tableRows(getShape(generated.slides[behaviorIndex], 7)).length, 1 + expected[count]);
    if (count % 2) assert(!plan.key.concat(plan.development).includes(`C${count}-B${(count + 1) / 2}`), "the middle behaviour is skipped");
    const finding = generated.plan.findIndex((item) => item.family === "key-findings" && item.competency === name);
    assert.equal(text(getShape(generated.slides[finding], 17)).split("\n").filter(Boolean).length, expected[count]);
  }
});

test("B: text fit uses explicit run sizes, never normAutofit, body text ≥ 10 pt", () => {
  for (const generated of [whole, split, variedDeck]) generated.slides.forEach((xml, index) => {
    assert.doesNotMatch(xml, /normAutofit|lnSpcReduction/u, `slide ${index + 1}`);
    if (["key-findings", "how-to-read", "methodology", "executive-summary", "behavior", "benchmark"].includes(generated.plan[index].family)) {
      const ids = { "key-findings": [17, 18], "how-to-read": [5, 6], methodology: [5, 6], "executive-summary": [4, 7, 11], behavior: [7], benchmark: [4, 17] }[generated.plan[index].family];
      for (const id of ids) for (const size of getShape(xml, id).matchAll(/<a:rPr\b[^>]*\ssz="(\d+)"/gu)) assert(Number(size[1]) >= 1000, `slide ${index + 1} id ${id} has ${size[1] / 100} pt`);
    }
  });
});

test("C: methodology omits a consultant fact left empty instead of printing a number-less line", () => {
  const method = whole.slides[2];
  const lines = text(getShape(method, 5)).split("\n");
  assert(lines.every((line) => /^\d/u.test(line) || /^Alte instrumente/u.test(line)), lines.join(" | "));
  assert(!lines.some((line) => /^consultanți|^zile|^exerciții/u.test(line)));
  assert.deepEqual(whole.plan[2].page.missingLabels, ["consultanți TREND implicați", "zile de evaluare", "număr de exerciții"]);
});

test("F23: sibling boxes share one size, and every title of a series has one size (≥ 60 % of the template)", () => {
  const largest = (shape) => Math.max(0, ...[...shape.matchAll(/<a:rPr\b[^>]*\ssz="(\d+)"[^>]*>(?:(?!<\/a:r>)[\s\S])*?<a:t>[^<]/gu)].map((match) => Number(match[1])));
  const siblings = { "how-to-read": [5, 6], methodology: [5, 6], "key-findings": [17, 18], benchmark: [4, 17] };
  const titles = { "key-findings": [21, 3750], "competency-participants": [5, 3600], behavior: [3, 3750] };
  for (const generated of [whole, split, variedDeck]) {
    const series = {};
    generated.plan.forEach((item, index) => {
      const xml = generated.slides[index];
      if (siblings[item.family]) { const sizes = siblings[item.family].map((id) => largest(getShape(xml, id))).filter(Boolean); assert.equal(new Set(sizes).size, 1, `slide ${index + 1} ${item.family} siblings ${sizes}`); }
      if (titles[item.family]) (series[item.family] ||= []).push(largest(getShape(xml, titles[item.family][0])));
    });
    for (const [family, sizes] of Object.entries(series)) { assert.equal(new Set(sizes).size, 1, `${family} titles ${[...new Set(sizes)]}`); assert(sizes[0] >= titles[family][1] * 0.6 - 50, `${family} title ≥ 60 %`); }
  }
});

test("F27: PNG/SVG items carry no caption; F29: app.xml counts the generated slides; F26: how-to-read follows the toggles", async () => {
  const payload = payloadOf(acceptance, { splitGroups: true });
  const artifacts = await buildBundleArtifacts(XLSX, payload);
  for (const [index, entry] of artifacts.manifest.items.entries()) {
    const titlePaths = artifacts.layouts[index].texts.filter((text) => text.caption).length;
    assert(titlePaths > 0);
    const svgPaths = (artifacts.svg[index].content.match(/<path\b/gu) || []).length;
    const itemTexts = artifacts.layouts[index].texts.filter((text) => !text.caption).length;
    assert.equal(svgPaths, itemTexts, `SVG ${entry.id} holds only the item's text`);
  }
  const app = await split.zip.file("docProps/app.xml").async("string");
  assert.match(app, new RegExp(`<Slides>${split.plan.length}</Slides>`, "u")); assert.doesNotMatch(app, /TitlesOfParts|<Notes>/u);
  const off = reportPlan(payloadOf(acceptance, { slideToggles: { range: false, keyFindings: false, behavior: false } }));
  const paragraphs = off.find((item) => item.family === "how-to-read").paragraphs.join(" ");
  assert.doesNotMatch(paragraphs, /Graficele de distribuție|Abilitățile cheie|Procentele/u);
  const on = reportPlan(payloadOf(acceptance)).find((item) => item.family === "how-to-read").paragraphs.join(" ");
  assert.match(on, /Graficele de distribuție/u); assert.match(on, /Procentele indică/u);
});

const { createFixture: insp5Fixture } = await import("./fixtures/grf-r-insp5-fixture.mjs");
const insp5 = insp5Fixture(XLSX);
const longLines = (count, prefix) => Array.from({ length: count }, (_, index) => `${prefix} ${index + 1}: text de test al consultantului, formulat suficient de lung pentru a verifica încadrarea în casetă la dimensiunea șablonului`).join("\n");
const filledDeck = await deck(payloadOf(insp5, { splitGroups: true, conclusionsStrengths: longLines(4, "Punct forte"), conclusionsDevelopment: longLines(4, "Arie"), conclusionsInterventions: longLines(6, "Intervenție"), groupConclusions: { MGR: { strengths: "Grup MGR: punct forte specific", interventions: "Grup MGR: intervenție" } } }));

test("F31/F23 ruling: text fills at the largest size that fits (template size when it fits); F32 per-group conclusions", () => {
  const size = (shape) => Math.max(0, ...[...shape.matchAll(/<a:rPr\b[^>]*\ssz="(\d+)"[^>]*>(?:(?!<\/a:r>)[\s\S])*?<a:t>[^<]/gu)].map((match) => Number(match[1])));
  const whole = filledDeck.plan.findIndex((item) => item.family === "conclusions" && !item.groupKey);
  assert.equal(size(getShape(filledDeck.slides[whole], 11)), 2000, "the interventions box keeps the template's 20 pt when its text fits");
  assert(size(getShape(filledDeck.slides[whole], 4)) >= 1000);
  const mgr = filledDeck.plan.findIndex((item) => item.family === "conclusions" && item.groupKey === "MGR");
  assert.equal(text(getShape(filledDeck.slides[mgr], 11)), "Grup MGR: intervenție");
  assert.equal(text(getShape(filledDeck.slides[mgr], 4)), "Grup MGR: punct forte specific");
  const spc = filledDeck.plan.findIndex((item) => item.family === "conclusions" && item.groupKey === "SPC");
  assert.equal(text(getShape(filledDeck.slides[spc], 11)), "", "a group without its own fields stays blank, never the whole-project text");
});

test("F33: on t7 the longest competency line ends before the mean's right tab", () => {
  for (const [index, item] of filledDeck.plan.entries()) {
    if (item.family !== "benchmark") continue;
    const shape = getShape(filledDeck.slides[index], 4);
    const tab = Number(shape.match(/<a:tab pos="(\d+)" algn="r"\/>/u)[1]);
    for (const paragraph of shape.match(/<a:p>[\s\S]*?<\/a:p>/gu)) {
      const last = [...paragraph.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].at(-1)[1]; const sizePt = Number(paragraph.match(/sz="(\d+)"/u)[1]) / 100;
      const [name, mean] = decode(last).split("\t");
      assert.match(mean, /^\d\.\d\d$/u);
      assert(name.length * sizePt * 0.5 * 12700 + (mean.length + 1) * sizePt * 0.5 * 12700 <= tab, `„${name}” runs into its mean`);
    }
  }
});

const insp5Deck = await deck(payloadOf(insp5, { splitGroups: true }));
test("F30/F34/F35: A3 legend sized to name every competency; one annex label size per series; wrapped names keep A4 to equal readable pages; bundle bars carry values only", async () => {
  const comparison = insp5Deck.plan.map((item, index) => [item, index]).filter(([item]) => item.family === "participant-comparison");
  for (const [item, index] of comparison) {
    const [{ chart }] = await insp5Deck.chartsOf(index);
    const legend = chart.match(/<c:legend>[\s\S]*?<\/c:legend>/u)[0];
    const size = Number(legend.match(/<a:defRPr\b[^>]*\bsz="(\d+)"/u)[1]) / 100; const h = Number(legend.match(/<c:h val="([^"]+)"/u)[1]);
    const frameCy = 9.82 * 72; const width = 0.998 * 17.94 * 72 - 20;
    let rows = 1; let line = 0; for (const name of item.competencies) { const w = name.length * size * 0.55 + size * 2; if (line && line + w > width) { rows += 1; line = w; } else line += w; }
    assert(size >= 10 && h * frameCy >= rows * size * 1.45, `legend ${size} pt × ${rows} rows fits its box`);
  }
  for (const family of ["participant-comparison", "competency-participants"]) assert.equal(new Set(insp5Deck.plan.filter((item) => item.family === family).map((item) => item.labelSize)).size, 1, `${family} label size`);
  const a4 = insp5Deck.plan.filter((item) => item.family === "competency-participants" && item.competencyIndex === 0).map((item) => item.rows.length);
  assert.deepEqual(a4, [11, 11]);
  const artifacts = await buildBundleArtifacts(XLSX, payloadOf(insp5, { splitGroups: true }));
  artifacts.manifest.items.forEach((entry, index) => {
    if (entry.family !== "participant-comparison") return;
    const names = new Set(insp5Deck.plan.find((item) => item.family === "participant-comparison").competencies);
    const nameTexts = artifacts.layouts[index].texts.filter((text) => !text.caption && names.has(text.line));
    assert(nameTexts.length <= names.size, "series names appear in the legend only, never inside bars");
  });
});

test("F36: each bundle image carries its own heading; files are named NN-item-scope (ASCII)", async () => {
  const artifacts = await buildBundleArtifacts(XLSX, payloadOf(insp5, { splitGroups: true }));
  for (const [index, entry] of artifacts.manifest.items.entries()) {
    assert.match(entry.file, /^\d{3}-[a-z-]+-[a-z0-9-]+$/u);
    const itemTexts = artifacts.layouts[index].texts.filter((text) => !text.caption).map((text) => text.line).join(" ");
    const first = entry.heading.split(/\s+/u).slice(0, 2).join(" ");
    assert(itemTexts.includes(first), `${entry.file} shows its heading „${entry.heading}” inside the item`);
    if (entry.family === "key-findings") assert.match(itemTexts, /medie \d\.\d\d · mediană \d\.\d\d/u);
    assert.doesNotMatch(artifacts.layouts[index].texts.filter((text) => text.caption).map((text) => text.line).join(" "), new RegExp(entry.title.slice(0, 20).replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"), "u"));
  }
  assert.equal(new Set(artifacts.svg.map((item) => item.name)).size, artifacts.svg.length);
});

test("F39: with every slide type off and no annex the how-to-read keeps only the opening and the closing sentence", () => {
  const off = reportPlan(payloadOf(insp5, { annex: "none", slideToggles: { range: false, competencyMean: false, benchmark: false, competencyDistribution: false, zone: false, observation: false, behavior: false, conclusions: false } }));
  assert.deepEqual(off.find((item) => item.family === "how-to-read").paragraphs, ["Rezultatele pe competențe sunt exprimate pe o scală de la 1 la 5, unde 1 reprezintă nivelul minim, iar 5 nivelul maxim.", "Rezultatele descriu grupul evaluat."]);
});

test("A2 follows rule 10: labels ≥ 10 pt, one size per series, equal readable pages (insp5: 11 + 11)", async () => {
  const pages = insp5Deck.plan.map((item, index) => [item, index]).filter(([item]) => item.family === "participant-mean");
  assert.deepEqual(pages.map(([item]) => item.rows.length), [11, 11]);
  assert.equal(new Set(pages.map(([item]) => item.labelSize)).size, 1);
  for (const [item, index] of pages) {
    const [{ chart }] = await insp5Deck.chartsOf(index);
    const size = Number(chart.match(/<c:catAx>[\s\S]*?<a:defRPr\b[^>]*\bsz="(\d+)"/u)[1]);
    assert(size >= 1000, `A2 label ${size / 100} pt`); assert.equal(size, Math.min(2400, item.labelSize));
  }
});

test("zero-width breaks stay in the deck's charts: bundle and audit carry clean names", async () => {
  const invisible = /[\u200B\u200C\u200D\u2060\uFEFF]/u;
  const payload = payloadOf(insp5, { splitGroups: true });
  const artifacts = await buildBundleArtifacts(XLSX, payload);
  assert.doesNotMatch(JSON.stringify(artifacts.manifest), invisible);
  for (const item of artifacts.svg) { assert.doesNotMatch(item.name, invisible); assert.doesNotMatch(item.content, invisible); }
  for (const layout of artifacts.layouts) for (const text of layout.texts) assert.doesNotMatch(text.line, invisible);
  const data = XLSX.read(artifacts.workbookBytes, { type: "array" });
  for (const name of data.SheetNames) for (const row of XLSX.utils.sheet_to_json(data.Sheets[name], { header: 1, defval: "" })) for (const cell of row) assert.doesNotMatch(String(cell), invisible);
  const { createAuditWorkbook } = await import("../src/rebuild-core.js");
  const audit = createAuditWorkbook(XLSX, payload);
  for (const name of audit.SheetNames) for (const row of XLSX.utils.sheet_to_json(audit.Sheets[name], { header: 1, defval: "" })) for (const cell of row) assert.doesNotMatch(String(cell), invisible);
});

const { createFixture: insp6Fixture } = await import("./fixtures/grf-r-insp6-fixture.mjs");
const insp6 = insp6Fixture(XLSX);
const insp6Payload = payloadOf(insp6, { splitGroups: true });
const insp6Deck = await deck(insp6Payload);

test("F41: a competency with one scored behaviour keeps table/braces/mean·median, loses its R5 boxes, has no behaviour slide, and the bundle has no empty rows", async () => {
  const { unrankedCompetencies } = await import("../src/rebuild-report-plan.js");
  const unranked = unrankedCompetencies(insp6Payload);
  assert.equal(unranked.length, 1);
  const index = insp6Deck.plan.findIndex((item) => item.family === "key-findings" && item.competency === unranked[0]);
  const xml = insp6Deck.slides[index];
  for (const id of [17, 18, 19, 20]) assert(!hasShapeId(xml, id), `id ${id} removed`);
  for (const id of [6, 8, 14, 12]) assert(hasShapeId(xml, id));
  assert.match(xml, /GRF-R new:subtitle[\s\S]*?medie \d\.\d\d · mediană/u);
  assert(!insp6Deck.plan.some((item) => item.family === "behavior" && item.competency === unranked[0]));
  const artifacts = await buildBundleArtifacts(XLSX, insp6Payload);
  const entry = artifacts.manifest.items.findIndex((item) => item.family === "key-findings" && item.heading.startsWith(unranked[0]));
  assert(!artifacts.layouts[entry].texts.some((text) => /Abilități cheie|Arii de dezvoltare/u.test(text.line)));
});
const hasShapeId = (xml, id) => new RegExp(`<p:cNvPr\\b[^>]*\\bid="${id}"`, "u").test(xml);

test("F42: tied strongest/weakest competencies are all named; F44: a missing score names the participant", () => {
  const summary = insp6Deck.plan.find((item) => item.family === "executive-summary").summary;
  const means = insp6Payload.calculations.filter((item) => item.mean !== null).map((item) => f2(item.mean));
  const lowest = means.slice().sort()[0];
  if (means.filter((mean) => mean === lowest).length > 1) assert.match(summary.competencyLines[1], / și .* \(\d\.\d\d\)$/u);
  assert.match(summary.competencyLines[1], new RegExp(`\\(${lowest.replace(".", "\\.")}\\)$`, "u"));
  const blank = insp6Payload.warnings.find((item) => item.code === "detailed-blank");
  assert(blank && /^Scor lipsă: Participant Sintetic B\d\d – „/u.test(blank.message), blank?.message);
});

test("F43: bundle file names keep group and page parts and stay unique", async () => {
  const artifacts = await buildBundleArtifacts(XLSX, insp6Payload);
  const names = artifacts.manifest.items.map((item) => item.file);
  assert.equal(new Set(names.map((name) => name.slice(4))).size, names.length, "unique even without the number");
  for (const item of artifacts.manifest.items) {
    if (item.group !== "whole-project") assert(item.file.endsWith(`-${item.groupLabel.toLowerCase()}`) || new RegExp(`-${item.groupLabel.toLowerCase()}-p\\d+$`, "u").test(item.file), item.file);
  }
  const paged = insp6Deck.plan.filter((item) => item.pages > 1 && BUNDLE_FAMILIES.has(item.family));
  for (const item of paged) assert(names.some((name) => name.endsWith(`-p${item.page}`)));
});

test("F45: a zone bar too short for its name carries the name outside its end", async () => {
  for (const [index, item] of insp6Deck.plan.entries()) {
    if (item.family !== "zone") continue;
    const [{ chart }] = await insp6Deck.chartsOf(index);
    item.values.forEach((values, seriesIndex) => values.forEach((value, point) => {
      if (value === null || value > 1.2) return;
      const series = chart.match(/<c:ser>[\s\S]*?<\/c:ser>/gu)[seriesIndex];
      assert.match(series, new RegExp(`<c:dLbl><c:idx val="${point}"/>[\\s\\S]*?<c:dLblPos val="outEnd"/>`, "u"), `${item.regions[seriesIndex]} at ${value}`);
    }));
  }
});

test("t7 ladder: every participant's name sits beside their overall mean", () => {
  const index = whole.plan.findIndex((item) => item.family === "benchmark");
  const rows = tableRows(getShape(whole.slides[index], 2));
  const cells = rows.map((row) => (row.match(/<a:tc\b[\s\S]*?<\/a:tc>/gu) || []).map((cell) => decode([...cell.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].map((match) => match[1]).join(""))));
  assert.equal(cells.length, 20);
  assert(cells.every(([name, value]) => /^Participant sintetic \d\d$/u.test(name) && /^\d\.\d\d$/u.test(value)));
  assert.deepEqual(cells.find(([name]) => name === "Participant sintetic 20"), ["Participant sintetic 20", "3.25"]);
  const sizes = rows.map((row) => Number((row.match(/<a:tc\b[\s\S]*?<\/a:tc>/u)[0].match(/\ssz="(\d+)"/u) || [0, 0])[1]));
  assert.equal(new Set(sizes).size, 1); assert(sizes[0] >= 900);
});
