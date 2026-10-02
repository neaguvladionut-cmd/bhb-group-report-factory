import { reportPlan, BUNDLE_FAMILIES, f2 } from "./rebuild-report-plan.js";
import { EMU, NEW_SHAPE_PREFIX, xmlEscape, getShape, updateShape, hasShape, hideShape, addToTree, xfrmOf, setXfrm, setParagraphs, appendToLastRun, setRunText, templateParagraphs, regularRPr, boldRPr, setRPrAttr, setRPrColor, fitText, withoutBullet, setBodyInsets, moveBand, plotLayoutFromBand, visualBox, groupScale, fitTitleOneLine, setVisualBox, fitTextSized, scaleRunSizes, tableRows, rowHeight, isShadedRow, setRowCells, setRowHeight, setRowId, replaceTableRows, rowLayout, placeBrace, fillChartXml, chartWorkbookRows, roundChartValue, newTextShape, newRectShape } from "./trend-fill.js";

const TEMPLATE_PATH = "./assets/trend/template-raport-de-grup-RO.pptx";
const FONT_PATH = "./assets/vendor/Poppins-Regular.ttf";
const relsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const asBytes = (base64) => typeof Buffer !== "undefined" ? Buffer.from(base64, "base64") : Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
const slideName = (index) => `ppt/slides/slide${index}.xml`;
const slideRelsName = (index) => `ppt/slides/_rels/slide${index}.xml.rels`;
const TITLE_ACCENT = "D49802"; // t12 title competency run (orange)
const NAVY = "003057"; const GREY = "A5A5A5"; const GOLD = "DBA900"; const SIXTH = "FF9D75";
// Per-competency colours: the template's five series colours, then FF9D75 (DoD 6).
const ZONE_CATEGORY_COLORS = ["003057", "4472C4", "00A19A", "DBA900", "545454", SIXTH];
const SERIES_SIXTH = ["", "", "", "", "", SIXTH];

export const FIXED_TEMPLATE_LABELS = [
  "LEGENDĂ", "Benchmark pe baza evaluărilor Trend", "Abilități cheie", "Abilități de dezvoltat", "Arii de dezvoltare",
  "Puncte forte și recomandări de grup", "Concluzii și recomandări", "MULȚUMIM!", "8th Menuetului Street",
  "Bucharest 013713, Romania", "office@trendconsult.eu", "www.trendconsult.eu", "Raport de grup"
];
// Shapes the fill map creates („New”), by family. Everything else on a slide is a template shape.
export const NEW_SHAPES = { cover: ["confidential"], "key-findings": ["subtitle"], "executive-summary": ["infographic"] };

async function templateBytes() {
  if (window.__GRF_TEMPLATE_BASE64__) return asBytes(window.__GRF_TEMPLATE_BASE64__);
  const response = await fetch(TEMPLATE_PATH, { cache: "no-store" });
  if (!response.ok) throw new Error("Nu am putut încărca șablonul Trend curățat.");
  return new Uint8Array(await response.arrayBuffer());
}
async function fontBytes() {
  if (window.__GRF_FONT_BASE64__) return asBytes(window.__GRF_FONT_BASE64__);
  const response = await fetch(FONT_PATH, { cache: "no-store" });
  if (!response.ok) throw new Error("Nu am putut încărca fontul Poppins vendorizat.");
  return new Uint8Array(await response.arrayBuffer());
}
const allSlideNumbers = (zip) => Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).map((name) => Number(name.match(/slide(\d+)\.xml/u)[1])).sort((a, b) => a - b);
const allChartNumbers = (zip) => Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/u.test(name)).map((name) => Number(name.match(/chart(\d+)\.xml/u)[1])).sort((a, b) => a - b);

// ---------------------------------------------------------------- chart data (deck and bundle share it)
const competencyLabel = (value) => String(value ?? "");
export function chartSpec(item) {
  if (item.family === "range") return { data: { categories: item.items.map((row) => competencyLabel(row.competency)), series: [{ name: "MIN", values: item.items.map((row) => row.min) }, { name: "MAX", values: item.items.map((row) => row.max) }, { name: "MEDIAN", values: item.items.map((row) => row.median) }] }, options: { fixedAxis: true }, band: { id: 2, axis: "y" } };
  if (item.family === "ranking") return { data: { categories: ["Media"], series: item.items.map((row) => ({ name: competencyLabel(row.competency), values: [row.mean] })) }, options: { fixedAxis: true, seriesColors: SERIES_SIXTH }, band: { id: 2, axis: "x" } };
  if (item.family === "population") return { data: { categories: item.rows.map((row) => row.competency), series: [{ name: "Low", values: item.rows.map((row) => row.below / 100) }, { name: "BENCH", values: item.rows.map((row) => row.in / 100) }, { name: "High", values: item.rows.map((row) => row.above / 100) }] }, options: {} };
  if (item.family === "zone") return { data: { categories: item.competencies, series: item.regions.map((region, index) => ({ name: region, values: item.values[index] })) }, options: { fixedAxis: true, categoryColors: ZONE_CATEGORY_COLORS }, band: { id: 2, axis: "y" } };
  if (item.family === "participant-mean") return { data: { categories: item.rows.map((row) => row.name), series: [{ name: "Media", values: item.rows.map((row) => row.mean) }] }, options: { fixedAxis: true }, band: { id: 3, axis: "x" } };
  if (item.family === "participant-comparison") return { data: { categories: item.rows.map((row) => row.name), series: item.competencies.map((competency) => ({ name: competency, values: item.rows.map((row) => row.scores[competency] ?? null) })) }, options: { fixedAxis: true, seriesColors: SERIES_SIXTH, labelSize: columnLabelSize(item.rows) }, band: { ids: [2, 10], axis: "y" } };
  if (item.family === "competency-participants") return { data: { categories: item.rows.map((row) => row.name), series: [{ name: item.competency, values: item.rows.map((row) => row.score) }] }, options: { fixedAxis: true, recolorTo: item.recolor || null, labelSize: columnLabelSize(item.rows) }, band: { id: 3, axis: "y" } };
  return null;
}
// Rule 10: category labels stay ≥ 10 pt; they step down from the template size only as far as needed.
function columnLabelSize(rows) {
  const slotPoints = 17.2 * 72 / Math.max(1, rows.length) * 0.9;
  const longestWord = Math.max(1, ...rows.flatMap((row) => String(row.name).split(/\s+/u).map((word) => word.length)));
  return Math.max(1000, Math.min(2000, Math.floor(slotPoints / (longestWord * 0.6)) * 100));
}


// ---------------------------------------------------------------- chart text fit (F21/F22)
const PT = EMU / 72;
const FOOTER_TOP = 10.45 * EMU; // the TREND footer mark starts below this line on every content slide
function wrappedLines(value, size, width) {
  const charWidth = size * 0.5; let lines = 1; let line = 0;
  for (const word of String(value).split(/\s+/u).filter(Boolean)) { const w = word.length * charWidth; const gap = line ? charWidth : 0; if (line && line + gap + w > width) { lines += 1; line = w; } else line += gap + w; while (line > width) { lines += 1; line -= width; } }
  return lines;
}
const axisLabelSize = (chartXml, axis) => Number(chartXml.match(new RegExp(`<c:${axis}>[\\s\\S]*?<a:defRPr\\b[^>]*\\bsz="(\\d+)"`, "u"))?.[1] || 1800);
/** Largest label size (≤ template, ≥ 10 pt) at which every label wraps into its box. */
function fitLabels(labels, { max, width, height, maxLines = 4 }) {
  for (let size = max; size >= 1000; size -= 100) {
    const points = size / 100;
    if (labels.every((label) => { const lines = wrappedLines(label, points, width); return lines <= maxLines && lines * points * 1.2 <= height; })) return size;
  }
  return 1000;
}
const templateLayout = (chartXml) => { const match = chartXml.match(/<c:plotArea><c:layout><c:manualLayout>[\s\S]*?<c:x val="([^"]+)"\/><c:y val="([^"]+)"\/><c:w val="([^"]+)"\/><c:h val="([^"]+)"\/>/u); return match ? { x: Number(match[1]), y: Number(match[2]), w: Number(match[3]), h: Number(match[4]) } : null; };
/**
 * Pinned plot areas (F21): the plot area is chosen first (room for the labels, clear of the footer) and the band is
 * placed from that plot area, so the band lands on the benchmark wherever the renderer draws the axis.
 */
function chartLayout(item, chartXml, slideXml, spec) {
  const options = {}; let slide = null; const { low, high } = bandRange(item);
  if (item.family === "range") {
    const frame = xfrmOf(getShape(slideXml, 22)); const band = visualBox(getShape(slideXml, 2));
    const perPoint = band.cy / (high - low); const top = band.y - (5 - high) * perPoint; // y(5) from the template rectangle
    const left = band.x; const width = band.cx; const slot = width / Math.max(1, spec.data.categories.length) / PT * 0.92;
    const size = fitLabels(spec.data.categories, { max: Math.min(2000, axisLabelSize(chartXml, "catAx")), width: slot, height: 4 * 20 * 1.2, maxLines: 4 });
    const lines = Math.max(...spec.data.categories.map((label) => wrappedLines(label, size / 100, slot)));
    const bottom = Math.min(top + 4 * perPoint, FOOTER_TOP - (lines * size / 100 * 1.25 + 10) * PT);
    options.plotLayout = { x: (left - frame.x) / frame.cx, y: (top - frame.y) / frame.cy, w: width / frame.cx, h: (bottom - top) / frame.cy };
    options.labelSize = size;
    slide = updateShape(slideXml, 2, (shape) => setVisualBox(shape, { ...visualBox(shape), y: top + (5 - high) / 4 * (bottom - top), cy: (high - low) / 4 * (bottom - top) }));
  }
  if (item.family === "participant-mean") {
    const frame = xfrmOf(getShape(slideXml, 2)); const band = visualBox(getShape(slideXml, 3));
    const perPoint = band.cx / (high - low); const right = band.x + (5 - low) * perPoint; // x(5) from the template rectangle
    const top = frame.y + 0.07 * frame.cy; const height = 0.9 * frame.cy; const slot = height / Math.max(1, item.rows.length) / PT;
    const size = Math.max(1000, Math.min(axisLabelSize(chartXml, "catAx"), Math.floor(slot / 2.6) * 100));
    // A name wraps only at spaces: the label column must hold its longest word, and half of the name at most.
    const needed = Math.max(...item.rows.map((row) => { const name = String(row.name); return Math.max(...name.split(/\s+/u).map((word) => word.length), Math.ceil(name.length / 2)); }));
    const labelWidth = (needed * size / 100 * 0.6 + 24) * PT;
    const left = Math.max(right - 4 * perPoint, frame.x + labelWidth);
    options.plotLayout = { x: (left - frame.x) / frame.cx, y: 0.07, w: (Math.min(right, frame.x + frame.cx * 0.99) - left) / frame.cx, h: 0.9 };
    const plotRight = left + options.plotLayout.w * frame.cx; const scale = (plotRight - left) / 4;
    options.labelSize = size; options.dataLabelSize = Math.max(1000, Math.min(2400, Math.round(size * 1.2 / 100) * 100));
    slide = updateShape(slideXml, 3, (shape) => setVisualBox(shape, { ...visualBox(shape), x: left + (low - 1) * scale, cx: (high - low) * scale, y: top, cy: height }));
  }
  if (item.family === "ranking") {
    const frame = xfrmOf(getShape(slideXml, 6)); const layout = templateLayout(chartXml) || { x: 0.02, y: 0.17, w: 0.96, h: 0.81 };
    const plotW = layout.w * frame.cx / PT; const plotH = layout.h * frame.cy / PT; const count = spec.data.series.length;
    const thickness = plotH / (count + 2.19);
    const labels = spec.data.series.map((series) => ({ text: `${series.name}; ${f2(series.values[0])}`, room: Math.max(60, (Number(series.values[0]) - 1) / 4 * plotW) }));
    let size = Math.min(2400, axisLabelSize(chartXml, "dLbls") || 2000);
    while (size > 1000 && !labels.every((label) => wrappedLines(label.text, size / 100, label.room * 0.9) * size / 100 * 1.2 <= thickness * 0.95)) size -= 100;
    options.dataLabelSize = size;
    // F22: when „competency; mean” cannot sit inside the bars at ≥ 10 pt, the labels carry the value only and the
    // competencies move to a legend under the plot (the band shrinks with the plot's height).
    // Renderers wrap a data label at roughly a fifth of the chart width; a label wider than that would wrap over its neighbours.
    const wrapWidth = frame.cx / PT / 5;
    if (!labels.every((label) => wrappedLines(label.text, size / 100, Math.min(label.room * 0.9, wrapWidth)) * size / 100 * 1.2 <= thickness * 0.95)) {
      const legendHeight = 0.14;
      options.valueOnlyLabels = true;
      options.legend = { x: layout.x, y: layout.y + layout.h - legendHeight + 0.02, w: layout.w, h: legendHeight, size: 1400 };
      options.plotLayout = { ...layout, h: layout.h - legendHeight };
      options.dataLabelSize = Math.min(2400, axisLabelSize(chartXml, "dLbls") || 2000);
      const plotTop = frame.y + layout.y * frame.cy; const plotHeight = options.plotLayout.h * frame.cy;
      slide = updateShape(slideXml, 2, (shape) => setVisualBox(shape, { ...visualBox(shape), y: plotTop, cy: plotHeight }));
    }
  }
  if (item.family === "population") {
    const frame = xfrmOf(getShape(slideXml, 6)); const layout = templateLayout(chartXml) || { x: 0.126, y: 0, w: 0.856, h: 0.96 };
    const width = layout.x * frame.cx / PT - 12; const slot = layout.h * frame.cy / PT / Math.max(1, spec.data.categories.length);
    options.labelSize = fitLabels(spec.data.categories, { max: axisLabelSize(chartXml, "catAx"), width, height: slot * 0.9, maxLines: 6 });
  }
  if (item.family === "zone") {
    const frame = xfrmOf(getShape(slideXml, 6)); const layout = templateLayout(chartXml) || { x: 0.04, y: 0.02, w: 0.96, h: 0.78 };
    const plotBottom = frame.y + (layout.y + layout.h) * frame.cy; const room = (FOOTER_TOP - plotBottom) / PT - 8;
    const slot = layout.w * frame.cx / PT / Math.max(1, spec.data.categories.length) * 0.92;
    options.labelSize = fitLabels(spec.data.categories, { max: axisLabelSize(chartXml, "catAx"), width: slot, height: room, maxLines: 4 });
  }
  return { options, slide };
}

// ---------------------------------------------------------------- slide fillers (fill map §1)
const groupSuffix = (item) => item.suffix || "";
const fill = (xml, id, paragraphs, options = {}) => { const scale = groupScale(xml, id); return updateShape(xml, id, (shape) => { const filled = setParagraphs(shape, paragraphs, options); return options.fit === false ? filled : fitText(filled, scale); }); };
// F23: every title of one series (all M5, all A4, all M13, all range/ranking/… slides, all dividers) gets ONE size —
// the smallest one-line fit in the series, ≥ 60 % of the template. Computed in a pre-pass (seriesTitleScales).
const competencyRuns = (shape, competency) => {
  const styles = templateParagraphs(shape)[0]; const base = styles.rPrs[0];
  const baseColor = base.match(/<a:srgbClr val="(\w+)"/u)?.[1];
  const accent = styles.rPrs.find((rPr) => (rPr.match(/<a:srgbClr val="(\w+)"/u)?.[1] || baseColor) !== baseColor) || setRPrColor(base, TITLE_ACCENT);
  return setParagraphs(shape, [[{ text: "Distribuția pe competențe – ", rPr: base }, { text: competency, rPr: accent }]]);
};
const withSuffix = (shape, item) => (groupSuffix(item) ? appendToLastRun(shape, groupSuffix(item)) : shape);
export const TITLE_RULES = {
  "key-findings": { id: 21, build: (shape, item) => competencyRuns(shape, item.competency) },
  "competency-participants": { id: 5, build: (shape, item) => competencyRuns(shape, item.competency) },
  behavior: { id: 3, build: (shape, item) => appendToLastRun(shape, ` ${item.competency}${groupSuffix(item)}`) },
  range: { id: 3, build: withSuffix }, ranking: { id: 4, build: withSuffix }, benchmark: { id: 21, build: withSuffix },
  population: { id: 4, build: withSuffix }, zone: { id: 4, build: withSuffix }, conclusions: { id: 5, build: withSuffix },
  "divider-results": { id: 8, build: withSuffix, box: true }, "divider-behaviors": { id: 8, build: withSuffix, box: true }, "divider-conclusions": { id: 8, build: withSuffix, box: true }
};
function titleFit(templateXml, item, forceScale = null) {
  const rule = TITLE_RULES[item.family]; const scale = groupScale(templateXml, rule.id);
  const built = rule.build(getShape(templateXml, rule.id), item);
  if (rule.box) { const fitted = forceScale === null ? fitTextSized(built, { ...scale, minScale: 0.6, minPoints: 0 }) : { xml: scaleRunSizes(built, forceScale), scale: forceScale }; return { ...fitted, lines: 1, lineHeight: 0 }; }
  return fitTitleOneLine(built, { ...scale, forceScale });
}
export function seriesTitleScales(plan, templateSlides) {
  const scales = {};
  for (const item of plan) if (TITLE_RULES[item.family]) { const { scale } = titleFit(templateSlides.get(item.templateIndex), item); scales[item.family] = Math.min(scales[item.family] ?? 1, scale); }
  return scales;
}
let SERIES_SCALES = {};
/** Writes the series-sized title; returns how many lines it takes and the line height (for compressBelow). */
function applyTitle(xml, item) {
  const rule = TITLE_RULES[item.family];
  const fitted = titleFit(xml, item, SERIES_SCALES[item.family] ?? null);
  return { xml: updateShape(xml, rule.id, () => fitted.xml), lines: fitted.lines, lineHeight: fitted.lineHeight };
}
const titleSuffix = (xml, id, item) => applyTitle(xml, item).xml;
const dividerSuffix = (xml, id, item) => applyTitle(xml, item).xml;
/** Move shapes below a growing title down by `delta`, compressing them into the same bottom edge (tables scale their rows). */
function compressBelow(xml, ids, delta) {
  const boxes = ids.map((id) => visualBox(getShape(xml, id)));
  const top = Math.min(...boxes.map((box) => box.y)); const bottom = Math.max(...boxes.map((box) => box.y + box.cy));
  const k = (bottom - top - delta) / (bottom - top);
  for (const id of ids) xml = updateShape(xml, id, (shape) => {
    const box = visualBox(shape); let next = setVisualBox(shape, { ...box, y: top + delta + (box.y - top) * k, cy: box.cy * k });
    if (/<a:tbl>/u.test(next)) next = next.replace(/(<a:tr\b[^>]*\bh=")(\d+)(")/gu, (whole, head, h, tail) => `${head}${Math.round(Number(h) * k)}${tail}`);
    return next;
  });
  return xml;
}
function competencyTitle(xml, id, competency, suffix = "", item = null) { return applyTitle(xml, item); }
const bandRange = (item) => ({ low: Number(item.low ?? 2.75), high: Number(item.high ?? 3.5) });
function applyBand(xml, spec, item) {
  if (!spec?.band) return xml;
  const { low, high } = bandRange(item);
  for (const id of spec.band.ids || [spec.band.id]) xml = updateShape(xml, id, (shape) => moveBand(shape, { low, high, axis: spec.band.axis }));
  return xml;
}
const benchmarkSentence = (item) => `între ${f2(bandRange(item).low)} și ${f2(bandRange(item).high)} – performanță la nivel mediu`;
let nextNewId = 9001;
const newId = () => nextNewId++;

function fillCover(xml, item) {
  xml = fill(xml, 15, [item.title]);
  xml = updateShape(xml, 23, (shape) => setRunText(shape, 0, item.year));
  if (item.annexMark) xml = updateShape(xml, 12, (shape) => appendToLastRun(shape, " – Anexă"));
  const logo = xfrmOf(getShape(xml, 20));
  const style = setRPrColor(templateParagraphs(getShape(xml, 12))[0].rPrs[0], NAVY);
  return addToTree(xml, fitTitleOneLine(newTextShape({ id: newId(), name: "confidential", x: logo.x, y: logo.y + logo.cy + 0.15 * EMU, cx: logo.cx, cy: 0.6 * EMU, rPr: style, text: "CONFIDENȚIAL", align: "ctr" }), { minScale: 0.5 }).xml);
}
function fillHowToRead(xml, item) {
  xml = updateShape(xml, 93, (shape) => setParagraphs(shape, [[{ text: item.title, run: 0 }]]));
  const body = (shape, paragraphs) => { const rPr = regularRPr(shape); return fitText(setParagraphs(shape, paragraphs.map((text) => [{ text, rPr }]))); };
  const split = Math.ceil(item.paragraphs.length / 2);
  xml = updateShape(xml, 5, (shape) => body(shape, item.paragraphs.slice(0, split)));
  return updateShape(xml, 6, (shape) => body(shape, item.paragraphs.slice(split)));
}
function fillMethodology(xml, item) {
  xml = updateShape(xml, 5, (shape) => { const bold = boldRPr(shape); const regular = regularRPr(shape); return fitText(setParagraphs(shape, item.page.facts.map((fact) => [{ text: fact.number ? `${fact.number} ` : "", rPr: bold }, { text: fact.text, rPr: regular }]))); });
  return updateShape(xml, 6, (shape) => { const regular = regularRPr(shape); return fitText(setParagraphs(shape, item.page.principles.map((text) => [{ text, rPr: regular }]))); });
}
function fillExecutiveSummary(xml, item) {
  const summary = item.summary;
  xml = fill(xml, 5, ["Executive Summary"], { fit: false });
  xml = fill(xml, 8, ["Imaginea de ansamblu"], { fit: false });
  xml = fill(xml, 9, ["Competențe"], { fit: false });
  xml = fill(xml, 10, ["Concluzii comportamentale"], { fit: false });
  const reserved = 1.55 * EMU;
  xml = updateShape(xml, 4, (shape) => { const box = xfrmOf(shape); const filled = setBodyInsets(setParagraphs(shape, [summary.sentence], { style: (index, styles) => withoutBullet(styles[0]) }), { bIns: reserved }); return fitText(filled, { height: box.cy }); });
  xml = fill(xml, 7, summary.competencyLines);
  xml = fill(xml, 11, summary.conclusions ? summary.conclusions.split(/\n+/u) : [""]);
  // „Imaginea de ansamblu” infographic (New): three percentages over a three-band benchmark bar.
  const box = xfrmOf(getShape(xml, 4)); const rPr = regularRPr(getShape(xml, 4));
  const inner = 0.3 * EMU; const width = (box.cx - 2 * inner) / 3; const top = box.y + box.cy - reserved + 0.1 * EMU;
  const bands = [["below", NAVY, `sub ${f2(summary.low)}`], ["in", GREY, `${f2(summary.low)}–${f2(summary.high)}`], ["above", GOLD, `peste ${f2(summary.high)}`]];
  const parts = bands.flatMap(([key, color, label], index) => {
    const x = box.x + inner + index * width;
    return [
      newTextShape({ id: newId(), name: `infographic ${key} share`, x, y: top, cx: width, cy: 0.5 * EMU, rPr: setRPrColor(setRPrAttr(setRPrAttr(rPr, "sz", "2800"), "b", "1"), NAVY), text: `${summary.shares[key]}%`, align: "ctr", anchor: "b", autofit: false }),
      newRectShape({ id: newId(), name: `infographic ${key} band`, x, y: top + 0.58 * EMU, cx: width, cy: 0.3 * EMU, fill: color }),
      newTextShape({ id: newId(), name: `infographic ${key} label`, x, y: top + 0.95 * EMU, cx: width, cy: 0.35 * EMU, rPr: setRPrColor(setRPrAttr(rPr, "sz", "1400"), NAVY), text: label, align: "ctr", autofit: false })
    ];
  });
  return addToTree(xml, parts.join(""));
}
function bandSpans(frameXml, frameBox, rows, bands) {
  const spans = {}; let y = frameBox.y;
  rows.forEach((row, index) => { const h = rowHeight(row); const key = bands[index]; if (!spans[key]) spans[key] = { top: y, bottom: y + h }; else spans[key].bottom = y + h; y += h; });
  return spans;
}
function fillScoreTable(xml, tableId, values, bandsOf, low, high) {
  const frame = getShape(xml, tableId); const box = xfrmOf(frame);
  const templateRows = tableRows(frame);
  const total = templateRows.reduce((sum, row) => sum + rowHeight(row), 0);
  const grey = templateRows.find(isShadedRow) || templateRows[0]; const white = templateRows.find((row) => !isShadedRow(row)) || templateRows[0];
  const baseSize = Number(white.match(/\ssz="(\d+)"/u)?.[1] || 1600);
  const layout = rowLayout(total, values.length, baseSize);
  const bands = values.map((value) => bandsOf(value));
  const rows = values.map((value, index) => setRowId(setRowHeight(setRowCells(bands[index] === "in" ? grey : white, ["", f2(value)], layout.size ? { size: layout.size } : {}), layout.height), 1000000 + index));
  xml = updateShape(xml, tableId, (shape) => replaceTableRows(shape, rows.length ? rows : [setRowCells(white, ["", ""])]));
  return { xml, spans: bandSpans(frame, box, rows, bands) };
}
const bandOf = (low, high) => (value) => (value > high ? "above" : value < low ? "below" : "in");
function fillKeyFindings(xml, item) {
  const { low, high } = bandRange(item);
  const title = competencyTitle(xml, 21, item.competency, "", item); xml = title.xml;
  const shift = title.lines > 1 ? title.lineHeight : 0;
  if (shift) xml = compressBelow(xml, [6, 19, 17, 20, 18], shift);
  const titleBox = xfrmOf(getShape(xml, 21)); const accent = setRPrColor(templateParagraphs(getShape(xml, 21))[0].rPrs.at(-1), TITLE_ACCENT);
  const table = fillScoreTable(xml, 6, item.scores, bandOf(low, high), low, high); xml = table.xml;
  xml = placeBrace(xml, 8, 11, table.spans.above, item.counts.above);
  xml = placeBrace(xml, 14, 16, table.spans.in, item.counts.in);
  xml = placeBrace(xml, 12, 13, table.spans.below, item.counts.below);
  xml = fill(xml, 17, item.strengths.length ? item.strengths : [""]);
  xml = fill(xml, 18, item.development.length ? item.development : [""]);
  const strengthsHeader = xfrmOf(getShape(xml, 19));
  return addToTree(xml, newTextShape({ id: newId(), name: "subtitle", x: strengthsHeader.x, y: titleBox.y + titleBox.cy - 0.18 * EMU + shift, cx: 10.8 * EMU, cy: 0.4 * EMU, rPr: setRPrAttr(accent, "sz", "2000"), text: `medie ${f2(item.mean)} · mediană ${f2(item.median)}`, align: "l", anchor: "ctr" }));
}
/**
 * F33: „{competency}<tab>{mean}” with the mean on a right tab at the box edge. Names are broken into lines (explicit
 * <a:br/>) that end before the mean's column, at the largest size (≤ template, ≥ 10 pt) where the list fits the box.
 */
function competencyMeansBox(shape, rows) {
  const box = xfrmOf(shape); const style = templateParagraphs(shape)[0];
  const margin = Number(style.pPr.match(/\smarL="(\d+)"/u)?.[1] || 0);
  const lineFactor = Number(style.pPr.match(/<a:lnSpc><a:spcPct val="(\d+)"/u)?.[1] || 100000) / 100000;
  const tabPos = box.cx - 2 * 91440 - margin - 0.05 * EMU;
  const base = Number(style.rPrs[0].match(/\ssz="(\d+)"/u)?.[1] || 2000);
  const wrap = (value, width, size) => { const lines = []; let line = ""; for (const word of value.split(/\s+/u)) { const next = line ? `${line} ${word}` : word; if (line && next.length * size * 0.5 > width) { lines.push(line); line = word; } else line = next; } lines.push(line); return lines; };
  let chosen = null;
  for (let size = base; size >= 1000; size -= 100) {
    const points = size / 100; const allowed = tabPos / 12700 - 4 * points * 0.5 - points * 1.2;
    const wrapped = rows.map((row) => wrap(row.competency, allowed, points));
    const height = wrapped.reduce((sum, lines) => sum + lines.length, 0) * points * 1.2 * lineFactor * 12700 + 2 * 45720;
    chosen = { size, wrapped };
    if (height <= box.cy) break;
  }
  const rPr = setRPrAttr(style.rPrs[0], "sz", String(chosen.size));
  const tab = `<a:tabLst><a:tab pos="${Math.round(tabPos)}" algn="r"/></a:tabLst>`;
  let pPr = style.pPr || "<a:pPr/>"; pPr = pPr.replace(/<a:tabLst\/>|<a:tabLst>[\s\S]*?<\/a:tabLst>/u, "");
  pPr = pPr.endsWith("/>") ? pPr.replace(/\/>$/u, `>${tab}</a:pPr>`) : (() => { const at = pPr.search(/<a:(?:defRPr|extLst)\b/u); return at >= 0 ? `${pPr.slice(0, at)}${tab}${pPr.slice(at)}` : pPr.replace(/<\/a:pPr>$/u, `${tab}</a:pPr>`); })();
  const brRPr = rPr.replace(/^<a:rPr/u, "<a:rPr");
  const paragraphs = rows.map((row, index) => { const lines = chosen.wrapped[index]; const body = lines.map((line, lineIndex) => lineIndex < lines.length - 1 ? `<a:r>${rPr}<a:t>${xmlEscape(line)}</a:t></a:r><a:br>${brRPr}</a:br>` : `<a:r>${rPr}<a:t>${xmlEscape(`${line}\t${f2(row.mean)}`)}</a:t></a:r>`).join(""); return `<a:p>${pPr}${body}</a:p>`; }).join("");
  const open = shape.search(/<p:txBody>/u); const close = shape.search(/<\/p:txBody>/u); const bodyXml = shape.slice(open, close);
  const first = bodyXml.search(/<a:p[\s>]/u);
  return `${shape.slice(0, open)}${bodyXml.slice(0, first)}${paragraphs}${shape.slice(close)}`;
}
function fillBenchmark(xml, item) {
  const { low, high } = bandRange(item); const table = item.table;
  xml = titleSuffix(xml, 21, item);
  const filled = fillScoreTable(xml, 2, table.rows.map((row) => row.value), bandOf(low, high), low, high); xml = filled.xml;
  xml = placeBrace(xml, 8, 11, filled.spans.above, `${table.shares.above}%`);
  xml = placeBrace(xml, 14, 16, filled.spans.in, `${table.shares.in}%`);
  xml = placeBrace(xml, 12, 13, filled.spans.below, `${table.shares.below}%`);
  xml = updateShape(xml, 4, (shape) => competencyMeansBox(shape, table.competencyMeans));
  xml = fill(xml, 19, [`Rezultate raportate la benchmark (${f2(low)}-${f2(high)})`], { fit: false });
  return fill(xml, 17, [`${table.shares.above}% au obținut o medie generală peste ${f2(high)}`, `${table.shares.in}% au obținut o medie generală între ${f2(low)} – ${f2(high)}`, `${table.shares.below}% au obținut o medie generală mai mică de ${f2(low)}`]);
}
function fillPopulation(xml, item) {
  const { low, high } = bandRange(item);
  xml = titleSuffix(xml, 4, item);
  xml = fill(xml, 14, [`Participanți care au obținut o medie între ${f2(low)} – ${f2(high)}`]);
  xml = fill(xml, 16, [`Participanți care au obținut o medie peste ${f2(high)}`]);
  xml = fill(xml, 19, [`Participanți care au obținut o medie mai mică de ${f2(low)}`]);
  return updateShape(xml, 26, (shape) => setRunText(shape, 3, benchmarkSentence(item)));
}
function fillBehavior(xml, item) {
  xml = applyTitle(xml, item).xml;
  const frame = getShape(xml, 7); const rows = tableRows(frame); const header = rows[0]; const bodyTemplate = rows[1];
  const bodyHeight = rows.slice(1).reduce((sum, row) => sum + rowHeight(row), 0);
  const count = Math.max(1, item.key.length, item.development.length);
  const body = Array.from({ length: count }, (_, index) => setRowId(setRowHeight(setRowCells(bodyTemplate, [item.key[index] || "", item.development[index] || ""], cellSize(item, count)), bodyHeight / count), 2000000 + index));
  return updateShape(xml, 7, (shape) => replaceTableRows(shape, [header, ...body]));
}
function cellSize(item, count) {
  // Rule 5/7 for the behaviour table: the longest cell must fit its row (≈ 8.5 in wide); step down to 9 pt minimum.
  const rowPoints = 7.4 * 72 / count; const longest = Math.max(1, ...[...item.key, ...item.development].map((value) => value.length));
  for (let size = 2000; size >= 1000; size -= 100) { const charsPerLine = 8.4 * 72 / (size / 100 * 0.53); const lines = Math.ceil(longest / charsPerLine); if (lines * size / 100 * 1.25 + 8 <= rowPoints) return { size }; }
  return { size: 1000 };
}
function fillConclusions(xml, item) {
  xml = titleSuffix(xml, 5, item);
  const lines = (value) => (value ? value.split(/\n+/u) : [""]);
  xml = fill(xml, 4, lines(item.strengths));
  xml = fill(xml, 7, lines(item.development));
  return fill(xml, 11, lines(item.interventions));
}
function fillSlide(xml, item) {
  switch (item.family) {
    case "cover": return fillCover(xml, item);
    case "how-to-read": return fillHowToRead(xml, item);
    case "methodology": return fillMethodology(xml, item);
    case "executive-summary": return fillExecutiveSummary(xml, item);
    case "key-findings": return fillKeyFindings(xml, item);
    case "divider-results": case "divider-behaviors": case "divider-conclusions": return dividerSuffix(xml, 8, item);
    case "range": return updateShape(titleSuffix(xml, 3, item), 26, (shape) => setRunText(shape, 3, benchmarkSentence(item)));
    case "ranking": case "zone": return titleSuffix(xml, 4, item);
    case "benchmark": return fillBenchmark(xml, item);
    case "population": return fillPopulation(xml, item);
    case "behavior": return fillBehavior(xml, item);
    case "conclusions": return fillConclusions(xml, item);
    case "appendix-divider": return fill(fill(xml, 8, [item.heading], { fit: false }), 9, [item.subheading], { fit: false });
    case "participant-mean": case "participant-comparison": return xml;
    case "competency-participants": { const title = competencyTitle(xml, 5, item.competency, "", item); return title.lines > 1 ? compressBelow(title.xml, [2, 3], title.lineHeight) : title.xml; }
    case "close": return xml;
    default: throw new Error(`No fill rule for plan family ${item.family}.`);
  }
}
// F23: sibling text boxes on one slide share ONE size — the smallest of their fits.
// Groups of boxes that sit side by side in one column/row of the template (the big right-hand box of t21 is its own group).
const SIBLINGS = { "how-to-read": [[5, 6]], methodology: [[5, 6]], "executive-summary": [[4, 7]], "key-findings": [[17, 18]], benchmark: [[4, 17]], conclusions: [[4, 7]] };
const largestSize = (shape) => Math.max(0, ...[...shape.matchAll(/<a:rPr\b[^>]*\ssz="(\d+)"[^>]*>(?:(?!<\/a:r>)[\s\S])*?<a:t>[^<]/gu)].map((match) => Number(match[1])));
function equaliseSiblings(xml, ids) {
  const sizes = ids.map((id) => largestSize(getShape(xml, id))).filter(Boolean);
  if (sizes.length < 2) return xml; const target = Math.min(...sizes);
  for (const id of ids) xml = updateShape(xml, id, (shape) => { const size = largestSize(shape); return size > target ? scaleRunSizes(shape, target / size) : shape; });
  return xml;
}
export function renderSlide(item, sourceXml) {
  let xml = fillSlide(sourceXml, item);
  for (const group of SIBLINGS[item.family] || []) xml = equaliseSiblings(xml, group);
  xml = applyBand(xml, chartSpec(item), item);
  return xml;
}

// ---------------------------------------------------------------- package machinery
function relationshipTargets(xml) { return [...xml.matchAll(/<Relationship\b([^>]*)\/>/gu)].map((match) => Object.fromEntries([...match[1].matchAll(/(Id|Type|Target|TargetMode)="([^"]*)"/gu)].map((item) => [item[1], item[2]]))); }
function relationshipTarget(source, target) {
  if (/^https?:|^mailto:/u.test(target)) return null;
  const base = source === "_rels/.rels" ? "" : source.split("/").slice(0, -1).join("/");
  const parts = `${base}/${target}`.replace(/\/+/gu, "/").split("/");
  const output = [];
  for (const part of parts) { if (part === "..") output.pop(); else if (part && part !== ".") output.push(part); }
  return output.join("/");
}
function normalizeSlideIdentity(xml, targetIndex) {
  const creationId = String(1000000000 + targetIndex);
  let shapeNumber = 0;
  return xml.replace(/<p14:creationId\b([^>]*?)val="[^"]+"([^>]*)\/>/u, `<p14:creationId$1val="${creationId}"$2/>`).replace(/<a16:creationId\b([^>]*?)id="[^"]+"([^>]*)\/>/gu, () => {
    const slide = targetIndex.toString(16).padStart(8, "0"); const shape = shapeNumber.toString(16).padStart(4, "0"); const tail = `${slide}${shapeNumber.toString(16).padStart(8, "0")}`.slice(-12); shapeNumber += 1;
    return `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{${slide}-${shape}-46C6-A399-${tail}}"/>`;
  });
}
function normalizeChartIdentity(xml, chartNumber) {
  let seriesNumber = 0;
  return xml.replace(/<c16:uniqueId\b[^>]*val="[^"]+"[^>]*\/>/gu, () => {
    const chart = chartNumber.toString(16).padStart(8, "0"); const series = seriesNumber.toString(16).padStart(4, "0"); const tail = `${chart}${seriesNumber.toString(16).padStart(8, "0")}`.slice(-12);
    seriesNumber += 1;
    return `<c16:uniqueId val="{${chart}-${series}-4000-8000-${tail}}"/>`;
  });
}
async function packageGraph(zip) {
  const reachable = new Set(); const queue = ["ppt/presentation.xml"]; const dangling = [];
  if (zip.file("_rels/.rels")) for (const rel of relationshipTargets(await zip.file("_rels/.rels").async("string"))) { const target = relationshipTarget("_rels/.rels", rel.Target); if (target) queue.push(target); }
  while (queue.length) {
    const part = queue.shift(); if (reachable.has(part)) continue;
    if (!zip.file(part)) { dangling.push(part); continue; }
    reachable.add(part);
    const relPart = `${part.split("/").slice(0, -1).join("/")}/_rels/${part.split("/").at(-1)}.rels`;
    if (zip.file(relPart)) { reachable.add(relPart); for (const rel of relationshipTargets(await zip.file(relPart).async("string"))) { if (rel.TargetMode === "External") continue; const target = relationshipTarget(part, rel.Target); if (target) queue.push(target); } }
  }
  return { reachable, dangling };
}
export async function selfCheckPptx(zip) {
  const errors = []; const names = Object.keys(zip.files).filter((name) => !name.endsWith("/"));
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  if (!contentTypes) errors.push("missing [Content_Types].xml");
  const overrides = contentTypes ? [...contentTypes.matchAll(/<Override\b[^>]*PartName="([^"]+)"/gu)].map((match) => match[1].replace(/^\//u, "")) : [];
  for (const part of overrides) if (!zip.file(part)) errors.push(`content type points to missing ${part}`);
  const graph = await packageGraph(zip); errors.push(...graph.dangling.map((part) => `relationship points to missing ${part}`));
  for (const part of names.filter((name) => /^(ppt\/(slides|charts|embeddings)\/)/u.test(name))) if (!graph.reachable.has(part)) errors.push(`orphan package part ${part}`);
  const creationIds = new Set(); const shapeCreationIds = new Set();
  for (const entry of names) {
    if (/ppt\/fonts\//u.test(entry)) errors.push("ppt/fonts part present");
    if (!entry.endsWith(".xml")) continue;
    const xml = await zip.file(entry).async("string");
    if (/<(?:p:|a:)[^>]*embeddedFontLst|<p:notesMasterIdLst/iu.test(xml)) errors.push(`forbidden IdLst entry in ${entry}`);
    if (!/^ppt\/slides\/slide\d+\.xml$/u.test(entry)) continue;
    const shapeIds = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="([^"]+)"/gu)].map((match) => match[1]); if (shapeIds.length !== new Set(shapeIds).size) errors.push(`duplicate shape id in ${entry}`);
    const rowIds = [...xml.matchAll(/<a16:rowId\b[^>]*val="([^"]+)"/gu)].map((match) => match[1]); if (rowIds.length !== new Set(rowIds).size) errors.push(`duplicate table row id in ${entry}`);
    for (const id of xml.matchAll(/<p14:creationId\b[^>]*val="([^"]+)"/gu)) if (creationIds.has(id[1])) errors.push(`duplicate slide creation id ${id[1]}`); else creationIds.add(id[1]);
    for (const id of xml.matchAll(/<a16:creationId\b[^>]*id="([^"]+)"/gu)) if (shapeCreationIds.has(id[1])) errors.push(`duplicate shape creation id ${id[1]}`); else shapeCreationIds.add(id[1]);
    for (const table of xml.matchAll(/<a:tbl>[\s\S]*?<\/a:tbl>/gu)) { const columns = (table[0].match(/<a:gridCol\b/gu) || []).length; for (const row of table[0].match(/<a:tr\b[\s\S]*?<\/a:tr>/gu) || []) if ((row.match(/<a:tc\b/gu) || []).length !== columns) errors.push(`table row cell count differs from grid in ${entry}`); }
  }
  if (errors.length) throw new Error(`PPTX self-check failed: ${errors.join("; ")}`);
  return { ok: true, parts: names.length, reachable: graph.reachable.size };
}

async function cloneSlide(zip, sourceIndex, targetIndex, chartOffset, embeddingOffset, source) {
  const sourceSlide = source.slides.get(sourceIndex); const sourceRels = source.slides.get(`${sourceIndex}.rels`); let nextChart = chartOffset; let nextEmbedding = embeddingOffset; let rels = sourceRels;
  for (const match of sourceRels.matchAll(/<Relationship\b[^>]*Target="\.\.\/charts\/chart(\d+)\.xml"[^>]*\/>/gu)) {
    const originalChart = Number(match[1]); nextChart += 1;
    const chartXmlSource = source.charts.get(originalChart);
    if (!chartXmlSource) throw new Error(`Template chart${originalChart}.xml is missing while cloning slide${sourceIndex}.xml.`);
    let chartRels = source.chartRels.get(originalChart) || `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    for (const related of relationshipTargets(chartRels).filter((entry) => /\/(?:chartStyle|chartColorStyle|themeOverride)$/u.test(entry.Type))) {
      const originalPart = relationshipTarget(`ppt/charts/chart${originalChart}.xml`, related.Target); const sourceBytes = originalPart ? source.parts.get(originalPart) : null;
      if (!originalPart || !sourceBytes) throw new Error(`Template chart support part ${originalPart || related.Target} is missing while cloning chart${originalChart}.xml.`);
      const directory = originalPart.split("/").slice(0, -1).join("/"); const filename = originalPart.split("/").at(-1).replace(/\.xml$/u, `-chart${nextChart}.xml`); const newPart = `${directory}/${filename}`; const newTarget = directory === "ppt/theme" ? `../theme/${filename}` : filename;
      zip.file(newPart, sourceBytes); chartRels = chartRels.replace(`Target="${related.Target}"`, `Target="${newTarget}"`);
    }
    for (const embed of chartRels.matchAll(/Target="\.\.\/embeddings\/([^"]+\.xlsx)"/gu)) { const newEmbedding = `Microsoft_Excel_Worksheet${nextEmbedding ? nextEmbedding : ""}.xlsx`; nextEmbedding += 1; chartRels = chartRels.replaceAll(`../embeddings/${embed[1]}`, `../embeddings/${newEmbedding}`); zip.file(`ppt/embeddings/${newEmbedding}`, source.embeddings.get(`ppt/embeddings/${embed[1]}`)); }
    zip.file(`ppt/charts/chart${nextChart}.xml`, normalizeChartIdentity(chartXmlSource, nextChart)); zip.file(`ppt/charts/_rels/chart${nextChart}.xml.rels`, chartRels); rels = rels.replace(`../charts/chart${originalChart}.xml`, `../charts/chart${nextChart}.xml`);
  }
  zip.file(slideName(targetIndex), normalizeSlideIdentity(sourceSlide, targetIndex)); zip.file(slideRelsName(targetIndex), rels); return { chartOffset: nextChart, embeddingOffset: nextEmbedding };
}
function workbookBytes(data) {
  const XLSX = window.XLSX; if (!XLSX) throw new Error("Lipsește biblioteca locală SheetJS pentru datele graficelor.");
  const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(chartWorkbookRows(data)), "Sheet1");
  return XLSX.write(book, { type: "array", bookType: "xlsx", compression: true });
}
async function updateCharts(zip, plan) {
  for (const [index, item] of plan.entries()) {
    const rels = await zip.file(slideRelsName(index + 1)).async("string");
    const charts = relationshipTargets(rels).filter((entry) => entry.Type === `${relsType}chart`);
    const spec = chartSpec(item);
    if (charts.length && !spec) throw new Error(`Slide ${index + 1} (${item.family}) has a chart but no chart data.`);
    for (const rel of charts) {
      const chartPath = relationshipTarget(slideName(index + 1), rel.Target);
      const options = { ...spec.options };
      const chartSource = await zip.file(chartPath).async("string");
      const slideXml = await zip.file(slideName(index + 1)).async("string");
      const layoutFix = chartLayout(item, chartSource, slideXml, spec);
      Object.assign(options, layoutFix.options);
      if (layoutFix.slide) zip.file(slideName(index + 1), layoutFix.slide);
      zip.file(chartPath, normalizeChartIdentity(fillChartXml(chartSource, spec.data, options), Number(chartPath.match(/chart(\d+)\.xml$/u)[1])));
      const chartRelsPath = `${chartPath.split("/").slice(0, -1).join("/")}/_rels/${chartPath.split("/").at(-1)}.rels`;
      const workbookRel = relationshipTargets(await zip.file(chartRelsPath).async("string")).find((entry) => entry.Type === `${relsType}package`);
      if (!workbookRel) throw new Error(`Chart ${chartPath} has no embedded workbook.`);
      zip.file(relationshipTarget(chartPath, workbookRel.Target), workbookBytes(spec.data));
    }
  }
}
function updateContentTypes(xml, zip) {
  const updated = xml.replace(/<Override\b[^>]*PartName="([^"]+)"[^>]*\/>/gu, (whole, partName) => zip.file(partName.replace(/^\//u, "")) ? whole : "").replace(/<Override\b[^>]*PartName="\/ppt\/(?:slides\/slide\d+\.xml|charts\/chart\d+\.xml)"[^>]*\/>/gu, ""); const overrides = [];
  for (const name of Object.keys(zip.files)) { if (/^ppt\/slides\/slide\d+\.xml$/u.test(name)) overrides.push(`<Override PartName="/${name}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`); if (/^ppt\/charts\/chart\d+\.xml$/u.test(name)) overrides.push(`<Override PartName="/${name}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`); }
  return updated.replace("</Types>", `${overrides.join("")}</Types>`);
}
async function prunePackage(zip) {
  const graph = await packageGraph(zip); for (const name of Object.keys(zip.files)) if (!name.endsWith("/") && name !== "[Content_Types].xml" && /^(ppt\/(slides|charts|embeddings)\/)/u.test(name) && !graph.reachable.has(name)) zip.remove(name);
  for (const name of Object.keys(zip.files)) if (/^ppt\/(?:charts|theme)\/(?:style|colors|themeOverride)\d*(?:-chart\d+)?\.xml$/u.test(name) && !graph.reachable.has(name)) zip.remove(name);
  zip.file("[Content_Types].xml", updateContentTypes(await zip.file("[Content_Types].xml").async("string"), zip));
}
// Sections are preserved: each generated slide joins its template slide's section; consecutive slides of one
// section form one run, and a repeated run gets its own section id.
async function templateSections(zip) {
  const presentation = await zip.file("ppt/presentation.xml").async("string");
  const rels = relationshipTargets(await zip.file("ppt/_rels/presentation.xml.rels").async("string"));
  const slideOfId = new Map([...presentation.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"[^>]*r:id="([^"]+)"/gu)].map((match) => [match[1], Number(rels.find((rel) => rel.Id === match[2])?.Target.match(/slide(\d+)\.xml/u)?.[1])]));
  const sections = new Map();
  for (const section of presentation.matchAll(/<p14:section\b[^>]*\bname="([^"]*)"[^>]*>([\s\S]*?)<\/p14:section>/gu)) for (const id of section[2].matchAll(/<p14:sldId id="(\d+)"/gu)) sections.set(slideOfId.get(id[1]), section[1]);
  return sections;
}
function rebuildSections(presentation, plan, sectionOfSlide) {
  if (!/<p14:sectionLst\b/u.test(presentation)) return presentation;
  const runs = [];
  plan.forEach((item, index) => { const name = sectionOfSlide.get(item.templateIndex) || "Raport"; if (runs.at(-1)?.name === name) runs.at(-1).ids.push(256 + index); else runs.push({ name, ids: [256 + index] }); });
  const sections = runs.map((run, index) => `<p14:section name="${xmlEscape(run.name)}" id="{${String(index + 1).padStart(8, "0")}-6A1E-4C2B-9F00-${String(index + 1).padStart(12, "0")}}"><p14:sldIdLst>${run.ids.map((id) => `<p14:sldId id="${id}"/>`).join("")}</p14:sldIdLst></p14:section>`).join("");
  return presentation.replace(/(<p14:sectionLst\b[^>]*>)[\s\S]*?(<\/p14:sectionLst>)/u, `$1${sections}$2`);
}
async function rebuildSlides(zip, plan) {
  const source = { slides: new Map(), charts: new Map(), chartRels: new Map(), embeddings: new Map(), parts: new Map(), sectionOfSlide: await templateSections(zip) };
  for (const index of allSlideNumbers(zip)) { source.slides.set(index, await zip.file(slideName(index)).async("string")); source.slides.set(`${index}.rels`, await zip.file(slideRelsName(index)).async("string")); zip.remove(slideName(index)); zip.remove(slideRelsName(index)); }
  for (const index of allChartNumbers(zip)) { source.charts.set(index, await zip.file(`ppt/charts/chart${index}.xml`).async("string")); const relsName = `ppt/charts/_rels/chart${index}.xml.rels`; if (zip.file(relsName)) source.chartRels.set(index, await zip.file(relsName).async("string")); }
  for (const name of Object.keys(zip.files).filter((entry) => /^ppt\/embeddings\/.*\.xlsx$/u.test(entry))) source.embeddings.set(name, await zip.file(name).async("uint8array"));
  for (const name of Object.keys(zip.files).filter((entry) => /^ppt\/(?:charts|theme)\/.*\.xml$/u.test(entry))) source.parts.set(name, await zip.file(name).async("uint8array"));
  let chartOffset = 0; let embeddingOffset = 0; nextNewId = 9001;
  SERIES_SCALES = seriesTitleScales(plan, source.slides);
  for (const [index, item] of plan.entries()) {
    if (!source.slides.has(item.templateIndex)) throw new Error(`Plan item ${item.family} names template slide ${item.templateIndex}, which the cleaned asset does not have.`);
    const clone = await cloneSlide(zip, item.templateIndex, index + 1, chartOffset, embeddingOffset, source); chartOffset = clone.chartOffset; embeddingOffset = clone.embeddingOffset;
    zip.file(slideName(index + 1), renderSlide(item, await zip.file(slideName(index + 1)).async("string")));
  }
  const presentation = await zip.file("ppt/presentation.xml").async("string"); const relsPath = "ppt/_rels/presentation.xml.rels"; let rels = await zip.file(relsPath).async("string");
  rels = rels.replace(/<Relationship\b[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide"[^>]*\/>/gu, "");
  const nextRel = Math.max(0, ...[...rels.matchAll(/Id="rId(\d+)"/gu)].map((match) => Number(match[1]))) + 1;
  rels = rels.replace("</Relationships>", `${plan.map((_, index) => `<Relationship Id="rId${nextRel + index}" Type="${relsType}slide" Target="slides/slide${index + 1}.xml"/>`).join("")}</Relationships>`);
  zip.file(relsPath, rels);
  zip.file("ppt/presentation.xml", rebuildSections(presentation.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/u, `<p:sldIdLst>${plan.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${nextRel + index}"/>`).join("")}</p:sldIdLst>`), plan, source.sectionOfSlide));
  await updateCharts(zip, plan); await prunePackage(zip);
  // F29: extended properties describe the generated deck, not the template.
  if (zip.file("docProps/app.xml")) zip.file("docProps/app.xml", (await zip.file("docProps/app.xml").async("string")).replace(/<Slides>\d+<\/Slides>/u, `<Slides>${plan.length}</Slides>`).replace(/<Notes>\d+<\/Notes>|<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>|<HeadingPairs>[\s\S]*?<\/HeadingPairs>/gu, ""));
}
export async function generateTrendPptx(payload, { scope = "whole" } = {}) {
  if (!window.JSZip) throw new Error("Lipsește biblioteca locală pentru pachetul PPTX.");
  const zip = await window.JSZip.loadAsync(await templateBytes());
  const plan = reportPlan(payload, { scope });
  await rebuildSlides(zip, plan); await selfCheckPptx(zip);
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", compression: "DEFLATE" });
}

// ---------------------------------------------------------------- BHB bundle (fill map §2)
// Each item is drawn from the same data the deck writes; its page/image is cropped to its own bounds + 24 px.
export function bundleItems(payload) {
  return reportPlan(payload, { scope: "whole" }).filter((item) => BUNDLE_FAMILIES.has(item.family)).map((item, index) => ({ id: String(index + 1).padStart(3, "0"), item, ...bundleContent(item) }));
}
const chartValueText = (value) => { const rounded = roundChartValue(value); return rounded === null ? "" : String(rounded); };
function bundleContent(item) {
  const spec = chartSpec(item);
  if (spec) {
    const percent = item.family === "population";
    const series = spec.data.series.map((entry) => ({ name: entry.name, values: entry.values.map(roundChartValue) }));
    return { kind: "chart", categories: spec.data.categories, series, scale: percent ? [0, 1] : [1, 5], percent, sourceValues: [...spec.data.categories, ...series.map((entry) => entry.name), ...series.flatMap((entry) => entry.values.map(chartValueText).filter(Boolean))] };
  }
  if (item.family === "key-findings") { const rows = [["Scoruri", item.scores.map(f2).join("  ")], ["Abilități cheie – Puncte forte", item.strengths.join("\n")], ["Arii de dezvoltare", item.development.join("\n")]]; return { kind: "table", subtitle: `medie ${f2(item.mean)} · mediană ${f2(item.median)}`, rows, sourceValues: [item.competency, f2(item.mean), f2(item.median), ...item.scores.map(f2), ...item.strengths, ...item.development] }; }
  if (item.family === "benchmark") { const table = item.table; const rows = table.competencyMeans.map((row) => [row.competency, f2(row.mean)]); rows.push([`Peste ${f2(table.high)}`, `${table.shares.above}%`], [`Între ${f2(table.low)} – ${f2(table.high)}`, `${table.shares.in}%`], [`Sub ${f2(table.low)}`, `${table.shares.below}%`], ["Medii individuale", table.rows.map((row) => f2(row.value)).join("  ")]); return { kind: "table", rows, sourceValues: [...table.competencyMeans.flatMap((row) => [row.competency, f2(row.mean)]), `${table.shares.above}%`, `${table.shares.in}%`, `${table.shares.below}%`, ...table.rows.map((row) => f2(row.value))] }; }
  if (item.family === "behavior") { const count = Math.max(item.key.length, item.development.length); const rows = [["Abilități cheie", "Abilități de dezvoltat"], ...Array.from({ length: count }, (_, index) => [item.key[index] || "", item.development[index] || ""])]; return { kind: "table", header: true, rows, sourceValues: [...item.key, ...item.development] }; }
  return { kind: "table", rows: [], sourceValues: [] };
}

async function fontPath() { if (!window.opentype) throw new Error("Lipsește opentype.js vendorizat."); const bytes = await fontBytes(); const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength); return window.opentype.parse(buffer); }
function wrapText(font, value, size, maxWidth) {
  const output = [];
  for (const paragraph of String(value || "").split("\n")) {
    const words = paragraph.split(/\s+/u).filter(Boolean);
    if (!words.length) { output.push(""); continue; }
    let line = "";
    for (const word of words) { const candidate = line ? `${line} ${word}` : word; if (!line || font.getAdvanceWidth(candidate, size) <= maxWidth) { line = candidate; continue; } output.push(line); line = word; }
    if (line) output.push(line);
  }
  return output;
}
const BHB = { navy: "#0A375B", aqua: "#09BAD2", ink: "#231F20", muted: "#64748B", rule: "#D9DEE5", band: "#E6EBF0" };
const SERIES_COLORS = ["#0A375B", "#09BAD2", "#27808F", "#A98138", "#64748B", "#FF9D75", "#4A8C61", "#8B6F47"];
/** Draws one item as primitive shapes + text runs (later outlined), returning its own bounds. */
function layoutItem(entry, font) {
  const shapes = []; const texts = []; const margin = 24; const width = 1400; let y = margin;
  const textLines = (value, x, size, maxWidth, color = BHB.ink, lineHeight = size * 1.3) => { const lines = wrapText(font, value, size, maxWidth); lines.forEach((line, index) => texts.push({ line, x, y: y + size + index * lineHeight, size, color })); return lines.length * lineHeight; };
  // Caption strip (outside the item proper): title, scope and benchmark.
  const captionStart = texts.length; const captionShapes = shapes.length;
  y += textLines(entry.item.title, margin, 30, width - 2 * margin, BHB.navy) + 6;
  y += textLines(`${entry.item.groupLabel ? `Grup: ${entry.item.groupLabel}` : "Întregul proiect"}${entry.item.low !== undefined ? ` · benchmark ${f2(entry.item.low)}–${f2(entry.item.high)}` : ""}${entry.subtitle ? ` · ${entry.subtitle}` : ""}`, margin, 18, width - 2 * margin, BHB.muted) + 16;
  shapes.push({ type: "rect", x: margin, y, w: width - 2 * margin, h: 2, fill: BHB.aqua }); y += 18;
  texts.slice(captionStart).forEach((entryText) => { entryText.caption = true; }); shapes.slice(captionShapes).forEach((shape) => { shape.caption = true; });
  const itemTop = y - margin; // the item proper starts here; the caption strip above is for the PDF page only
  if (entry.kind === "chart") {
    const labelWidth = 380; const plotX = margin + labelWidth + 16; const plotW = width - plotX - margin - 70; const [min, max] = entry.scale;
    const scaleX = (value) => plotX + (Math.max(min, Math.min(max, value)) - min) / (max - min) * plotW;
    if (entry.item.low !== undefined && !entry.percent) shapes.push({ type: "band", x: scaleX(entry.item.low), w: scaleX(entry.item.high) - scaleX(entry.item.low), top: y });
    const bandIndex = shapes.length - 1;
    entry.categories.forEach((category, categoryIndex) => {
      const blockTop = y; const labelHeight = textLines(category, margin, 18, labelWidth, BHB.ink, 22);
      if (entry.percent) {
        let x = plotX; const barY = blockTop + 2;
        entry.series.forEach((series, seriesIndex) => { const value = series.values[categoryIndex] || 0; const w = value * plotW; shapes.push({ type: "rect", x, y: barY, w: Math.max(0, w), h: 26, fill: [BHB.navy, "#A5A5A5", "#DBA900"][seriesIndex] }); if (value >= 0.08) texts.push({ line: `${Math.round(value * 100)}%`, x: x + 6, y: barY + 19, size: 15, color: seriesIndex === 1 ? BHB.ink : "#FFFFFF" }); x += w; });
        y = blockTop + Math.max(labelHeight, 32) + 10; return;
      }
      let rowY = blockTop;
      entry.series.forEach((series, seriesIndex) => {
        const value = series.values[categoryIndex]; if (value === null || value === undefined) { rowY += 26; return; }
        shapes.push({ type: "rect", x: plotX, y: rowY + 3, w: Math.max(2, scaleX(value) - plotX), h: 18, fill: entry.series.length > 1 ? SERIES_COLORS[seriesIndex % SERIES_COLORS.length] : BHB.navy });
        texts.push({ line: f2(value), x: scaleX(value) + 8, y: rowY + 18, size: 15, color: BHB.ink });
        if (entry.series.length > 1 && entry.categories.length > 1) texts.push({ line: series.name, x: plotX + 6, y: rowY + 17, size: 12, color: "#FFFFFF", clip: Math.max(0, scaleX(value) - plotX - 12) });
        rowY += 26;
      });
      y = Math.max(rowY, blockTop + labelHeight) + 12;
    });
    if (bandIndex >= 0 && shapes[bandIndex]?.type === "band") { shapes[bandIndex].h = y - shapes[bandIndex].top; shapes[bandIndex].y = shapes[bandIndex].top; }
    const ticks = entry.percent ? [0, 0.25, 0.5, 0.75, 1] : [1, 2, 3, 4, 5];
    ticks.forEach((tick) => texts.push({ line: entry.percent ? `${tick * 100}%` : f2(tick), x: scaleX(tick) - 16, y: y + 18, size: 13, color: BHB.muted }));
    y += 28;
    if (entry.series.length > 1 && entry.categories.length === 1) entry.series.forEach((series, index) => { shapes.push({ type: "rect", x: margin, y: y + 4, w: 14, h: 14, fill: SERIES_COLORS[index % SERIES_COLORS.length] }); y += textLines(`${series.name}`, margin + 22, 15, width - 2 * margin - 22, BHB.ink, 20) + 2; });
    if (entry.series.length > 1 && entry.categories.length > 1 && !entry.percent) { entry.series.forEach((series, index) => { shapes.push({ type: "rect", x: margin, y: y + 4, w: 14, h: 14, fill: SERIES_COLORS[index % SERIES_COLORS.length] }); y += textLines(series.name, margin + 22, 15, width - 2 * margin - 22, BHB.ink, 20) + 2; }); }
    if (entry.percent) { ["Sub benchmark", "În benchmark", "Peste benchmark"].forEach((label, index) => { shapes.push({ type: "rect", x: margin + index * 300, y: y + 4, w: 14, h: 14, fill: [BHB.navy, "#A5A5A5", "#DBA900"][index] }); texts.push({ line: label, x: margin + index * 300 + 22, y: y + 17, size: 15, color: BHB.ink }); }); y += 24; }
  } else {
    const columns = Math.max(1, ...entry.rows.map((row) => row.length)); const columnWidth = (width - 2 * margin) / columns;
    entry.rows.forEach((row, rowIndex) => {
      const top = y; let height = 0;
      row.forEach((cell, cellIndex) => { const saved = y; y = top + 8; const used = textLines(cell, margin + cellIndex * columnWidth + 10, entry.header && rowIndex === 0 ? 18 : 16, columnWidth - 20, entry.header && rowIndex === 0 ? "#FFFFFF" : BHB.ink, 22); height = Math.max(height, used + 16); y = saved; });
      if (entry.header && rowIndex === 0) shapes.splice(shapes.length, 0, { type: "rect", x: margin, y: top, w: width - 2 * margin, h: height, fill: BHB.navy, under: true });
      shapes.push({ type: "rect", x: margin, y: top + height, w: width - 2 * margin, h: 1, fill: BHB.rule });
      y = top + height;
    });
  }
  y += margin;
  // Crop: the item's own bounds plus the 24 px margin (F15).
  const right = Math.max(...texts.map((entryText) => entryText.x + font.getAdvanceWidth(entryText.line, entryText.size)), ...shapes.map((shape) => shape.x + (shape.w || 0)), 0);
  return { width: Math.ceil(Math.min(width, right + margin)), height: Math.ceil(y), shapes, texts, itemTop: Math.floor(itemTop) };
}
/** F27: the PNG and SVG carry the item only; the caption strip stays on the PDF page and in the manifest. */
function itemOnly(layout) {
  const shift = layout.itemTop;
  return { width: layout.width, height: layout.height - shift, shapes: layout.shapes.filter((shape) => !shape.caption).map((shape) => ({ ...shape, y: shape.y - shift, top: shape.top === undefined ? undefined : shape.top - shift })), texts: layout.texts.filter((entry) => !entry.caption).map((entry) => ({ ...entry, y: entry.y - shift })) };
}
function svgItem(layout, font) {
  const shapes = layout.shapes.map((shape) => shape.type === "band" ? `<rect x="${shape.x.toFixed(1)}" y="${shape.y.toFixed(1)}" width="${shape.w.toFixed(1)}" height="${shape.h.toFixed(1)}" fill="${BHB.band}"/>` : `<rect x="${shape.x.toFixed(1)}" y="${shape.y.toFixed(1)}" width="${Math.max(0, shape.w).toFixed(1)}" height="${shape.h.toFixed(1)}" fill="${shape.fill}"/>`);
  const ordered = [...shapes.filter((_, index) => layout.shapes[index].type === "band" || layout.shapes[index].under), ...shapes.filter((_, index) => layout.shapes[index].type !== "band" && !layout.shapes[index].under)];
  const paths = layout.texts.map((entry) => `<path d="${font.getPath(entry.line, entry.x, entry.y, entry.size).toPathData(2)}" fill="${entry.color}"/>`);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${layout.width}" height="${layout.height}" viewBox="0 0 ${layout.width} ${layout.height}" data-font="Poppins" data-text-outlined="true"><rect width="${layout.width}" height="${layout.height}" fill="#FFFFFF"/>${ordered.join("")}${paths.join("")}</svg>`;
}
export async function buildBundleArtifacts(XLSX, payload) {
  const font = await fontPath(); const rawFont = await fontBytes();
  const items = bundleItems(payload);
  const manifest = []; const svg = []; const layouts = [];
  for (const entry of items) {
    const layout = layoutItem(entry, font); layouts.push(layout);
    const image = itemOnly(layout);
    svg.push({ name: `SVG/${entry.id}-${entry.item.family}.svg`, content: svgItem(image, font), width: image.width, height: image.height });
    manifest.push({ id: entry.id, title: entry.item.title, family: entry.item.family, group: entry.item.groupKey || "whole-project", groupLabel: entry.item.groupLabel || "", benchmark: entry.item.low !== undefined ? `${f2(entry.item.low)}–${f2(entry.item.high)}` : "", caption: entry.subtitle || "", deckSlide: entry.item.number, width: image.width, height: image.height, pdfPage: { width: layout.width, height: layout.height }, sourceValues: entry.sourceValues.map(String) });
  }
  const workbook = XLSX.utils.book_new();
  for (const [index, entry] of items.entries()) {
    const rows = entry.kind === "chart" ? [["", ...entry.series.map((series) => series.name)], ...entry.categories.map((category, categoryIndex) => [category, ...entry.series.map((series) => series.values[categoryIndex] ?? "")])] : entry.rows;
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[manifest[index].title], [manifest[index].group], ...rows]), `Item-${entry.id}`);
  }
  return { manifest: { version: "AC-GRF-R-1.1", bijection: true, items: manifest }, svg, layouts, workbookBytes: XLSX.write(workbook, { type: "array", bookType: "xlsx", compression: true }), pdfBytes: makePdf(layouts, rawFont, font) };
}
async function pngFromSvg(svg, width, height) { const image = new Image(); const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`; await new Promise((resolve, reject) => { image.onload = resolve; image.onerror = reject; image.src = url; }); const canvas = document.createElement("canvas"); canvas.width = width * 3; canvas.height = height * 3; const context = canvas.getContext("2d"); context.scale(3, 3); context.drawImage(image, 0, 0, width, height); return canvas.toDataURL("image/png").split(",")[1]; }
export async function generateBundle(payload) {
  if (!window.JSZip || !window.XLSX) throw new Error("Lipsește biblioteca locală pentru bundle.");
  const artifacts = await buildBundleArtifacts(window.XLSX, payload); const zip = new window.JSZip();
  for (const item of artifacts.svg) { zip.file(item.name, item.content); zip.file(item.name.replace(/^SVG\//u, "PNG/").replace(/\.svg$/u, ".png"), await pngFromSvg(item.content, item.width, item.height), { base64: true }); }
  zip.file("data.xlsx", artifacts.workbookBytes); zip.file("manifest.json", JSON.stringify(artifacts.manifest, null, 2)); zip.file("01-report-items.pdf", artifacts.pdfBytes);
  return zip.generateAsync({ type: "blob", mimeType: "application/zip", compression: "DEFLATE" });
}
function downloadBlob(blob, filename) { const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = filename; document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
export async function downloadTrendPptx(payload, filename = "trend-group-report.pptx", options = {}) { downloadBlob(await generateTrendPptx(payload, options), filename); }
export async function downloadBundle(payload, filename = "bhb-report-bundle.zip") { downloadBlob(await generateBundle(payload), filename); }

function pdfNumber(value) { return Number(value || 0).toFixed(2); }
const pdfColor = (hex) => { const value = hex.replace("#", ""); return [0, 2, 4].map((offset) => (parseInt(value.slice(offset, offset + 2), 16) / 255).toFixed(3)).join(" "); };
function pdfPath(font, entry, pageHeight) {
  const commands = []; let currentX = 0; let currentY = 0; const flip = (value) => pdfNumber(pageHeight - value);
  for (const command of font.getPath(entry.line, entry.x, entry.y, entry.size).commands) {
    if (command.type === "M") { currentX = command.x; currentY = command.y; commands.push(`${pdfNumber(currentX)} ${flip(currentY)} m`); }
    else if (command.type === "L") { currentX = command.x; currentY = command.y; commands.push(`${pdfNumber(currentX)} ${flip(currentY)} l`); }
    else if (command.type === "C") { currentX = command.x; currentY = command.y; commands.push(`${pdfNumber(command.x1)} ${flip(command.y1)} ${pdfNumber(command.x2)} ${flip(command.y2)} ${pdfNumber(currentX)} ${flip(currentY)} c`); }
    else if (command.type === "Q") { const endX = command.x; const endY = command.y; commands.push(`${pdfNumber(currentX + (2 / 3) * (command.x1 - currentX))} ${flip(currentY + (2 / 3) * (command.y1 - currentY))} ${pdfNumber(endX + (2 / 3) * (command.x1 - endX))} ${flip(endY + (2 / 3) * (command.y1 - endY))} ${pdfNumber(endX)} ${flip(endY)} c`); currentX = endX; currentY = endY; }
    else if (command.type === "Z") commands.push("h");
  }
  return commands.length ? `${pdfColor(entry.color)} rg ${commands.join(" ")} f` : "";
}
function makePdf(layouts, rawFont, font) {
  const objects = []; const add = (value) => { objects.push(value); return objects.length; };
  const fontFile = add({ stream: rawFont, dict: `<< /Length ${rawFont.length} /Length1 ${rawFont.length} >>` });
  const descriptor = add(`<< /Type /FontDescriptor /FontName /Poppins /Flags 4 /FontBBox [0 -200 1200 1000] /ItalicAngle 0 /Ascent 1000 /Descent -250 /CapHeight 700 /StemV 80 /FontFile2 ${fontFile} 0 R >>`);
  const cid = add(`<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Poppins /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${descriptor} 0 R /CIDToGIDMap /Identity >>`);
  const cmapBytes = new TextEncoder().encode("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Poppins-UTF16 def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <ffff>\nendcodespacerange\n1 beginbfrange\n<0000> <ffff> <0000>\nendbfrange\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend");
  const cmap = add({ stream: cmapBytes, dict: `<< /Length ${cmapBytes.length} >>` });
  const type0 = add(`<< /Type /Font /Subtype /Type0 /BaseFont /Poppins /Encoding /Identity-H /DescendantFonts [${cid} 0 R] /ToUnicode ${cmap} 0 R >>`);
  const pages = add(""); const pageIds = [];
  for (const layout of layouts) {
    const h = layout.height;
    const rects = [...layout.shapes.filter((shape) => shape.type === "band" || shape.under), ...layout.shapes.filter((shape) => shape.type !== "band" && !shape.under)].map((shape) => `${pdfColor(shape.type === "band" ? BHB.band : shape.fill)} rg ${pdfNumber(shape.x)} ${pdfNumber(h - shape.y - shape.h)} ${pdfNumber(Math.max(0, shape.w))} ${pdfNumber(shape.h)} re f`);
    const content = `q 1 1 1 rg 0 0 ${layout.width} ${h} re f ${rects.join(" ")} ${layout.texts.map((entry) => pdfPath(font, entry, h)).filter(Boolean).join(" ")} Q`;
    const contentBytes = new TextEncoder().encode(content);
    const contentId = add({ stream: contentBytes, dict: `<< /Length ${contentBytes.length} >>` });
    pageIds.push(add(`<< /Type /Page /Parent ${pages} 0 R /MediaBox [0 0 ${layout.width} ${h}] /Resources << /Font << /F1 ${type0} 0 R >> >> /Contents ${contentId} 0 R >>`));
  }
  objects[pages - 1] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;
  const catalog = add(`<< /Type /Catalog /Pages ${pages} 0 R >>`);
  const chunks = [new TextEncoder().encode("%PDF-1.7\n")]; const offsets = [0]; let length = chunks[0].length;
  objects.forEach((object, index) => { offsets[index + 1] = length; const head = new TextEncoder().encode(`${index + 1} 0 obj\n`); const body = object?.stream ? new Uint8Array([...new TextEncoder().encode(`${object.dict}\nstream\n`), ...object.stream, ...new TextEncoder().encode("\nendstream\nendobj\n")]) : new TextEncoder().encode(`${object}\nendobj\n`); chunks.push(head, body); length += head.length + body.length; });
  const xref = length;
  chunks.push(new TextEncoder().encode(`xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objects.length + 1} /Root ${catalog} 0 R >>\nstartxref\n${xref}\n%%EOF`));
  const output = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0)); let offset = 0;
  for (const chunk of chunks) { output.set(chunk, offset); offset += chunk.length; }
  return output;
}
