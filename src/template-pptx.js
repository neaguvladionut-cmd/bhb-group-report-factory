import { reportPlan } from "./rebuild-report-plan.js";

const TEMPLATE_PATH = "./assets/trend/template-raport-de-grup-RO.pptx";
const FONT_PATH = "./assets/vendor/Poppins-Regular.ttf";
const relsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const xmlEscape = (value) => String(value ?? "").replace(/[&<>"']/gu, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&apos;" })[character]);
const asBytes = (base64) => typeof Buffer !== "undefined" ? Buffer.from(base64, "base64") : Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
const slideName = (index) => `ppt/slides/slide${index}.xml`;
const slideRelsName = (index) => `ppt/slides/_rels/slide${index}.xml.rels`;

export const FIXED_TEMPLATE_LABELS = [
  "LEGENDĂ", "Benchmark pe baza evaluărilor Trend", "Abilități cheie", "Abilități de dezvoltat",
  "Puncte forte și recomandări de grup", "Concluzii și recomandări", "MULȚUMIM!", "8th Menuetului Street",
  "Bucharest 013713, Romania", "office@trendconsult.eu", "www.trendconsult.eu"
];

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

function slideValues(item, payload) {
  if (item.family === "cover") return ["Confidențial", payload.metadata.clientName || payload.metadata.projectName, payload.metadata.projectName, payload.metadata.reportDate || ""];
  if (item.family === "how-to-read") return [item.title, item.copy];
  if (item.family === "methodology") return [item.title, ...(item.page?.left || []), ...(item.page?.right || [])];
  if (item.family === "executive-summary") return [item.title, `${item.summary.population} participanți`, item.summary.conclusion, ...(item.summary.distribution || []).map((entry) => `${entry.label}: ${entry.value}%`), payload.metadata.conclusions || ""];
  if (item.family === "key-findings") return [item.title, `${item.item.mean?.toFixed(2) || "—"} · mediană ${item.item.median?.toFixed(2) || "—"} · n ${item.item.n}`, ...(item.insight?.key || []).map((row) => row.score2 || row.behavior), ...(item.insight?.development || []).map((row) => row.score0 || row.behavior)];
  if (["range", "ranking"].includes(item.family)) return [item.title, ...(item.items || []).map((entry) => `${entry.competency}: ${entry.mean?.toFixed(2) || "—"}`)];
  if (item.family === "competency-distribution" && item.appendix) return [item.title, ...(item.items || []).map((record) => `${record.name}: ${record.scores?.[item.item.competency] ?? "—"}`)];
  if (item.family === "competency-distribution") return [item.title, ...(item.items || []).map((entry) => `${entry.competency}: ${entry.mean?.toFixed(2) || "—"}`)];
  if (item.family === "benchmark") return [item.title, `Sub ${payload.bands.low}: ${payload.bands.below}`, `În interval: ${payload.bands.typical}`, `Peste ${payload.bands.high}: ${payload.bands.above}`];
  if (item.family === "zone") return [item.title, ...(item.items || []).map((entry) => `${entry.region}: ${entry.mean?.toFixed(2) || "—"}`)];
  if (item.family === "observation") return [item.title, `Medie ${item.item.mean?.toFixed(2) || "—"}`, `Mediană ${item.item.median?.toFixed(2) || "—"}`, `N ${item.item.n}`];
  if (item.family === "behavior") return [item.title, ...(item.insight?.key || []).map((row) => row.score2 || row.behavior), ...(item.insight?.development || []).map((row) => row.score0 || row.behavior)];
  if (item.family === "participant-comparison") return [item.title, ...(item.items || []).map((record) => `${record.name}: ${Object.values(record.scores || {}).join(" · ")}`)];
  if (item.family === "conclusions") return [item.title, item.copy || payload.metadata.conclusions || "Completează concluziile consultantului."];
  return [item.title];
}

export function fixedLabelsFor(item) {
  if (item.family === "benchmark") return ["LEGENDĂ", "Benchmark pe baza evaluărilor Trend"];
  if (item.family === "behavior") return ["Puncte forte și recomandări de grup", "Abilități cheie", "Abilități de dezvoltat"];
  if (item.family === "conclusions") return ["Concluzii și recomandări"];
  if (item.family === "close") return FIXED_TEMPLATE_LABELS.slice(6);
  return [];
}

function textShape(role, value, y, size, color = "231F20") {
  const lines = String(value || "").split("\n");
  const runs = lines.map((line, index) => `${index ? "<a:br/>" : ""}<a:r><a:rPr lang="ro-RO" sz="${size}"/><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:t>${xmlEscape(line)}</a:t></a:r>`).join("");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${90000 + y}" name="GRF-R role:${role}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="${y}"/><a:ext cx="16459200" cy="${role === "title" ? 700000 : 5900000}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p>${runs}</a:p></p:txBody></p:sp>`;
}
function labelShape(labels) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="99999" name="GRF-R fixed template labels"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="6500000"/><a:ext cx="16459200" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="ro-RO" sz="900"/><a:t>${labels.map(xmlEscape).join(" • ")}</a:t></a:r></a:p></p:txBody></p:sp>`;
}
export function fillFixedTemplateLabels(xml, labels = []) {
  const missing = labels.filter((label) => !xml.includes(xmlEscape(label)));
  return missing.length ? xml.replace("</p:spTree>", `${labelShape(missing)}</p:spTree>`) : xml;
}
function renderSlide(item, payload, sourceXml) {
  const values = slideValues(item, payload).filter((value) => value !== undefined && value !== null && String(value).trim()).map(String);
  const title = item.title;
  if (values[0] === title) values.shift();
  let xml = sourceXml.replace(/<a:t(?:\s[^>]*)?>[\s\S]*?<\/a:t>/gu, "<a:t></a:t>");
  xml = xml.replace("</p:spTree>", `${textShape("title", title, 550000, 2400)}${textShape("body", values.join("\n"), 1450000, 1250)}</p:spTree>`);
  return fillFixedTemplateLabels(xml, fixedLabelsFor(item));
}

function chartSeriesFor(item, payload) {
  if (item.family === "benchmark") return { categories: [`Sub ${payload.bands.low}`, `În intervalul ${payload.bands.low}–${payload.bands.high}`, `Peste ${payload.bands.high}`], series: [{ name: "Participanți", values: [payload.bands.below, payload.bands.typical, payload.bands.above] }] };
  if (item.family === "zone") return { categories: (item.items || []).map((entry) => entry.region), series: [{ name: "Medie", values: (item.items || []).map((entry) => entry.mean ?? 0) }] };
  if (item.family === "competency-distribution" && item.appendix) return { categories: (item.items || []).map((record) => record.name), series: [{ name: item.item.competency, values: (item.items || []).map((record) => record.scores?.[item.item.competency] ?? 0) }] };
  if (["range", "ranking", "competency-distribution"].includes(item.family)) return { categories: (item.items || []).map((entry) => entry.competency), series: [{ name: "Medie", values: (item.items || []).map((entry) => entry.mean ?? 0) }] };
  if (item.family === "behavior") { const rows = [...(item.insight?.key || []), ...(item.insight?.development || [])]; return { categories: rows.map((entry) => entry.behavior), series: [{ name: "Sumă scoruri", values: rows.map((entry) => entry.sum ?? 0) }] }; }
  const rows = payload.calculations.filter((entry) => entry.mean !== null);
  return { categories: rows.map((entry) => entry.competency), series: [{ name: "Medie", values: rows.map((entry) => entry.mean ?? 0) }] };
}
function cacheXml(kind, values) {
  const numeric = kind === "numCache";
  return `<c:${kind}><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${xmlEscape(numeric ? Number(value || 0).toFixed(4) : value)}</c:v></c:pt>`).join("")}</c:${kind}>`;
}
function refXml(kind, formula, values) { return `<c:${kind}Ref><c:f>${xmlEscape(formula)}</c:f>${cacheXml(kind === "str" ? "strCache" : "numCache", values)}</c:${kind}Ref>`; }
function updateChartXml(xml, data) {
  let seriesIndex = 0;
  return xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/gu, (seriesXml) => {
    const series = data.series[Math.min(seriesIndex, data.series.length - 1)] || { name: "Serie", values: [] };
    const end = Math.max(2, data.categories.length + 1);
    let updated = seriesXml.replace(/<c:tx>[\s\S]*?<\/c:tx>/u, `<c:tx><c:v>${xmlEscape(series.name)}</c:v></c:tx>`);
    updated = updated.replace(/<c:cat>[\s\S]*?<\/c:cat>/u, `<c:cat>${refXml("str", `Sheet1!$A$2:$A$${end}`, data.categories)}</c:cat>`);
    updated = updated.replace(/<c:val>[\s\S]*?<\/c:val>/u, `<c:val>${refXml("num", `Sheet1!$${String.fromCharCode(66 + seriesIndex)}$2:$${String.fromCharCode(66 + seriesIndex)}$${end}`, series.values)}</c:val>`);
    seriesIndex += 1;
    return updated;
  });
}
async function updateEmbeddedWorkbook(zip, filename, data) {
  if (!window.XLSX || !filename || !zip.file(filename)) return;
  const workbook = window.XLSX.read(await zip.file(filename).async("array"), { type: "array" });
  const sheetName = workbook.SheetNames[0] || "Sheet1";
  const rows = [["Categorie", ...data.series.map((series) => series.name)]];
  data.categories.forEach((category, index) => rows.push([category, ...data.series.map((series) => series.values[index] ?? 0)]));
  workbook.Sheets[sheetName] = window.XLSX.utils.aoa_to_sheet(rows);
  zip.file(filename, window.XLSX.write(workbook, { type: "array", bookType: "xlsx", compression: true }));
}
function relationshipTargets(xml) { return [...xml.matchAll(/<Relationship\b([^>]*)\/>/gu)].map((match) => Object.fromEntries([...match[1].matchAll(/(Id|Type|Target|TargetMode)="([^"]*)"/gu)].map((item) => [item[1], item[2]]))); }
function relationshipTarget(source, target) {
  if (/^https?:|^mailto:/u.test(target)) return null;
  const base = source === "_rels/.rels" ? "" : source.split("/").slice(0, -1).join("/");
  const parts = `${base}/${target}`.replace(/\/+/gu, "/").split("/");
  const output = [];
  for (const part of parts) { if (part === "..") output.pop(); else if (part && part !== ".") output.push(part); }
  return output.join("/");
}
async function packageGraph(zip) {
  const reachable = new Set(); const queue = ["ppt/presentation.xml"]; const dangling = [];
  if (zip.file("_rels/.rels")) for (const rel of relationshipTargets(await zip.file("_rels/.rels").async("string"))) { const target = relationshipTarget("_rels/.rels", rel.Target); if (target) queue.push(target); }
  while (queue.length) {
    const part = queue.shift(); if (reachable.has(part)) continue;
    if (!zip.file(part)) { dangling.push(part); continue; }
    reachable.add(part);
    const relPart = `${part.split("/").slice(0, -1).join("/")}/_rels/${part.split("/").at(-1)}.rels`;
    if (zip.file(relPart)) { reachable.add(relPart); for (const rel of relationshipTargets(await zip.file(relPart).async("string"))) { const target = relationshipTarget(part, rel.Target); if (target) queue.push(target); } }
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
  for (const entry of names) { if (/ppt\/fonts\//u.test(entry)) errors.push("ppt/fonts part present"); if (entry.endsWith(".xml")) { const xml = await zip.file(entry).async("string"); if (/<(?:p:|a:)[^>]*embeddedFontLst|<p:notesMasterIdLst/iu.test(xml)) errors.push(`forbidden IdLst entry in ${entry}`); } }
  if (errors.length) throw new Error(`PPTX self-check failed: ${errors.join("; ")}`);
  return { ok: true, parts: names.length, reachable: graph.reachable.size };
}

async function cloneSlide(zip, sourceIndex, targetIndex, chartOffset, embeddingOffset, sourceSlides) {
  const sourceSlide = sourceSlides.get(sourceIndex); const sourceRels = sourceSlides.get(`${sourceIndex}.rels`); let nextChart = chartOffset; let nextEmbedding = embeddingOffset; let rels = sourceRels;
  for (const match of sourceRels.matchAll(/<Relationship\b[^>]*Target="\.\.\/charts\/chart(\d+)\.xml"[^>]*\/>/gu)) {
    const originalChart = Number(match[1]); nextChart += 1;
    const chartFile = zip.file(`ppt/charts/chart${originalChart}.xml`);
    if (!chartFile) throw new Error(`Template chart${originalChart}.xml is missing while cloning slide${sourceIndex}.xml.`);
    const chartXml = await chartFile.async("string"); const chartRelsName = `ppt/charts/_rels/chart${originalChart}.xml.rels`;
    let chartRels = zip.file(chartRelsName) ? await zip.file(chartRelsName).async("string") : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    for (const embed of chartRels.matchAll(/Target="\.\.\/embeddings\/([^"]+\.xlsx)"/gu)) { const originalEmbedding = `ppt/embeddings/${embed[1]}`; const newEmbedding = `Microsoft_Excel_Worksheet${nextEmbedding ? nextEmbedding : ""}.xlsx`; nextEmbedding += 1; zip.file(`ppt/embeddings/${newEmbedding}`, await zip.file(originalEmbedding).async("uint8array")); chartRels = chartRels.replaceAll(`../embeddings/${embed[1]}`, `../embeddings/${newEmbedding}`); }
    zip.file(`ppt/charts/chart${nextChart}.xml`, chartXml); zip.file(`ppt/charts/_rels/chart${nextChart}.xml.rels`, chartRels); rels = rels.replace(`../charts/chart${originalChart}.xml`, `../charts/chart${nextChart}.xml`);
  }
  zip.file(slideName(targetIndex), sourceSlide); zip.file(slideRelsName(targetIndex), rels); return { chartOffset: nextChart, embeddingOffset: nextEmbedding };
}
async function updateCharts(zip, plan, payload) {
  for (const [index, item] of plan.entries()) { const rels = await zip.file(slideRelsName(index + 1)).async("string"); const data = chartSeriesFor(item, payload); for (const rel of relationshipTargets(rels).filter((entry) => entry.Type === `${relsType}chart`)) { const chartPath = relationshipTarget(slideName(index + 1), rel.Target); if (!chartPath || !zip.file(chartPath)) continue; zip.file(chartPath, updateChartXml(await zip.file(chartPath).async("string"), data)); const chartRelsPath = `${chartPath.split("/").slice(0, -1).join("/")}/_rels/${chartPath.split("/").at(-1)}.rels`; const chartRels = zip.file(chartRelsPath) ? await zip.file(chartRelsPath).async("string") : ""; const workbookRel = relationshipTargets(chartRels).find((entry) => entry.Type === `${relsType}package`); if (workbookRel) await updateEmbeddedWorkbook(zip, relationshipTarget(chartPath, workbookRel.Target), data); } }
}
function updateContentTypes(xml, zip) {
  let updated = xml.replace(/<Override\b[^>]*PartName="([^"]+)"[^>]*\/>/gu, (whole, partName) => zip.file(partName.replace(/^\//u, "")) ? whole : "").replace(/<Override\b[^>]*PartName="\/ppt\/(?:slides\/slide\d+\.xml|charts\/chart\d+\.xml)"[^>]*\/>/gu, ""); const overrides = [];
  for (const name of Object.keys(zip.files)) { if (/^ppt\/slides\/slide\d+\.xml$/u.test(name)) overrides.push(`<Override PartName="/${name}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`); if (/^ppt\/charts\/chart\d+\.xml$/u.test(name)) overrides.push(`<Override PartName="/${name}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`); }
  return updated.replace("</Types>", `${overrides.join("")}</Types>`);
}
async function prunePackage(zip) {
  const graph = await packageGraph(zip); for (const name of Object.keys(zip.files)) if (!name.endsWith("/") && name !== "[Content_Types].xml" && /^(ppt\/(slides|charts|embeddings)\/)/u.test(name) && !graph.reachable.has(name)) zip.remove(name);
  zip.file("[Content_Types].xml", updateContentTypes(await zip.file("[Content_Types].xml").async("string"), zip));
}
async function rebuildSlides(zip, plan, payload) {
  const sourceSlides = new Map(); for (const index of allSlideNumbers(zip)) { sourceSlides.set(index, await zip.file(slideName(index)).async("string")); sourceSlides.set(`${index}.rels`, await zip.file(slideRelsName(index)).async("string")); zip.remove(slideName(index)); zip.remove(slideRelsName(index)); }
  let chartOffset = 0; let embeddingOffset = 0;
  for (const [index, item] of plan.entries()) { const sourceIndex = sourceSlides.has(item.templateIndex) ? item.templateIndex : 2; const clone = await cloneSlide(zip, sourceIndex, index + 1, chartOffset, embeddingOffset, sourceSlides); chartOffset = clone.chartOffset; embeddingOffset = clone.embeddingOffset; zip.file(slideName(index + 1), renderSlide(item, payload, await zip.file(slideName(index + 1)).async("string"))); }
  const presentation = await zip.file("ppt/presentation.xml").async("string"); const relsPath = "ppt/_rels/presentation.xml.rels"; let rels = await zip.file(relsPath).async("string"); rels = rels.replace(/<Relationship\b[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide"[^>]*\/>/gu, ""); const nextRel = Math.max(0, ...[...rels.matchAll(/Id="rId(\d+)"/gu)].map((match) => Number(match[1]))) + 1; const slideRelationships = plan.map((_, index) => `<Relationship Id="rId${nextRel + index}" Type="${relsType}slide" Target="slides/slide${index + 1}.xml"/>`).join(""); rels = rels.replace("</Relationships>", `${slideRelationships}</Relationships>`); const slideIds = plan.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${nextRel + index}"/>`).join(""); zip.file(relsPath, rels); zip.file("ppt/presentation.xml", presentation.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/u, `<p:sldIdLst>${slideIds}</p:sldIdLst>`)); await updateCharts(zip, plan, payload); await prunePackage(zip);
}
export async function generateTrendPptx(payload, { scope = "whole" } = {}) { if (!window.JSZip) throw new Error("Lipsește biblioteca locală pentru pachetul PPTX."); const zip = await window.JSZip.loadAsync(await templateBytes()); const plan = reportPlan(payload, { scope }); await rebuildSlides(zip, plan, payload); await selfCheckPptx(zip); return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", compression: "DEFLATE" }); }

function itemText(item, payload) { return slideValues(item, payload).filter((value) => value !== undefined && value !== null && String(value).trim()).slice(0, 8).join(" · "); }
async function fontPath() { if (!window.opentype) throw new Error("Lipsește opentype.js vendorizat."); const bytes = await fontBytes(); const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength); return window.opentype.parse(buffer); }
function svgItem(item, payload, index, font) { const title = String(item.title || ""); const body = itemText(item, payload); const accent = index % 2 ? "09BAD2" : "39B54A"; const pathFor = (value, x, y, size) => font.getPath(value, x, y, size).toPathData(2); return `<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900" data-font="Poppins" data-text-outlined="true"><rect width="1600" height="900" fill="#231F20"/><rect x="0" y="0" width="1600" height="36" fill="#${accent}"/><rect x="96" y="112" width="1408" height="660" rx="24" fill="#ffffff"/><path d="${pathFor(title, 140, 208, 42)}" fill="#231F20"/><path d="${pathFor(body, 140, 300, 28)}" fill="#231F20"/></svg>`; }
async function pngFromSvg(svg) { const image = new Image(); const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`; await new Promise((resolve, reject) => { image.onload = resolve; image.onerror = reject; image.src = url; }); const canvas = document.createElement("canvas"); canvas.width = 1600; canvas.height = 900; canvas.getContext("2d").drawImage(image, 0, 0); return canvas.toDataURL("image/png").split(",")[1]; }
const pdfHex = (value) => [...String(value || "")].map((character) => character.codePointAt(0).toString(16).padStart(4, "0")).join("");
export async function buildBundleArtifacts(XLSX, payload) { const font = await fontPath(); const rawFont = await fontBytes(); const plan = reportPlan(payload, { scope: "whole" }).filter((item) => !["cover", "close"].includes(item.family)); const manifest = []; const pdfItems = []; const svg = []; for (const [index, item] of plan.entries()) { const id = String(index + 1).padStart(3, "0"); const sourceValues = slideValues(item, payload).filter((value) => value !== undefined && value !== null).map(String); svg.push({ name: `SVG/${id}-${id}.svg`, content: svgItem(item, payload, index, font) }); manifest.push({ id, title: item.title, group: item.groupKey || "whole-project", sourceValues }); pdfItems.push({ title: item.title, text: itemText(item, payload) }); } const workbook = XLSX.utils.book_new(); for (const item of manifest) XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["Item", "Titlu", "Grup", "Valori"], [item.id, item.title, item.group, item.sourceValues.join(" · ")]]), `Item-${item.id}`); return { manifest: { version: "AC-GRF-R-1.0", bijection: true, items: manifest }, svg, workbookBytes: XLSX.write(workbook, { type: "array", bookType: "xlsx", compression: true }), pdfBytes: makePdf(pdfItems, rawFont) }; }
export async function generateBundle(payload) { if (!window.JSZip || !window.XLSX) throw new Error("Lipsește biblioteca locală pentru bundle."); const artifacts = await buildBundleArtifacts(window.XLSX, payload); const zip = new window.JSZip(); for (const item of artifacts.svg) zip.file(item.name, item.content); const plan = reportPlan(payload, { scope: "whole" }).filter((item) => !["cover", "close"].includes(item.family)); for (const [index, item] of plan.entries()) zip.file(`PNG/${String(index + 1).padStart(3, "0")}-${String(index + 1).padStart(3, "0")}.png`, await pngFromSvg(artifacts.svg[index].content), { base64: true }); zip.file("data.xlsx", artifacts.workbookBytes); zip.file("manifest.json", JSON.stringify(artifacts.manifest, null, 2)); zip.file("01-report-items.pdf", artifacts.pdfBytes); return zip.generateAsync({ type: "blob", mimeType: "application/zip", compression: "DEFLATE" }); }
function downloadBlob(blob, filename) { const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = filename; document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
export async function downloadTrendPptx(payload, options) { downloadBlob(await generateTrendPptx(payload, options), "trend-group-report.pptx"); }
export async function downloadBundle(payload) { downloadBlob(await generateBundle(payload), "bhb-report-bundle.zip"); }
function makePdf(items, font) {
  const objects = [];
  const add = (value) => { objects.push(value); return objects.length; };
  const fontFile = add({ stream: font, dict: `<< /Length ${font.length} /Length1 ${font.length} >>` });
  const descriptor = add(`<< /Type /FontDescriptor /FontName /Poppins /Flags 4 /FontBBox [0 -200 1200 1000] /ItalicAngle 0 /Ascent 1000 /Descent -250 /CapHeight 700 /StemV 80 /FontFile2 ${fontFile} 0 R >>`);
  const cid = add(`<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Poppins /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${descriptor} 0 R /CIDToGIDMap /Identity >>`);
  const cmapText = `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Poppins-UTF16 def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <ffff>\nendcodespacerange\n1 beginbfrange\n<0000> <ffff> <0000>\nendbfrange\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend`;
  const cmapBytes = new TextEncoder().encode(cmapText);
  const cmap = add({ stream: cmapBytes, dict: `<< /Length ${cmapBytes.length} >>` });
  const type0 = add(`<< /Type /Font /Subtype /Type0 /BaseFont /Poppins /Encoding /Identity-H /DescendantFonts [${cid} 0 R] /ToUnicode ${cmap} 0 R >>`);
  const pages = add("");
  const pageIds = [];
  for (const item of items) {
    const content = `BT /F1 24 Tf 48 720 Td <${pdfHex(item.title)}> Tj 0 -42 Td /F1 12 Tf <${pdfHex(item.text)}> Tj ET`;
    const contentBytes = new TextEncoder().encode(content);
    const contentId = add({ stream: contentBytes, dict: `<< /Length ${contentBytes.length} >>` });
    pageIds.push(add(`<< /Type /Page /Parent ${pages} 0 R /MediaBox [0 0 792 612] /Resources << /Font << /F1 ${type0} 0 R >> >> /Contents ${contentId} 0 R >>`));
  }
  objects[pages - 1] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;
  const catalog = add(`<< /Type /Catalog /Pages ${pages} 0 R >>`);
  const chunks = [new TextEncoder().encode("%PDF-1.7\n")];
  const offsets = [0];
  let length = chunks[0].length;
  objects.forEach((object, index) => {
    offsets[index + 1] = length;
    const head = new TextEncoder().encode(`${index + 1} 0 obj\n`);
    const body = object?.stream ? new Uint8Array([...new TextEncoder().encode(`${object.dict}\nstream\n`), ...object.stream, ...new TextEncoder().encode("\nendstream\nendobj\n")]) : new TextEncoder().encode(`${object}\nendobj\n`);
    chunks.push(head, body);
    length += head.length + body.length;
  });
  const xref = length;
  chunks.push(new TextEncoder().encode(`xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objects.length + 1} /Root ${catalog} 0 R >>\nstartxref\n${xref}\n%%EOF`));
  return chunks.reduce((all, chunk) => new Uint8Array([...all, ...chunk]), new Uint8Array());
}
