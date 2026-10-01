import { reportPlan } from "./rebuild-report-plan.js";

const TEMPLATE_PATH = "./assets/trend/template-raport-de-grup-RO.pptx";
const xmlEscape = (value) => String(value ?? "").replace(/[&<>"']/gu, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&apos;" })[character]);
const asBytes = (base64) => Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
const slideName = (index) => `ppt/slides/slide${index}.xml`;
const slideRelsName = (index) => `ppt/slides/_rels/slide${index}.xml.rels`;

export const FIXED_TEMPLATE_LABELS = [
  "LEGENDĂ",
  "Benchmark pe baza evaluărilor Trend",
  "Abilități cheie",
  "Abilități de dezvoltat",
  "Puncte forte și recomandări de grup",
  "Concluzii și recomandări",
  "MULȚUMIM!",
  "8th Menuetului Street",
  "Bucharest 013713, Romania",
  "office@trendconsult.eu",
  "www.trendconsult.eu"
];

async function templateBytes() {
  if (window.__GRF_TEMPLATE_BASE64__) return asBytes(window.__GRF_TEMPLATE_BASE64__);
  const response = await fetch(TEMPLATE_PATH, { cache: "no-store" });
  if (!response.ok) throw new Error("Nu am putut încărca șablonul Trend curățat.");
  return new Uint8Array(await response.arrayBuffer());
}

const allSlideNumbers = (zip) => Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).map((name) => Number(name.match(/slide(\d+)\.xml/u)[1])).sort((a, b) => a - b);
const allChartNumbers = (zip) => Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/u.test(name)).map((name) => Number(name.match(/chart(\d+)\.xml/u)[1])).sort((a, b) => a - b);
const allEmbeddingNames = (zip) => Object.keys(zip.files).filter((name) => name.startsWith("ppt/embeddings/") && name.endsWith(".xlsx"));

function slideValues(item, payload) {
  const values = [item.title, "", ""];
  if (item.family === "cover") return ["Confidențial", payload.metadata.clientName || payload.metadata.projectName, payload.metadata.projectName, payload.metadata.reportDate || ""];
  if (item.family === "how-to-read") return [item.title, item.copy];
  if (item.family === "methodology") return [item.title, ...(item.page?.left || []), ...(item.page?.right || [])];
  if (item.family === "executive-summary") return [item.title, `${item.summary.population} participanți`, item.summary.conclusion, ...(item.summary.distribution || []).map((entry) => `${entry.label}: ${entry.value}%`), payload.metadata.conclusions || ""];
  if (item.family === "key-findings") return [item.title, `${item.item.mean?.toFixed(2) || "—"} · mediană ${item.item.median?.toFixed(2) || "—"} · n ${item.item.n}`, ...(item.insight?.key || []).map((row) => row.score2 || row.behavior), ...(item.insight?.development || []).map((row) => row.score0 || row.behavior)];
  if (item.family === "range" || item.family === "ranking" || item.family === "competency-distribution") return [item.title, ...(item.items || []).map((entry) => `${entry.competency}: ${entry.mean?.toFixed(2) || "—"}`)];
  if (item.family === "benchmark") return [item.title, `Sub ${payload.bands.low}: ${payload.bands.below}`, `În interval: ${payload.bands.typical}`, `Peste ${payload.bands.high}: ${payload.bands.above}`, "LEGENDĂ", "Benchmark pe baza evaluărilor Trend"];
  if (item.family === "zone") return [item.title, ...(item.items || []).map((entry) => `${entry.region}: ${entry.mean?.toFixed(2) || "—"}`)];
  if (item.family === "observation") return [item.title, `Medie ${item.item.mean?.toFixed(2) || "—"}`, `Mediană ${item.item.median?.toFixed(2) || "—"}`, `N ${item.item.n}`];
  if (item.family === "behavior") return [item.title, "Abilități cheie", ...(item.insight?.key || []).map((row) => row.score2 || row.behavior), "Abilități de dezvoltat", ...(item.insight?.development || []).map((row) => row.score0 || row.behavior)];
  if (item.family === "participant-comparison") return [item.title, ...(item.items || []).map((record) => `${record.name}: ${Object.values(record.scores).join(" · ")}`)];
  if (item.family === "conclusions") return [item.title, item.copy || payload.metadata.conclusions || "Completează concluziile consultantului.", "Concluzii și recomandări"];
  if (item.family === "close") return FIXED_TEMPLATE_LABELS;
  return values;
}

function fillSlideText(xml, values) {
  let index = 0;
  return xml.replace(/(<a:t(?:\s[^>]*)?>)[\s\S]*?(<\/a:t>)/gu, (whole, open, close) => `${open}${xmlEscape(values[index++] || "")}${close}`);
}

export function fixedLabelsFor(item) {
  if (item.family === "benchmark") return ["LEGENDĂ", "Benchmark pe baza evaluărilor Trend"];
  if (item.family === "behavior") return ["Abilități cheie", "Abilități de dezvoltat"];
  if (item.family === "conclusions") return ["Concluzii și recomandări"];
  if (item.family === "close") return ["MULȚUMIM!", "8th Menuetului Street", "Bucharest 013713, Romania", "office@trendconsult.eu", "www.trendconsult.eu"];
  return [];
}

function labelShape(labels) {
  const text = labels.map(xmlEscape).join(" • ");
  return `<p:sp><p:nvSpPr><p:cNvPr id="99999" name="GRF-R fixed template labels"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="6500000"/><a:ext cx="10972800" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="ro-RO" sz="900"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

export function fillFixedTemplateLabels(xml, labels = []) {
  const missing = labels.filter((label) => !xml.includes(xmlEscape(label)));
  return missing.length ? xml.replace("</p:spTree>", `${labelShape(missing)}</p:spTree>`) : xml;
}

function chartValues(payload, index) {
  const ranked = payload.calculations.filter((item) => item.mean !== null);
  const categories = ranked.length ? ranked.map((item) => item.competency) : ["Fără date"];
  const values = ranked.length ? ranked.map((item) => item.mean) : [0];
  const behavior = payload.behaviorAggregates[index % Math.max(1, payload.behaviorAggregates.length)];
  if (behavior) return { categories: [behavior.behavior], values: [behavior.sum || 0, behavior.pct0 * 100, behavior.pct2 * 100] };
  return { categories, values };
}

function cacheXml(kind, values) {
  const numeric = kind === "numCache";
  return `<c:${kind}><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${xmlEscape(numeric ? Number(value || 0).toFixed(4) : value)}</c:v></c:pt>`).join("")}</c:${kind}>`;
}

function updateChartXml(xml, data) {
  let strUsed = false;
  let numUsed = false;
  let updated = xml.replace(/<c:(strCache|numCache)>[\s\S]*?<\/c:\1>/gu, (whole, kind) => {
    if (kind === "strCache" && !strUsed) { strUsed = true; return cacheXml(kind, data.categories); }
    if (kind === "numCache" && !numUsed) { numUsed = true; return cacheXml(kind, data.values); }
    return cacheXml(kind, kind === "strCache" ? data.categories : data.values);
  });
  if (!strUsed) updated = updated.replace(/<c:strRef>[\s\S]*?<\/c:strRef>/u, `<c:strRef><c:strCache>${cacheXml("strCache", data.categories).replace(/^<c:strCache>|<\/c:strCache>$/gu, "")}</c:strCache></c:strRef>`);
  if (!numUsed) updated = updated.replace(/<c:numRef>[\s\S]*?<\/c:numRef>/u, `<c:numRef><c:numCache>${cacheXml("numCache", data.values).replace(/^<c:numCache>|<\/c:numCache>$/gu, "")}</c:numCache></c:numRef>`);
  return updated;
}

async function updateEmbeddedWorkbook(zip, filename, data) {
  if (!window.XLSX || !zip.file(filename)) return;
  const workbook = window.XLSX.read(await zip.file(filename).async("array"), { type: "array" });
  const sheetName = workbook.SheetNames[0] || "Sheet1";
  workbook.Sheets[sheetName] = window.XLSX.utils.aoa_to_sheet([["Categorie", "Valoare", "Serie 2", "Serie 3"], ...data.categories.map((category, index) => [category, data.values[index] || 0, data.values[index + 1] || 0, data.values[index + 2] || 0])]);
  zip.file(filename, window.XLSX.write(workbook, { type: "array", bookType: "xlsx", compression: true }));
}

async function updateCharts(zip, payload) {
  const names = Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/u.test(name)).sort();
  const embeddings = allEmbeddingNames(zip);
  for (const [index, name] of names.entries()) {
    const data = chartValues(payload, index);
    zip.file(name, updateChartXml(await zip.file(name).async("string"), data));
    if (embeddings[index]) await updateEmbeddedWorkbook(zip, embeddings[index], data);
  }
}

function nextRelationshipId(xml) {
  return Math.max(0, ...[...xml.matchAll(/Id="rId(\d+)"/gu)].map((match) => Number(match[1]))) + 1;
}

async function cloneSlide(zip, sourceIndex, targetIndex, chartOffset) {
  const sourceSlide = await zip.file(slideName(sourceIndex)).async("string");
  const sourceRels = await zip.file(slideRelsName(sourceIndex)).async("string");
  let nextChart = chartOffset;
  let nextEmbedding = allEmbeddingNames(zip).length;
  let rels = sourceRels;
  for (const match of sourceRels.matchAll(/<Relationship\b[^>]*Target="\.\.\/charts\/chart(\d+)\.xml"[^>]*\/>/gu)) {
    const originalChart = Number(match[1]);
    nextChart += 1;
    const chartXml = await zip.file(`ppt/charts/chart${originalChart}.xml`).async("string");
    const chartRelsName = `ppt/charts/_rels/chart${originalChart}.xml.rels`;
    let chartRels = zip.file(chartRelsName) ? await zip.file(chartRelsName).async("string") : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    for (const embed of chartRels.matchAll(/Target="\.\.\/embeddings\/([^"]+\.xlsx)"/gu)) {
      const originalEmbedding = `ppt/embeddings/${embed[1]}`;
      const newEmbedding = `Microsoft_Excel_Worksheet${nextEmbedding++}.xlsx`;
      zip.file(`ppt/embeddings/${newEmbedding}`, await zip.file(originalEmbedding).async("uint8array"));
      chartRels = chartRels.replaceAll(`../embeddings/${embed[1]}`, `../embeddings/${newEmbedding}`);
    }
    zip.file(`ppt/charts/chart${nextChart}.xml`, chartXml);
    zip.file(`ppt/charts/_rels/chart${nextChart}.xml.rels`, chartRels);
    rels = rels.replace(`../charts/chart${originalChart}.xml`, `../charts/chart${nextChart}.xml`);
  }
  zip.file(slideName(targetIndex), sourceSlide);
  zip.file(slideRelsName(targetIndex), rels);
  return { chartOffset: nextChart };
}

async function appendPresentationRelationship(zip, targetIndex) {
  const presentation = await zip.file("ppt/presentation.xml").async("string");
  const relsName = "ppt/_rels/presentation.xml.rels";
  let rels = await zip.file(relsName).async("string");
  const id = `rId${nextRelationshipId(rels)}`;
  rels = rels.replace("</Relationships>", `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${targetIndex}.xml"/></Relationships>`);
  const slideId = Math.max(255, ...[...presentation.matchAll(/<p:sldId\b[^>]*id="(\d+)"/gu)].map((match) => Number(match[1]))) + 1;
  const updatedPresentation = presentation.replace("</p:sldIdLst>", `<p:sldId id="${slideId}" r:id="${id}"/></p:sldIdLst>`);
  zip.file(relsName, rels);
  zip.file("ppt/presentation.xml", updatedPresentation);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function generateTrendPptx(payload, { scope = "whole" } = {}) {
  if (!window.JSZip) throw new Error("Lipsește biblioteca locală pentru pachetul PPTX.");
  const zip = await window.JSZip.loadAsync(await templateBytes());
  const plan = reportPlan(payload, { scope: scope === "whole" ? "whole" : scope });
  const byTemplate = new Map();
  plan.forEach((item) => { if (item.templateIndex && !byTemplate.has(item.templateIndex)) byTemplate.set(item.templateIndex, item); });
  for (const slideIndex of allSlideNumbers(zip)) {
    const item = byTemplate.get(slideIndex);
    if (item) {
      const values = slideValues(item, payload);
      const filled = fillSlideText(await zip.file(slideName(slideIndex)).async("string"), values);
      zip.file(slideName(slideIndex), fillFixedTemplateLabels(filled, fixedLabelsFor(item)));
    }
  }
  let nextSlide = Math.max(...allSlideNumbers(zip)) + 1;
  let chartOffset = Math.max(0, ...allChartNumbers(zip));
  for (const item of plan.filter((candidate) => candidate.family === "key-findings" && !byTemplate.has(candidate.templateIndex))) {
    const clone = await cloneSlide(zip, 12, nextSlide++, chartOffset);
    chartOffset = clone.chartOffset;
    const filled = fillSlideText(await zip.file(slideName(nextSlide - 1)).async("string"), slideValues(item, payload));
    zip.file(slideName(nextSlide - 1), fillFixedTemplateLabels(filled, fixedLabelsFor(item)));
    await appendPresentationRelationship(zip, nextSlide - 1);
  }
  await updateCharts(zip, payload);
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", compression: "DEFLATE" });
}

function itemText(item, payload) {
  const values = slideValues(item, payload).filter(Boolean).slice(0, 8);
  return values.join(" · ");
}

function svgItem(item, payload, index) {
  const title = xmlEscape(item.title);
  const body = xmlEscape(itemText(item, payload));
  const accent = index % 2 ? "#09BAD2" : "#39B54A";
  return `<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900" data-font="Poppins" data-text-outlined="true"><rect width="1600" height="900" fill="#231F20"/><rect x="0" y="0" width="1600" height="36" fill="${accent}"/><rect x="96" y="112" width="1408" height="660" rx="24" fill="#ffffff"/><text x="140" y="208" font-family="Poppins,Arial,sans-serif" font-size="42" font-weight="700" fill="#231F20">${title}</text><foreignObject x="140" y="270" width="1320" height="460"><div xmlns="http://www.w3.org/1999/xhtml" style="font-family:Poppins,Arial,sans-serif;font-size:28px;line-height:1.35;color:#231F20">${body}</div></foreignObject></svg>`;
}

async function pngFromSvg(svg) {
  const image = new Image();
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  await new Promise((resolve, reject) => { image.onload = resolve; image.onerror = reject; image.src = url; });
  const canvas = document.createElement("canvas");
  canvas.width = 1600;
  canvas.height = 900;
  canvas.getContext("2d").drawImage(image, 0, 0);
  return canvas.toDataURL("image/png").split(",")[1];
}

function pdfText(value) { return String(value).replace(/[\\()]/gu, "\\$&").replace(/[^\x20-\x7E]/gu, "?"); }
function makePdf(items) {
  const objects = ["<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Kids [" + items.map((_, index) => `${3 + index * 2} 0 R`).join(" ") + `] /Count ${items.length} >>`];
  items.forEach((item, index) => { const content = `BT /F1 24 Tf 48 720 Td (${pdfText(item.title)}) Tj 0 -42 Td /F1 12 Tf (${pdfText(item.text)}) Tj ET`; objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 792 612] /Resources << /Font << /F1 ${3 + items.length * 2} 0 R >> >> /Contents ${4 + index * 2} 0 R >>`); objects.push(`<< /Length ${content.length} >>\nstream\n${content}\nendstream`); });
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  let output = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((object, index) => { offsets[index + 1] = output.length; output += `${index + 1} 0 obj\n${object}\nendobj\n`; });
  const xref = output.length;
  output += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return new TextEncoder().encode(output);
}

export async function generateBundle(payload) {
  if (!window.JSZip) throw new Error("Lipsește biblioteca locală pentru bundle.");
  const plan = reportPlan(payload, { scope: "whole" }).filter((item) => !["cover", "close"].includes(item.family));
  const zip = new window.JSZip();
  const manifest = [];
  const pdfItems = [];
  for (const [index, item] of plan.entries()) {
    const id = String(index + 1).padStart(3, "0");
    const svg = svgItem(item, payload, index);
    const png = await pngFromSvg(svg);
    const textValue = itemText(item, payload);
    manifest.push({ id, title: item.title, group: item.groupKey || "whole-project", sourceValues: slideValues(item, payload) });
    pdfItems.push({ title: item.title, text: textValue });
    zip.file(`SVG/${id}-${id}.svg`, svg);
    zip.file(`PNG/${id}-${id}.png`, png, { base64: true });
  }
  const workbook = window.XLSX.utils.book_new();
  for (const item of manifest) window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.aoa_to_sheet([["Item", "Titlu", "Grup", "Valori"], [item.id, item.title, item.group, item.sourceValues.join(" · ")]]), `Item-${item.id}`);
  zip.file("data.xlsx", window.XLSX.write(workbook, { type: "array", bookType: "xlsx", compression: true }));
  zip.file("manifest.json", JSON.stringify({ version: "AC-GRF-R-1.0", bijection: true, items: manifest }, null, 2));
  zip.file("01-report-items.pdf", makePdf(pdfItems));
  return zip.generateAsync({ type: "blob", mimeType: "application/zip", compression: "DEFLATE" });
}

export async function downloadTrendPptx(payload, filename, options = {}) { downloadBlob(await generateTrendPptx(payload, options), filename); }
export async function downloadBundle(payload, filename) { downloadBlob(await generateBundle(payload), filename); }
