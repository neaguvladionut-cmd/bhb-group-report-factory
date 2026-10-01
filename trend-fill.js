// BP-GRF-R fill layer: writes plan data into the named shapes of a cloned Trend template slide
// (BP-GRF-R-fill-map.md). String-level OOXML surgery so it runs the same in the browser and in Node.
// Rule 1: text goes into the template shape's own paragraphs, copying its pPr and run rPr; the only
// shapes created are the map's „New” shapes, each named "GRF-R new:*" and styled from a template shape.

export const EMU = 914400;
export const NEW_SHAPE_PREFIX = "GRF-R new:";
export const xmlEscape = (value) => String(value ?? "").replace(/[&<>"']/gu, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&apos;" })[character]);
const f2 = (value) => Number(value).toFixed(2);

// ---------- shape location ----------
export function shapeRange(xml, id) {
  const marker = new RegExp(`<p:cNvPr\\b[^>]*\\bid="${id}"`, "u").exec(xml);
  if (!marker) return null;
  const open = [...xml.slice(0, marker.index).matchAll(/<p:(sp|graphicFrame|grpSp|pic|cxnSp)\b/gu)].at(-1);
  if (!open) return null;
  const tag = open[1];
  const re = new RegExp(`<p:${tag}\\b|</p:${tag}>`, "gu");
  re.lastIndex = open.index;
  let depth = 0;
  for (let match = re.exec(xml); match; match = re.exec(xml)) {
    depth += match[0].startsWith("</") ? -1 : 1;
    if (!depth) return { start: open.index, end: match.index + match[0].length, tag };
  }
  return null;
}
export function getShape(xml, id) { const range = shapeRange(xml, id); if (!range) throw new Error(`Template shape id ${id} is missing.`); return xml.slice(range.start, range.end); }
export function updateShape(xml, id, transform) { const range = shapeRange(xml, id); if (!range) throw new Error(`Template shape id ${id} is missing.`); return `${xml.slice(0, range.start)}${transform(xml.slice(range.start, range.end))}${xml.slice(range.end)}`; }
export function hasShape(xml, id) { return Boolean(shapeRange(xml, id)); }
export function hideShape(shapeXml) { return shapeXml.replace(/<p:cNvPr\b([^>]*?)(\s*\/?>)/u, (whole, attrs, close) => `<p:cNvPr${attrs.replace(/\shidden="[^"]*"/u, "")} hidden="1"${close}`); }
export function addToTree(xml, shapeXml) { return xml.replace("</p:spTree>", `${shapeXml}</p:spTree>`); }

export function xfrmOf(shapeXml) {
  const off = shapeXml.match(/<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>/u); const ext = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/u);
  if (!off || !ext) return null;
  return { x: Number(off[1]), y: Number(off[2]), cx: Number(ext[1]), cy: Number(ext[2]) };
}
export function setXfrm(shapeXml, { x, y, cx, cy }) {
  return shapeXml.replace(/<a:off x="-?\d+" y="-?\d+"\s*\/>/u, `<a:off x="${Math.round(x)}" y="${Math.round(y)}"/>`).replace(/<a:ext cx="\d+" cy="\d+"\s*\/>/u, `<a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/>`);
}

// ---------- paragraphs and runs ----------
const txBodyRange = (shapeXml) => { const open = shapeXml.search(/<(?:p|a):txBody>/u); const close = shapeXml.search(/<\/(?:p|a):txBody>/u); return open < 0 ? null : { open, close }; };
export const paragraphsOf = (xml) => xml.match(/<a:p>[\s\S]*?<\/a:p>|<a:p\b[^>]*>[\s\S]*?<\/a:p>/gu) || [];
const cleanRPr = (rPr) => rPr.replace(/\s(?:err|dirty|noProof)="[^"]*"/gu, "");
export function runRPrs(paragraphXml) {
  return [...paragraphXml.matchAll(/<a:r>\s*(<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>)?[\s\S]*?<\/a:r>/gu)].map((match) => cleanRPr(match[1] || "<a:rPr lang=\"ro-RO\"/>"));
}
export const runTexts = (paragraphXml) => [...paragraphXml.matchAll(/<a:r>[\s\S]*?<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>[\s\S]*?<\/a:r>/gu)].map((match) => match[1]);
const pPrOf = (paragraphXml) => paragraphXml.match(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/u)?.[0] || "";
const endRPrOf = (paragraphXml) => paragraphXml.match(/<a:endParaRPr\b[^>]*\/>|<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/u)?.[0] || "";
const endAsRPr = (endRPr) => endRPr ? cleanRPr(endRPr.replace(/^<a:endParaRPr/u, "<a:rPr").replace(/<\/a:endParaRPr>$/u, "</a:rPr>")) : "<a:rPr lang=\"ro-RO\"/>";
export function setRPrAttr(rPr, name, value) {
  const open = rPr.match(/^<a:rPr\b[^>]*?(\/?)>/u);
  let tag = open[0].replace(new RegExp(`\\s${name}="[^"]*"`, "u"), "");
  if (value !== null && value !== undefined) tag = tag.replace(/^<a:rPr\b/u, `<a:rPr ${name}="${value}"`);
  return tag + rPr.slice(open[0].length);
}
export function setRPrColor(rPr, hex) {
  const selfClosing = /\/>$/u.test(rPr) && !/<\/a:rPr>$/u.test(rPr);
  const fill = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  if (selfClosing) return rPr.replace(/\/>$/u, `>${fill}</a:rPr>`);
  const body = rPr.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/u, "");
  return body.replace(/^(<a:rPr\b[^>]*>)(<a:ln\b[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>)?/u, (whole, head, ln) => `${head}${ln || ""}${fill}`);
}
const rPrSize = (rPr) => Number(rPr.match(/\ssz="(\d+)"/u)?.[1] || 0);

/** Template paragraph styles of a shape: [{pPr, rPrs:[...], endRPr}] */
export function templateParagraphs(shapeXml) {
  return paragraphsOf(shapeXml).map((paragraph) => { const rPrs = runRPrs(paragraph); const endRPr = endRPrOf(paragraph); return { pPr: pPrOf(paragraph), rPrs: rPrs.length ? rPrs : [endAsRPr(endRPr)], endRPr }; });
}
export function regularRPr(shapeXml) { for (const style of templateParagraphs(shapeXml)) for (const rPr of style.rPrs) if (!/\sb="1"/u.test(rPr)) return rPr; return setRPrAttr(templateParagraphs(shapeXml)[0]?.rPrs[0] || "<a:rPr/>", "b", null); }
export function boldRPr(shapeXml) { for (const style of templateParagraphs(shapeXml)) for (const rPr of style.rPrs) if (/\sb="1"/u.test(rPr)) return rPr; return setRPrAttr(regularRPr(shapeXml), "b", "1"); }

const run = (rPr, value) => `<a:r>${rPr}<a:t>${xmlEscape(value)}</a:t></a:r>`;
/**
 * Replace a shape's paragraphs. `paragraphs` is an array; each entry is a string or an array of
 * segments {text, rPr?, run?, bold?, color?}. Paragraph k copies the template paragraph `styleOf(k)`.
 */
export function setParagraphs(shapeXml, paragraphs, { style = (index, styles) => styles[Math.min(index, styles.length - 1)] } = {}) {
  const styles = templateParagraphs(shapeXml);
  if (!styles.length) throw new Error("Template shape has no paragraph to copy the style from.");
  const output = (paragraphs.length ? paragraphs : [""]).map((paragraph, index) => {
    const template = style(index, styles);
    const segments = Array.isArray(paragraph) ? paragraph : [{ text: paragraph }];
    const runs = segments.filter((segment) => String(segment.text ?? "") !== "").map((segment) => {
      let rPr = segment.rPr || template.rPrs[Math.min(segment.run ?? 0, template.rPrs.length - 1)];
      if (segment.bold === true) rPr = setRPrAttr(rPr, "b", "1");
      if (segment.bold === false) rPr = setRPrAttr(rPr, "b", "0");
      if (segment.color) rPr = setRPrColor(rPr, segment.color);
      if (segment.size) rPr = setRPrAttr(rPr, "sz", String(segment.size));
      return run(rPr, segment.text);
    }).join("");
    const end = template.endRPr || endAsRPr("").replace("<a:rPr", "<a:endParaRPr");
    return `<a:p>${template.pPr}${runs}${end.startsWith("<a:endParaRPr") ? end : ""}</a:p>`;
  }).join("");
  const range = txBodyRange(shapeXml);
  const body = shapeXml.slice(range.open, range.close);
  const firstParagraph = body.search(/<a:p[\s>]/u);
  return `${shapeXml.slice(0, range.open)}${body.slice(0, firstParagraph)}${output}${shapeXml.slice(range.close)}`;
}
/** Append text to the last run that has visible text (group suffix, competency in a fixed title). */
export function appendToLastRun(shapeXml, suffix, { newRun = false } = {}) {
  if (!suffix) return shapeXml;
  const runs = [...shapeXml.matchAll(/<a:r>([\s\S]*?)<a:t(\s[^>]*)?>([\s\S]*?)<\/a:t>([\s\S]*?)<\/a:r>/gu)];
  const target = [...runs].reverse().find((match) => match[3].trim()) || runs.at(-1);
  if (!target) throw new Error("Template title has no run to append to.");
  const trimmed = target[3].replace(/\s+$/u, "");
  const replacement = newRun ? `<a:r>${target[1]}<a:t${target[2] || ""}>${target[3]}</a:t>${target[4]}</a:r><a:r>${target[1]}<a:t>${xmlEscape(suffix)}</a:t></a:r>` : `<a:r>${target[1]}<a:t${target[2] || ""}>${trimmed}${xmlEscape(suffix)}</a:t>${target[4]}</a:r>`;
  // Drop whitespace-only runs after the target so the suffix sits at the end of the title.
  const head = shapeXml.slice(0, target.index);
  const tail = shapeXml.slice(target.index + target[0].length).replace(/<a:r>(?:(?!<\/a:r>)[\s\S])*?<a:t(?:\s[^>]*)?>\s*<\/a:t>(?:(?!<\/a:r>)[\s\S])*?<\/a:r>/gu, "");
  return `${head}${replacement}${tail}`;
}
export function setRunText(shapeXml, runIndex, value) {
  let index = 0;
  return shapeXml.replace(/(<a:r>[\s\S]*?<a:t(?:\s[^>]*)?>)([\s\S]*?)(<\/a:t>[\s\S]*?<\/a:r>)/gu, (whole, head, current, tail) => (index++ === runIndex ? `${head}${xmlEscape(value)}${tail}` : whole));
}
export const shapeText = (shapeXml) => paragraphsOf(shapeXml).map((paragraph) => runTexts(paragraph).join("")).join("\n");

// ---------- rule 7: text fit ----------
const bodyPrOf = (shapeXml) => shapeXml.match(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/u)?.[0] || "";
const insetOf = (bodyPr, name, fallback) => Number(bodyPr.match(new RegExp(`\\s${name}="(\\d+)"`, "u"))?.[1] ?? fallback);
export function textHeightNeeded(shapeXml, scale = 1, widthEmu = null, spacing = 1) {
  const geometry = xfrmOf(shapeXml); const bodyPr = bodyPrOf(shapeXml);
  const width = (widthEmu ?? geometry.cx) - insetOf(bodyPr, "lIns", 91440) - insetOf(bodyPr, "rIns", 91440);
  let height = 0;
  for (const paragraph of paragraphsOf(shapeXml)) {
    const sizes = runRPrs(paragraph).map(rPrSize).filter(Boolean);
    const size = (sizes.length ? Math.max(...sizes) : rPrSize(endRPrOf(paragraph).replace("endParaRPr", "rPr")) || 1800) / 100 * scale;
    const pPr = pPrOf(paragraph);
    const margin = Number(pPr.match(/\smarL="(\d+)"/u)?.[1] || 0);
    const lineFactor = Number(pPr.match(/<a:lnSpc><a:spcPct val="(\d+)"/u)?.[1] || 100000) / 100000;
    const before = Number(pPr.match(/<a:spcBef><a:spcPts val="(\d+)"/u)?.[1] || 0) / 100 * scale;
    const after = Number(pPr.match(/<a:spcAft><a:spcPts val="(\d+)"/u)?.[1] || 0) / 100 * scale;
    const content = runTexts(paragraph).join("").replace(/&[a-z]+;/gu, "x");
    const usable = Math.max(1, (width - margin) / 12700); // points
    const rPrs = runRPrs(paragraph); const letterSpacing = Math.max(0, ...rPrs.map((rPr) => Number(rPr.match(/\sspc="(-?\d+)"/u)?.[1] || 0))) / 100 * scale;
    const charWidth = size * (rPrs.some((rPr) => /\sb="1"/u.test(rPr)) ? 0.57 : 0.53) + letterSpacing;
    let lines = 0;
    for (const part of content.split("\t").join("    ").split("\n")) {
      let line = 0; lines += 1;
      for (const word of part.split(/\s+/u).filter(Boolean)) {
        const wordWidth = word.length * charWidth; const space = line ? charWidth * 0.55 : 0;
        if (line && line + space + wordWidth > usable) { lines += 1; line = wordWidth; } else line += space + wordWidth;
        while (line > usable) { lines += 1; line -= usable; }
      }
    }
    height += lines * size * 1.2 * lineFactor * spacing + before + after;
  }
  return height * 12700 + insetOf(bodyPr, "tIns", 45720) + insetOf(bodyPr, "bIns", 45720);
}
/** Scale that enclosing groups apply to a shape's child coordinates (text keeps its point size). */
export function groupScale(xml, id) {
  const range = shapeRange(xml, id); let scaleX = 1; let scaleY = 1;
  if (!range) return { scaleX, scaleY };
  const groupOpen = /<p:grpSp>/gu; let match;
  while ((match = groupOpen.exec(xml)) && match.index < range.start) {
    const group = shapeRange(xml.slice(0, match.index) + xml.slice(match.index), (xml.slice(match.index).match(/<p:cNvPr\b[^>]*\bid="(\d+)"/u) || [])[1]);
    if (!group || group.end < range.end) continue;
    const head = xml.slice(group.start, group.end).match(/<p:grpSpPr>[\s\S]*?<\/a:xfrm>/u)?.[0] || "";
    const ext = head.match(/<a:ext cx="(\d+)" cy="(\d+)"/u); const chExt = head.match(/<a:chExt cx="(\d+)" cy="(\d+)"/u);
    if (ext && chExt && Number(chExt[1]) && Number(chExt[2])) { scaleX *= Number(ext[1]) / Number(chExt[1]); scaleY *= Number(ext[2]) / Number(chExt[2]); }
  }
  return { scaleX, scaleY };
}
export function fitText(shapeXml, { minScale = 0.7, height = null, scaleX = 1, scaleY = 1 } = {}) {
  const geometry = xfrmOf(shapeXml); if (!geometry) return shapeXml;
  const available = (height ?? geometry.cy) * scaleY;
  const width = geometry.cx * scaleX;
  const textHeightNeeded_ = textHeightNeeded;
  const textHeightNeeded__ = (xml, scale, ignored, spacing = 1) => textHeightNeeded_(xml, scale, width, spacing);
  if (textHeightNeeded__(shapeXml, 1) <= available) return shapeXml;
  let scale = 1;
  while (scale > minScale && textHeightNeeded__(shapeXml, scale) > available) scale = Math.round((scale - 0.025) * 1000) / 1000;
  scale = Math.max(minScale, scale);
  // Still too tall at the 70 % floor: let PowerPoint also reduce line spacing (up to 20 %), as its own autofit does.
  let reduction = 0;
  while (reduction < 0.2 && textHeightNeeded__(shapeXml, scale, null, 1 - reduction) > available) reduction = Math.round((reduction + 0.05) * 100) / 100;
  return setAutofit(shapeXml, `<a:normAutofit fontScale="${Math.round(scale * 100000)}"${reduction ? ` lnSpcReduction="${Math.round(reduction * 100000)}"` : ""}/>`);
}
export function setAutofit(shapeXml, autofit) {
  const bodyPr = bodyPrOf(shapeXml);
  const open = bodyPr.match(/^<a:bodyPr\b[^>]*?(?=\/?>)/u)[0];
  let inner = bodyPr.endsWith("</a:bodyPr>") ? bodyPr.replace(/^<a:bodyPr\b[^>]*>/u, "").replace(/<\/a:bodyPr>$/u, "") : "";
  inner = inner.replace(/<a:(?:spAutoFit|noAutofit|normAutofit)\b[^>]*\/>/gu, "").replace(/<a:normAutofit\b[^>]*>[\s\S]*?<\/a:normAutofit>/gu, "");
  const warp = inner.match(/^<a:prstTxWarp\b[\s\S]*?<\/a:prstTxWarp>/u)?.[0] || "";
  return shapeXml.replace(bodyPr, `${open}>${warp}${autofit}${inner.slice(warp.length)}</a:bodyPr>`);
}
export function fontScaleOf(shapeXml) { const value = shapeXml.match(/<a:normAutofit fontScale="(\d+)"/u)?.[1]; return value ? Number(value) / 100000 : 1; }
export function setBodyInsets(shapeXml, insets) {
  return shapeXml.replace(/<a:bodyPr\b([^>]*?)(\/?)>/u, (whole, attrs, slash) => { let next = attrs; for (const [name, value] of Object.entries(insets)) { next = next.replace(new RegExp(`\\s${name}="[^"]*"`, "u"), ""); next += ` ${name}="${Math.round(value)}"`; } return `<a:bodyPr${next}${slash}>`; });
}

// ---------- rule 4: benchmark bands ----------
/** Move a band rectangle drawn for 2.75–3.50 on a fixed 1–5 axis to [low, high]. */
const quarterTurn = (shapeXml) => { const rot = Number(shapeXml.match(/<a:xfrm\b[^>]*\brot="(-?\d+)"/u)?.[1] || 0); return Math.round(rot / 5400000) % 2 !== 0; };
/** The box a shape occupies on screen (a 90°/270° rotation swaps its width and height about the centre). */
export function visualBox(shapeXml) {
  const box = xfrmOf(shapeXml); if (!quarterTurn(shapeXml)) return box;
  const centerX = box.x + box.cx / 2; const centerY = box.y + box.cy / 2;
  return { x: centerX - box.cy / 2, y: centerY - box.cx / 2, cx: box.cy, cy: box.cx };
}
export function setVisualBox(shapeXml, visual) {
  if (!quarterTurn(shapeXml)) return setXfrm(shapeXml, visual);
  const centerX = visual.x + visual.cx / 2; const centerY = visual.y + visual.cy / 2;
  return setXfrm(shapeXml, { x: centerX - visual.cy / 2, y: centerY - visual.cx / 2, cx: visual.cy, cy: visual.cx });
}
/** Rule 4: move a band rectangle drawn for 2.75–3.50 on a fixed 1–5 axis to [low, high]. */
export function moveBand(shapeXml, { low, high, axis }) {
  const box = visualBox(shapeXml); const templateLow = 2.75; const templateHigh = 3.5;
  if (axis === "x") { const perPoint = box.cx / (templateHigh - templateLow); return setVisualBox(shapeXml, { x: box.x + (low - templateLow) * perPoint, y: box.y, cx: (high - low) * perPoint, cy: box.cy }); }
  const perPoint = box.cy / (templateHigh - templateLow);
  return setVisualBox(shapeXml, { x: box.x, y: box.y + (templateHigh - high) * perPoint, cx: box.cx, cy: (high - low) * perPoint });
}
/** Inner plot area (fractions of the chart frame) that puts 1 and 5 where the band's mapping says. */
export function plotLayoutFromBand(bandShape, frameShape, { low, high, axis, cross }) {
  const band = visualBox(bandShape); const frame = xfrmOf(frameShape);
  if (axis === "x") { const perPoint = band.cx / (high - low); const x1 = band.x - (low - 1) * perPoint; return { x: (x1 - frame.x) / frame.cx, w: 4 * perPoint / frame.cx, y: cross[0], h: cross[1] }; }
  const perPoint = band.cy / (high - low); const y5 = band.y - (5 - high) * perPoint; return { y: (y5 - frame.y) / frame.cy, h: 4 * perPoint / frame.cy, x: cross[0], w: cross[1] };
}
export function setPlotLayout(chartXml, { x, y, w, h }) {
  const layout = `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="${x.toFixed(4)}"/><c:y val="${y.toFixed(4)}"/><c:w val="${w.toFixed(4)}"/><c:h val="${h.toFixed(4)}"/></c:manualLayout></c:layout>`;
  return chartXml.replace(/<c:plotArea>\s*(?:<c:layout\/>|<c:layout>[\s\S]*?<\/c:layout>)?/u, `<c:plotArea>${layout}`);
}
// ---------- rule 5/6: tables and braces ----------
export const tableRows = (frameXml) => frameXml.match(/<a:tr\b[\s\S]*?<\/a:tr>/gu) || [];
export const rowHeight = (rowXml) => Number(rowXml.match(/<a:tr\b[^>]*\bh="(\d+)"/u)?.[1] || 0);
// Grey rows (inside the benchmark) carry a cell fill; white rows have <a:noFill/> (t7 D9D9D9, t12 D0CECE).
export const isShadedRow = (rowXml) => (rowXml.match(/<a:tcPr\b[\s\S]*?<\/a:tcPr>/gu) || []).some((tcPr) => /<a:solidFill>/u.test(tcPr.replace(/<a:ln[LRTB]\b[\s\S]*?<\/a:ln[LRTB]>/gu, "")));
export function setCellText(cellXml, value, { size = null } = {}) {
  const txBody = cellXml.match(/<a:txBody>[\s\S]*?<\/a:txBody>/u)[0];
  const paragraph = paragraphsOf(txBody)[0] || "<a:p></a:p>";
  const rPr0 = runRPrs(paragraph)[0] || endAsRPr(endRPrOf(paragraph));
  const rPr = size ? setRPrAttr(rPr0, "sz", String(size)) : rPr0;
  const end = endRPrOf(paragraph);
  const lines = String(value ?? "").split("\n");
  const paragraphs = lines.map((line) => `<a:p>${pPrOf(paragraph)}${line ? run(rPr, line) : ""}${end ? (size ? end.replace(/\ssz="\d+"/u, ` sz="${size}"`) : end) : ""}</a:p>`).join("");
  const head = txBody.slice(0, txBody.search(/<a:p[\s>]/u));
  return cellXml.replace(txBody, `${head}${paragraphs}</a:txBody>`);
}
export function setRowCells(rowXml, values, options = {}) {
  let index = 0;
  return rowXml.replace(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/gu, (cell) => { const value = values[index]; index += 1; return value === undefined ? cell : setCellText(cell, value, options); });
}
export const setRowHeight = (rowXml, height) => rowXml.replace(/(<a:tr\b[^>]*\bh=")\d+(")/u, `$1${Math.round(height)}$2`);
export const setRowId = (rowXml, value) => rowXml.replace(/(<a16:rowId\b[^>]*\bval=")\d+(")/u, `$1${value}$2`);
export function replaceTableRows(frameXml, rows) {
  const first = frameXml.search(/<a:tr\b/u); const last = frameXml.lastIndexOf("</a:tr>") + "</a:tr>".length;
  const height = rows.reduce((sum, row) => sum + rowHeight(row), 0);
  const updated = `${frameXml.slice(0, first)}${rows.join("")}${frameXml.slice(last)}`;
  return updated.replace(/(<p:xfrm>[\s\S]*?<a:ext cx="\d+" cy=")\d+(")/u, `$1${height}$2`);
}
/** Rows of fixed total height: never below 0.18 in; beyond that the font steps down to 9 pt (rule 5). */
export function rowLayout(totalHeight, count, baseSize) {
  const minimum = Math.round(0.18 * EMU);
  if (!count) return { height: totalHeight, size: null };
  const height = totalHeight / count;
  if (height >= minimum) return { height, size: null };
  return { height: minimum, size: Math.max(900, Math.min(baseSize || 1600, 900)) };
}
export function placeBrace(xml, braceId, labelId, span, label) {
  if (!span) return updateShape(updateShape(xml, braceId, hideShape), labelId, hideShape);
  xml = updateShape(xml, braceId, (shape) => { const box = xfrmOf(shape); return setXfrm(shape, { ...box, y: span.top, cy: Math.max(1, span.bottom - span.top) }); });
  return updateShape(xml, labelId, (shape) => { const box = xfrmOf(shape); const filled = setParagraphs(shape, [String(label)]); return setXfrm(filled, { ...box, y: (span.top + span.bottom) / 2 - box.cy / 2 }); });
}

// ---------- rule 8: native charts ----------
export function columnName(index) { let value = index + 1; let output = ""; while (value) { const remainder = (value - 1) % 26; output = String.fromCharCode(65 + remainder) + output; value = Math.floor((value - 1) / 26); } return output; }
const strCache = (values) => `<c:strCache><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${xmlEscape(value)}</c:v></c:pt>`).join("")}</c:strCache>`;
const numCache = (values, formatCode) => `<c:numCache><c:formatCode>${xmlEscape(formatCode)}</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => (value === null || value === undefined || !Number.isFinite(Number(value)) ? "" : `<c:pt idx="${index}"><c:v>${Number(value)}</c:v></c:pt>`)).join("")}</c:numCache>`;
export const roundChartValue = (value) => (value === null || value === undefined || !Number.isFinite(Number(value)) ? null : Math.round(Number(value) * 100) / 100);
function recolor(seriesXml, from, to) { return from ? seriesXml.replaceAll(`<a:srgbClr val="${from}"`, `<a:srgbClr val="${to}"`) : seriesXml; }
function seriesFillColor(seriesXml) { return seriesXml.match(/<\/c:tx>\s*<c:spPr>[\s\S]*?<a:solidFill>\s*<a:srgbClr val="([0-9A-F]{6})"/iu)?.[1] || null; }
function setSeriesFill(seriesXml, hex) {
  const current = seriesFillColor(seriesXml);
  if (current) return seriesXml.replace(/(<\/c:tx>\s*<c:spPr>[\s\S]*?<a:solidFill>\s*)<a:srgbClr val="[0-9A-F]{6}"/iu, `$1<a:srgbClr val="${hex}"`);
  return seriesXml.replace(/(<\/c:tx>\s*<c:spPr>[\s\S]*?<a:solidFill>\s*)<a:schemeClr val="\w+"\s*\/>/u, `$1<a:srgbClr val="${hex}"/>`).replace(/(<\/c:tx>\s*<c:spPr>[\s\S]*?<a:solidFill>\s*)<a:schemeClr val="\w+">[\s\S]*?<\/a:schemeClr>/u, `$1<a:srgbClr val="${hex}"/>`);
}
/**
 * Replace categories, series names and values of a template chart (cache + formulas), keeping its
 * formatting. data = {categories, series:[{name, values}]}; options: {sheet, seriesColors (for series
 * beyond the template's), categoryColors (per-point colours, chart6), recolor:{from,to}, minLabelSize}.
 */
export function fillChartXml(xml, data, { sheet = "Sheet1", seriesColors = [], categoryColors = null, recolorTo = null, minLabelSize = null, labelSize = null, dataLabelSize = null, fixedAxis = false, plotLayout = null } = {}) {
  const templateSeries = xml.match(/<c:ser>[\s\S]*?<\/c:ser>/gu) || [];
  if (!templateSeries.length) throw new Error("Template chart has no series.");
  const first = xml.indexOf(templateSeries[0]); const last = xml.lastIndexOf(templateSeries.at(-1)) + templateSeries.at(-1).length;
  const count = data.categories.length; const end = count + 1;
  const categoryRef = `<c:strRef><c:f>${sheet}!$A$2:$A$${Math.max(2, end)}</c:f>${strCache(data.categories)}</c:strRef>`;
  const series = data.series.map((entry, index) => {
    let body = templateSeries[Math.min(index, templateSeries.length - 1)];
    const column = columnName(index + 1);
    const formatCode = body.match(/<c:val>[\s\S]*?<c:formatCode>([^<]*)<\/c:formatCode>/u)?.[1] || "General";
    body = body.replace(/<c:idx val="\d+"\/>/u, `<c:idx val="${index}"/>`).replace(/<c:order val="\d+"\/>/u, `<c:order val="${index}"/>`);
    body = body.replace(/<c:tx>[\s\S]*?<\/c:tx>/u, `<c:tx><c:strRef><c:f>${sheet}!$${column}$1</c:f>${strCache([entry.name])}</c:strRef></c:tx>`);
    if (/<c:cat>/u.test(body)) body = body.replace(/<c:cat>[\s\S]*?<\/c:cat>/u, `<c:cat>${categoryRef}</c:cat>`);
    else body = body.replace(/<c:val>/u, `<c:cat>${categoryRef}</c:cat><c:val>`);
    body = body.replace(/<c:val>[\s\S]*?<\/c:val>/u, `<c:val><c:numRef><c:f>${sheet}!$${column}$2:$${column}$${Math.max(2, end)}</c:f>${numCache(entry.values.map(roundChartValue), formatCode)}</c:numRef></c:val>`);
    // Per-point formatting beyond the new category count is dropped (no orphan dPt/dLbl).
    body = body.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/gu, (point) => (Number(point.match(/<c:idx val="(\d+)"/u)?.[1]) >= count ? "" : point));
    body = body.replace(/<c:dLbl>[\s\S]*?<\/c:dLbl>/gu, (label) => (Number(label.match(/<c:idx val="(\d+)"/u)?.[1]) >= count ? "" : label));
    if (categoryColors) {
      const sample = body.match(/<c:dPt>[\s\S]*?<\/c:dPt>/u)?.[0];
      if (sample) {
        const points = data.categories.map((_, pointIndex) => sample.replace(/<c:idx val="\d+"\/>/u, `<c:idx val="${pointIndex}"/>`).replace(/<a:srgbClr val="[0-9A-F]{6}"/iu, `<a:srgbClr val="${categoryColors[pointIndex % categoryColors.length]}"`)).join("");
        body = body.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/gu, "").replace(/(<\/c:spPr>)(\s*<c:invertIfNegative\b[^>]*\/>)?/u, `$1$2${points}`);
      }
    }
    if (index >= templateSeries.length && seriesColors[index]) body = setSeriesFill(body, seriesColors[index]);
    if (recolorTo) body = recolor(body, seriesFillColor(body), recolorTo);
    return body;
  });
  let output = `${xml.slice(0, first)}${series.join("")}${xml.slice(last)}`;
  if (fixedAxis) output = output.replace(/(<c:valAx>[\s\S]*?<c:scaling>)([\s\S]*?)(<\/c:scaling>)/u, (whole, open, inner, close) => `${open}${(inner.match(/<c:orientation\b[^>]*\/>/u) || ['<c:orientation val="minMax"/>'])[0]}<c:max val="5"/><c:min val="1"/>${close}`);
  if (labelSize) output = output.replace(/(<c:catAx>[\s\S]*?<\/c:catAx>)/u, (axis) => axis.replace(/(<a:defRPr\b[^>]*\bsz=")(\d+)(")/u, (whole, head, size, tail) => `${head}${Math.min(Number(size), labelSize)}${tail}`));
  if (dataLabelSize) output = output.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/gu, (labels) => labels.replace(/(<a:defRPr\b[^>]*\bsz=")(\d+)(")/gu, (whole, head, size, tail) => `${head}${Math.min(Number(size), dataLabelSize)}${tail}`));
  if (plotLayout) output = setPlotLayout(output, plotLayout);
  if (minLabelSize) output = output.replace(/(<c:catAx>[\s\S]*?<\/c:catAx>)/u, (axis) => axis.replace(/(<a:defRPr\b[^>]*\bsz=")(\d+)(")/u, (whole, head, size, tail) => `${head}${Math.max(Number(size), minLabelSize)}${tail}`));
  return output;
}
export function chartWorkbookRows(data) {
  return [["", ...data.series.map((entry) => entry.name)], ...data.categories.map((category, index) => [category, ...data.series.map((entry) => roundChartValue(entry.values[index]))])];
}

// ---------- new shapes (fill map §1 „New”) ----------
export function newTextShape({ id, name, x, y, cx, cy, rPr, text: value, align = "l", anchor = "t", autofit = true }) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${NEW_SHAPE_PREFIX}${xmlEscape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="${anchor}">${autofit ? "<a:normAutofit/>" : "<a:noAutofit/>"}</a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${align}"/>${run(rPr, value)}</a:p></p:txBody></p:sp>`;
}
export function newRectShape({ id, name, x, y, cx, cy, fill }) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${NEW_SHAPE_PREFIX}${xmlEscape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="ro-RO"/></a:p></p:txBody></p:sp>`;
}
/** A paragraph style without its bullet (for a sentence written into a bulleted template box). */
export function withoutBullet(style) {
  let pPr = style.pPr || "<a:pPr/>";
  pPr = pPr.replace(/<a:bu(?:Font|Char|AutoNum|Clr|SzPct|SzPts|ClrTx|SzTx|FontTx|Blip)\b[^>]*\/>|<a:bu(?:Clr|Blip)\b[^>]*>[\s\S]*?<\/a:bu(?:Clr|Blip)>|<a:buNone\/>/gu, "");
  pPr = pPr.replace(/\s(?:marL|indent)="-?\d+"/gu, "").replace(/^<a:pPr\b/u, '<a:pPr marL="0" indent="0"');
  if (pPr.endsWith("/>")) pPr = pPr.replace(/\/>$/u, "><a:buNone/></a:pPr>");
  else { const at = pPr.search(/<a:(?:tabLst|defRPr|extLst)\b/u); pPr = at >= 0 ? `${pPr.slice(0, at)}<a:buNone/>${pPr.slice(at)}` : pPr.replace(/<\/a:pPr>$/u, "<a:buNone/></a:pPr>"); }
  return { ...style, pPr };
}
