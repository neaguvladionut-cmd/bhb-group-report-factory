import { reportPlan } from "./report-plan.js";

const PAGE = { width: 960, height: 540 };
const COLORS = {
  navy: "003057", ink: "14233b", aqua: "1aa6a8", green: "4a8c61", blue: "0a375b",
  pale: "eef7f7", line: "cbd9df", white: "ffffff", muted: "526579", dark: "082c4b"
};
const safeText = value => String(value ?? "").replace(/[\u0000-\u001f<>]/gu, " ").replace(/\s+/gu, " ").trim();
function decodeFontData(value) {
  if (value instanceof Uint8Array) return value;
  if (typeof value !== "string" || typeof atob !== "function") throw new Error("Datele fontului Noto Sans nu sunt disponibile");
  const binary = atob(value), bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}
const asNumber = value => Number.isFinite(Number(value)) ? Number(value) : 0;
const hex = (value, rgb) => { const code = String(value || "000000").replace("#", "").padEnd(6, "0"); return rgb(parseInt(code.slice(0, 2), 16) / 255, parseInt(code.slice(2, 4), 16) / 255, parseInt(code.slice(4, 6), 16) / 255); };

function wrapText(value, font, size, maxWidth) {
  const lines = [];
  for (const paragraph of String(value ?? "").split(/\r?\n/u)) {
    const words = safeText(paragraph).split(/\s+/u).filter(Boolean); let line = "";
    for (const word of words) {
      const candidate = line ? `${line} ${word}` : word;
      if (line && font.widthOfTextAtSize(candidate, size) > maxWidth) { lines.push(line); line = word; }
      else line = candidate;
    }
    lines.push(line);
  }
  return lines.length ? lines : [""];
}
function text(page, value, x, y, options) {
  const { font, size = 16, color, maxWidth = 800, lineGap = 5 } = options;
  const lines = wrapText(value, font, size, maxWidth);
  lines.forEach((line, index) => page.drawText(line, { x, y: y - index * (size + lineGap), size, font, color }));
  return y - lines.length * (size + lineGap);
}
function rect(page, x, y, width, height, color, opacity = 1) { page.drawRectangle({ x, y, width, height, color, opacity }); }
function line(page, x1, y1, x2, y2, color, thickness = 1) { page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, color, thickness }); }
function titleFor(slide) { return safeText(slide.title || slide.family); }

async function browserAsset(pdf, id) {
  if (typeof document === "undefined") return null;
  const image = document.getElementById(id);
  if (!image?.naturalWidth) return null;
  try {
    const canvas = document.createElement("canvas"); canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
    canvas.getContext("2d").drawImage(image, 0, 0);
    const bytes = Uint8Array.from(atob(canvas.toDataURL("image/png").split(",")[1]), char => char.charCodeAt(0));
    return pdf.embedPng(bytes);
  } catch { return null; }
}
function backgroundId(renderer, family) {
  if (renderer === "trend") return { cover: "trend-cover-frame-asset", methodology: "trend-methodology-frame-asset", divider: "trend-divider-frame-asset", range: "trend-range-frame-asset", ranking: "trend-ranking-frame-asset", benchmark: "trend-benchmark-frame-asset", "key-findings": "trend-key-findings-frame-asset", behavior: "trend-behavior-frame-asset", conclusions: "trend-conclusions-frame-asset", "appendix-divider": "trend-summary-divider-frame-asset", "benchmark-summary": "trend-benchmark-summary-frame-asset", "participant-comparison": "trend-participant-frame-asset", "competency-distribution": "trend-competency-frame-asset", close: "trend-closing-frame-asset" }[family] || "trend-divider-frame-asset";
  return { cover: "pptx-cover-asset", methodology: "pptx-content-asset", divider: "pptx-divider-asset", range: "pptx-report-background-asset", ranking: "pptx-report-background-asset", benchmark: "pptx-report-background-asset", "key-findings": "pptx-observations-background-asset", behavior: "pptx-behavior-background-asset", conclusions: "pptx-content-asset", "appendix-divider": "pptx-distribution-divider-asset", "benchmark-summary": "pptx-report-background-asset", "participant-comparison": "pptx-report-background-asset", "competency-distribution": "pptx-report-background-asset", close: "pptx-closing-asset" }[family] || "pptx-report-background-asset";
}
async function paintBackground(page, pdf, rgb, renderer, family) {
  const image = await browserAsset(pdf, backgroundId(renderer, family));
  if (image) page.drawImage(image, { x: 0, y: 0, width: PAGE.width, height: PAGE.height, opacity: .92 });
  else rect(page, 0, 0, PAGE.width, PAGE.height, hex(renderer === "trend" ? "f4f6f8" : "f7fafb", rgb));
}
function metadataLines(payload) {
  const metadata = payload.metadata || {};
  return [["Client", metadata.clientName], ["Data raportului", metadata.reportDate], ["Context", metadata.context]].filter(([, value]) => safeText(value));
}
function drawFooter(page, fonts, slide) {
  line(page, 48, 28, 912, 28, fonts.muted, 1);
  text(page, slide.deliverable === "appendix" ? "ANEXĂ · BHB Profiler" : "BHB Profiler · Raport de grup", 48, 14, { font: fonts.regular, size: 9, color: fonts.muted, maxWidth: 600 });
  text(page, `${slide.number} / ${slide.total}`, 850, 14, { font: fonts.bold, size: 9, color: fonts.muted, maxWidth: 60 });
}
function drawHeading(page, fonts, slide, renderer) {
  text(page, renderer === "trend" ? "TREND" : "BUSINESS HEALTH BAR", 48, 500, { font: fonts.bold, size: 10, color: renderer === "trend" ? fonts.blue : fonts.aqua, maxWidth: 400 });
  const title = titleFor(slide); let size = 25;
  while (fonts.bold.widthOfTextAtSize(title, size) > 820 && size > 15) size -= 1;
  text(page, title, 48, 464, { font: fonts.bold, size, color: fonts.navy, maxWidth: 820, lineGap: 3 });
}
function drawBars(page, fonts, items, x, y, width) {
  const rows = items.slice(0, 8); rows.forEach((item, index) => {
    const label = safeText(item.competency || item.name || item.title || "");
    const score = Math.max(0, Math.min(5, asNumber(item.mean ?? item.score)));
    const yy = y - index * 42;
    text(page, label, x, yy + 15, { font: fonts.regular, size: 11, color: fonts.ink, maxWidth: width * .48 });
    rect(page, x + width * .5, yy + 8, width * .45, 12, fonts.line);
    rect(page, x + width * .5, yy + 8, width * .45 * score / 5, 12, fonts.aqua);
    text(page, score ? score.toFixed(2) : "—", x + width * .96, yy + 11, { font: fonts.bold, size: 10, color: fonts.navy, maxWidth: 30 });
  });
}
function drawPageBody(page, pdf, fonts, slide, payload, renderer) {
  const family = slide.family;
  if (family === "cover") {
    rect(page, 0, 0, 650, PAGE.height, fonts.white, .96);
    if (renderer === "trend") rect(page, 0, 0, 650, PAGE.height, fonts.white, .8);
    text(page, renderer === "trend" ? "TREND" : "BUSINESS HEALTH BAR", 58, 474, { font: fonts.bold, size: 12, color: fonts.aqua, maxWidth: 500 });
    const project = payload.metadata?.projectName || slide.title || "Raport de grup", meta = metadataLines(payload);
    let titleSize = 38, metadataSize = 15;
    while ((wrapText(project, fonts.bold, titleSize, 550).length * (titleSize + 5)) + meta.reduce((sum, [, value]) => sum + wrapText(value, fonts.regular, metadataSize, 520).length * (metadataSize + 4) + 7, 0) > 350 && titleSize > 24) titleSize -= 2;
    while ((wrapText(project, fonts.bold, titleSize, 550).length * (titleSize + 5)) + meta.reduce((sum, [, value]) => sum + wrapText(value, fonts.regular, metadataSize, 520).length * (metadataSize + 4) + 7, 0) > 350 && metadataSize > 10) metadataSize -= 1;
    let y = text(page, project, 58, 420, { font: fonts.bold, size: titleSize, color: fonts.ink, maxWidth: 550, lineGap: 5 }) - 24;
    meta.forEach(([label, value]) => { y = text(page, `${label}: ${value}`, 58, y, { font: fonts.regular, size: metadataSize, color: fonts.ink, maxWidth: 520, lineGap: 4 }) - 7; });
    text(page, "CONFIDENȚIAL", 58, 48, { font: fonts.bold, size: 10, color: fonts.muted, maxWidth: 180 });
    return;
  }
  if (family === "divider" || family === "appendix-divider") {
    rect(page, 0, 0, PAGE.width, PAGE.height, fonts.navy);
    text(page, slide.title, 76, 350, { font: fonts.bold, size: 34, color: fonts.white, maxWidth: 700, lineGap: 6 });
    text(page, slide.deliverable === "appendix" ? "Rezultate detaliate pentru consultare" : "Rezultatele evaluării grupului", 78, 265, { font: fonts.regular, size: 17, color: fonts.white, maxWidth: 600 });
    return;
  }
  drawHeading(page, fonts, slide, renderer);
  if (family === "methodology") {
    const columns = slide.page || {};
    text(page, (columns.left || []).join("\n"), 58, 395, { font: fonts.regular, size: 14, color: fonts.ink, maxWidth: 360, lineGap: 9 });
    text(page, (columns.right || []).join("\n"), 520, 395, { font: fonts.regular, size: 14, color: fonts.ink, maxWidth: 360, lineGap: 9 });
  } else if (family === "executive-summary") {
    const summary = slide.summary || {};
    text(page, `Populație evaluată: ${summary.population || payload.participantCounts?.included || 0}`, 58, 402, { font: fonts.bold, size: 19, color: fonts.navy, maxWidth: 400 });
    text(page, `Cel mai puternic: ${summary.strongest?.competency || "—"}\nAtenție: ${summary.weakest?.competency || "—"}`, 58, 350, { font: fonts.regular, size: 15, color: fonts.ink, maxWidth: 370, lineGap: 9 });
    text(page, (summary.distribution || []).map(item => `${item.label}: ${item.value}%`).join("\n"), 520, 402, { font: fonts.regular, size: 15, color: fonts.ink, maxWidth: 350, lineGap: 9 });
    text(page, `Puncte forte: ${(summary.strengths || []).map(item => item.behavior).join("; ") || "Nu există observații suficiente."}\nArii de dezvoltare: ${(summary.development || []).map(item => item.behavior).join("; ") || "Nu există observații suficiente."}`, 58, 235, { font: fonts.regular, size: 14, color: fonts.ink, maxWidth: 800, lineGap: 8 });
  } else if (["range", "ranking"].includes(family)) drawBars(page, fonts, slide.items || [], 58, 400, 820);
  else if (family === "benchmark") {
    const bands = payload.bands || {}; const total = asNumber(bands.n) || 1;
    text(page, `Sub ${bands.low ?? "—"}: ${Math.round(asNumber(bands.below) / total * 100)}%`, 70, 360, { font: fonts.bold, size: 20, color: fonts.blue, maxWidth: 260 });
    text(page, `În interval: ${Math.round(asNumber(bands.typical) / total * 100)}%`, 350, 360, { font: fonts.bold, size: 20, color: fonts.aqua, maxWidth: 260 });
    text(page, `Peste ${bands.high ?? "—"}: ${Math.round(asNumber(bands.above) / total * 100)}%`, 630, 360, { font: fonts.bold, size: 20, color: fonts.green, maxWidth: 260 });
    rect(page, 70, 290, 790, 26, fonts.line); rect(page, 70, 290, 790 * asNumber(bands.below) / total, 26, fonts.blue); rect(page, 70 + 790 * asNumber(bands.below) / total, 290, 790 * asNumber(bands.typical) / total, 26, fonts.aqua); rect(page, 70 + 790 * (asNumber(bands.below) + asNumber(bands.typical)) / total, 290, 790 * asNumber(bands.above) / total, 26, fonts.green);
  } else if (family === "key-findings") {
    const finding = slide.item || {}, insight = slide.insight || {};
    text(page, `Media: ${asNumber(finding.mean).toFixed(2)} · Mediană: ${asNumber(finding.median).toFixed(2)}`, 58, 395, { font: fonts.bold, size: 18, color: fonts.navy, maxWidth: 500 });
    text(page, `Puncte forte\n${(insight.strengths || []).map(item => item.behavior).join("\n") || "—"}`, 58, 340, { font: fonts.regular, size: 14, color: fonts.ink, maxWidth: 365, lineGap: 8 });
    text(page, `Arii de dezvoltare\n${(insight.development || []).map(item => item.behavior).join("\n") || "—"}`, 520, 340, { font: fonts.regular, size: 14, color: fonts.ink, maxWidth: 365, lineGap: 8 });
  } else if (family === "behavior") {
    rect(page, 0, 0, PAGE.width, PAGE.height, fonts.dark);
    text(page, titleFor(slide), 58, 460, { font: fonts.bold, size: 24, color: fonts.white, maxWidth: 820 });
    const insight = slide.insight || {};
    text(page, `Puncte forte\n${(insight.strengths || []).map(item => `${item.behavior} · ${asNumber(item.mean).toFixed(2)}`).join("\n") || "—"}`, 58, 390, { font: fonts.regular, size: 15, color: fonts.white, maxWidth: 380, lineGap: 10 });
    text(page, `Atenție\n${(insight.development || []).map(item => `${item.behavior} · ${asNumber(item.mean).toFixed(2)}`).join("\n") || "—"}`, 510, 390, { font: fonts.regular, size: 15, color: fonts.white, maxWidth: 380, lineGap: 10 });
  } else if (family === "conclusions") {
    text(page, slide.copy || payload.metadata?.conclusions || "", 58, 395, { font: fonts.regular, size: 18, color: fonts.ink, maxWidth: 820, lineGap: 10 });
  } else if (family === "benchmark-summary") {
    drawBars(page, fonts, slide.items || [], 58, 400, 820);
  } else if (family === "participant-comparison") {
    const competencies = (slide.ranked || payload.calculations || []).slice(0, 5).map(item => item.competency);
    const colX = [58, 270, 440, 610, 780, 900];
    text(page, "Participant", colX[0], 420, { font: fonts.bold, size: 11, color: fonts.navy, maxWidth: 190 });
    competencies.forEach((competency, column) => text(page, competency, colX[column + 1], 420, { font: fonts.bold, size: 10, color: fonts.navy, maxWidth: 150 }));
    (slide.items || []).slice(0, 6).forEach((item, index) => { const y = 385 - index * 48; if (index % 2 === 0) rect(page, 48, y - 14, 864, 30, fonts.pale); text(page, item.name, colX[0], y, { font: fonts.bold, size: 11, color: fonts.ink, maxWidth: 190 }); competencies.forEach((competency, column) => { const value = item.scores?.[competency]; text(page, Number.isFinite(value) ? Number(value).toFixed(2) : "—", colX[column + 1], y, { font: fonts.regular, size: 11, color: fonts.ink, maxWidth: 150 }); }); });
  } else if (family === "competency-distribution") {
    const competency = slide.item?.competency || "";
    (slide.items || []).slice(0, 8).forEach((item, index) => { const y = 390 - index * 38, score = asNumber(item.scores?.[competency]); text(page, item.name, 58, y, { font: fonts.regular, size: 11, color: fonts.ink, maxWidth: 260 }); rect(page, 340, y - 1, 470, 12, fonts.line); rect(page, 340, y - 1, 470 * score / 5, 12, fonts.aqua); text(page, score.toFixed(2), 830, y, { font: fonts.bold, size: 11, color: fonts.navy, maxWidth: 50 }); });
  } else if (family === "close") {
    text(page, renderer === "trend" ? "Raport TREND pregătit" : "Raport BHB pregătit", 58, 390, { font: fonts.bold, size: 24, color: fonts.navy, maxWidth: 700 });
    text(page, "Fișierul conține raportul principal și anexa selectată pentru livrare.", 58, 335, { font: fonts.regular, size: 17, color: fonts.ink, maxWidth: 700, lineGap: 8 });
  }
  else text(page, "Rezultat disponibil în raportul combinat.", 58, 390, { font: fonts.regular, size: 18, color: fonts.ink, maxWidth: 820 });
}

export async function createPdf(payload, { scope = "whole", renderer = payload?.metadata?.renderer || "bhb" } = {}) {
  const lib = globalThis.PDFLib;
  if (!lib?.PDFDocument) throw new Error("Biblioteca PDF locală nu este disponibilă");
  const fontkit = globalThis.fontkit, fontData = globalThis.__grfFontData;
  if (!fontkit?.create || !fontData?.regular || !fontData?.bold) throw new Error("Fonturile Noto Sans pentru diacritice nu sunt disponibile");
  const { PDFDocument, rgb } = lib;
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  const fonts = { regular: await pdf.embedFont(decodeFontData(fontData.regular), { subset: true }), bold: await pdf.embedFont(decodeFontData(fontData.bold), { subset: true }), white: hex(COLORS.white, rgb), navy: hex(COLORS.navy, rgb), ink: hex(COLORS.ink, rgb), aqua: hex(COLORS.aqua, rgb), blue: hex(COLORS.blue, rgb), green: hex(COLORS.green, rgb), line: hex(COLORS.line, rgb), muted: hex(COLORS.muted, rgb), dark: hex(COLORS.dark, rgb) };
  const plan = reportPlan(payload, { scope });
  for (const slide of plan) { const page = pdf.addPage(); page.setSize(PAGE.width, PAGE.height); await paintBackground(page, pdf, rgb, renderer, slide.family); drawPageBody(page, pdf, fonts, slide, payload, renderer); if (!['cover', 'divider', 'appendix-divider', 'behavior'].includes(slide.family)) drawFooter(page, fonts, slide); }
  return pdf.save({ useObjectStreams: false });
}

export const pdfPageSize = () => ({ ...PAGE });
