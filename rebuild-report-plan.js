// BP-GRF-R report plan. One plan drives the preview, the Trend deck and the bundle (fill map §1).
// Every item names the template slide it is cloned from (`templateIndex`) and carries the data the
// fill layer writes into that slide's named shapes; nothing here is rendered free-form.

const text = (value) => String(value ?? "").trim();
export const f2 = (value) => value !== null && value !== "" && Number.isFinite(Number(value)) ? Number(value).toFixed(2) : "";
export const pct = (count, total) => total ? Math.round(count / total * 100) : 0;
const average = (values) => values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
export const medianOf = (values) => {
  const sorted = values.filter(Number.isFinite).slice().sort((a, b) => a - b);
  if (!sorted.length) return null;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};
export const overallMean = (record) => average(Object.values(record.scores || {}).filter(Number.isFinite));
const band = (value, low, high) => value > high ? "above" : value < low ? "below" : "in";

export const TEMPLATE_SLIDES = {
  cover: 1, "how-to-read": 2, methodology: 2, "executive-summary": 21, "key-findings": 12,
  "divider-results": 3, range: 4, ranking: 6, benchmark: 7, population: 8, zone: 10,
  "divider-behaviors": 18, behavior: 19, "divider-conclusions": 20, conclusions: 21,
  "appendix-divider": 11, "participant-mean": 5, "participant-comparison": 9, "competency-participants": 13, close: 22
};
// Fixed template titles (the template's own wording, whitespace-normalised).
export const TEMPLATE_TITLES = {
  "divider-results": "Distribuția rezultatelor",
  range: "Distribuția rezultatelor – pe competențe (mediană)",
  ranking: "Distribuția rezultatelor – media pe competențe",
  benchmark: "Distribuția pe media competențelor",
  population: "Distribuția pe media competențelor – toată populația",
  zone: "Distribuția rezultatelor – pe competențe/ pe zone",
  "divider-behaviors": "Puncte forte și recomandări de grup",
  "divider-conclusions": "Concluzii și Recomandări",
  conclusions: "Concluzii și recomandări",
  "participant-mean": "Distribuția rezultatelor – media pe competențe",
  "participant-comparison": "Distribuția rezultatelor – pe competențe/ per participant",
  close: "MULȚUMIM!"
};
// Families that are a chart or a table: one bundle item each (fill map §2).
export const BUNDLE_FAMILIES = new Set(["key-findings", "range", "ranking", "benchmark", "population", "zone", "behavior", "participant-mean", "participant-comparison", "competency-participants"]);

// Rule 10: readable participants per slide, split into the fewest slides, sizes differing by at most one.
export const PARTICIPANT_START_CAPS = { "participant-mean": 25, "participant-comparison": 5, "competency-participants": 20 };
// A participant label wraps at spaces and hyphens onto at most four lines: it needs the width of its longest
// segment, and at least a quarter of the whole name (F34).
export const labelChars = (name) => Math.max(...String(name).split(/[\s-]+/u).map((part) => part.length + 1), Math.ceil(String(name).length / 4));
export function readableCapacity(family, { seriesCount = 1, longestLabel = 0 } = {}) {
  if (family === "participant-mean") { // horizontal bars on a plot ≈ 8.0 in tall, gap 150 %: label ≥ 10 pt, bar ≥ 0.12 in
    const slot = Math.max(10 * 1.25 / 72, 0.12 * 2.5);
    return Math.max(1, Math.floor(8.0 / slot));
  }
  const plotWidth = 17.2; // column charts t9 / t13–t17, gap 219 %, overlap −27 %
  const barSlot = 0.12 * (seriesCount * 1.27 - 0.27 + 2.19);
  const labelSlot = Math.max(1, longestLabel) * 10 * 0.55 / 72; // longestLabel = labelChars(): a 10 pt label wraps onto ≤ 4 lines
  return Math.max(1, Math.floor(plotWidth / Math.max(barSlot, labelSlot)));
}
export function participantsPerSlide(family, options = {}) { return Math.max(1, Math.min(PARTICIPANT_START_CAPS[family] || 20, readableCapacity(family, options))); }
export function splitEqual(items, cap) {
  if (!items.length) return [];
  const pages = Math.ceil(items.length / Math.max(1, cap));
  const base = Math.floor(items.length / pages); const extra = items.length % pages;
  const output = []; let start = 0;
  for (let page = 0; page < pages; page += 1) { const size = base + (page < extra ? 1 : 0); output.push(items.slice(start, start + size)); start += size; }
  return output;
}
export const pageGroups = (items, size) => items.length ? splitEqual(items, size) : [[]];
export const participantChartPageSize = () => PARTICIPANT_START_CAPS["competency-participants"];
export const participantComparisonPageSize = PARTICIPANT_START_CAPS["participant-comparison"];

// R5 as amended by D313: rank by the sum of present scores; ties by import order.
export function rankBehaviors(rows = []) {
  const groups = new Map();
  rows.forEach((row, index) => {
    if (!groups.has(row.competency)) groups.set(row.competency, []);
    groups.get(row.competency).push({ ...row, sourceIndex: row.sourceIndex ?? index });
  });
  return [...groups.entries()].map(([competency, items]) => {
    const ranked = items.slice().sort((a, b) => (b.sum ?? 0) - (a.sum ?? 0) || a.sourceIndex - b.sourceIndex);
    let keyCount = 0; let developmentCount = 0; let medianBehavior = null;
    if (ranked.length >= 10) { keyCount = 5; developmentCount = 5; }
    else if (ranked.length >= 6) { keyCount = 3; developmentCount = 3; }
    else {
      // Below 6 (Vlad, 2026-10-01): an even count splits exactly in half; an odd count skips the middle behaviour.
      const half = Math.floor(ranked.length / 2);
      medianBehavior = ranked.length % 2 ? ranked[half] : null;
      keyCount = half;
      developmentCount = half;
    }
    const key = ranked.slice(0, keyCount);
    const development = ranked.slice(ranked.length - developmentCount);
    return { competency, ranked, key, development, strengths: key, median: medianBehavior, all: ranked };
  });
}
export const behaviorInsights = (rows) => rankBehaviors(rows).map((item) => ({ ...item, strengths: item.key, development: item.development }));
export const behaviorPageGroups = (rows) => pageGroups(rows, 6);

const range = (low, high) => `${f2(low)}–${f2(high)}`;
// §10a, approved standard text; the last sentence follows the annex setting.
// F26: a sentence that explains a chart type is left out when that chart type is toggled off.
export function howToReadParagraphs(low = 2.75, high = 3.5, annex = "end", toggles = {}) {
  const on = (key) => toggles[key] !== false;
  const last = annex === "none" ? "Rezultatele descriu grupul evaluat." : `Rezultatele descriu grupul evaluat. Rezultatele individuale se regăsesc ${annex === "separate" ? "în anexa transmisă separat" : "în anexă"}.`;
  const second = ["Media arată nivelul general al grupului; mediana este scorul participantului aflat la mijlocul grupului și este mai puțin influențată de rezultatele extreme.", on("range") ? "Graficele de distribuție arată, pentru fiecare competență, cel mai mic și cel mai mare scor obținut, mediana și intervalele în care se situează jumătatea superioară și cea inferioară a participanților." : ""].filter(Boolean).join(" ");
  const keyFindings = on("keyFindings") && on("observation");
  const third = keyFindings || on("behavior") ? ["Abilitățile cheie sunt comportamentele cel mai bine demonstrate în cadrul fiecărei competențe; abilitățile de dezvoltat sunt cele mai puțin demonstrate.", keyFindings ? "Procentele indică ponderea participanților care au demonstrat pe deplin comportamentul, respectiv care nu l-au demonstrat." : ""].filter(Boolean).join(" ") : "";
  return [
    `Rezultatele pe competențe sunt exprimate pe o scală de la 1 la 5, unde 1 reprezintă nivelul minim, iar 5 nivelul maxim. Banda gri din grafice marchează intervalul de referință (benchmark) de ${range(low, high)}, care corespunde unei performanțe la nivel mediu în evaluările TREND: rezultatele din bandă sunt la nivel mediu, cele de deasupra ei peste medie, iar cele de dedesubt sub medie.`,
    second, third, last
  ].filter(Boolean);
}
export const HOW_TO_READ = (low = 2.75, high = 3.5, annex = "end") => howToReadParagraphs(low, high, annex).join(" ");

// Methodology (feedback text). Counts the data knows are prefilled; the rest come from the builder.
export const METHODOLOGY_PRINCIPLES = [
  "Competențele evaluate au fost operaționalizate în comportamente observabile, iar evaluatorii au utilizat aceleași grile de evaluare. Scorurile se bazează pe evidențe comportamentale observate în cadrul exercițiilor, nu pe impresii generale.",
  "Fiecare comportament a fost evaluat în etapa de analiză a competențelor. Evaluatorii și-au calibrat observațiile pe baza evidențelor colectate, iar rezultatele au fost ulterior integrate în Business Health Bar Profiler.",
  "Rezultatele agregate la nivel de competență au fost exprimate pe o scală de la 1 la 5, unde 1 reprezintă nivelul minim, iar 5 nivelul maxim."
];
export function methodologyColumns(payload) {
  const metadata = payload.metadata || {};
  const schemas = payload.schemas || [];
  const evaluators = new Set(schemas.flatMap((schema) => schema.methodology?.evaluators || []));
  const dates = new Set(schemas.flatMap((schema) => schema.methodology?.dates || []));
  const teamSizes = schemas.flatMap((schema) => schema.methodology?.teamSizes || []);
  const consultants = text(metadata.evaluators) || String(evaluators.size || teamSizes.filter((value) => value >= 2).sort((a, b) => b - a)[0] || "");
  const days = text(metadata.days) || String(dates.size || "");
  const exerciseList = text(metadata.exercises);
  const exerciseCount = text(metadata.exerciseCount) || (exerciseList ? String(exerciseList.split(/[;,]\s*|\n/u).filter((item) => text(item)).length) : "");
  const client = text(metadata.clientName || metadata.projectName);
  const facts = [
    { number: String(payload.participantCounts?.included ?? ""), text: `participanți${client ? ` ${client}` : ""}` },
    { number: consultants, text: "consultanți TREND implicați" },
    { number: days, text: days === "1" ? "zi de evaluare" : "zile de evaluare" },
    { number: String((payload.behaviorAggregates || []).length), text: "comportamente specifice observate" },
    { number: exerciseCount, text: `exerciții concepute pentru a evidenția nivelul competențelor evaluate${exerciseList ? `: ${exerciseList}` : ""}` }
  ];
  // A fact the consultant left empty is omitted, never printed as a number-less sentence (ruling 2026-10-01).
  const missing = { evaluators: !consultants, days: !days, exercises: !exerciseCount };
  for (let index = facts.length - 1; index >= 1; index -= 1) if (!facts[index].number) facts.splice(index, 1);
  if (text(metadata.otherInstruments)) facts.push({ number: "", text: `Alte instrumente folosite: ${text(metadata.otherInstruments)}` });
  const principles = text(metadata.methodologyText) ? text(metadata.methodologyText).split(/\n+/u).map(text).filter(Boolean) : METHODOLOGY_PRINCIPLES;
  const line = (fact) => `${fact.number ? `${fact.number} ` : ""}${fact.text}`;
  return {
    facts, principles,
    left: ["METODOLOGIE", ...facts.map(line)], right: ["PROCESUL DE EVALUARE", ...principles],
    missing: { ...missing, otherInstruments: !text(metadata.otherInstruments) },
    missingLabels: [missing.evaluators && "consultanți TREND implicați", missing.days && "zile de evaluare", missing.exercises && "număr de exerciții"].filter(Boolean)
  };
}
export const methodologyPages = (payload) => [{ page: methodologyColumns(payload) }];

const scoredCalculations = (view) => view.calculations.filter((item) => item.mean !== null);
const byMeanDesc = (items) => items.slice().sort((a, b) => b.mean - a.mean);

export function executiveSummary(payload) {
  const bands = payload.bands || { below: 0, typical: 0, above: 0, n: 0, low: 2.75, high: 3.5 };
  const sorted = byMeanDesc(scoredCalculations(payload));
  const shares = { below: pct(bands.below, bands.n), in: pct(bands.typical, bands.n), above: pct(bands.above, bands.n) };
  const n = payload.participantCounts.included;
  return {
    population: n, shares, low: bands.low, high: bands.high,
    distribution: [{ label: `Sub ${f2(bands.low)}`, value: shares.below }, { label: `În intervalul ${range(bands.low, bands.high)}`, value: shares.in }, { label: `Peste ${f2(bands.high)}`, value: shares.above }],
    sentence: `Evaluarea celor ${n} participanți: ${shares.in}% dintre participanți se încadrează în intervalul benchmarkului (${range(bands.low, bands.high)}), ${shares.above}% îl depășesc, iar ${shares.below}% se situează sub nivelul său inferior.`,
    strongest: sorted[0] || null, weakest: sorted.at(-1) || null,
    competencyLines: sorted.length ? [`Cel mai bine reprezentată: ${sorted[0].competency} (${f2(sorted[0].mean)})`, `Principala oportunitate: ${sorted.at(-1).competency} (${f2(sorted.at(-1).mean)})`] : [],
    conclusions: text(payload.metadata?.executiveConclusions ?? payload.metadata?.conclusions)
  };
}

const behaviourLine = (row, field, share) => `${text(row[field]) || row.behavior} (${Math.round((share || 0) * 100)}%)`;
export function competencyFindings(payload) {
  const insights = new Map(behaviorInsights(payload.behaviorAggregates || []).map((item) => [item.competency, item]));
  const { low, high } = payload.bands;
  return byMeanDesc(scoredCalculations(payload)).map((item) => {
    const insight = insights.get(item.competency) || { key: [], development: [] };
    const scores = payload.records.map((record) => record.scores?.[item.competency]).filter(Number.isFinite).sort((a, b) => b - a);
    const counts = { above: 0, in: 0, below: 0 }; scores.forEach((score) => { counts[band(score, low, high)] += 1; });
    return {
      item, insight, competency: item.competency, mean: item.mean, median: medianOf(scores), scores, counts,
      strengths: insight.key.map((row) => behaviourLine(row, "score2", row.pct2)),
      development: insight.development.map((row) => behaviourLine(row, "score0", row.pct0))
    };
  });
}

function groupBehaviorAggregates(payload, identities) {
  const grouped = new Map();
  payload.behaviorRecords.filter((row) => identities.has(row.identity)).forEach((row) => {
    const key = `${row.competency}\u0000${row.behavior}`;
    if (!grouped.has(key)) grouped.set(key, { competency: row.competency, subcompetency: row.subcompetency, behavior: row.behavior, values: [], sourceIndex: row.sourceIndex });
    grouped.get(key).values.push(row.score);
  });
  return [...grouped.values()].map((row) => {
    const descriptor = payload.behaviorAggregates.find((candidate) => candidate.competency === row.competency && candidate.behavior === row.behavior) || {};
    return { competency: row.competency, subcompetency: row.subcompetency, behavior: row.behavior, n: row.values.length, sum: row.values.reduce((sum, value) => sum + value, 0), mean: average(row.values), pct0: row.values.length ? row.values.filter((value) => value === 0).length / row.values.length : 0, pct2: row.values.length ? row.values.filter((value) => value === 2).length / row.values.length : 0, score0: descriptor.score0 || "", score2: descriptor.score2 || "", sourceIndex: row.sourceIndex };
  });
}
function scoreStats(records, competency) {
  const values = records.map((record) => record.scores[competency]).filter(Number.isFinite);
  return { competency, n: values.length, mean: average(values), min: values.length ? Math.min(...values) : null, median: medianOf(values), max: values.length ? Math.max(...values) : null };
}
export function viewForGroup(payload, group) {
  const identities = new Set(group.records.map((record) => record.identity));
  const records = payload.records.filter((record) => identities.has(record.identity));
  const competencies = payload.competencies || payload.calculations.map((item) => item.competency);
  const calculations = competencies.map((competency) => scoreStats(records, competency));
  const overallScores = records.map(overallMean).filter(Number.isFinite);
  const bands = { ...payload.bands, below: overallScores.filter((score) => score < payload.bands.low).length, typical: overallScores.filter((score) => score >= payload.bands.low && score <= payload.bands.high).length, above: overallScores.filter((score) => score > payload.bands.high).length, n: overallScores.length };
  const zones = [...new Set(records.map((record) => record.region).filter(Boolean))].map((region) => ({ region, records: records.filter((record) => record.region === region) }));
  return { ...payload, records, participantCounts: { ...payload.participantCounts, included: records.length }, calculations, bands, behaviorAggregates: groupBehaviorAggregates(payload, identities), zones, zoneCalculations: zones.flatMap((zone) => competencies.map((competency) => ({ region: zone.region, ...scoreStats(zone.records, competency) }))), regionReadiness: { available: zones.length, blank: records.filter((record) => !record.region).length, disabledReason: zones.length ? "" : "Nu există valori de regiune în exportul detaliat." } };
}

function benchmarkTable(view) {
  const { low, high } = view.bands;
  const rows = view.records.map(overallMean).filter(Number.isFinite).sort((a, b) => b - a).map((value) => ({ value, band: band(value, low, high) }));
  const counts = { above: 0, in: 0, below: 0 }; rows.forEach((row) => { counts[row.band] += 1; });
  const shares = { above: pct(counts.above, rows.length), in: pct(counts.in, rows.length), below: pct(counts.below, rows.length) };
  return { rows, counts, shares, low, high, competencyMeans: byMeanDesc(scoredCalculations(view)).map((item) => ({ competency: item.competency, mean: item.mean })) };
}
function populationRows(view) {
  const { low, high } = view.bands;
  return byMeanDesc(scoredCalculations(view)).map((item) => {
    const scores = view.records.map((record) => record.scores?.[item.competency]).filter(Number.isFinite);
    const counts = { below: 0, in: 0, above: 0 }; scores.forEach((score) => { counts[band(score, low, high)] += 1; });
    return { competency: item.competency, below: pct(counts.below, scores.length), in: pct(counts.in, scores.length), above: pct(counts.above, scores.length), n: scores.length };
  });
}

function addTemplateSection(slides, view, groupKey = "") {
  const groupLabel = groupKey ? text(view.metadata.groupNames?.[groupKey]) || groupKey : "";
  const suffix = groupKey ? ` · ${groupLabel}` : "";
  const toggles = view.metadata.slideToggles || {};
  const on = (key) => toggles[key] !== false;
  const add = (family, data = {}) => slides.push({ family, templateIndex: TEMPLATE_SLIDES[family], title: `${TEMPLATE_TITLES[family] || ""}${suffix}`, groupKey, groupLabel, suffix, deliverable: "main", ...data });
  const calculations = byMeanDesc(scoredCalculations(view));
  const results = [];
  if (on("range")) results.push(["range", { items: calculations }]);
  if (on("competencyMean")) results.push(["ranking", { items: calculations.slice().reverse() }]);
  if (on("benchmark")) results.push(["benchmark", { table: benchmarkTable(view) }]);
  if (on("competencyDistribution")) results.push(["population", { rows: populationRows(view) }]);
  if (on("zone") && view.regionReadiness?.available) {
    const regions = [...new Set((view.zoneCalculations || []).map((entry) => entry.region))];
    results.push(["zone", { regions, competencies: calculations.map((item) => item.competency), values: regions.map((region) => calculations.map((item) => (view.zoneCalculations || []).find((entry) => entry.region === region && entry.competency === item.competency)?.mean ?? null)) }]);
  }
  if (results.length) { add("divider-results"); results.forEach(([family, data]) => add(family, { ...data, low: view.bands.low, high: view.bands.high })); }
  const insights = on("behavior") ? behaviorInsights(view.behaviorAggregates).filter((insight) => insight.key.length || insight.development.length) : [];
  if (insights.length) {
    add("divider-behaviors");
    const order = new Map(calculations.map((item, index) => [item.competency, index]));
    insights.sort((a, b) => (order.get(a.competency) ?? 99) - (order.get(b.competency) ?? 99)).forEach((insight) => add("behavior", { title: `Comportamente cheie – ${insight.competency}${suffix}`, competency: insight.competency, key: insight.key.map((row) => text(row.score2) || row.behavior), development: insight.development.map((row) => text(row.score0) || row.behavior) }));
  }
  if (on("conclusions")) {
    add("divider-conclusions");
    // F32: a group's conclusions slide uses that group's own fields; the whole project uses the project fields.
    const own = groupKey ? view.metadata.groupConclusions?.[groupKey] || {} : { strengths: view.metadata.conclusionsStrengths, development: view.metadata.conclusionsDevelopment, interventions: view.metadata.conclusionsInterventions };
    add("conclusions", { strengths: text(own.strengths), development: text(own.development), interventions: text(own.interventions) });
  }
}

function appendixSlides(payload) {
  const slides = [];
  const add = (family, data = {}) => slides.push({ family, templateIndex: TEMPLATE_SLIDES[family], title: TEMPLATE_TITLES[family] || "", groupKey: "", deliverable: "appendix", low: payload.bands.low, high: payload.bands.high, ...data });
  const competencies = (payload.competencies || payload.calculations.map((item) => item.competency)).filter((competency) => payload.records.some((record) => Number.isFinite(record.scores?.[competency])));
  const participants = payload.records.map((record) => ({ name: record.name, mean: overallMean(record), scores: record.scores || {} })).filter((row) => Number.isFinite(row.mean)).sort((a, b) => b.mean - a.mean);
  const longestLabel = Math.max(0, ...participants.map((row) => labelChars(row.name)));
  add("appendix-divider", { title: "Anexă – rezultate individuale", heading: "Anexă", subheading: "rezultate individuale" });
  const meanPages = splitEqual(participants, participantsPerSlide("participant-mean", { longestLabel }));
  meanPages.forEach((rows, page) => add("participant-mean", { rows, page: page + 1, pages: meanPages.length }));
  const comparisonPages = splitEqual(participants, participantsPerSlide("participant-comparison", { seriesCount: competencies.length, longestLabel }));
  comparisonPages.forEach((rows, page) => add("participant-comparison", { rows, competencies, page: page + 1, pages: comparisonPages.length }));
  competencies.forEach((competency, competencyIndex) => {
    const rows = participants.filter((row) => Number.isFinite(row.scores[competency])).map((row) => ({ name: row.name, score: row.scores[competency] })).sort((a, b) => b.score - a.score);
    const pages = splitEqual(rows, participantsPerSlide("competency-participants", { longestLabel }));
    pages.forEach((pageRows, page) => add("competency-participants", { title: `Distribuția pe competențe – ${competency}`, competency, competencyIndex, templateIndex: competencyIndex < 5 ? 13 + competencyIndex : 13, recolor: competencyIndex >= 5 ? "FF9D75" : "", rows: pageRows, page: page + 1, pages: pages.length }));
  });
  // F34: one label size per annex series — the smallest per-page fit (≤ 20 pt, ≥ 10 pt).
  for (const family of ["participant-comparison", "competency-participants"]) {
    const pages = slides.filter((slide) => slide.family === family);
    const fit = (page) => { const slot = 17.2 * 72 / Math.max(1, page.rows.length) * 0.9; const need = Math.max(1, ...page.rows.map((row) => labelChars(row.name))); return Math.max(1000, Math.min(2000, Math.floor(slot / (need * 0.55)) * 100)); };
    const size = Math.min(2000, ...pages.map(fit));
    pages.forEach((page) => { page.labelSize = size; });
  }
  return slides;
}

export function reportPlan(payload, { scope = "whole" } = {}) {
  const slides = [];
  const metadata = payload.metadata || {};
  const annex = metadata.annex || "end";
  const toggles = metadata.slideToggles || {};
  const client = text(metadata.clientName || metadata.projectName);
  const program = text(metadata.program) || "Centru de Dezvoltare";
  const year = text(metadata.reportDate).match(/\b(\d{4})\b/u)?.[1] || String(new Date().getFullYear());
  const add = (family, data = {}) => slides.push({ family, templateIndex: TEMPLATE_SLIDES[family], title: TEMPLATE_TITLES[family] || "", groupKey: "", deliverable: "main", ...data });
  add("cover", { title: `${client} – ${program}`, client, program, year });
  add("how-to-read", { title: "CUM CITIM ACEST RAPORT", paragraphs: howToReadParagraphs(payload.bands.low, payload.bands.high, annex, toggles) });
  add("methodology", { title: "PRIVIRE DE ANSAMBLU ASUPRA PROIECTULUI - METODOLOGIE", page: methodologyColumns(payload) });
  add("executive-summary", { title: "Executive Summary", summary: executiveSummary(payload) });
  if (toggles.keyFindings !== false && toggles.observation !== false) competencyFindings(payload).forEach((finding) => add("key-findings", { ...finding, title: `Distribuția pe competențe – ${finding.competency}`, low: payload.bands.low, high: payload.bands.high }));
  addTemplateSection(slides, payload);
  if (metadata.splitGroups && (payload.groups || []).length >= 2) payload.groups.forEach((group) => addTemplateSection(slides, viewForGroup(payload, group), group.code));
  const appendix = annex === "none" ? [] : appendixSlides(payload);
  const close = { family: "close", templateIndex: 22, title: "MULȚUMIM!", groupKey: "", deliverable: "main" };
  let selected;
  if (scope === "main") selected = [...slides, close];
  else if (scope === "appendix") selected = [{ ...slides[0], annexMark: true, deliverable: "appendix" }, ...appendix, { ...close, deliverable: "appendix" }];
  else selected = [...slides, ...appendix, close];
  return selected.map((slide, index) => ({ ...slide, number: index + 1, total: selected.length }));
}
