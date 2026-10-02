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
  "appendix-divider": 11, "divider-observations": 11, "participant-mean": 5, "participant-comparison": 9, "competency-participants": 13, close: 22
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
  "divider-observations": "Analiza observațiilor – pe competențe",
  "participant-comparison": "Distribuția rezultatelor – pe competențe/ per participant",
  close: "MULȚUMIM!"
};
// Families that are a chart or a table: one bundle item each (fill map §2).
export const BUNDLE_FAMILIES = new Set(["key-findings", "range", "ranking", "benchmark", "population", "zone", "behavior", "participant-mean", "participant-comparison", "competency-participants"]);

// Rule 10: readable participants per slide, split into the fewest slides, sizes differing by at most one.
// Vlad 2026-10-02: ~11 participants per annex slide, labels ≥ 12 pt (A3 keeps its 5 groups of bars).
export const PARTICIPANT_START_CAPS = { "participant-mean": 11, "participant-comparison": 5, "competency-participants": 11 };
export const MIN_PARTICIPANT_LABEL_PT = 12;
// A participant label wraps at spaces and hyphens onto at most four lines: it needs the width of its longest
// segment, and at least a quarter of the whole name (F34).
export const A2_LABEL_COLUMN_IN = 4.3;
/** Lines a participant name takes at `size` pt in `width` pt, breaking at spaces and after hyphens. */
export function wrapLabel(name, size, width) {
  const parts = String(name).split(/(?<=-)|\s+/u).filter(Boolean); const charWidth = size * 0.55; let lines = 1; let line = 0;
  for (const part of parts) { const w = part.length * charWidth; if (line && line + w > width) { lines += 1; line = w; } else line += w + (part.endsWith("-") ? 0 : charWidth); while (line > width + charWidth) { lines += 1; line -= width; } }
  return lines;
}
export const labelChars = (name) => Math.max(...String(name).split(/[\s-]+/u).map((part) => part.length + 1), Math.ceil(String(name).length / 4));
export function readableCapacity(family, { seriesCount = 1, longestLabel = 0 } = {}) {
  if (family === "participant-mean") { // horizontal bars on a plot ≈ 8.0 in tall, gap 150 %: label ≥ 10 pt, bar ≥ 0.12 in
    const slot = Math.max(10 * 1.25 / 72, 0.12 * 2.5);
    return Math.max(1, Math.floor(8.0 / slot));
  }
  const plotWidth = 17.2; // column charts t9 / t13–t17, gap 219 %, overlap −27 %
  const barSlot = 0.12 * (seriesCount * 1.27 - 0.27 + 2.19);
  const labelSlot = Math.max(1, longestLabel) * MIN_PARTICIPANT_LABEL_PT * 0.55 / 72; // longestLabel = labelChars(): a 12 pt label wraps onto ≤ 4 lines
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

// R5 (Vlad 2026-10-02): each behaviour is ranked by its MEAN score (0–2) over the participants who have a score for it
// (missing never counts; an all-unscored behaviour has no aggregate and stays out of the ranking and the count).
// Top list: mean desc, then share scoring 2 desc, then column order. Bottom list: mean asc, then share scoring 0
// desc, then column order. Cuts: ≥ 10 → 5 + 5, 6–9 → 3 + 3, < 6 → even halves, odd middle skipped. No behaviour
// appears in both lists.
const meanOf = (row) => (Number.isFinite(row.mean) ? row.mean : row.sum ?? 0);
export const topOrder = (a, b) => meanOf(b) - meanOf(a) || (b.pct2 ?? 0) - (a.pct2 ?? 0) || a.sourceIndex - b.sourceIndex;
export const bottomOrder = (a, b) => meanOf(a) - meanOf(b) || (b.pct0 ?? 0) - (a.pct0 ?? 0) || a.sourceIndex - b.sourceIndex;
export function rankBehaviors(rows = []) {
  const groups = new Map();
  rows.forEach((row, index) => {
    if (!groups.has(row.competency)) groups.set(row.competency, []);
    groups.get(row.competency).push({ ...row, sourceIndex: row.sourceIndex ?? index });
  });
  return [...groups.entries()].map(([competency, items]) => {
    const ranked = items.slice().sort(topOrder);
    const count = ranked.length < 6 ? Math.floor(ranked.length / 2) : ranked.length >= 10 ? 5 : 3;
    const key = ranked.slice(0, count);
    const development = items.slice().sort(bottomOrder).filter((row) => !key.includes(row)).slice(0, count);
    const leftover = ranked.filter((row) => !key.includes(row) && !development.includes(row));
    return { competency, ranked, key, development, strengths: key, median: ranked.length < 6 && ranked.length % 2 ? leftover[0] || null : null, all: ranked };
  });
}
export const behaviorInsights = (rows) => rankBehaviors(rows).map((item) => ({ ...item, strengths: item.key, development: item.development }));
export const behaviorPageGroups = (rows) => pageGroups(rows, 6);

const range = (low, high) => `${f2(low)}–${f2(high)}`;
// §10a, approved standard text; the last sentence follows the annex setting.
// F26/F39: each sentence appears only when a generated slide shows what it explains; with nothing generated, only
// the generic opening and the annex sentence remain. `families` = the plan families that are generated.
const ALL_FAMILIES = new Set(["key-findings", "range", "ranking", "benchmark", "population", "zone", "behavior", "participant-mean", "participant-comparison", "competency-participants"]);
export function howToReadParagraphs(low = 2.75, high = 3.5, annex = "end", families = ALL_FAMILIES) {
  const has = (...names) => names.some((name) => families.has(name));
  const opening = "Rezultatele pe competențe sunt exprimate pe o scală de la 1 la 5, unde 1 reprezintă nivelul minim, iar 5 nivelul maxim.";
  const bandText = has("range", "ranking", "zone", "participant-mean", "participant-comparison", "competency-participants") ? ` Banda gri din grafice marchează intervalul de referință (benchmark) de ${range(low, high)}, care corespunde unei performanțe la nivel mediu în evaluările TREND: rezultatele din bandă sunt la nivel mediu, cele de deasupra ei peste medie, iar cele de dedesubt sub medie.` : "";
  const second = [has("key-findings", "range", "ranking", "benchmark", "participant-mean") ? "Media arată nivelul general al grupului; mediana este scorul participantului aflat la mijlocul grupului și este mai puțin influențată de rezultatele extreme." : "", has("range") ? "Graficele de distribuție arată, pentru fiecare competență, cel mai mic și cel mai mare scor obținut, mediana și intervalele în care se situează jumătatea superioară și cea inferioară a participanților." : ""].filter(Boolean).join(" ");
  const third = has("key-findings", "behavior") ? "Abilitățile cheie sunt comportamentele cel mai bine demonstrate în cadrul fiecărei competențe; abilitățile de dezvoltat sunt cele mai puțin demonstrate." : "";
  const last = annex === "none" ? "Rezultatele descriu grupul evaluat." : `Rezultatele descriu grupul evaluat. Rezultatele individuale se regăsesc ${annex === "separate" ? "în anexa transmisă separat" : "în anexă"}.`;
  return [`${opening}${bandText}`, second, third, last].filter(Boolean);
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

const joinRo = (names) => (names.length < 2 ? names.join("") : `${names.slice(0, -1).join(", ")} și ${names.at(-1)}`);
const tied = (sorted, mean) => joinRo(sorted.filter((item) => f2(item.mean) === f2(mean)).map((item) => item.competency));
/** Competencies with fewer than two scored behaviours: no R5 ranking (F41). */
export function unrankedCompetencies(payload) {
  const counts = new Map((payload.competencies || []).map((competency) => [competency, 0]));
  for (const row of payload.behaviorAggregates || []) counts.set(row.competency, (counts.get(row.competency) || 0) + 1);
  return [...counts.entries()].filter(([, count]) => count < 2).map(([competency]) => competency);
}
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
    // F42: a tie names every tied competency („X și Y (3.20)”).
    competencyLines: sorted.length ? [`Cel mai bine reprezentată: ${tied(sorted, sorted[0].mean)} (${f2(sorted[0].mean)})`, `Principala oportunitate: ${tied(sorted, sorted.at(-1).mean)} (${f2(sorted.at(-1).mean)})`] : [],
    conclusions: text(payload.metadata?.executiveConclusions ?? payload.metadata?.conclusions)
  };
}

// Vlad 2026-10-02: no percentages after the behaviours — the declined text or the behaviour only.
const behaviourLine = (row, field) => text(row[field]) || row.behavior;
export function competencyFindings(payload) {
  const insights = new Map(behaviorInsights(payload.behaviorAggregates || []).map((item) => [item.competency, item]));
  const { low, high } = payload.bands;
  return byMeanDesc(scoredCalculations(payload)).map((item) => {
    const insight = insights.get(item.competency) || { key: [], development: [] };
    const ladder = payload.records.filter((record) => Number.isFinite(record.scores?.[item.competency])).map((record) => ({ name: record.name, value: record.scores[item.competency] })).sort((a, b) => b.value - a.value);
    const scores = ladder.map((entry) => entry.value);
    const counts = { above: 0, in: 0, below: 0 }; scores.forEach((score) => { counts[band(score, low, high)] += 1; });
    return {
      item, insight, competency: item.competency, mean: item.mean, median: medianOf(scores), scores, ladder, counts,
      // F41: fewer than two scored behaviours leave R5 nothing to rank; the slide drops its two R5 boxes.
      noRanking: !insight.key.length && !insight.development.length,
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
  const rows = view.records.map((record) => ({ name: record.name, value: overallMean(record) })).filter((row) => Number.isFinite(row.value)).sort((a, b) => b.value - a.value).map((row) => ({ ...row, band: band(row.value, low, high) }));
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

// Section builders (fill map §1 as amended by Vlad 2026-10-02: template order). Each runs for the whole project
// first, then once per CODE group in place.
function sectionContext(view, groupKey) {
  const groupLabel = groupKey ? text(view.metadata.groupNames?.[groupKey]) || groupKey : "";
  const suffix = groupKey ? ` · ${groupLabel}` : "";
  const toggles = view.metadata.slideToggles || {};
  return { groupLabel, suffix, on: (key) => toggles[key] !== false, calculations: byMeanDesc(scoredCalculations(view)) };
}
const pusher = (slides, groupKey, groupLabel, suffix, deliverable) => (family, data = {}) => slides.push({ family, templateIndex: TEMPLATE_SLIDES[family], title: `${TEMPLATE_TITLES[family] || ""}${suffix}`, groupKey, groupLabel, suffix, deliverable, ...data });
function addResultsSection(slides, view, groupKey = "") {
  const { groupLabel, suffix, on, calculations } = sectionContext(view, groupKey);
  const add = pusher(slides, groupKey, groupLabel, suffix, "main");
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
}
function addKeyFindings(slides, view, groupKey = "") {
  const { groupLabel, suffix, on } = sectionContext(view, groupKey);
  if (!on("keyFindings") || !on("observation")) return;
  const add = pusher(slides, groupKey, groupLabel, suffix, "appendix");
  competencyFindings(view).forEach((finding) => add("key-findings", { ...finding, title: `Distribuția pe competențe – ${finding.competency}${suffix}`, low: view.bands.low, high: view.bands.high }));
}
function addBehaviorSection(slides, view, groupKey = "") {
  const { groupLabel, suffix, on, calculations } = sectionContext(view, groupKey);
  const insights = on("behavior") ? behaviorInsights(view.behaviorAggregates).filter((insight) => insight.key.length || insight.development.length) : [];
  if (!insights.length) return;
  const add = pusher(slides, groupKey, groupLabel, suffix, "appendix");
  add("divider-behaviors");
  const order = new Map(calculations.map((item, index) => [item.competency, index]));
  insights.sort((a, b) => (order.get(a.competency) ?? 99) - (order.get(b.competency) ?? 99)).forEach((insight) => add("behavior", { title: `Comportamente cheie – ${insight.competency}${suffix}`, competency: insight.competency, key: insight.key.map((row) => text(row.score2) || row.behavior), development: insight.development.map((row) => text(row.score0) || row.behavior) }));
}
function addConclusionsSection(slides, view, groupKey = "") {
  const { groupLabel, suffix, on } = sectionContext(view, groupKey);
  if (!on("conclusions")) return;
  const add = pusher(slides, groupKey, groupLabel, suffix, "main");
  add("divider-conclusions");
  // F32: a group's conclusions slide uses that group's own fields; the whole project uses the project fields.
  const own = groupKey ? view.metadata.groupConclusions?.[groupKey] || {} : { strengths: view.metadata.conclusionsStrengths, development: view.metadata.conclusionsDevelopment, interventions: view.metadata.conclusionsInterventions };
  add("conclusions", { strengths: text(own.strengths), development: text(own.development), interventions: text(own.interventions) });
}

function participantSlides(payload) {
  const slides = [];
  const add = (family, data = {}) => slides.push({ family, templateIndex: TEMPLATE_SLIDES[family], title: TEMPLATE_TITLES[family] || "", groupKey: "", deliverable: "appendix", low: payload.bands.low, high: payload.bands.high, ...data });
  const competencies = (payload.competencies || payload.calculations.map((item) => item.competency)).filter((competency) => payload.records.some((record) => Number.isFinite(record.scores?.[competency])));
  const participants = payload.records.map((record) => ({ name: record.name, mean: overallMean(record), scores: record.scores || {} })).filter((row) => Number.isFinite(row.mean)).sort((a, b) => b.mean - a.mean);
  const longestLabel = Math.max(0, ...participants.map((row) => labelChars(row.name)));
  // Rule 10 for A2: each name wraps (at spaces and after hyphens) inside a label column of ≤ 4.3 in; a slide holds as
  // many participants as keep every label ≥ 10 pt on its own rows, split into equal pages.
  const plotPoints = 0.9 * 9.25 * 72; const columnPoints = A2_LABEL_COLUMN_IN * 72;
  const slotFor = (size) => Math.max(...participants.map((row) => wrapLabel(row.name, size, columnPoints))) * size * 1.3 + 6;
  const meanCap = Math.max(1, Math.min(PARTICIPANT_START_CAPS["participant-mean"], Math.floor(plotPoints / Math.max(slotFor(MIN_PARTICIPANT_LABEL_PT), 0.3 * 72))));
  const meanPages = splitEqual(participants, meanCap);
  let meanSize = 24; while (meanSize > MIN_PARTICIPANT_LABEL_PT && !meanPages.every((rows) => slotFor(meanSize) <= plotPoints / Math.max(1, rows.length))) meanSize -= 1;
  meanPages.forEach((rows, page) => add("participant-mean", { rows, page: page + 1, pages: meanPages.length, labelSize: meanSize * 100 }));
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
    const fit = (page) => { const slot = 17.2 * 72 / Math.max(1, page.rows.length) * 0.9; const need = Math.max(1, ...page.rows.map((row) => labelChars(row.name))); return Math.max(MIN_PARTICIPANT_LABEL_PT * 100, Math.min(2000, Math.floor(slot / (need * 0.55)) * 100)); };
    const size = Math.min(2000, ...pages.map(fit));
    pages.forEach((page) => { page.labelSize = size; });
  }
  return slides;
}

export function reportPlan(payload, { scope = "whole" } = {}) {
  const metadata = payload.metadata || {};
  const annex = metadata.annex || "end";
  const client = text(metadata.clientName || metadata.projectName);
  const program = text(metadata.program) || "Centru de Dezvoltare";
  const year = text(metadata.reportDate).match(/\b(\d{4})\b/u)?.[1] || String(new Date().getFullYear());
  const groups = metadata.splitGroups && (payload.groups || []).length >= 2 ? payload.groups.map((group) => [viewForGroup(payload, group), group.code]) : [];
  const eachScope = (builder, target) => { builder(target, payload); groups.forEach(([view, code]) => builder(target, view, code)); };
  const single = (family, data = {}, deliverable = "main") => ({ family, templateIndex: TEMPLATE_SLIDES[family], title: TEMPLATE_TITLES[family] || "", groupKey: "", deliverable, ...data });
  // MAIN: cover, how to read, methodology, executive summary, „Distribuția rezultatelor” (whole, then groups).
  const main = [single("cover", { title: `${client} – ${program}`, client, program, year }), single("how-to-read", { title: "CUM CITIM ACEST RAPORT", paragraphs: [] }), single("methodology", { title: "PRIVIRE DE ANSAMBLU ASUPRA PROIECTULUI - METODOLOGIE", page: methodologyColumns(payload) }), single("executive-summary", { title: "Executive Summary", summary: executiveSummary(payload) })];
  eachScope(addResultsSection, main);
  // ANEXĂ: opener, per-person charts, „Analiza observațiilor” (key findings, then charts by participant), behaviours.
  const annexBlock = [];
  if (annex !== "none") {
    const people = participantSlides(payload);
    const findings = []; eachScope(addKeyFindings, findings);
    const behaviours = []; eachScope(addBehaviorSection, behaviours);
    const byPerson = people.filter((slide) => slide.family === "competency-participants");
    annexBlock.push(single("appendix-divider", { title: "Anexă – rezultate individuale", heading: "Anexă", subheading: "rezultate individuale" }, "appendix"), ...people.filter((slide) => slide.family !== "competency-participants"));
    if (findings.length || byPerson.length) annexBlock.push(single("divider-observations", {}, "appendix"), ...findings, ...byPerson);
    annexBlock.push(...behaviours);
  }
  // END: „Concluzii și recomandări” (whole, then groups) — always the last section before the closing slide.
  const end = []; eachScope(addConclusionsSection, end);
  const generated = new Set([...main, ...annexBlock, ...end].map((slide) => slide.family));
  main[1].paragraphs = howToReadParagraphs(payload.bands.low, payload.bands.high, annexBlock.length ? annex : "none", generated);
  const close = single("close", { title: "MULȚUMIM!" });
  let selected;
  if (scope === "main") selected = [...main, ...end, close];
  else if (scope === "appendix") selected = [{ ...main[0], annexMark: true, deliverable: "appendix" }, ...annexBlock, { ...close, deliverable: "appendix" }];
  else selected = [...main, ...annexBlock, ...end, close];
  return selected.map((slide, index) => ({ ...slide, number: index + 1, total: selected.length }));
}
