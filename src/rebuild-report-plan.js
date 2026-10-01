const text = (value) => String(value ?? "").trim();
const overall = (record) => { const values = Object.values(record.scores || {}).filter(Number.isFinite); return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null; };
const percent = (value, total) => total ? Math.round(value / total * 100) : 0;

export const participantChartPageSize = () => 8;
export const participantComparisonPageSize = 6;
export const pageGroups = (items, size) => { const pages = []; for (let index = 0; index < items.length; index += size) pages.push(items.slice(index, index + size)); return pages.length ? pages : [[]]; };

export function rankBehaviors(rows = []) {
  const groups = new Map();
  rows.forEach((row, index) => {
    if (!groups.has(row.competency)) groups.set(row.competency, []);
    groups.get(row.competency).push({ ...row, sourceIndex: row.sourceIndex ?? index });
  });
  return [...groups.entries()].map(([competency, items]) => {
    const ranked = items.slice().sort((a, b) => (b.sum ?? 0) - (a.sum ?? 0) || a.sourceIndex - b.sourceIndex);
    let keyCount = 0;
    let developmentCount = 0;
    let medianBehavior = null;
    if (ranked.length >= 10) { keyCount = 5; developmentCount = 5; }
    else if (ranked.length >= 6) { keyCount = 3; developmentCount = 3; }
    else {
      const medianIndex = Math.floor(ranked.length / 2);
      medianBehavior = ranked[medianIndex] || null;
      keyCount = medianIndex;
      developmentCount = Math.max(0, ranked.length - medianIndex - 1);
    }
    const key = ranked.slice(0, keyCount);
    const development = ranked.slice(ranked.length - developmentCount);
    return { competency, ranked, key, development, strengths: key, median: medianBehavior, all: ranked };
  });
}

export const behaviorInsights = (rows) => rankBehaviors(rows).map((item) => ({ ...item, strengths: item.key, development: item.development }));
export const behaviorPageGroups = (rows) => pageGroups(rows, 6);

export const HOW_TO_READ = (low = 2.75, high = 3.5, annex = "end") => `Rezultatele pe competențe sunt exprimate pe o scală de la 1 la 5, unde 1 reprezintă nivelul minim, iar 5 nivelul maxim. Banda gri din grafice marchează intervalul de referință (benchmark) de ${low.toFixed(2).replace(".", ",")}–${high.toFixed(2).replace(".", ",")}. Rezultatele din bandă sunt la nivel mediu, cele de deasupra ei peste medie, iar cele de dedesubt sub medie. Media arată nivelul general al grupului, iar mediana este mai puțin influențată de rezultatele extreme. Abilitățile cheie sunt comportamentele cel mai bine demonstrate; abilitățile de dezvoltat sunt cele mai puțin demonstrate. Rezultatele individuale se regăsesc ${annex === "separate" ? "în anexa transmisă separat" : annex === "none" ? "în livrarea individuală a instrumentului" : "în anexă"}.`;

export function methodologyColumns(payload) {
  const schemas = payload.schemas || [];
  const evaluators = new Set(schemas.flatMap((schema) => schema.methodology?.evaluators || []));
  const dates = new Set(schemas.flatMap((schema) => schema.methodology?.dates || []));
  const teamSizes = schemas.flatMap((schema) => schema.methodology?.teamSizes || []);
  const consultants = evaluators.size || teamSizes.filter((value) => value >= 2).sort((a, b) => b - a)[0] || "—";
  const days = dates.size || "—";
  const metadata = payload.metadata || {};
  return {
    left: ["METODOLOGIE", `${payload.participantCounts.included} participanți ${text(metadata.clientName || metadata.projectName)}`.trim(), `${consultants} consultanți Trend implicați`, `${days} ${days === 1 ? "zi de evaluare" : "zile de evaluare"}`, `${payload.behaviorAggregates.length} comportamente specifice observate`, text(metadata.exercises) ? `Exerciții: ${metadata.exercises}` : "Exerciții: de completat de consultant", text(metadata.otherInstruments) ? `Alte instrumente: ${metadata.otherInstruments}` : "Alte instrumente: de completat de consultant"],
    right: ["PROCESUL DE EVALUARE", "Fiecare participant a fost observat de o echipă de consultanți.", "Fiecare comportament a fost evaluat folosind o grilă comună.", "Observațiile au fost calibrate și integrate în platforma BHB Profiler.", "Platforma a generat rezultatele pentru fiecare competență.", "Scala utilizată: de la 1 la 5, unde 1 reprezintă nivelul minim și 5 nivelul maxim."],
    missing: { exercises: !text(metadata.exercises), otherInstruments: !text(metadata.otherInstruments) },
    facts: { days, consultants }
  };
}

export const methodologyPages = (payload) => [{ page: methodologyColumns(payload) }];

export function executiveSummary(payload, calculations = payload.calculations.filter((item) => item.mean !== null).slice().sort((a, b) => b.mean - a.mean)) {
  const bands = payload.bands || { below: 0, typical: 0, above: 0, n: 0, low: 2.75, high: 3.5 };
  const insights = behaviorInsights(payload.behaviorAggregates || []);
  return {
    population: payload.participantCounts.included,
    distribution: [{ label: `Sub ${bands.low}`, value: percent(bands.below, bands.n) }, { label: `În intervalul ${bands.low}–${bands.high}`, value: percent(bands.typical, bands.n) }, { label: `Peste ${bands.high}`, value: percent(bands.above, bands.n) }],
    strongest: calculations[0] || null,
    weakest: calculations.at(-1) || null,
    strengths: insights.flatMap((item) => item.key).slice(0, 3),
    development: insights.flatMap((item) => item.development).slice(0, 3),
    conclusion: calculations[0] && calculations.at(-1) ? `Grupul prezintă cel mai ridicat nivel mediu la ${calculations[0].competency} și cea mai importantă zonă de dezvoltare la ${calculations.at(-1).competency}.` : "Completează concluzia comportamentală după revizuirea rezultatelor."
  };
}

export const competencyFindings = (payload, calculations = payload.calculations.filter((item) => item.mean !== null).slice().sort((a, b) => b.mean - a.mean)) => {
  const insights = new Map(behaviorInsights(payload.behaviorAggregates || []).map((item) => [item.competency, item]));
  return calculations.map((item) => ({ item, insight: insights.get(item.competency) || { key: [], development: [], strengths: [] } }));
};

function viewForGroup(payload, group) {
  const identities = new Set(group.records.map((record) => record.identity));
  const records = payload.records.filter((record) => identities.has(record.identity));
  const competencies = payload.competencies || payload.calculations.map((item) => item.competency);
  return { ...payload, records, participantCounts: { ...payload.participantCounts, included: records.length }, calculations: competencies.map((competency) => { const values = records.map((record) => record.scores[competency]).filter(Number.isFinite); return { competency, n: values.length, mean: values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null, min: values.length ? Math.min(...values) : null, median: values.length ? values.slice().sort((a, b) => a - b)[Math.floor(values.length / 2)] : null, max: values.length ? Math.max(...values) : null }; }), behaviorAggregates: payload.behaviorAggregates.filter((row) => payload.behaviorRecords.some((record) => identities.has(record.identity) && record.competency === row.competency && record.behavior === row.behavior)) };
}

function addTemplateSection(slides, view, groupKey = "") {
  const groupLabel = groupKey ? view.metadata.groupNames?.[groupKey] || groupKey : "întregul proiect";
  const calculations = view.calculations.filter((item) => item.mean !== null).slice().sort((a, b) => b.mean - a.mean);
  const toggles = view.metadata.slideToggles || {};
  const add = (family, title, data = {}, deliverable = "main") => slides.push({ family, title, groupKey, deliverable, ...data });
  add("divider", `Distribuția rezultatelor${groupKey ? ` · ${groupLabel}` : ""}`, { templateIndex: 3 });
  if (toggles.range !== false) add("range", `Mediană și plajă${groupKey ? ` · ${groupLabel}` : ""}`, { items: calculations, templateIndex: 4 });
  if (toggles.competencyMean !== false) add("ranking", `Media pe competențe${groupKey ? ` · ${groupLabel}` : ""}`, { items: calculations, templateIndex: 6 });
  if (toggles.benchmark !== false) add("benchmark", `Distribuția pe benchmark${groupKey ? ` · ${groupLabel}` : ""}`, { templateIndex: 7 });
  if (toggles.competencyDistribution !== false) add("competency-distribution", `Distribuția pe competențe${groupKey ? ` · ${groupLabel}` : ""}`, { items: calculations, templateIndex: 8 });
  if (toggles.zone !== false && view.regionReadiness.available) add("zone", `Rezultate pe regiuni${groupKey ? ` · ${groupLabel}` : ""}`, { items: view.zoneCalculations, templateIndex: 10 });
  add("divider", `Analiza observațiilor${groupKey ? ` · ${groupLabel}` : ""}`, { templateIndex: 11 });
  if (toggles.observation !== false) calculations.forEach((item) => add("observation", `Observații · ${item.competency}${groupKey ? ` · ${groupLabel}` : ""}`, { item, templateIndex: 12 }));
  add("divider", `Puncte forte și recomandări${groupKey ? ` · ${groupLabel}` : ""}`, { templateIndex: 18 });
  if (toggles.behavior !== false) behaviorInsights(view.behaviorAggregates).forEach((insight) => add("behavior", `Comportamente cheie · ${insight.competency}${groupKey ? ` · ${groupLabel}` : ""}`, { insight, templateIndex: 19 }));
  if (toggles.conclusions !== false) add("conclusions", `Concluzii și recomandări${groupKey ? ` · ${groupLabel}` : ""}`, { copy: text(view.metadata.conclusions), templateIndex: 21 });
}

export function reportPlan(payload, { scope = "whole" } = {}) {
  const slides = [];
  const add = (family, title, data = {}, deliverable = "main") => slides.push({ family, title, deliverable, ...data });
  const annex = payload.metadata.annex || "end";
  add("cover", payload.metadata.clientName || payload.metadata.projectName, { templateIndex: 1 });
  add("how-to-read", "Cum citim acest raport", { copy: HOW_TO_READ(payload.bands.low, payload.bands.high, annex), templateIndex: 2 });
  add("methodology", "Metodologie", { page: methodologyColumns(payload), templateIndex: 2 });
  add("executive-summary", "Imaginea de ansamblu", { summary: executiveSummary(payload), templateIndex: 2 });
  competencyFindings(payload).forEach((finding) => add("key-findings", `Constatări cheie · ${finding.item.competency}`, { ...finding, templateIndex: 12 }));
  addTemplateSection(slides, payload);
  const groupSections = payload.metadata.splitGroups && payload.groups.length >= 2;
  if (groupSections) payload.groups.forEach((group) => addTemplateSection(slides, viewForGroup(payload, group), group.code));
  if (annex !== "none") {
    add("appendix-divider", "Anexă · rezultate individuale", {}, "appendix");
    const ranked = payload.calculations.filter((item) => item.mean !== null).slice().sort((a, b) => b.mean - a.mean);
    const comparison = payload.records.slice().sort((a, b) => overall(b) - overall(a));
    pageGroups(comparison, participantComparisonPageSize).forEach((items) => add("participant-comparison", "Rezultate pe competențe / participant", { items, ranked }, "appendix"));
    ranked.forEach((item) => pageGroups(comparison.filter((record) => Number.isFinite(record.scores[item.competency])), participantChartPageSize()).forEach((items) => add("competency-distribution", `Rezultate individuale · ${item.competency}`, { item, items, appendix: true }, "appendix")));
  }
  add("close", "MULȚUMIM!", { templateIndex: 22 });
  let selected = slides;
  if (scope === "main") selected = slides.filter((slide) => slide.deliverable !== "appendix");
  if (scope === "appendix") selected = [...slides.filter((slide) => slide.deliverable === "appendix"), slides.find((slide) => slide.family === "close")].filter(Boolean);
  return selected.map((slide, index) => ({ ...slide, number: index + 1, total: selected.length }));
}
