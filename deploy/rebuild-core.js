export const CALCULATION_VERSION = "AC-GRF-R-1.0";
export const EVAL_SHEET_HEADERS = ["competency", "subcompetency", "behavior", "objective_text_score_0", "objective_text_score_-1", "objective_text_score_1", "objective_text_score_2"];

const text = (value) => String(value ?? "").replace(/[\u00a0\u2007\u202f]/gu, " ").replace(/\s+/gu, " ").trim();
const folded = (value) => text(value).normalize("NFD").replace(/[\u0300-\u036f]/gu, "").toLocaleLowerCase("ro");
const exact = (value) => text(value);
const key = (...values) => values.map(folded).join("|");
const safe = (value) => /^[=+\-@]/u.test(text(value)) ? `'${text(value)}` : text(value);
const numberValue = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (text(value) === "") return null;
  const parsed = Number(String(value).replace(",", "."));
  return Number.isFinite(parsed) ? parsed : NaN;
};
const median = (values) => {
  const sorted = values.filter(Number.isFinite).slice().sort((a, b) => a - b);
  if (!sorted.length) return null;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};
const fingerprint = (bytes) => {
  const view = new Uint8Array(bytes);
  let hash = 2166136261;
  for (const byte of view) hash = Math.imul(hash ^ byte, 16777619);
  return `${view.byteLength} bytes · ${String(hash >>> 0).padStart(10, "0")}`;
};
const issue = (list, severity, code, message, rowNumber, details = {}) => {
  const id = details.id || [code, folded(details.sourceName), rowNumber || "", details.identity || "", folded(details.field)].join(":");
  list.push({ severity, code, message, rowNumber, ...details, id });
};

export function mergeSelectedFiles(existing = [], incoming = []) {
  const merged = new Map(existing.map((file) => [folded(file.name), file]));
  incoming.forEach((file) => merged.set(folded(file.name), file));
  return [...merged.values()];
}

function rowsFromWorkbook(XLSX, bytes, sourceName = "") {
  const csv = /\.csv$/iu.test(sourceName);
  const input = csv ? new TextDecoder("utf-8").decode(bytes).replace(/^\uFEFF/u, "") : bytes;
  const workbook = XLSX.read(input, { type: csv ? "string" : bytes instanceof ArrayBuffer ? "array" : "buffer", cellFormula: false, cellHTML: false, raw: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("Fișierul nu conține nicio foaie.");
  return { sheetName, rows: XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true, blankrows: false }) };
}

const findHeader = (rows, predicate) => rows.findIndex((row) => predicate(row.map(folded)));
const methodFacts = (rows, headerRow) => {
  const header = (rows[headerRow] || []).map(folded);
  const values = (name) => { const index = header.indexOf(name); return index < 0 ? [] : rows.slice(headerRow + 1).map((row) => text(row[index])).filter(Boolean); };
  const evaluatorIndexes = ["principal evaluator", "secondary evaluator", "evaluator 3", "evaluator 4", "evaluatori"].map((name) => header.indexOf(name)).filter((index) => index >= 0);
  const teamSizes = rows.slice(headerRow + 1).map((row) => new Set(evaluatorIndexes.flatMap((index) => text(row[index]).split(/[,;/]/u).map(text).filter(Boolean))).size).filter(Boolean);
  return { evaluators: ["principal evaluator", "secondary evaluator", "evaluator 3", "evaluator 4", "evaluatori"].flatMap(values), dates: [...values("invited at"), ...values("date")], locations: [...values("certification location"), ...values("regiune")], teamSizes };
};

export function detectSchema(XLSX, bytes, sourceName = "export.xlsx") {
  const { sheetName, rows } = rowsFromWorkbook(XLSX, bytes, sourceName);
  const summary = findHeader(rows, (header) => header.includes("name") && header.includes("cod cp") && header.some((value) => value && !["code", "name", "cod cp", "job", "email"].includes(value)));
  if (summary >= 0) return { kind: "ac-summary-1-5", sourceName, sheetName, headerRow: summary + 1, rows, methodology: methodFacts(rows, summary) };
  const detailed = findHeader(rows, (header) => header.includes("name the person evaluated") && header.includes("cod ac") && header.includes("competente"));
  if (detailed >= 0) return { kind: "ac-detailed-0-2", sourceName, sheetName, headerRow: detailed + 1, rows, methodology: methodFacts(rows, detailed) };
  const descriptors = findHeader(rows, (header) => EVAL_SHEET_HEADERS.every((name) => header.includes(name)));
  if (descriptors >= 0) return { kind: "devplan-descriptors", sourceName, sheetName, headerRow: descriptors + 1, rows, methodology: { evaluators: [], dates: [], locations: [], teamSizes: [] } };
  return { kind: "unsupported", sourceName, sheetName, headerRow: 0, rows, methodology: methodFacts(rows, 0) };
}

function summaryColumns(header) {
  const meta = new Set(["code", "name", "job", "email", "certification location", "invited at", "principal evaluator", "secondary evaluator", "evaluator 3", "evaluator 4", "cod cp"]);
  return header.map((value, index) => ({ value: text(value), index })).filter(({ value }) => value && !meta.has(folded(value)));
}

function parseSummary(schema, corrections = {}) {
  const issues = [];
  const header = schema.rows[schema.headerRow - 1].map(text);
  const foldedHeader = header.map(folded);
  const nameIndex = foldedHeader.indexOf("name");
  const assessmentIndex = foldedHeader.indexOf("cod cp");
  const codeIndex = foldedHeader.indexOf("code");
  const columns = summaryColumns(header);
  if (nameIndex < 0 || assessmentIndex < 0 || !columns.length) issue(issues, "blocker", "summary-headers", "Lipsesc coloanele obligatorii pentru exportul AC de sinteză.", undefined, { sourceName: schema.sourceName });
  const records = schema.rows.slice(schema.headerRow).filter((row) => row.some((value) => text(value))).map((row, offset) => {
    const rowNumber = schema.headerRow + offset + 1;
    const name = text(row[nameIndex]);
    const assessment = text(row[assessmentIndex]);
    const identity = key(name, assessment);
    const code = codeIndex >= 0 ? exact(row[codeIndex]) : "";
    const scores = {};
    if (!name) issue(issues, "blocker", "required-value", "Numele persoanei lipsește.", rowNumber, { sourceName: schema.sourceName, identity });
    if (!assessment) issue(issues, "blocker", "required-value", "Codul evaluării lipsește.", rowNumber, { sourceName: schema.sourceName, identity });
    for (const column of columns) {
      const correctionKey = `summary:${schema.sourceName}:${rowNumber}:${column.value}`;
      const correction = corrections?.values?.[correctionKey];
      const original = row[column.index];
      const score = correction?.mode === "value" ? numberValue(correction.value) : numberValue(original);
      const details = { correctionKey, sourceName: schema.sourceName, field: column.value, identity, originalValue: original, kind: "summary-score", subject: column.value };
      if (score === null) issue(issues, "warning", "summary-unrated", `${name || "Participantul"} nu are scor pentru ${column.value}; valoarea nu intră în calcule.`, rowNumber, details);
      else if (!Number.isFinite(score) || score < 1 || score > 5) issue(issues, "blocker", "summary-score", `Scor invalid (1–5) pentru ${column.value}.`, rowNumber, details);
      else scores[column.value] = score;
    }
    return { name, assessment, identity, code, scores, rowNumber, source: schema.sourceName, inclusion: Object.keys(scores).length ? Object.keys(scores).length < columns.length ? "included-partial" : "included" : "excluded-unrated" };
  });
  return { records, competencies: columns.map(({ value }) => value), issues };
}

function parseDetailed(schema, corrections = {}) {
  const issues = [];
  const rows = schema.rows;
  const groups = rows[schema.headerRow - 1].map(text);
  const subcompetencies = rows[schema.headerRow].map(text);
  const behaviors = rows[schema.headerRow + 1].map(text);
  const normalizedGroups = groups.map(folded);
  const nameIndex = normalizedGroups.indexOf("name the person evaluated");
  const assessmentIndex = normalizedGroups.indexOf("cod ac");
  const codeIndex = normalizedGroups.indexOf("code");
  const regionIndex = normalizedGroups.indexOf("regiune");
  const behaviorStart = behaviors.findIndex((value) => folded(value) === "behavior");
  const columns = [];
  let currentCompetency = "";
  for (let index = 0; index < groups.length; index += 1) {
    if (index > behaviorStart && groups[index]) currentCompetency = groups[index];
    if (index > behaviorStart && currentCompetency && behaviors[index]) columns.push({ index, competency: currentCompetency, subcompetency: subcompetencies[index], behavior: behaviors[index] });
  }
  // F18 (ruling 2026-10-02): export artefacts are not client text. A leading score-scale annotation „(0-2)” is
  // always removed; a trailing „ /TOKEN” only when every imported behaviour ends with the same token.
  const trailing = columns.map((column) => column.behavior.match(/\s+\/([^\s/]+)$/u)?.[1] || null);
  const sharedToken = trailing.length && trailing.every((token) => token && token === trailing[0]) ? trailing[0] : null;
  for (const column of columns) {
    column.behaviorRaw = column.behavior;
    let clean = column.behavior.replace(/^\s*\(\s*0\s*[-–—]\s*2\s*\)\s*/u, "");
    if (sharedToken) clean = clean.replace(new RegExp(`\\s+/${sharedToken.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&")}$`, "u"), "");
    column.behavior = clean.trim() || column.behavior;
  }
  columns.forEach((column, order) => { column.order = order; });
  if (nameIndex < 0 || assessmentIndex < 0 || behaviorStart < 0 || !columns.length) issue(issues, "blocker", "detailed-headers", "Lipsesc coloanele obligatorii pentru exportul AC detaliat.", undefined, { sourceName: schema.sourceName });
  const records = [];
  const behaviorRecords = [];
  for (const [offset, row] of rows.slice(schema.headerRow + 2).filter((candidate) => candidate.some((value) => text(value))).entries()) {
    const rowNumber = schema.headerRow + offset + 3;
    const name = text(row[nameIndex]);
    const assessment = text(row[assessmentIndex]);
    const identity = key(name, assessment);
    const code = codeIndex >= 0 ? exact(row[codeIndex]) : "";
    const region = regionIndex >= 0 ? exact(row[regionIndex]) : "";
    const scores = {};
    if (!name) issue(issues, "blocker", "required-value", "Numele persoanei lipsește.", rowNumber, { sourceName: schema.sourceName, identity });
    if (!assessment) issue(issues, "blocker", "required-value", "Codul evaluării lipsește.", rowNumber, { sourceName: schema.sourceName, identity });
    if (!region) issue(issues, "warning", "region-missing", `${name || "Participantul"} nu are regiune; rămâne în proiect, dar nu apare pe slide-ul pe regiuni.`, rowNumber, { sourceName: schema.sourceName, identity, field: "regiune" });
    for (const column of columns) {
      const correctionKey = `detailed:${schema.sourceName}:${rowNumber}:${column.competency}:${column.behavior}`;
      const correction = corrections?.values?.[correctionKey];
      const original = row[column.index];
      const score = correction?.mode === "value" ? numberValue(correction.value) : numberValue(original);
      const details = { correctionKey, sourceName: schema.sourceName, field: column.behavior, competency: column.competency, behavior: column.behavior, identity, originalValue: original, kind: "detailed-score", subject: column.behavior };
      if (score === null) { issue(issues, "warning", "detailed-blank", `Scor lipsă: ${name || "participant fără nume"} – „${column.behavior}”`, rowNumber, details); continue; }
      if (!Number.isInteger(score) || score < 0 || score > 2) { issue(issues, "blocker", "detailed-score", `Scor invalid: sunt acceptate numai valorile întregi 0, 1 sau 2 pentru ${column.behavior}.`, rowNumber, details); continue; }
      if (!scores[column.competency]) scores[column.competency] = [];
      scores[column.competency].push(score);
      behaviorRecords.push({ name, assessment, identity, code, region, competency: column.competency, subcompetency: column.subcompetency, behavior: column.behavior, behaviorRaw: column.behaviorRaw, score, rowNumber, source: schema.sourceName, sourceIndex: column.order, recordIndex: behaviorRecords.length });
    }
    records.push({ name, assessment, identity, code, region, scores, rowNumber, source: schema.sourceName });
  }
  // F40: a behaviour with no score in any row is reported once (behaviour + number of participants), not row by row.
  for (const column of columns) {
    if (behaviorRecords.some((record) => record.competency === column.competency && record.behavior === column.behavior)) continue;
    const blanks = issues.filter((item) => item.code === "detailed-blank" && item.competency === column.competency && item.behavior === column.behavior);
    if (!blanks.length) continue;
    for (const blank of blanks) issues.splice(issues.indexOf(blank), 1);
    issue(issues, "warning", "detailed-all-blank", `Comportament fără niciun scor în export: „${column.behavior}” (${blanks.length} participanți fără scor). Este exclus din clasament; rămâne în șablonul CSV.`, undefined, { sourceName: schema.sourceName, field: column.behavior, competency: column.competency, behavior: column.behavior, id: `detailed-all-blank:${folded(column.competency)}:${folded(column.behavior)}` });
  }
  const behaviorCatalog = columns.filter((column, index, all) => all.findIndex((candidate) => candidate.competency === column.competency && candidate.behavior === column.behavior) === index).map((column) => ({ competency: column.competency, subcompetency: column.subcompetency || "", behavior: column.behavior, behaviorRaw: column.behaviorRaw, sourceIndex: column.order }));
  return { records, behaviorRecords, behaviorCatalog, competencies: [...new Set(columns.map(({ competency }) => competency))], issues };
}

function parseDescriptors(schema) {
  const issues = [];
  const header = schema.rows[schema.headerRow - 1].map(text);
  const indexes = Object.fromEntries(EVAL_SHEET_HEADERS.map((name) => [name, header.indexOf(name)]));
  const records = schema.rows.slice(schema.headerRow).filter((row) => row.some((value) => text(value))).map((row, offset) => {
    const rowNumber = schema.headerRow + offset + 1;
    const record = Object.fromEntries(EVAL_SHEET_HEADERS.map((name) => [name, text(row[indexes[name]])]));
    if (!record.behavior) issue(issues, "warning", "descriptor-incomplete", "Rândul de declinații fără comportament este ignorat.", rowNumber, { sourceName: schema.sourceName });
    return { ...record, rowNumber, source: schema.sourceName };
  }).filter((record) => record.behavior);
  return { records, issues };
}

const scoreStats = (records, competency) => {
  const values = records.map((record) => record.scores[competency]).filter(Number.isFinite);
  return { competency, n: values.length, mean: values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null, min: values.length ? Math.min(...values) : null, median: median(values), max: values.length ? Math.max(...values) : null };
};

// D315 / D314 scope: a CSV row that names a competency matches only within it; a row with a blank competency
// matches on the behaviour text alone; more than one candidate falls back to the imported text with a warning.
function descriptorFor(descriptors, competency, behavior, warnings, behaviorRaw = behavior) {
  // F38: the raw imported text is the cross-tool key; the cleaned text is the fallback.
  const rawMatches = descriptors.filter((item) => exact(item.behavior) === exact(behaviorRaw));
  const sameText = rawMatches.length ? rawMatches : descriptors.filter((item) => exact(item.behavior) === exact(behavior));
  const scoped = sameText.filter((item) => item.competency && folded(item.competency) === folded(competency));
  const blank = sameText.filter((item) => !item.competency);
  const candidates = scoped.length ? scoped : blank;
  if (candidates.length > 1) {
    warnings.push({ severity: "warning", code: "descriptor-ambiguous", message: `Declinația pentru „${behavior}” este ambiguă; se păstrează textul importat.`, sourceName: candidates.map((item) => item.source).join(", "), field: behavior, identity: key(competency, behavior), id: `descriptor-ambiguous:${key(competency, behavior)}` });
    return { score0: "", score2: "", descriptorSource: "fallback-ambiguous" };
  }
  return candidates[0] ? { score0: candidates[0].objective_text_score_0, score2: candidates[0].objective_text_score_2, scoreMinus1: candidates[0]["objective_text_score_-1"], score1: candidates[0].objective_text_score_1, descriptorSource: candidates[0].source } : { score0: "", score2: "", descriptorSource: "fallback-import" };
}

function behaviorAggregates(behaviorRecords, includedIdentities, descriptors, warnings) {
  const grouped = new Map();
  for (const row of behaviorRecords) {
    if (!includedIdentities.has(row.identity)) continue;
    const mapKey = `${row.competency}\u0000${row.behavior}`;
    if (!grouped.has(mapKey)) grouped.set(mapKey, { competency: row.competency, subcompetency: row.subcompetency, behavior: row.behavior, behaviorRaw: row.behaviorRaw, values: [], sourceIndex: row.sourceIndex });
    grouped.get(mapKey).values.push(row.score);
  }
  return [...grouped.values()].map((item) => {
    const values = item.values;
    const descriptor = descriptorFor(descriptors, item.competency, item.behavior, warnings, item.behaviorRaw);
    return { behaviorRaw: item.behaviorRaw, sourceIndex: item.sourceIndex, competency: item.competency, subcompetency: item.subcompetency, behavior: item.behavior, n: values.length, sum: values.reduce((sum, value) => sum + value, 0), mean: values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null, pct0: values.length ? values.filter((value) => value === 0).length / values.length : 0, pct2: values.length ? values.filter((value) => value === 2).length / values.length : 0, score0: descriptor.score0, score2: descriptor.score2, descriptorSource: descriptor.descriptorSource, sourceIndex: item.sourceIndex };
  });
}

function sheet(XLSX, rows) { return XLSX.utils.aoa_to_sheet(rows.map((row) => row.map((value) => typeof value === "string" ? safe(value) : value))); }

export function buildPayload(XLSX, files, metadata = {}, corrections = {}, reviewState = {}) {
  const schemas = files.map((file) => ({ ...detectSchema(XLSX, file.bytes, file.name), fingerprint: fingerprint(file.bytes) }));
  const blockers = [];
  const warnings = [];
  const summaries = schemas.filter((schema) => schema.kind === "ac-summary-1-5");
  const detailedSources = schemas.filter((schema) => schema.kind === "ac-detailed-0-2");
  const descriptorSources = schemas.filter((schema) => schema.kind === "devplan-descriptors");
  schemas.filter((schema) => schema.kind === "unsupported").forEach((schema) => issue(blockers, "blocker", "unsupported", `${schema.sourceName} nu este un export AC recunoscut.`, undefined, { sourceName: schema.sourceName }));
  if (!summaries.length) issue(blockers, "blocker", "summary-required", "Încarcă exportul AC de sinteză.");
  if (summaries.length > 1) issue(blockers, "blocker", "summary-source-ambiguity", `Există mai multe exporturi de sinteză: ${summaries.map((schema) => schema.sourceName).join(", ")}. Niciunul nu este ales automat.`);
  if (!detailedSources.length) issue(blockers, "blocker", "detailed-required", "Încarcă exportul AC detaliat pentru comportamente și zone.");
  if (detailedSources.length > 1) issue(blockers, "blocker", "detailed-source-ambiguity", `Există mai multe exporturi detaliate: ${detailedSources.map((schema) => schema.sourceName).join(", ")}. Niciunul nu este ales automat.`);
  const summary = summaries[0];
  const detailed = detailedSources[0];
  const parsedSummary = summary ? parseSummary(summary, corrections) : { records: [], competencies: [], issues: [] };
  const parsedDetailed = detailed ? parseDetailed(detailed, corrections) : { records: [], behaviorRecords: [], behaviorCatalog: [], competencies: [], issues: [] };
  blockers.push(...parsedSummary.issues.filter((item) => item.severity === "blocker"), ...parsedDetailed.issues.filter((item) => item.severity === "blocker"));
  warnings.push(...parsedSummary.issues.filter((item) => item.severity === "warning"), ...parsedDetailed.issues.filter((item) => item.severity === "warning"));
  const descriptorParsed = descriptorSources.map(parseDescriptors);
  const descriptors = descriptorParsed.flatMap((result) => result.records);
  warnings.push(...descriptorParsed.flatMap((result) => result.issues));
  const summaryCount = new Map();
  parsedSummary.records.forEach((record) => summaryCount.set(record.identity, (summaryCount.get(record.identity) || 0) + 1));
  for (const [identity, count] of summaryCount) if (count > 1) issue(blockers, "blocker", "duplicate-identity", `Identitatea persoană + cod evaluare este duplicată în sinteză: ${identity}.`, undefined, { identity, sourceName: summary?.sourceName });
  const detailedCount = new Map();
  parsedDetailed.records.forEach((record) => detailedCount.set(record.identity, (detailedCount.get(record.identity) || 0) + 1));
  for (const [identity, count] of detailedCount) if (count > 1) issue(blockers, "blocker", "detailed-duplicate-identity", `Identitatea persoană + cod evaluare este duplicată în detaliat: ${identity}.`, undefined, { identity, sourceName: detailed?.sourceName });
  const summaryIdentities = new Set(parsedSummary.records.map((record) => record.identity));
  const detailedIdentities = new Set(parsedDetailed.records.map((record) => record.identity));
  for (const identity of detailedIdentities) if (!summaryIdentities.has(identity)) issue(blockers, "blocker", "detailed-only-identity", `Identitatea ${identity} există în detaliat, dar nu și în sinteză.`, undefined, { identity, sourceName: detailed?.sourceName });
  for (const identity of summaryIdentities) if (!detailedIdentities.has(identity)) warnings.push({ severity: "warning", code: "detailed-missing-identity", message: `Identitatea ${identity} nu are rând în exportul detaliat; nu contribuie la comportamente.`, identity, sourceName: detailed?.sourceName || "", id: `detailed-missing-identity:${identity}` });
  if (!text(metadata.projectName)) issue(blockers, "blocker", "project-required", "Numele proiectului este obligatoriu.");
  const detailByIdentity = new Map(parsedDetailed.records.map((record) => [record.identity, record]));
  const records = parsedSummary.records.filter((record) => record.inclusion !== "excluded-unrated").map((record) => ({ ...record, code: record.code || detailByIdentity.get(record.identity)?.code || "", region: detailByIdentity.get(record.identity)?.region || "" }));
  const auditRecords = parsedSummary.records.map((record) => ({ ...record, code: record.code || detailByIdentity.get(record.identity)?.code || "", region: detailByIdentity.get(record.identity)?.region || "" }));
  const includedIdentities = new Set(records.map((record) => record.identity));
  const codeValues = [...new Set(records.map((record) => record.code).filter(Boolean))];
  const codeMissing = records.filter((record) => !record.code);
  if (codeMissing.length) {
    const names = codeMissing.map((record) => record.name);
    const listed = names.slice(0, 8).join(", ") + (names.length > 8 ? ` și încă ${names.length - 8}` : "");
    warnings.push({ severity: "warning", code: "code-missing", message: `${codeMissing.length} participanți nu au CODE (${listed}); sunt incluși doar în vederea întregului proiect.`, identities: codeMissing.map((record) => record.identity), id: "code-missing" });
  }
  const zones = [...new Set(records.map((record) => record.region).filter(Boolean))].map((region) => ({ region, records: records.filter((record) => record.region === region) }));
  const regionMissing = records.filter((record) => !record.region);
  const low = Number(metadata.benchmarkLow ?? 2.75);
  const high = Number(metadata.benchmarkHigh ?? 3.5);
  if (!Number.isFinite(low) || !Number.isFinite(high) || low < 1 || high > 5 || low >= high) issue(blockers, "blocker", "benchmark-invalid", "Intervalul benchmark trebuie să fie între 1 și 5, cu pragul inferior mai mic decât pragul superior.");
  const calculations = parsedSummary.competencies.map((competency) => scoreStats(records, competency));
  const overallScores = records.map((record) => Object.values(record.scores).reduce((sum, value, _, values) => sum + value / values.length, 0));
  const bands = { low, high, below: overallScores.filter((score) => score < low).length, typical: overallScores.filter((score) => score >= low && score <= high).length, above: overallScores.filter((score) => score > high).length, n: overallScores.length };
  // F37: a CSV row that matches no imported behaviour is reported (row and text); its texts are not used.
  for (const record of descriptors) {
    const matched = parsedDetailed.behaviorCatalog.some((row) => [exact(row.behaviorRaw), exact(row.behavior)].includes(exact(record.behavior)) && (!record.competency || folded(record.competency) === folded(row.competency)));
    if (!matched) warnings.push({ severity: "warning", code: "descriptor-unmatched", message: `Rândul ${record.rowNumber} din fișierul de declinații nu corespunde niciunui comportament importat: „${record.behavior}”.`, sourceName: record.source, rowNumber: record.rowNumber, field: record.behavior, id: `descriptor-unmatched:${folded(record.source)}:${record.rowNumber}` });
  }
  const descriptorWarnings = [];
  const aggregates = behaviorAggregates(parsedDetailed.behaviorRecords, includedIdentities, descriptors, descriptorWarnings);
  warnings.push(...descriptorWarnings);
  const acknowledged = new Set(reviewState.acknowledgedWarningIds || []);
  const finalWarnings = warnings.map((item) => ({ ...item, reviewed: acknowledged.has(item.id) }));
  const groups = codeValues.map((code) => ({ code, name: code, records: records.filter((record) => record.code === code) }));
  const zoneCalculations = zones.flatMap((zone) => parsedSummary.competencies.map((competency) => ({ region: zone.region, ...scoreStats(zone.records, competency) })));
  // F20: after import the template lists every imported behaviour with the declined texts already known.
  const descriptorTemplate = parsedDetailed.behaviorCatalog.map((row) => { const known = descriptorFor(descriptors, row.competency, row.behavior, [], row.behaviorRaw); return { competency: row.competency, subcompetency: row.subcompetency || "", behavior: row.behaviorRaw || row.behavior, objective_text_score_0: known.score0 || "", "objective_text_score_-1": known.scoreMinus1 || "", objective_text_score_1: known.score1 || "", objective_text_score_2: known.score2 || "" }; });
  return {
    calculationVersion: CALCULATION_VERSION,
    createdAt: new Date().toISOString(),
    metadata: { ...metadata, projectName: text(metadata.projectName), clientName: text(metadata.clientName), reportDate: text(metadata.reportDate), annex: metadata.annex || "end", splitGroups: Boolean(metadata.splitGroups), groupNames: metadata.groupNames || {}, slideToggles: metadata.slideToggles || {} },
    schemas, sourceShape: { summary: summaries.length, detailed: detailedSources.length, descriptors: descriptorSources.length }, competencies: parsedSummary.competencies, records, auditRecords, behaviorRecords: parsedDetailed.behaviorRecords, behaviorAggregates: aggregates,
    calculations, groups, zones, zoneCalculations, codeReadiness: { available: codeValues.length, blank: codeMissing.length, groups: codeValues, splitAvailable: codeValues.length >= 2 }, regionReadiness: { available: zones.length, blank: regionMissing.length, disabledReason: zones.length ? "" : "Nu există valori de regiune în exportul detaliat." },
    participantCounts: { total: parsedSummary.records.length, included: records.length, excludedUnrated: parsedSummary.records.filter((record) => record.inclusion === "excluded-unrated").length }, bands, blockers, warnings: finalWarnings, warningReviews: finalWarnings, corrections: Object.values(corrections.values || {}).filter((item) => item?.mode === "value"), readiness: blockers.length === 0 && finalWarnings.every((item) => item.reviewed), descriptorTemplate
  };
}

export function createEvaluationSheetTemplate(XLSX, payload) {
  const rows = [EVAL_SHEET_HEADERS, ...(payload.descriptorTemplate || []).map((row) => EVAL_SHEET_HEADERS.map((header) => row[header] || ""))];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet(XLSX, rows), "Evaluation sheet");
  return workbook;
}

export function createAuditWorkbook(XLSX, payload) {
  const workbook = XLSX.utils.book_new();
  const add = (name, rows) => XLSX.utils.book_append_sheet(workbook, sheet(XLSX, rows), name);
  add("Rezumat", [["BHB · Group Report Factory", ""], ["Versiune calcul", payload.calculationVersion], ["Proiect", payload.metadata.projectName], ["Generat", payload.createdAt], ["Participanți în fișier", payload.participantCounts.total], ["Participanți incluși", payload.participantCounts.included], ["Participanți fără scor excluși", payload.participantCounts.excludedUnrated], ["Grupuri CODE", payload.codeReadiness.groups.join(", ")], ["Zone", payload.regionReadiness.available], ["Prag benchmark inferior", payload.bands.low], ["Prag benchmark superior", payload.bands.high], ["Blocaje", payload.blockers.length], ["Avertismente", payload.warnings.length], ["Avertismente revizuite", payload.warnings.filter((item) => item.reviewed).length]]);
  add("Surse", [["Fișier", "Tip detectat", "Foaie", "Rând header", "Amprentă"], ...payload.schemas.map((schema) => [schema.sourceName, schema.kind, schema.sheetName, schema.headerRow, schema.fingerprint])]);
  add("Validare", [["Severitate", "Cod", "ID avertisment", "Revizuit", "Sursă", "Rând", "Identitate stabilă", "Câmp", "Mesaj"], ...[...payload.blockers, ...payload.warnings].map((item) => [item.severity, item.code, item.severity === "warning" ? item.id : "", item.severity === "warning" ? (item.reviewed ? "Da" : "Nu") : "", item.sourceName || "", item.rowNumber || "", item.identity || "", item.field || "", item.message])]);
  add("Date normalizate", [["Nume", "Cod evaluare", "CODE", "Regiune", "Identitate stabilă", "Sursă", "Rând", "Includere", ...payload.competencies], ...payload.auditRecords.map((record) => [record.name, record.assessment, record.code, record.region, record.identity, record.source, record.rowNumber, record.inclusion, ...payload.competencies.map((competency) => record.scores[competency] ?? "")])]);
  add("Calcule", [["Competență", "N", "Medie 1–5", "Min", "Mediană", "Max"], ...payload.calculations.map((item) => [item.competency, item.n, item.mean, item.min, item.median, item.max])]);
  add("Comportamente", [["Competență", "Comportament", "Text importat (brut)", "N", "Sumă scoruri prezente", "Medie 0–2", "% scor 0", "% scor 2", "Descriptor scor 0", "Descriptor scor 2", "Sursă descriptor"], ...payload.behaviorAggregates.map((item) => [item.competency, item.behavior, item.behaviorRaw || item.behavior, item.n, item.sum, item.mean, item.pct0, item.pct2, item.score0, item.score2, item.descriptorSource])]);
  add("Grupuri", [["CODE", "N", "Nume afișat"], ...payload.groups.map((group) => [group.code, group.records.length, payload.metadata.groupNames?.[group.code] || group.code])]);
  add("Zone", [["Regiune", "N", ...payload.competencies], ...payload.zones.map((zone) => [zone.region, zone.records.length, ...payload.competencies.map((competency) => scoreStats(zone.records, competency).mean ?? "")])]);
  // R5 ranking basis (Vlad 2026-10-02): mean over scored participants, share scoring 2, share scoring 0; ordered within
  // each competency by mean desc, then share 2 desc, then column order.
  const ranking = payload.behaviorAggregates.slice().sort((a, b) => payload.competencies.indexOf(a.competency) - payload.competencies.indexOf(b.competency) || (b.mean ?? 0) - (a.mean ?? 0) || b.pct2 - a.pct2 || a.sourceIndex - b.sourceIndex);
  add("Clasament", [["Competență", "Comportament", "Medie 0–2", "% scor 2", "% scor 0", "N scorați", "Sumă"], ...ranking.map((item) => [item.competency, item.behavior, item.mean, item.pct2, item.pct0, item.n, item.sum])]);
  for (const name of workbook.SheetNames) workbook.Sheets[name]["!cols"] = Array.from({ length: 16 }, (_, index) => ({ wch: index === 0 ? 32 : 20 }));
  return workbook;
}
