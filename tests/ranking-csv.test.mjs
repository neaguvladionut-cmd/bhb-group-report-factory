// Inspector recheck 3 findings F17–F20, F28 (rulings 2026-10-02), on the inspector-shaped synthetic fixture.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import test from "node:test";
import { buildPayload, createEvaluationSheetTemplate, createAuditWorkbook, EVAL_SHEET_HEADERS } from "../src/rebuild-core.js";
import { reportPlan } from "../src/rebuild-report-plan.js";
import { ALLMISS, CODED, COMPETENCIES, SHARED, createFixture } from "./fixtures/grf-r-insp4-fixture.mjs";

const root = resolve(new URL("..", import.meta.url).pathname);
const loadUmd = async (path) => { const module = { exports: {} }; new Function("module", "exports", "require", await readFile(path, "utf8"))(module, module.exports, undefined); return module.exports; };
const XLSX = await loadUmd(resolve(root, "src/assets/vendor/xlsx.full.min.js"));
const fixture = createFixture(XLSX);
const book = (rows) => { const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Sheet1"); return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }); };
const payload = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: fixture.summary }, { name: "detail.xlsx", bytes: fixture.detailed }, { name: "evaluation-sheet-template.csv", bytes: fixture.csv }], fixture.metadata, {}, { acknowledgedWarningIds: [] });
const plan = reportPlan(payload);
const behaviorSlide = (competency) => plan.find((item) => item.family === "behavior" && item.competency === competency && !item.groupKey);

test("F17: an R5 tie is broken by the behaviour's column order, even when its first participant row is blank", () => {
  const slide = behaviorSlide(COMPETENCIES[1].name);
  assert.equal(slide.key.length, 3);
  assert.match(slide.key[2], /Negociază soluții/u);
  assert(!slide.key.some((line) => /Ascultă activ/u.test(line)));
});

test("F28 (ruling): an all-missing behaviour is excluded from the ranking and from the count that picks the cut", () => {
  const competency = COMPETENCIES[2].name; // 6 imported, 5 scored → median split 2 + 2
  const slide = behaviorSlide(competency);
  assert.equal(slide.key.length, 2); assert.equal(slide.development.length, 2);
  assert(![...slide.key, ...slide.development].some((line) => line.includes(ALLMISS)));
  assert(!payload.behaviorAggregates.some((row) => row.behavior === ALLMISS));
});

test("F18: the (0-2) annotation is stripped; a non-uniform trailing /TOKEN is kept", () => {
  const coded = payload.behaviorAggregates.find((row) => row.behaviorRaw === CODED);
  assert.equal(coded.behavior, "Formulează obiective măsurabile și verifică înțelegerea lor în echipă /PPCd");
  const audit = createAuditWorkbook(XLSX, payload);
  const rows = XLSX.utils.sheet_to_json(audit.Sheets.Comportamente, { header: 1 });
  assert(rows.some((row) => row[2] === CODED), "the audit keeps the raw imported text");
  const visible = (item) => JSON.stringify([item.title, item.key, item.development, item.strengths]);
  for (const item of plan) assert(!visible(item).includes("(0-2)"));
});

test("F18: a trailing /TOKEN shared by every behaviour is an export code and is stripped everywhere", () => {
  const summary = book([["CODE", "name", "cod cp", "A", "B"], ["", "Test Unu", "T-1", 3, 4], ["", "Test Doi", "T-2", 4, 2]]);
  const detailed = book([["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", "A", "", "B"], ["", "", "", "", "Subcompetente", "S", "S", "S"], ["", "", "", "", "behavior", "(0–2) Ascultă activ /PPCd", "Rezumă clar /PPCd", "Decide la timp /PPCd"], ["", "Test Unu", "R", "T-1", "", 2, 1, 0], ["", "Test Doi", "R", "T-2", "", 1, 2, 2]]);
  const uniform = buildPayload(XLSX, [{ name: "s.xlsx", bytes: summary }, { name: "d.xlsx", bytes: detailed }], { projectName: "Uniform" }, {}, { acknowledgedWarningIds: [] });
  assert.deepEqual(uniform.behaviorAggregates.map((row) => row.behavior), ["Ascultă activ", "Rezumă clar", "Decide la timp"]);
  const template = XLSX.utils.sheet_to_json(createEvaluationSheetTemplate(XLSX, uniform).Sheets["Evaluation sheet"], { header: 1 });
  // F38 (ruling): the template's behavior column is the RAW export text, the cross-tool key with the Dev Plan.
  assert.deepEqual(template.slice(1).map((row) => row[2]), ["(0–2) Ascultă activ /PPCd", "Rezumă clar /PPCd", "Decide la timp /PPCd"]);
  for (const item of reportPlan(uniform)) assert.doesNotMatch(JSON.stringify([item.title, item.key, item.development, item.strengths]), /\/PPCd|\(0–2\)/u);
});

test("F19: CSV scope — a named competency matches only within it; a blank competency matches on the text alone; ambiguity falls back", () => {
  const shared = payload.behaviorAggregates.filter((row) => row.behavior === SHARED);
  assert.equal(shared.find((row) => row.competency === COMPETENCIES[0].name).score2, "Ai comunicat deschis așteptările, oferind context");
  assert.equal(shared.find((row) => row.competency === COMPETENCIES[1].name).score2, "", "a row naming competency 1 must not decline the same text under competency 2");
  const blank = payload.behaviorAggregates.find((row) => row.behavior === COMPETENCIES[3].beh[0][0]);
  assert.equal(blank.score2, "Ai oferit feedback specific și echilibrat");
  const ambiguous = payload.behaviorAggregates.find((row) => row.behavior === COMPETENCIES[4].beh[0][0]);
  assert.equal(ambiguous.descriptorSource, "fallback-ambiguous"); assert.equal(ambiguous.score2, "");
  assert(payload.warnings.some((item) => item.code === "descriptor-ambiguous"));
});

test("F20: the CSV template has headers only before import and every behaviour with known texts after it", () => {
  const empty = buildPayload(XLSX, [], {}, {}, {});
  const before = XLSX.utils.sheet_to_json(createEvaluationSheetTemplate(XLSX, empty).Sheets["Evaluation sheet"], { header: 1 });
  assert.deepEqual(before, [EVAL_SHEET_HEADERS]);
  const after = XLSX.utils.sheet_to_json(createEvaluationSheetTemplate(XLSX, payload).Sheets["Evaluation sheet"], { header: 1, defval: "" });
  assert.equal(after.length - 1, COMPETENCIES.reduce((sum, competency) => sum + competency.beh.length, 0));
  const declined = after.slice(1).filter((row) => row[3] || row[6]);
  assert(declined.length >= 4, `known declinations are carried (${declined.length})`);
  const sharedUnderC2 = after.find((row) => row[0] === COMPETENCIES[1].name && row[2] === SHARED);
  assert.equal(sharedUnderC2[6], "");
});

test("F37/F38/F40: unmatched CSV row warned with its row; CSV template keys on the raw text; an all-blank behaviour is one warning", async () => {
  const { createFixture: insp5Fixture } = await import("./fixtures/grf-r-insp5-fixture.mjs");
  for (const token of [false, true]) {
    const source = insp5Fixture(XLSX, { token });
    const result = buildPayload(XLSX, [{ name: "summary.xlsx", bytes: source.summary }, { name: "detail.xlsx", bytes: source.detailed }, { name: "evaluation-sheet-template.csv", bytes: source.csv }], source.metadata, {}, { acknowledgedWarningIds: [] });
    const unmatched = result.warnings.filter((item) => item.code === "descriptor-unmatched");
    assert.equal(unmatched.length, 1); assert.match(unmatched[0].message, /Rândul \d+ .*„Comportament care nu există în export/u);
    const blanks = result.warnings.filter((item) => item.code === "detailed-all-blank");
    assert.equal(blanks.length, 1); assert.match(blanks[0].message, /\(22 participanți fără scor\)/u);
    assert(!result.warnings.some((item) => item.code === "detailed-blank" && /niciun scor/u.test(item.behavior)));
    const template = XLSX.utils.sheet_to_json(createEvaluationSheetTemplate(XLSX, result).Sheets["Evaluation sheet"], { header: 1, defval: "" });
    if (token) assert(template.slice(1).every((row) => row[2].endsWith(" /XYZd")), "raw export text, codes included");
    assert(template.some((row) => /^\(0-2\) Adaptează mesajul/u.test(row[2])));
    const adapted = result.behaviorAggregates.find((row) => row.behaviorRaw.startsWith("(0-2) Adaptează"));
    assert.equal(adapted.score2, "Ai adaptat mesajul interlocutorului", "a CSV row with the raw text matches first");
    if (token) assert(result.behaviorAggregates.every((row) => !row.behavior.endsWith("/XYZd")), "a uniform code is stripped from client text");
  }
});

test("F56/F54: missing regions and missing competency scores are one grouped warning each; the audit lists every item", () => {
  const regions = payload.warnings.filter((item) => item.code === "region-missing");
  const unrated = payload.warnings.filter((item) => item.code === "summary-unrated");
  assert.equal(regions.length, 1); assert.equal(unrated.length, 1);
  assert.match(regions[0].message, /^\d+ participanți fără regiune în /u); assert.equal(regions[0].pairs.length, regions[0].count);
  assert.match(unrated[0].message, /^Scoruri pe competență lipsă în .*: 1 \(participant – competență\): Participant Sintetic \d\d – /u);
  const audit = createAuditWorkbook(XLSX, payload);
  const rows = XLSX.utils.sheet_to_json(audit.Sheets.Lipsuri, { header: 1 }).slice(1);
  for (const warning of payload.warnings.filter((item) => Array.isArray(item.pairs))) assert.equal(rows.filter((row) => row[0] === warning.code).length, warning.pairs.length, `${warning.code}: every item in the audit`);
  // F54: a long list is truncated in the warning but complete in the audit.
  const many = Array.from({ length: 14 }, (_, index) => ({ name: `Test ${String(index + 1).padStart(2, "0")}`, id: `T-${index + 1}` }));
  const summary = book([["CODE", "name", "cod cp", "A"], ...many.map((person) => ["", person.name, person.id, 3])]);
  const detailed = book([["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", "A", ""], ["", "", "", "", "Subcompetente", "S", "S"], ["", "", "", "", "behavior", "B1", "B2"], ...many.map((person, index) => ["", person.name, "R", person.id, "", index === 0 ? 2 : "", 1])]);
  const result = buildPayload(XLSX, [{ name: "s.xlsx", bytes: summary }, { name: "d.xlsx", bytes: detailed }], { projectName: "Lipsuri" }, {}, {});
  const blank = result.warnings.find((item) => item.code === "detailed-blank");
  assert.equal(blank.count, 13); assert.match(blank.message, /și încă 3/u);
  const lipsuri = XLSX.utils.sheet_to_json(createAuditWorkbook(XLSX, result).Sheets.Lipsuri, { header: 1 }).slice(1).filter((row) => row[0] === "detailed-blank");
  assert.equal(lipsuri.length, 13, "the audit lists all 13 pairs");
});
