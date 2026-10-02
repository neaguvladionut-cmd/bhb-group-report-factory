/* D210-labelled synthetic data only. No client data. Second fixture for the fill map: six competencies with
   distinct per-competency scores (so sorting, series order and colours are visible), 23 participants (rule 10
   splits), four regions (more zone series than the template's three) and 10/9/6/5/4/3 behaviours. */
import { EVAL_SHEET_HEADERS } from "../../src/rebuild-core.js";

export const competencies = [
  { name: "Orientare spre client și calitatea serviciului", count: 10 },
  { name: "Comunicare și influență", count: 9 },
  { name: "Gândire analitică", count: 6 },
  { name: "Planificare și organizare", count: 5 },
  { name: "Lucru în echipă", count: 4 },
  { name: "Inițiativă și asumarea responsabilității pentru rezultate", count: 3 }
];
const regions = ["Nord-Est", "Vest", "Sud", "Centru"];
const scoreSteps = [2.25, 3.5, 4.25, 4.75, 3.25, 2.5, 4, 3.75, 1.5, 4.5, 3, 5, 2.75]; // quarter steps
export const participants = Array.from({ length: 23 }, (_, index) => {
  const number = index + 1;
  return {
    code: number <= 11 ? "ALFA" : number <= 21 ? "BETA" : "",
    name: `Persoană test ${String(number).padStart(2, "0")}`,
    region: number === 23 ? "" : regions[index % regions.length],
    assessment: `VAR-${String(number).padStart(2, "0")}`,
    scores: competencies.map((_, competencyIndex) => scoreSteps[(index * 5 + competencyIndex * 3 + (competencyIndex === 0 ? 4 : 0)) % scoreSteps.length])
  };
});
export const behaviorRows = competencies.flatMap(({ name, count }, competencyIndex) => Array.from({ length: count }, (_, index) => ({ competency: name, subcompetency: `Sub ${competencyIndex + 1}`, behavior: `Comportament ${competencyIndex + 1}.${index + 1} – formulare observabilă pentru verificarea tabelelor`, index, competencyIndex })));

export function fixtureRows() {
  const summary = [["CODE", "name", "cod cp", ...competencies.map(({ name }) => name)], ...participants.map((participant) => [participant.code, participant.name, participant.assessment, ...participant.scores])];
  const detailed = [
    ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", ...behaviorRows.map((row) => row.competency)],
    ["", "", "", "", "Subcompetente", ...behaviorRows.map((row) => row.subcompetency)],
    ["", "", "", "", "behavior", ...behaviorRows.map((row) => row.behavior)],
    ...participants.map((participant, participantIndex) => [participant.code, participant.name, participant.region, participant.assessment, "", ...behaviorRows.map((row) => (participantIndex + row.index * 2 + row.competencyIndex) % 3)])
  ];
  return { summary, detailed };
}
export function evaluationCsv() {
  const lines = [EVAL_SHEET_HEADERS.join(",")];
  for (const row of behaviorRows.filter((_, index) => index % 2 === 0)) lines.push([row.competency, row.subcompetency, row.behavior, `Să exersezi: ${row.behavior}`, "", "", `Ai demonstrat: ${row.behavior}`].map((value) => `"${value}"`).join(","));
  return `${lines.join("\r\n")}\r\n`;
}
export function createFixture(XLSX) {
  const { summary, detailed } = fixtureRows();
  const workbook = (rows) => { const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1"); return XLSX.write(book, { type: "buffer", bookType: "xlsx" }); };
  return { summary: workbook(summary), detailed: workbook(detailed), csv: new TextEncoder().encode(evaluationCsv()), metadata: { projectName: "Proiect sintetic variat", clientName: "Client sintetic variat", annex: "end", splitGroups: true, evaluators: "4", days: "2", exercises: "un exercițiu de grup, un studiu de caz cu prezentare, o simulare managerială", conclusionsInterventions: "Text de test al consultantului pentru intervenții." } };
}
