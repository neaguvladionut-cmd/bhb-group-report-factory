/* D210-labelled synthetic acceptance data only. No client data. */
import { EVAL_SHEET_HEADERS } from "../../src/rebuild-core.js";

export const competencies = [
  { name: "Leadership strategică și coordonare operațională", count: 10 },
  { name: "Colaborare interculturală și comunicare incluzivă", count: 9 },
  { name: "Rezolvare analitică a problemelor complexe", count: 6 },
  { name: "Adaptare la schimbare și priorități concurente", count: 5 }
];

const longLabels = {
  "Leadership strategică și coordonare operațională": [
    "Comportament comun pentru două competențe cu formulare românească intenționat foarte lungă",
    ...Array.from({ length: 9 }, (_, index) => `Leadership comportamental ${index + 2} cu descriere românească extinsă pentru verificarea încadrării în suprafață`)
  ],
  "Colaborare interculturală și comunicare incluzivă": [
    "Comportament comun pentru două competențe cu formulare românească intenționat foarte lungă",
    ...Array.from({ length: 8 }, (_, index) => `Colaborare comportamentală ${index + 2} cu descriere românească extinsă pentru verificarea încadrării în suprafață`)
  ],
  "Rezolvare analitică a problemelor complexe": Array.from({ length: 6 }, (_, index) => index === 5 ? "Comportament fără scoruri importate pentru verificarea rândului all-missing" : `Rezolvare comportamentală ${index + 1} cu etichetă românească lungă și diacritice`),
  "Adaptare la schimbare și priorități concurente": Array.from({ length: 5 }, (_, index) => `Adaptare comportamentală ${index + 1} cu etichetă românească lungă și diacritice`)
};

export const participants = Array.from({ length: 20 }, (_, index) => {
  const number = index + 1;
  return {
    code: number <= 12 ? "NORD" : number <= 19 ? "SUD" : "",
    name: `Participant sintetic ${String(number).padStart(2, "0")}`,
    region: number <= 6 ? "București" : number <= 12 ? "Cluj-Napoca" : number <= 18 ? "Timișoara" : "",
    assessment: `SYN-${String(number).padStart(2, "0")}`,
    score: number <= 12 ? 4 : number <= 19 ? 2 : 3
  };
});

export const behaviorRows = competencies.flatMap(({ name }) => longLabels[name].map((behavior, index) => ({ competency: name, subcompetency: `Subcompetență ${name.slice(0, 18)}`, behavior, index })));

// Varies by behaviour so the shares of 2 and 0 differ between behaviours (key findings show distinct %).
export const behaviourScore = (participant, participantIndex, row) => (participant.score * (row.index % 3 + 1) + Math.floor(participantIndex * (row.index + 2) / 3)) % 3;

export function fixtureRows() {
  const summaryHeaders = ["CODE", "name", "cod cp", ...competencies.map(({ name }) => name)];
  const headers = ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", ...behaviorRows.map((row) => row.competency)];
  const subcompetencies = ["", "", "", "", "Subcompetente", ...behaviorRows.map((row) => row.subcompetency)];
  const behaviors = ["", "", "", "", "behavior", ...behaviorRows.map((row) => row.behavior)];
  const summary = [summaryHeaders, ...participants.map((participant) => [participant.code, participant.name, participant.assessment, participant.score, participant.score, participant.score, participant.score])];
  const detailed = [
    headers,
    subcompetencies,
    behaviors,
    ...participants.map((participant, participantIndex) => [participant.code, participant.name, participant.region, participant.assessment, "", ...behaviorRows.map((row) => row.competency === competencies[2].name && row.behavior === longLabels[competencies[2].name].at(-1) ? "" : behaviourScore(participant, participantIndex, row))])
  ];
  return { summary, detailed, behaviorRows };
}

export function evaluationCsv() {
  const lines = [EVAL_SHEET_HEADERS.join(",")];
  for (const row of behaviorRows) lines.push([row.competency, row.subcompetency, row.behavior, `Să exersezi ${row.behavior}`, "", "", `Ai demonstrat ${row.behavior}`].map((value) => `"${String(value).replaceAll('"', '""')}"`).join(","));
  return `\ufeff${lines.join("\r\n")}\r\n`;
}

export function createFixture(XLSX) {
  const { summary, detailed } = fixtureRows();
  const workbook = (rows) => { const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1"); return XLSX.write(book, { type: "buffer", bookType: "xlsx" }); };
  return {
    summary: workbook(summary),
    detailed: workbook(detailed),
    csv: new TextEncoder().encode(evaluationCsv()),
    metadata: { projectName: "Proiect sintetic GRF-R", clientName: "Client sintetic cu etichete românești lungi", annex: "end", splitGroups: true }
  };
}
