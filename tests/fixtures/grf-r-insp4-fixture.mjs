/* D210-labelled SYNTHETIC fixture, shape of Inspector recheck 3 (2026-10-02). No client data, no real names.
   Six competencies (10/9/6/5/4/3 behaviours), DIR 11 / SRD 7 / blank 2, three regions + blanks, an all-missing
   behaviour, ties at the R5 cuts (one tie whose first participant row is blank), one „(0-2) … /PPCd” label,
   long hyphenated names and a partial BOM/CRLF evaluation CSV with scope cases. */
export const SHARED = "Comunică deschis așteptările și oferă context pentru deciziile echipei";
export const CODED = "(0-2) Formulează obiective măsurabile și verifică înțelegerea lor în echipă /PPCd";
export const ALLMISS = "Comportament fără niciun scor importat (rând complet gol în export)";
// Competency → behaviours as [label, n2, n1, n0] (counts over the 20 detailed rows; blanks reduce n).
export const COMPETENCIES = [
  { name: "Orientare strategică și gândire sistemică în contexte incerte și schimbătoare", beh: [
    ["Anticipează efectele de durată ale deciziilor asupra organizației", 14, 4, 2],
    [SHARED, 12, 6, 2],
    ["Identifică tiparele comune din informații fragmentate și contradictorii", 10, 8, 2],
    ["Corelează obiectivele echipei cu direcția strategică a companiei", 9, 8, 3],
    ["Formulează scenarii alternative înainte de a alege o direcție", 6, 12, 2],   // sum 24  (tie A)
    ["Prioritizează inițiativele după impactul lor pe termen lung", 8, 8, 4],     // sum 24  (tie B, later column)
    ["Evaluează riscurile și oportunitățile pieței în mod structurat", 6, 8, 6],
    ["Comunică viziunea într-un mod care mobilizează oamenii", 4, 10, 6],
    ["Revizuiește planurile pe baza semnalelor noi din mediu", 3, 8, 9],
    ["Ține cont de interdependențele dintre departamente și procese", 2, 6, 12] ] },
  { name: "Colaborare transfuncțională și influențare fără autoritate formală", beh: [
    ["Construiește relații de încredere cu colegii din alte departamente", 13, 5, 2], // 31
    [SHARED, 11, 6, 3],                                                              // 28
    ["Negociază soluții reciproc avantajoase în situații tensionate", 9, 8, 3],      // 26 tie A (row 1 blank)
    ["Ascultă activ și reformulează punctele de vedere ale celorlalți", 10, 6, 4],   // 26 tie B (later column)
    ["Implică persoanele potrivite la momentul potrivit", 7, 9, 4],
    ["Gestionează constructiv dezacordurile din echipă", 6, 8, 6],                   // 20 tie C
    ["Oferă sprijin colegilor fără să fie solicitat", 5, 10, 5],                    // 20 tie D (later)
    ["Influențează deciziile prin argumente bazate pe date", 3, 9, 8],
    ["Recunoaște public contribuția celorlalți", 2, 7, 11] ] },
  { name: "Luarea deciziilor în condiții de presiune, ambiguitate și informații incomplete", beh: [
    ["Ia decizii la timp chiar și cu informații incomplete", 12, 6, 2],
    ["Cântărește alternativele pe baza unor criterii explicite", 10, 6, 4],
    [ALLMISS, 0, 0, 0],
    ["Își asumă responsabilitatea pentru consecințele deciziilor", 8, 6, 6],
    ["Revine asupra deciziilor atunci când apar date noi relevante", 5, 8, 7],
    ["Explică echipei raționamentul din spatele deciziilor", 3, 6, 11] ] },
  { name: "Dezvoltarea oamenilor și feedback constructiv", beh: [
    ["Oferă feedback specific, echilibrat și orientat spre comportament", 11, 6, 3], // 28
    [CODED, 9, 8, 3],                                                                // 26
    ["Deleagă sarcini care dezvoltă competențele colegilor", 8, 6, 6],               // 22 (median)
    ["Creează oportunități de învățare din greșeli", 6, 8, 6],                       // 20
    ["Urmărește progresul individual și ajustează sprijinul oferit", 4, 6, 10] ] },  // 14
  { name: "Orientare către client și rezultate", beh: [
    ["Înțelege nevoile reale ale clientului dincolo de cerința explicită", 12, 4, 4],
    ["Stabilește standarde ridicate de calitate pentru livrabile", 9, 6, 5],
    ["Urmărește indicatorii de rezultat și corectează abaterile", 7, 6, 7],
    ["Menține angajamentele asumate față de client", 4, 8, 8] ] },
  { name: "Reziliență și adaptabilitate", beh: [
    ["Își păstrează calmul și claritatea în situații de criză", 11, 6, 3],
    ["Se adaptează rapid la priorități schimbate", 8, 6, 6],
    ["Învață din eșecuri și revine cu soluții noi", 4, 8, 8] ] }
];

const NAMES = [
  "Participant Sintetic 01", "Participant Sintetic 02", "Persoana-Sintetică-Cu-Nume-Compus-Foarte-Lung 03", "Participant Sintetic 04",
  "Participant Sintetic 05", "Participant Sintetic 06", "Participant Sintetic 07", "Participantă Sintetică Ștefănescu-Țărănuș 08",
  "Participant Sintetic 09", "Participant Sintetic 10", "Participant Sintetic 11", "Participant Sintetic 12",
  "Participant Sintetic 13", "Persoana-Sintetică-Dublu-Nume-Bărbulescu-Constantinescu 14", "Participant Sintetic 15", "Participant Sintetic 16",
  "Participant Sintetic 17", "Participant Sintetic 18", "Participant Sintetic 19", "Participant Sintetic 20"];
// CODE: DIR ×11, SRD ×7, blank ×2.  Regions: 3 named + 3 blank.
const CODES = ["DIR","DIR","DIR","DIR","DIR","DIR","DIR","DIR","DIR","DIR","DIR","SRD","SRD","SRD","SRD","SRD","SRD","SRD","",""];
const REGIONS = ["București-Ilfov","Cluj","Iași","București-Ilfov","Cluj","Iași","București-Ilfov","Cluj","","Iași","București-Ilfov","Cluj","Iași","București-Ilfov","Cluj","","Iași","București-Ilfov","Cluj",""];
// Summary 1–5, varied per competency so means, medians and shares differ; medians strictly inside min–max.
const SUMMARY = [
  [5,4,4,5,3,4], [4,4,3,5,4,5], [4,5,4,4,3,4], [5,3,4,4,4,3], [3,4,3,4,5,4], [4,3,5,3,4,4], [3,4,3,4,3,null],
  [4,2,4,3,4,3], [3,3,2,4,3,3], [2,4,3,3,4,3], [4,3,4,2,3,3],
  [2,2,3,2,2,3], [3,1,2,3,2,2], [1,2,2,2,3,2], [2,3,1,2,2,1], [3,2,2,1,2,2], [2,1,3,3,1,2], [1,2,2,2,2,3],
  [3,3,3,3,3,3], [2,3,2,3,3,2]];
export const participants = NAMES.map((name, index) => ({ name, code: CODES[index], region: REGIONS[index], assessment: `SYN-AC-${String(index + 1).padStart(2, "0")}`, scores: SUMMARY[index] }));

// Behaviour score vectors: n2 twos, n1 ones, n0 zeros, rotated by behaviour so participants differ.
function vector(n2, n1, n0, shift, blankRow1) {
  const base = [...Array(n2).fill(2), ...Array(n1).fill(1), ...Array(n0).fill(0)];
  while (base.length < 20) base.push("");
  const out = Array.from({ length: 20 }, (_, i) => base[(i + shift) % 20]);
  if (blankRow1) { // move participant 1's value away so row 1 is blank but the sum is unchanged
    const value = out[0]; out[0] = ""; const free = out.findIndex((v, i) => i > 0 && v === "");
    if (free >= 0) out[free] = value; else throw new Error("no room");
  }
  return out;
}
export const columns = COMPETENCIES.flatMap((c, ci) => c.beh.map(([label, n2, n1, n0], bi) => {
  const blankRow1 = ci === 1 && bi === 2; // C2 tie A: participant 1 blank
  const total = n2 + n1 + n0;
  const counts = total === 20 && blankRow1 ? [n2, n1, n0 - 1] : [n2, n1, n0];
  // keep a blank slot for the row-1 move when needed
  const v = label === ALLMISS ? Array(20).fill("") : vector(...counts, (ci * 7 + bi * 3) % 20, blankRow1);
  if (blankRow1 && total === 20) { /* sum preserved: we removed one 0, sum unchanged */ }
  return { competency: c.name, sub: `Subcompetența ${ci + 1}.${Math.floor(bi / 3) + 1}`, behavior: label, values: v };
}));

export function rows() {
  const summary = [["CODE", "name", "job", "email", "cod cp", ...COMPETENCIES.map((c) => c.name)],
    ...participants.map((p) => [p.code, p.name, "Rol sintetic", "", p.assessment, ...p.scores.map((s) => s ?? "")])];
  const head = ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente"];
  let last = "";
  const groupRow = [...head, ...columns.map((col) => { const v = col.competency !== last ? col.competency : ""; last = col.competency; return v; })];
  const subRow = ["", "", "", "", "Subcompetente", ...columns.map((col) => col.sub)];
  const behRow = ["", "", "", "", "behavior", ...columns.map((col) => col.behavior)];
  const detailed = [groupRow, subRow, behRow, ...participants.map((p, i) => [p.code, p.name, p.region, p.assessment, "", ...columns.map((col) => col.values[i])])];
  return { summary, detailed };
}
const q = (v) => `"${String(v).replaceAll('"', '""')}"`;
export function csv() {
  const H = ["competency", "subcompetency", "behavior", "objective_text_score_0", "objective_text_score_-1", "objective_text_score_1", "objective_text_score_2"];
  const c = COMPETENCIES;
  const lines = [H.join(",")];
  // Shared behaviour declined ONLY under C1.
  lines.push([c[0].name, "Subcompetența 1.1", SHARED, "Să comunici deschis așteptările, oferind context", "", "", "Ai comunicat deschis așteptările, oferind context"].map(q).join(","));
  // C1 top behaviour, with a comma and a newline inside quotes.
  lines.push([c[0].name, "Subcompetența 1.1", c[0].beh[0][0], "Să anticipezi efectele deciziilor, inclusiv\npe termen lung", "", "", "Ai anticipat efectele deciziilor, inclusiv pe termen lung"].map(q).join(","));
  // C4 behaviour with NO competency named (text-alone match expected by DoD 14).
  lines.push(["", "", c[3].beh[0][0], "Să oferi feedback specific și echilibrat", "", "", "Ai oferit feedback specific și echilibrat"].map(q).join(","));
  // C5 behaviour duplicated with different texts → ambiguous → fallback + warning.
  lines.push([c[4].name, "", c[4].beh[0][0], "Să înțelegi nevoile clientului (varianta A)", "", "", "Ai înțeles nevoile clientului (varianta A)"].map(q).join(","));
  lines.push([c[4].name, "", c[4].beh[0][0], "Să înțelegi nevoile clientului (varianta B)", "", "", "Ai înțeles nevoile clientului (varianta B)"].map(q).join(","));
  // C2 bottom behaviour: only score-0 text, score-2 empty (partial declination).
  lines.push([c[1].name, "", c[1].beh[8][0], "Să recunoști public contribuția celorlalți", "", "", ""].map(q).join(","));
  // C6 middle: unmatched row (no such behaviour).
  lines.push([c[5].name, "", "Comportament inexistent în export", "Să nu apară", "", "", "Nu trebuie să apară"].map(q).join(","));
  return `﻿${lines.join("\r\n")}\r\n`;
}
export function createFixture(XLSX) {
  const { summary, detailed } = rows();
  const book = (r) => { const b = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(b, XLSX.utils.aoa_to_sheet(r), "Sheet1"); return XLSX.write(b, { type: "buffer", bookType: "xlsx" }); };
  return { summary: book(summary), detailed: book(detailed), csv: new TextEncoder().encode(csv()), metadata: { projectName: "Proiect sintetic inspecție", clientName: "Client sintetic cu denumire românească foarte lungă", annex: "end", splitGroups: true } };
}
