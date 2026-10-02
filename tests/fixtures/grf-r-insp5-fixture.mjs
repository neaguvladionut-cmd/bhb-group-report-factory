/* D210-labelled SYNTHETIC fixture, shape of Inspector recheck 4 (2026-10-02). No client data, no real names.
   22 participants (MGR 12 / SPC 7 / blank 3), six competencies (10/9/6/5/4/2 behaviours), an all-blank column,
   ties at the cuts, a „(0-2)” label, long hyphenated names and a partial CSV with scope, ambiguity and an unmatched
   row. `createFixture(XLSX, { token: true })` appends „ /XYZd” to every behaviour (a uniform export code). */
const N = 22;
export const S = "Clarifică rolurile și responsabilitățile fiecărui membru înainte de începerea lucrului";
// [label, n2, n1, n0, blankRow1]
export const COMPS = [
  { name: "Gândire analitică și sinteză a informațiilor complexe din surse multiple și contradictorii", beh: [
    ["Descompune problemele complexe în componente ușor de analizat și de prioritizat", 16, 4, 2],      // 36
    [S, 14, 6, 2],                                                                                     // 34
    ["Verifică acuratețea datelor înainte de a formula concluzii pentru echipă", 12, 7, 3],             // 31
    ["Identifică relațiile cauză-efect dintre fenomenele observate în organizație", 10, 9, 3],          // 29
    ["Sintetizează informațiile esențiale într-un mesaj clar, structurat și ușor de urmărit", 9, 9, 3, true], // 27 tie A (row1 blank, earlier col)
    ["Formulează ipoteze alternative și le testează în mod sistematic", 6, 10, 6],                     // 22
    ["Folosește indicatori cantitativi pentru a susține recomandările făcute", 4, 10, 8],              // 18
    ["Recunoaște limitele propriei analize și cere opinii suplimentare", 3, 9, 10],                    // 15
    ["Compară soluțiile posibile pe baza unor criterii explicite și transparente", 8, 11, 3],          // 27 tie B (later col)
    ["Documentează pașii analizei astfel încât să poată fi reluați de alții", 2, 8, 12] ] },          // 12
  { name: "Coordonarea echipei și delegarea responsabilităților în proiecte transversale", beh: [
    ["Stabilește obiective clare și termene realiste pentru fiecare etapă a proiectului", 13, 7, 2],   // 33
    [S, 11, 8, 3],                                                                                     // 30
    ["Monitorizează progresul echipei fără a interveni excesiv în detalii", 8, 10, 4],                 // 26 tie (earlier)
    ["Ajustează alocarea sarcinilor în funcție de încărcarea fiecăruia", 7, 10, 5],                     // 24
    ["Recunoaște și valorizează contribuția fiecărui coleg în ședințele de echipă", 8, 10, 3, true],   // 26 tie (later, row1 blank)
    ["Gestionează conflictele de prioritate dintre membrii echipei", 4, 9, 9],                         // 17 tie
    ["Comunică schimbările de direcție cu suficient timp înainte", 3, 11, 8],                          // 17 tie (later)
    ["Oferă autonomie oamenilor experimentați în modul de lucru", 5, 10, 7],                            // 20
    ["Organizează ședințe scurte, cu agendă și decizii consemnate", 1, 8, 13] ] },                     // 10
  { name: "Comunicare persuasivă și influențare în relația cu partenerii externi", beh: [
    ["(0-2) Adaptează mesajul la nivelul de cunoaștere și interesele interlocutorului", 11, 8, 3],     // 30
    ["Susține argumentele cu exemple concrete și date verificabile", 6, 10, 6],                        // 22 tie (earlier)
    ["Gestionează obiecțiile cu calm și fără a deveni defensiv", 4, 10, 8],                             // 18
    ["Ascultă activ și reformulează pentru a confirma înțelegerea", 10, 8, 4],                          // 28
    ["Construiește acorduri pornind de la interesele comune ale părților", 7, 8, 7],                    // 22 tie (later)
    ["Încheie discuțiile cu pași următori clari și asumați de ambele părți", 2, 10, 10] ] },           // 14
  { name: "Inițiativă și orientare spre îmbunătățirea continuă", beh: [
    ["Propune îmbunătățiri ale proceselor fără să aștepte solicitări formale", 9, 8, 5],               // 26
    ["Își asumă sarcini suplimentare atunci când echipa are nevoie", 5, 10, 7],                         // 20
    ["Comportament fără niciun scor în exportul sintetic (coloană complet goală)", 0, 0, 0],           // all-missing
    ["Testează idei noi la scară mică înainte de a le extinde", 3, 9, 10],                              // 15
    ["Caută activ feedback asupra propriilor rezultate", 7, 9, 6] ] },                                  // 23
  { name: "Integritate și responsabilitate profesională", beh: [
    ["Își respectă angajamentele chiar și sub presiunea termenelor", 8, 8, 6],                          // 24
    ["Recunoaște deschis propriile greșeli și propune corecții", 5, 9, 8],                              // 19 tie (earlier)
    ["Tratează informațiile confidențiale cu discreție", 6, 7, 9],                                      // 19 tie (later)
    ["Semnalează la timp riscurile etice observate în activitate", 2, 7, 13] ] },                      // 11
  { name: "Reziliență la stres", beh: [
    ["Își menține eficiența în perioadele cu volum mare de muncă", 9, 9, 4],                            // 27
    ["Își recuperează rapid energia după un eșec sau o critică", 4, 8, 10] ] }                         // 16
];
const NAMES = ["Candidat Fictiv A01", "Candidat Fictiv A02", "Candidată-Fictivă-Cu-Prenume-Compus Lungulescu-Ștrengărescu 03", "Candidat Fictiv A04",
  "Candidat Fictiv A05", "Candidat Fictiv A06", "Candidat Fictiv A07", "Candidat Fictiv A08", "Candidat Fictiv A09", "Candidat Fictiv A10",
  "Candidat Fictiv A11", "Candidat Fictiv A12", "Candidat Fictiv A13", "Candidat Fictiv A14", "Candidat-Fictiv-Bărbuță-Constantinovici-Dumitrescu 15",
  "Candidat Fictiv A16", "Candidat Fictiv A17", "Candidat Fictiv A18", "Candidat Fictiv A19", "Candidat Fictiv A20", "Candidat Fictiv A21", "Candidat Fictiv A22"];
// MGR 12, SPC 7, blank 3
const CODES = ["MGR","SPC","MGR","MGR","SPC","MGR","","MGR","SPC","MGR","MGR","SPC","MGR","","MGR","SPC","MGR","SPC","MGR","","MGR","SPC"];
const REGIONS = ["Timiș","Brașov","Constanța","Dolj","Timiș","","Brașov","Constanța","Dolj","Timiș","Brașov","","Constanța","Dolj","Timiș","Brașov","Constanța","","Dolj","Timiș","Brașov","Constanța"];
const SUMMARY = [
  [4,3,5,2,4,3],[3,4,4,3,5,2],[5,2,3,4,3,4],[2,3,4,5,4,3],[4,5,3,3,2,4],[3,2,2,4,3,5],[5,4,4,2,3,3],[2,3,3,3,4,4],
  [3,1,4,2,3,2],[4,3,2,5,4,3],[1,4,3,3,2,3],[3,3,5,4,3,null],[4,2,3,3,5,4],[2,5,2,1,3,3],[3,3,4,4,4,2],[5,2,3,3,2,3],
  [3,4,1,2,3,4],[2,3,3,4,3,3],[4,4,4,3,1,5],[3,2,3,5,4,2],[2,3,2,3,3,3],[3,4,3,2,4,3]];
export const participants = NAMES.map((name, i) => ({ name, code: CODES[i], region: REGIONS[i], assessment: `FICT-AC-${String(i + 1).padStart(3, "0")}`, scores: SUMMARY[i] }));
function vector(n2, n1, n0, shift, blankRow1) {
  const base = [...Array(n2).fill(2), ...Array(n1).fill(1), ...Array(n0).fill(0)];
  while (base.length < N) base.push("");
  let out = Array.from({ length: N }, (_, i) => base[(i + shift) % N]);
  if (blankRow1 && out[0] !== "") { const free = out.findIndex((v, i) => i > 0 && v === ""); [out[0], out[free]] = ["", out[0]]; }
  return out;
}
export function columns(transform = (l) => l) {
  return COMPS.flatMap((c, ci) => c.beh.map(([label, n2, n1, n0, b1], bi) => ({ competency: c.name, sub: `Subcompetența ${ci + 1}.${Math.floor(bi / 4) + 1}`, behavior: transform(label), clean: label.replace(/^\(0-2\)\s*/, ""),
    values: n2 + n1 + n0 === 0 ? Array(N).fill("") : vector(n2, n1, n0, (ci * 5 + bi * 3) % N, b1) })));
}
export function rows(transform) {
  const cols = columns(transform);
  const summary = [["CODE", "name", "job", "email", "cod cp", ...COMPS.map((c) => c.name)], ...participants.map((p) => [p.code, p.name, "Funcție fictivă", "", p.assessment, ...p.scores.map((s) => s ?? "")])];
  let last = ""; const head = ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente"];
  const groupRow = [...head, ...cols.map((c) => { const v = c.competency !== last ? c.competency : ""; last = c.competency; return v; })];
  const detailed = [groupRow, ["", "", "", "", "Subcompetente", ...cols.map((c) => c.sub)], ["", "", "", "", "behavior", ...cols.map((c) => c.behavior)],
    ...participants.map((p, i) => [p.code, p.name, p.region, p.assessment, "", ...cols.map((c) => c.values[i])])];
  return { summary, detailed, cols };
}
const q = (v) => `"${String(v).replaceAll('"', '""')}"`;
export function csv(raw = (l) => l) {
  const H = ["competency", "subcompetency", "behavior", "objective_text_score_0", "objective_text_score_-1", "objective_text_score_1", "objective_text_score_2"];
  const c = COMPS; const L = [H.join(",")];
  L.push([c[0].name, "Subcompetența 1.1", raw(S), "Să clarifici rolurile, înainte de start", "", "", "Ai clarificat rolurile, înainte de start"].map(q).join(","));         // scoped C1 only
  L.push(["", "", raw(c[2].beh[3][0]), "Să asculți activ și să reformulezi", "", "", "Ai ascultat activ și ai reformulat,\nconfirmând înțelegerea"].map(q).join(",")); // blank competency, newline
  L.push(["", "", raw(c[4].beh[0][0]), "Să îți respecți angajamentele (A)", "", "", "Ți-ai respectat angajamentele (A)"].map(q).join(","));                         // ambiguous 1
  L.push(["", "", raw(c[4].beh[0][0]), "Să îți respecți angajamentele (B)", "", "", "Ți-ai respectat angajamentele (B)"].map(q).join(","));                         // ambiguous 2
  L.push([c[0].name, "", raw(c[0].beh[9][0]), "Să documentezi pașii analizei", "", "", ""].map(q).join(","));                                                      // partial, bottom
  L.push([c[1].name, "", raw(c[1].beh[0][0]), "", "", "", "Ai stabilit obiective clare și termene realiste"].map(q).join(","));                                     // partial, top
  L.push([c[2].name, "", raw(c[2].beh[0][0]), "Să adaptezi mesajul", "", "", "Ai adaptat mesajul interlocutorului"].map(q).join(","));                            // (0-2) raw label
  L.push([c[3].name, "", "Comportament care nu există în export", "Să faci X", "", "", "Ai făcut X"].map(q).join(","));                                            // unmatched
  return "﻿" + L.join("\r\n") + "\r\n";
}
export function createFixture(XLSX, { token = false } = {}) {
  const transform = token ? (label) => `${label} /XYZd` : (label) => label;
  const { summary, detailed } = rows(transform);
  const book = (aoa) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Sheet1"); return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }); };
  return { summary: book(summary), detailed: book(detailed), csv: new TextEncoder().encode(csv(transform)), metadata: { projectName: "Proiect fictiv inspecție 4", clientName: "Client fictiv pentru verificarea completă a raportului", annex: "end", splitGroups: true } };
}
