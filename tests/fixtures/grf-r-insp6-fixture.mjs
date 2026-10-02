/* D210-labelled SYNTHETIC fixture, shape of Inspector recheck 5 (2026-10-02). No client data, no real names.
   25 participants (OPS/COM/blank), a competency with a single behaviour, tied competency means, a zone mean at 1.00,
   long competency names (file-name truncation) and a partial CSV. `{ token: true }` appends „ /QRSd” to every
   behaviour. */
export const N = 25;
export const SHARED = "Solicită clarificări atunci când informațiile primite sunt incomplete sau contradictorii";
// [label, n2, n1, n0, blankRow1]
export const COMPS = [
  { name: "Planificarea și organizarea resurselor operaționale în contexte cu priorități multiple și termene suprapuse", beh: [
    ["Stabilește priorități clare pe baza impactului și a urgenței fiecărei sarcini", 15, 7, 3],          // 37
    ["Elaborează planuri de lucru detaliate, cu etape, responsabili și termene verificabile", 13, 8, 4],    // 34
    [SHARED, 12, 8, 5],                                                                                 // 32
    ["Anticipează riscurile care pot întârzia livrarea și pregătește alternative", 11, 7, 7],            // 29
    ["Alocă resursele în funcție de competențele și disponibilitatea oamenilor", 9, 9, 6, true],         // 27 tie rank5 (earlier, row1 blank)
    ["Revizuiește periodic planul și îl ajustează la schimbările de context", 8, 11, 6],                 // 27 tie rank6 (later) -> skipped
    ["(0-2) Urmărește încadrarea în buget și semnalează abaterile din timp", 7, 10, 8],                   // 24 tie rank7 (earlier) -> bottom
    ["Documentează deciziile de planificare pentru a putea fi consultate ulterior", 6, 12, 7],           // 24 tie (later)
    ["Coordonează dependențele dintre echipe fără blocaje inutile", 5, 9, 11],                          // 19
    ["Folosește instrumente digitale de planificare în mod consecvent", 3, 10, 12],                     // 16
    ["Închide formal fiecare etapă, cu lecțiile învățate consemnate", 2, 8, 15] ] },                     // 12
  { name: "Relaționare și colaborare interdepartamentală", beh: [
    ["Inițiază contacte cu colegii din alte departamente pentru a rezolva probleme comune", 14, 8, 3],  // 36
    [SHARED, 12, 9, 4],                                                                                 // 33
    ["Împărtășește informațiile relevante fără să i se ceară", 10, 10, 5],                              // 30
    ["Respectă angajamentele asumate față de alte echipe", 11, 6, 8],                                   // 28
    ["Recunoaște public contribuția colegilor", 8, 10, 6],                                               // 26 tie at cut 5/6 (earlier)
    ["Gestionează constructiv dezacordurile, cu argumente, nu cu poziții", 7, 12, 6],                    // 26 tie (later)
    ["Oferă sprijin colegilor aflați sub presiune", 6, 10, 9],                                          // 22
    ["Cere feedback de la parteneri asupra colaborării", 4, 11, 10],                                    // 19
    ["Adaptează stilul de comunicare la interlocutor", 4, 9, 12],                                       // 17
    ["Construiește relații de lucru durabile, bazate pe încredere", 2, 9, 14] ] },                      // 13
  { name: "Orientare către client", beh: [
    ["Identifică nevoile reale ale clientului dincolo de cererea exprimată", 14, 7, 4],                  // 35
    ["Răspunde solicitărilor clienților în termenele promise", 12, 9, 4],                               // 33
    ["Tratează reclamațiile ca pe o oportunitate de îmbunătățire", 9, 10, 6],                            // 28 tie 3/4 (earlier)
    ["Verifică satisfacția clientului după livrare", 10, 8, 7],                                          // 28 tie (later) -> skipped median
    ["Explică clientului opțiunile, cu avantaje și limite", 7, 10, 8],                                  // 24
    ["Menține clientul informat despre stadiul solicitării", 6, 9, 10],                                  // 21
    ["Propune soluții care depășesc așteptările clientului", 3, 8, 14] ] },                              // 14
  { name: "Luarea deciziilor în condiții de incertitudine și presiune de timp", beh: [
    ["Analizează consecințele fiecărei variante înainte de a decide", 12, 9, 4],                         // 33
    ["Decide în timp util, chiar și cu informații incomplete", 8, 10, 7],                                // 26 tie (earlier)
    ["Comportament fără niciun scor în exportul sintetic al inspecției 5", 0, 0, 0],                     // all-missing
    ["Își asumă responsabilitatea pentru deciziile luate", 9, 8, 8],                                    // 26 tie (later) -> median skipped
    ["Consultă persoanele relevante înainte de deciziile importante", 6, 9, 10],                        // 21
    ["Revine asupra unei decizii atunci când apar date noi", 4, 8, 13] ] },                             // 16
  { name: "Integritate", beh: [
    ["Respectă regulile, chiar și atunci când nu este supravegheat", 13, 8, 4],                          // 34
    ["Semnalează deschis practicile neconforme", 7, 9, 9],                                               // 23
    ["Își recunoaște greșelile și propune corecții", 5, 8, 12] ] },                                      // 18
  { name: "Adaptabilitate", beh: [
    ["Se adaptează rapid la schimbările de prioritate, rămânând eficient", 10, 9, 6] ] }                  // 29
];
const NAMES = Array.from({ length: N }, (_, i) => `Participant Sintetic B${String(i + 1).padStart(2, "0")}`);
NAMES[3] = "Persoană-Fictivă-Cu-Prenume-Compus Verificărescu-Exemplificărescu-Testulescu 04";
NAMES[11] = "Participantă Sintetică cu un nume foarte lung pentru verificarea încadrării etichetelor 12";
NAMES[19] = "Ionela-Fictivă Ștefănescu-Brâncoveanu-Dumitrașcu 20";
// OPS 13, COM 9, blank 3
const CODES = ["OPS","COM","OPS","OPS","COM","","OPS","COM","OPS","OPS","COM","OPS","","COM","OPS","OPS","COM","OPS","COM","OPS","","COM","OPS","COM","OPS"];
const REGIONS = ["Cluj","Iași","Sibiu","Arad","","Cluj","Iași","Sibiu","Arad","Cluj","","Iași","Sibiu","Arad","Cluj","Iași","","Sibiu","Arad","Cluj","Iași","Sibiu","Arad","Cluj","Iași"];
const SUMMARY = [
  [4,3,5,3,4,3],[3,4,3,2,5,4],[5,3,4,4,3,2],[2,3,3,5,4,4],[3,5,4,3,2,4],[4,2,3,3,4,5],[5,4,4,3,3,5],[3,3,2,4,3,4],
  [3,1,4,2,3,3],[4,3,3,5,4,2],[2,4,3,3,2,3],[3,3,5,4,4,null],[4,2,3,3,5,4],[2,5,2,1,3,3],[3,4,4,4,4,2],[5,2,3,3,2,3],
  [3,4,1,2,3,4],[2,3,3,4,3,2],[4,4,4,3,1,5],[3,2,3,5,4,4],[2,3,2,3,3,3],[3,4,3,2,4,3],[4,3,4,4,5,4],[1,3,3,3,2,3],[3,3,4,3,3,5]];
export const participants = NAMES.map((name, i) => ({ name, code: CODES[i], region: REGIONS[i], assessment: `SINT-AC5-${String(i + 1).padStart(3, "0")}`, scores: SUMMARY[i] }));
function vector(n2, n1, n0, shift, blankRow1) {
  const base = [...Array(n2).fill(2), ...Array(n1).fill(1), ...Array(n0).fill(0)];
  while (base.length < N) base.push("");
  let out = Array.from({ length: N }, (_, i) => base[(i + shift) % N]);
  if (blankRow1 && out[0] !== "") { const free = out.findIndex((v, i) => i > 0 && v === ""); [out[0], out[free]] = ["", out[0]]; }
  return out;
}
export function columns(transform = (l) => l) {
  return COMPS.flatMap((c, ci) => c.beh.map(([label, n2, n1, n0, b1], bi) => ({ competency: c.name, sub: `Subcompetența ${ci + 1}.${Math.floor(bi / 4) + 1}`, behavior: transform(label), clean: label.replace(/^\(0-2\)\s*/, ""),
    values: n2 + n1 + n0 === 0 ? Array(N).fill("") : vector(n2, n1, n0, (ci * 7 + bi * 2) % N, b1) })));
}
export function rows(transform) {
  const cols = columns(transform);
  const summary = [["CODE", "name", "job", "email", "cod cp", ...COMPS.map((c) => c.name)], ...participants.map((p) => [p.code, p.name, "Funcție sintetică", "", p.assessment, ...p.scores.map((s) => s ?? "")])];
  let last = ""; const head = ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente"];
  const groupRow = [...head, ...cols.map((c) => { const v = c.competency !== last ? c.competency : ""; last = c.competency; return v; })];
  const detailed = [groupRow, ["", "", "", "", "Subcompetente", ...cols.map((c) => c.sub)], ["", "", "", "", "behavior", ...cols.map((c) => c.behavior)],
    ...participants.map((p, i) => [p.code, p.name, p.region, p.assessment, "", ...cols.map((c) => c.values[i])])];
  return { summary, detailed, cols };
}
const q = (v) => `"${String(v).replaceAll('"', '""')}"`;
export function csv(raw = (l) => l, rawForSome = raw) {
  const H = ["competency", "subcompetency", "behavior", "objective_text_score_0", "objective_text_score_-1", "objective_text_score_1", "objective_text_score_2"];
  const c = COMPS; const L = [H.join(",")];
  L.push([c[1].name, "Subcompetența 2.1", raw(SHARED), "Să soliciți clarificări, la timp", "", "", "Ai solicitat clarificări, la timp"].map(q).join(","));               // scoped to C2 only
  L.push(["", "", rawForSome(c[0].beh[0][0]), "Să stabilești priorități clare", "", "", "Ai stabilit priorități clare, pe baza impactului"].map(q).join(","));      // blank competency
  L.push(["", "", raw(c[2].beh[0][0]), "Să identifici nevoile reale (A)", "", "", "Ai identificat nevoile reale (A)"].map(q).join(","));                              // ambiguous 1
  L.push(["", "", raw(c[2].beh[0][0]), "Să identifici nevoile reale (B)", "", "", "Ai identificat nevoile reale (B)"].map(q).join(","));                              // ambiguous 2
  L.push([c[0].name, "", rawForSome(c[0].beh[10][0]), "Să închizi formal fiecare etapă,\ncu lecțiile consemnate", "", "", ""].map(q).join(","));                      // partial bottom, newline
  L.push([c[0].name, "", raw(c[0].beh[6][0]), "Să urmărești bugetul", "", "", "Ai urmărit bugetul"].map(q).join(","));                                              // (0-2) raw label
  L.push([c[4].name, "", raw(c[4].beh[2][0]), "Să îți recunoști greșelile", "", "", "Ți-ai recunoscut greșelile"].map(q).join(","));                               // integritate bottom
  L.push([c[3].name, "", "Comportament inexistent în exportul inspecției 5", "Să faci Y", "", "", "Ai făcut Y"].map(q).join(","));                                   // unmatched
  return "﻿" + L.join("\r\n") + "\r\n";
}
export function createFixture(XLSX, { token = false } = {}) {
  const transform = token ? (label) => `${label} /QRSd` : (label) => label;
  const { summary, detailed } = rows(transform);
  const book = (aoa) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Sheet1"); return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }); };
  const clean = (label) => label.replace(/^\(0-2\)\s*/u, "");
  return { summary: book(summary), detailed: book(detailed), csv: new TextEncoder().encode(token ? csv(transform, clean) : csv()), metadata: { projectName: "Proiect sintetic inspecție 5", clientName: "Client sintetic inspecție 5", annex: "end", splitGroups: true } };
}
