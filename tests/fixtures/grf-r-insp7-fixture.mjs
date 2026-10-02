/* D210-labelled SYNTHETIC fixture, shape of the final Inspector recheck (2026-10-02). No client data, no real names.
   33 participants with quarter-step summary scores (some overall means exactly at 2.75 and 3.50), scores at the 1.00
   axis minimum, scattered missing detailed scores and a partial CSV. */
export const N = 33;
// [label, n2, n1, n0]  (n2+n1+n0 < N → the rest are missing)
export const COMPS = [
  { name: "Gândire strategică", beh: [
    ["Formulează obiective pe termen lung aliniate cu direcția organizației", 20, 10, 3],
    ["Identifică tendințele pieței care pot influența activitatea", 18, 12, 3],
    ["Leagă deciziile curente de prioritățile strategice", 17, 12, 4],
    ["Evaluează impactul pe termen lung al opțiunilor analizate", 16, 13, 4],
    ["Propune direcții noi de dezvoltare pentru echipă", 14, 16, 3],        // 44/33 pct2 14 — earlier, loses spread tie
    ["Comunică viziunea astfel încât să fie înțeleasă de toți", 15, 14, 4],  // 44/33 pct2 15 — later, wins top 5
    ["Analizează riscurile strategice înainte de a se angaja", 8, 15, 10],    // 31/33 pct0 10
    ["Revizuiește strategia pe baza rezultatelor obținute", 7, 17, 9],        // 31/33 pct0 9
    ["Prioritizează inițiativele după valoarea adusă", 5, 13, 15],
    ["Integrează perspective din mai multe departamente", 4, 12, 17] ] },
  { name: "Comunicare și influență", beh: [
    ["Argumentează convingător folosind date concrete", 10, 2, 1],              // 22/13 = 1.69 mean top; sum 22 lowest
    ["Ascultă activ și reformulează pentru a verifica înțelegerea", 19, 10, 4],
    ["Adaptează mesajul la nivelul interlocutorului", 15, 16, 2],               // 46/33 pct2 15 earlier → loses
    ["Prezintă ideile structurat și concis", 17, 12, 4],                       // 46/33 pct2 17 later → top 3
    ["Obține sprijinul celorlalți pentru propunerile sale", 12, 15, 6],
    ["Gestionează obiecțiile fără a deveni defensiv", 4, 14, 15],
    ["Oferă feedback constructiv în mod direct", 6, 18, 9],                    // 30/33 pct0 9 earlier → loses bottom 3rd
    ["Negociază soluții acceptabile pentru ambele părți", 8, 14, 11],          // 30/33 pct0 11 later → bottom 3
    ["Folosește exemple relevante pentru a clarifica mesajul", 3, 12, 18] ] },
  { name: "Orientare spre rezultate", beh: [
    ["Își stabilește ținte ambițioase, dar realiste", 18, 10, 4],               // 32 scored
    ["Urmărește progresul față de obiective în mod regulat", 14, 15, 4],
    ["Persistă în fața obstacolelor", 12, 17, 4],                               // 41/33 pct2 12 earlier
    ["Ia inițiativa fără să aștepte instrucțiuni", 14, 13, 6],                 // 41/33 pct2 14 later → top
    ["Optimizează procesele pentru a obține rezultate mai bune", 8, 16, 9],
    ["Își asumă responsabilitatea pentru rezultatele echipei", 6, 14, 13] ] },
  { name: "Dezvoltarea oamenilor", beh: [
    ["Oferă oportunități de învățare membrilor echipei", 16, 12, 5],
    ["Recunoaște progresul colegilor", 11, 16, 6],                              // 38/33 pct2 11 earlier → middle
    ["Deleagă sarcini care dezvoltă competențe noi", 12, 14, 7],               // 38/33 pct2 12 later → top 2
    ["Discută periodic planurile de dezvoltare", 7, 15, 11],
    ["Acționează ca model de comportament", 6, 13, 14] ] },
  { name: "Lucrul în echipă", beh: [
    ["Contribuie activ la atingerea obiectivelor comune", 17, 11, 5],
    ["Comportament neobservat în această sesiune sintetică", 0, 0, 0],          // all-unscored
    ["Împărtășește informații utile cu colegii", 13, 15, 5],
    ["Sprijină colegii aflați sub presiune", 11, 15, 7],
    ["Gestionează conflictele în mod constructiv", 9, 15, 9],
    ["Respectă angajamentele asumate față de echipă", 8, 14, 11],
    ["Solicită părerea colegilor înainte de a decide", 5, 15, 13] ] },
  { name: "Adaptare la schimbare", beh: [
    ["Acceptă schimbările și caută beneficiile acestora", 15, 12, 6],
    ["Își ajustează rapid planurile când contextul se modifică", 10, 15, 8],
    ["Rămâne eficient în situații ambigue", 9, 14, 10],
    ["Propune soluții noi la probleme vechi", 6, 15, 12] ] }
];
const NAMES = ["Andreea-Maria Popescu", "Ioan Dumitrescu", "Elena Stoica", "Mihai-Alexandru Ionescu", "Cristina Georgescu", "Radu Marinescu",
  "Alina Constantin", "Bogdan Stan", "Diana-Ioana Munteanu", "Florin Rusu", "Gabriela Toma", "Lucian Dobre", "Raluca Matei", "Sorin Lazăr",
  "Oana-Cristina Barbu", "Vlad Ciobanu", "Irina Zamfir", "Cătălin Nistor", "Simona Preda", "Adrian Mocanu", "Larisa Vasilescu-Pop",
  "Tudor Enache", "Monica Stănescu", "Gheorghe Pavel", "Corina Dinu", "Ștefan Ilie", "Bianca-Elena Surdu", "Marius Oprea", "Roxana Chiriac",
  "Daniel Neagoe", "Teodora Rădulescu", "Paul Avram", "Ana Iordache"];
const CODES = ["VNZ","LOG","VNZ","","LOG","VNZ","LOG","VNZ","LOG","VNZ","","LOG","VNZ","LOG","VNZ","LOG","VNZ","LOG","","VNZ","LOG","VNZ","LOG","VNZ","LOG","VNZ","LOG","","VNZ","LOG","VNZ","LOG","VNZ"];
const REGIONS = ["Brașov","Constanța","Timiș","Bacău","","Brașov","Constanța","Timiș","Bacău","Brașov","Constanța","","Timiș","Bacău","Brașov","Constanța","Timiș","Bacău","Brașov","","Constanța","Timiș","Bacău","Brașov","Constanța","Timiș","Bacău","Brașov","Constanța","","Timiș","Bacău","Brașov"];
// deterministic quarter-step summary scores
let seed = 7; const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
const SUMMARY = NAMES.map(() => COMPS.map(() => Math.min(5, Math.max(1, Math.round((1.5 + rnd() * 3.25) * 4) / 4))));
SUMMARY[5] = [3.5, 3.5, 3.5, 3.5, 3.5, 3.5];                 // overall exactly 3.50
SUMMARY[12] = [1, 1, 2, 4.5, 3.25, 4.75];                     // overall exactly 2.75 (sum 16.5 / 6)
SUMMARY[20] = [3.25, 3.75, 3.5, 3.5, 3.25, 3.75];             // overall 3.50
SUMMARY[27] = [2.75, null, 2.75, 3, 2.5, 2.75];               // one blank summary score
SUMMARY[30] = [5, 4.75, 4.5, 5, 4.75, 5];
SUMMARY[8] = [1.5, 1.25, 1, 1.75, 1.5, 1];
export const participants = NAMES.map((name, i) => ({ name, code: CODES[i], region: REGIONS[i], assessment: `SIN7-${String(i + 1).padStart(3, "0")}`, scores: SUMMARY[i] }));
function vector(n2, n1, n0, shift) {
  const base = [...Array(n2).fill(2), ...Array(n1).fill(1), ...Array(n0).fill(0)];
  const out = Array(N).fill(""); const perm = Array.from({ length: N }, (_, i) => (i * 7 + shift) % N);
  base.forEach((value, k) => { out[perm[k]] = value; });
  return out;
}
export function columns() {
  return COMPS.flatMap((c, ci) => c.beh.map(([label, n2, n1, n0], bi) => ({ competency: c.name, sub: `Subcompetența ${ci + 1}.${Math.floor(bi / 3) + 1}`, behavior: label,
    values: n2 + n1 + n0 === 0 ? Array(N).fill("") : vector(n2, n1, n0, (ci * 5 + bi * 3) % N) })));
}
export function rows() {
  const cols = columns();
  const summary = [["CODE", "name", "job", "email", "cod cp", ...COMPS.map((c) => c.name)], ...participants.map((p) => [p.code, p.name, "Funcție sintetică", "", p.assessment, ...p.scores.map((s) => s ?? "")])];
  let last = ""; const head = ["CODE", "name the person evaluated", "regiune", "cod ac", "Competente"];
  const groupRow = [...head, ...cols.map((c) => { const v = c.competency !== last ? c.competency : ""; last = c.competency; return v; })];
  const detailed = [groupRow, ["", "", "", "", "Subcompetente", ...cols.map((c) => c.sub)], ["", "", "", "", "behavior", ...cols.map((c) => c.behavior)],
    ...participants.map((p, i) => [p.code, p.name, p.region, p.assessment, "", ...cols.map((c) => c.values[i])])];
  return { summary, detailed, cols };
}
const q = (v) => `"${String(v).replaceAll('"', '""')}"`;
export function csv() {
  const H = ["competency", "subcompetency", "behavior", "objective_text_score_0", "objective_text_score_-1", "objective_text_score_1", "objective_text_score_2"];
  const c = COMPS; const L = [H.join(",")];
  L.push([c[0].name, "", c[0].beh[0][0], "Să formulezi obiective pe termen lung", "", "", "Ai formulat obiective pe termen lung, aliniate cu direcția organizației"].map(q).join(","));
  L.push(["", "", c[1].beh[8][0], "Să folosești exemple relevante, ca să clarifici mesajul", "", "", "Ai folosit exemple relevante"].map(q).join(","));
  L.push([c[2].name, "", c[2].beh[3][0], "Să iei inițiativa", "", "", "Ai luat inițiativa, fără să aștepți instrucțiuni"].map(q).join(","));
  return "﻿" + L.join("\r\n") + "\r\n";
}
export function createFixture(XLSX) {
  const { summary, detailed } = rows();
  const book = (aoa) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Sheet1"); return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }); };
  return { summary: book(summary), detailed: book(detailed), csv: new TextEncoder().encode(csv()), metadata: { projectName: "Proiect sintetic final", clientName: "Client sintetic final", annex: "end", splitGroups: true } };
}
