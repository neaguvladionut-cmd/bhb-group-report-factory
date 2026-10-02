if (!window.XLSX || !window.JSZip) throw new Error("Lipsesc bibliotecile locale necesare");

let files = [];
let payload = null;
let activeDownloadUrl = "";
const state = { step: 1, acknowledged: new Set(), corrections: { values: {} }, busy: false, receipts: [], projectName: "", reportDate: "", previewOpen: false };
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const text = (value) => String(value ?? "").trim();
const warningKey = (item) => item.id || `${item.code}:${item.rowNumber || ""}:${item.message}`;
const esc = (value) => String(value ?? "").replace(/[&<>"']/gu, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
const slug = (value) => text(value).normalize("NFD").replace(/[\u0300-\u036f]/gu, "").toLocaleLowerCase("ro").replace(/[^a-z0-9]+/gu, "-").replace(/^-|-$/gu, "") || "raport-grup";
const today = () => new Intl.DateTimeFormat("ro-RO").format(new Date());
const announce = (message) => { const node = $("#live-status"); if (node) node.textContent = message; };

function metadata() {
  const groupNames = Object.fromEntries($$(`[data-group-name]`).map((input) => [input.dataset.groupName, input.value]));
  const groupConclusions = {};
  $$(`[data-group-conclusion]`).forEach((input) => { (groupConclusions[input.dataset.groupConclusion] ||= {})[input.dataset.field] = input.value; });
  const slideToggles = Object.fromEntries($$(`[data-slide-toggle]`).map((input) => [input.dataset.slideToggle, input.checked]));
  return {
    projectName: $("#project-name")?.value || state.projectName,
    clientName: $("#client-name")?.value || "",
    reportDate: $("#report-date")?.value || state.reportDate || today(),
    context: $("#report-context")?.value || "",
    program: $("#program")?.value || "Centru de Dezvoltare",
    exercises: $("#exercise-list")?.value || "",
    exerciseCount: $("#exercise-count")?.value || "",
    otherInstruments: $("#other-instruments")?.value || "",
    conclusions: $("#conclusions")?.value || "",
    executiveConclusions: $("#conclusions")?.value || "",
    conclusionsStrengths: $("#conclusions-strengths")?.value || "",
    conclusionsDevelopment: $("#conclusions-development")?.value || "",
    conclusionsInterventions: $("#conclusions-interventions")?.value || "",
    methodologyText: $("#methodology-text")?.value ?? "",
    evaluators: $("#evaluators")?.value || "",
    days: $("#days")?.value || "",
    benchmarkLow: $("#benchmark-low")?.value ?? "2.75",
    benchmarkHigh: $("#benchmark-high")?.value ?? "3.5",
    annex: document.querySelector(`input[name="annex"]:checked`)?.value || "end",
    splitGroups: Boolean($("#split-groups")?.checked), groupNames, groupConclusions, slideToggles
  };
}

function download(blob, name) {
  if (activeDownloadUrl) URL.revokeObjectURL(activeDownloadUrl);
  activeDownloadUrl = URL.createObjectURL(blob);
  const link = $("#download-fallback");
  if (!link) return;
  link.href = activeDownloadUrl; link.download = name; link.textContent = `Dacă descărcarea nu a pornit, apasă aici pentru ${name}.`; link.hidden = false; link.click();
}
window.__grfDownload = download;

function requiredFilesPresent() { const kinds = new Set((payload?.schemas || []).map((schema) => schema.kind)); return kinds.has("ac-summary-1-5") && kinds.has("ac-detailed-0-2"); }
function warningComplete() { return Boolean(payload) && payload.warnings.every((item) => item.reviewed); }
function ready(step) { if (step === 1) return files.length > 0; if (step >= 2) return Boolean(payload?.readiness && warningComplete()); return false; }
function invalidate() { recompute(); }

function deriveProjectName() {
  const first = files[0]?.name || "proiect-nou";
  return first.replace(/\.[^.]+$/u, "").replace(/(?:sinteza|summary|detaliat|detail|export|raport|ac)[_-]*/giu, " ").replace(/[._-]+/gu, " ").replace(/\s+/gu, " ").trim() || "Proiect Trend";
}

function friendlyKind(kind) { return { "ac-summary-1-5": "Export de sinteză", "ac-detailed-0-2": "Export detaliat", "devplan-descriptors": "Declinații", unsupported: "Fișier nerecunoscut" }[kind] || "Fișier"; }
function sourceDates() { return [...new Set((payload?.schemas || []).flatMap((schema) => schema.methodology?.dates || []))].filter(Boolean); }
function periodLabel() { const dates = sourceDates(); if (!dates.length) return "Nu apare în export"; if (dates.length === 1) return dates[0]; return `${dates[0]} – ${dates.at(-1)}`; }
function addFieldListeners() { $$(`[data-derived-field]`).forEach((input) => input.addEventListener("change", recompute)); }
function renderFileCards() {
  const schemas = payload?.schemas || [];
  const matched = new Map(schemas.map((schema) => [schema.kind, schema]));
  [["summary", "ac-summary-1-5", "Export recunoscut"], ["detailed", "ac-detailed-0-2", "Export recunoscut"], ["descriptors", "devplan-descriptors", "Declinații recunoscute"]].forEach(([cardKind, schemaKind, readyLabel]) => {
    const card = $(`[data-file-kind="${cardKind}"]`); const status = $(`[data-file-status="${cardKind}"]`); const schema = matched.get(schemaKind);
    card?.classList.toggle("is-found", Boolean(schema));
    if (status) status.textContent = schema ? `${readyLabel} · ${schema.sourceName}` : cardKind === "descriptors" ? "Poți continua fără el" : "Așteaptă fișierul";
  });
  const note = $("#upload-note"); if (note) note.textContent = requiredFilesPresent() ? "Exporturile obligatorii au fost recunoscute. Confirmă faptele propuse și mergi la verificare." : "Adaugă exportul de sinteză și exportul detaliat pentru a continua.";
}

function renderFound() {
  if (!payload || !files.length) { $("#found-panel")?.setAttribute("hidden", ""); return; }
  $("#found-panel")?.removeAttribute("hidden");
  const stats = [
    [payload.participantCounts?.included ?? 0, "participanți incluși"],
    [payload.competencies?.length ?? 0, "competențe"],
    [payload.behaviorAggregates?.length ?? 0, "comportamente"],
    [payload.regionReadiness?.available ?? 0, "regiuni"],
    [payload.codeReadiness?.available ?? 0, "grupuri"],
    [periodLabel(), "perioada evaluării"]
  ];
  const statsNode = $("#findings-stats");
  if (statsNode) statsNode.innerHTML = stats.map(([value, label]) => `<div class="stat"><strong>${esc(value)}</strong><span>${esc(label)}</span></div>`).join("");
  const factsNode = $("#source-facts");
  if (factsNode) factsNode.innerHTML = (payload.schemas || []).map((schema) => `<div class="source-fact"><strong>${esc(friendlyKind(schema.kind))}</strong><span>${esc(schema.sourceName)} · ${Math.max(0, schema.rows.length - schema.headerRow)} rânduri</span></div>`).join("");
  const values = metadata();
  const fields = [
    ["project-name", "Proiect", "Numele propus pentru copertă.", values.projectName || state.projectName, "Apare în: copertă și numele fișierelor.", "din numele fișierului", "text"],
    ["client-name", "Client", "Poți confirma numele folosit pe copertă.", values.clientName, "Apare în: copertă și titlul raportului.", "de confirmat de consultant", "text"],
    ["report-date", "Data raportului", "Data propusă pentru livrare.", values.reportDate || today(), "Apare în: copertă și chitanță.", "propunere", "text"],
    ["program", "Program", "Denumirea programului din livrare.", values.program || "Centru de Dezvoltare", "Apare în: copertă și titlul raportului.", "standard Trend", "text"],
    ["report-context", "Context copertă", "O formulare scurtă, dacă este necesară.", values.context, "Apare în: copertă.", "opțional", "text"]
  ];
  const derived = $("#derived-fields");
  if (derived) derived.innerHTML = fields.map(([id, label, help, value, where, source, type]) => `<div class="derived-field"><label for="${id}">${esc(label)}</label><p class="field-help">${esc(help)} Exemplu: ${esc(id === "report-date" ? today() : id === "program" ? "Centru de Dezvoltare" : id === "client-name" ? "PPC" : "Proiect Nord")}.</p><input id="${id}" data-derived-field="true" type="${type}" value="${esc(value)}" placeholder="${id === "client-name" ? "Exemplu: PPC" : "Completează dacă este necesar"}"><span class="source-chip">${esc(source)}</span><span class="field-where">${esc(where)}</span></div>`).join("");
  addFieldListeners();
}

const issueTitles = {
  "summary-score": "Scor invalid (1–5)", "detailed-score": "Scor de comportament neacceptat", "summary-headers": "Exportul de sinteză nu are coloanele necesare", "detailed-headers": "Exportul detaliat nu are coloanele necesare", "summary-required": "Lipsește exportul de sinteză", "detailed-required": "Lipsește exportul detaliat", "unsupported": "Fișier nerecunoscut", "project-required": "Numele proiectului lipsește", "benchmark-invalid": "Intervalul benchmark nu este valid", "duplicate-identity": "Persoană și cod de evaluare duplicate", "detailed-duplicate-identity": "Rând duplicat în exportul detaliat", "detailed-only-identity": "Persoană prezentă numai în exportul detaliat", "summary-source-ambiguity": "Sunt două exporturi de sinteză", "detailed-source-ambiguity": "Sunt două exporturi detaliate", "required-value": "Lipsește o valoare obligatorie"
};
const issueExplanations = {
  "summary-score": "În exportul de sinteză există scoruri în afara intervalului acceptat. Rândurile afectate nu pot intra în calcule.", "detailed-score": "În exportul detaliat există observații care nu pot fi interpretate ca scoruri valide. Rândurile afectate nu intră în analiza comportamentelor.", "summary-headers": "Fișierul nu seamănă cu un export de sinteză recunoscut.", "detailed-headers": "Fișierul nu seamănă cu un export detaliat recunoscut.", "summary-required": "Nu am găsit exportul care alimentează mediile pe competențe.", "detailed-required": "Nu am găsit exportul care alimentează comportamentele și regiunile.", "unsupported": "Fișierul nu poate fi folosit pentru acest raport.", "project-required": "Raportul are nevoie de un nume de proiect pe copertă.", "benchmark-invalid": "Pragurile nu formează un interval coerent.", "required-value": "Un rând important nu are o valoare necesară.", "duplicate-identity": "Aceeași persoană și același cod apar de mai multe ori.", "detailed-duplicate-identity": "Aceeași persoană și același cod apar de mai multe ori.", "detailed-only-identity": "O persoană din detaliat nu are pereche în sinteză.", "summary-source-ambiguity": "Nu alegem automat între exporturi care par echivalente.", "detailed-source-ambiguity": "Nu alegem automat între exporturi care par echivalente."
};
const issueFixes = {
  "summary-score": "Corectează scorurile în exportul de sinteză și înlocuiește fișierul.", "detailed-score": "Corectează observațiile în exportul detaliat și înlocuiește fișierul.", "summary-headers": "Deschide exportul din AC și alege fișierul de sinteză complet.", "detailed-headers": "Deschide exportul din AC și alege fișierul detaliat complet.", "summary-required": "Adaugă exportul de sinteză.", "detailed-required": "Adaugă exportul detaliat.", "unsupported": "Alege un export AC de sinteză sau detaliat.", "project-required": "Confirmă numele proiectului în secțiunea Încarcă.", "benchmark-invalid": "Alege un prag inferior mai mic decât pragul superior, între 1 și 5.", "required-value": "Completează valoarea în exportul original și înlocuiește fișierul.", "duplicate-identity": "Păstrează un singur rând pentru fiecare persoană și cod de evaluare.", "detailed-duplicate-identity": "Păstrează un singur rând pentru fiecare persoană și cod de evaluare.", "detailed-only-identity": "Regenerază exporturile pereche din același proiect.", "summary-source-ambiguity": "Păstrează un singur export de sinteză.", "detailed-source-ambiguity": "Păstrează un singur export detaliat."
};
const warningTitles = { "summary-unrated": "Scoruri pe competență lipsă", "region-missing": "Participanți fără regiune", "code-missing": "Participanți fără grup", "detailed-blank": "Observații lipsă", "detailed-all-blank": "Comportament fără observații", "descriptor-incomplete": "Rând incomplet în declinații", "descriptor-unmatched": "Declinație fără potrivire", "descriptor-ambiguous": "Declinație cu mai multe potriviri", "detailed-missing-identity": "Persoană fără rând în detaliat" };
const warningExplanations = { "summary-unrated": "Valorile lipsă nu intră în calcule și nu sunt tratate ca zero.", "region-missing": "Participantul rămâne în proiect, dar nu apare în compararea pe regiuni.", "code-missing": "Participantul rămâne în vederea întregului proiect și nu poate fi alocat unui grup.", "detailed-blank": "Observațiile lipsă rămân lipsă; restul observațiilor intră în calcul.", "detailed-all-blank": "Comportamentul este exclus din clasament până când există observații.", "descriptor-incomplete": "Rândul nu poate aduce o formulare în raport.", "descriptor-unmatched": "Formularea nu se potrivește cu un comportament importat.", "descriptor-ambiguous": "Păstrăm textul importat ca să nu alegem greșit.", "detailed-missing-identity": "Persoana nu contribuie la comportamentele din raport." };
const warningFixes = { "summary-unrated": "Confirmă că lipsa este intenționată; nu completa automat.", "region-missing": "Completează regiunea în export dacă ai nevoie de comparația pe regiuni.", "code-missing": "Completează grupul în export dacă persoana trebuie să aparțină unui grup.", "detailed-blank": "Revizuiește exportul dacă lipsurile nu sunt intenționate.", "detailed-all-blank": "Completează observațiile pentru comportament în exportul detaliat.", "descriptor-incomplete": "Completează rândul sau descarcă un nou șablon.", "descriptor-unmatched": "Păstrează numai formulările pentru comportamentele importate.", "descriptor-ambiguous": "Păstrează o singură potrivire clară.", "detailed-missing-identity": "Regenerază exportul detaliat din același proiect." };

function groupedIssues(items, severity) {
  const groups = new Map();
  for (const item of items) { const key = item.code || "other"; if (!groups.has(key)) groups.set(key, { key, severity, items: [] }); groups.get(key).items.push(item); }
  return [...groups.values()];
}
function sourceRowCount(sourceName) { const schema = payload?.schemas?.find((item) => item.sourceName === sourceName); return schema ? Math.max(1, schema.rows.length - schema.headerRow) : 0; }
function likelyCause(group) {
  const bySource = new Map();
  group.items.forEach((item) => { const source = item.sourceName || "fișierul importat"; if (!bySource.has(source)) bySource.set(source, new Set()); if (item.rowNumber) bySource.get(source).add(item.rowNumber); });
  for (const [source, rows] of bySource) if (rows.size / sourceRowCount(source) >= .8) return `Probabil fișierul este greșit sau este alt tip de export: ${source}. Înlocuiește-l cu exportul potrivit.`;
  return "";
}
function affectedRows(group) {
  const rows = [];
  for (const item of group.items) {
    if (Array.isArray(item.pairs) && item.pairs.length) item.pairs.forEach((pair, index) => rows.push({ source: item.sourceName || "", row: item.rowNumber || "", label: pair || item.message, key: `${item.id}-${index}` }));
    else if (Array.isArray(item.identities) && item.identities.length) item.identities.forEach((identity, index) => rows.push({ source: item.sourceName || "", row: item.rowNumber || "", label: identity, key: `${item.id}-${index}` }));
    else rows.push({ source: item.sourceName || "", row: item.rowNumber || "", label: item.message, field: item.field || "", key: warningKey(item) });
  }
  return rows;
}
function rowMatches(row, query) { return !query || [row.source, row.row, row.label, row.field].join(" ").toLocaleLowerCase("ro").includes(query.toLocaleLowerCase("ro")); }
function renderIssueGroup(group) {
  const title = group.severity === "blocker" ? (issueTitles[group.key] || "Blocaj de verificare") : (warningTitles[group.key] || "Avertisment");
  const explanation = group.severity === "blocker" ? (issueExplanations[group.key] || group.items[0]?.message || "Verificarea nu a putut fi încheiată.") : (warningExplanations[group.key] || group.items[0]?.message || "Apare o diferență care cere confirmare.");
  const fix = group.severity === "blocker" ? (issueFixes[group.key] || "Corectează exportul și încarcă-l din nou.") : (warningFixes[group.key] || "Verifică exportul și confirmă dacă situația este intenționată.");
  const rows = affectedRows(group); const cause = likelyCause(group); const id = `issue-${group.severity}-${group.key}`;
  const article = document.createElement("article"); article.className = `issue-group ${group.severity}`; article.dataset.issueGroup = id;
  article.innerHTML = `<div class="issue-group-head"><div class="issue-group-title"><small>${group.severity === "blocker" ? "BLOCAJ" : "AVERTISMENT"}</small><h3>${esc(title)}</h3><span class="issue-count">${rows.length}</span></div><div class="issue-group-actions">${group.severity === "warning" ? `<button type="button" class="button button-secondary button-small" data-confirm-group="${esc(group.key)}">Confirmă acest grup</button>` : ""}</div></div><p class="issue-group-summary">${esc(explanation)}</p><p class="issue-fix">Cum repari: ${esc(fix)}</p>${cause ? `<p class="issue-fix">${esc(cause)}</p>` : ""}<div class="issue-tools"><label>Caută în rândurile afectate<input type="search" data-issue-search="${esc(id)}" placeholder="nume, fișier sau rând" autocomplete="off"></label><span class="source-chip">${esc(new Set(rows.map((row) => row.source).filter(Boolean)).size ? [...new Set(rows.map((row) => row.source).filter(Boolean))].join(" · ") : "exportul importat")}</span></div><div class="issue-items" data-issue-items="${esc(id)}"></div>`;
  const list = article.querySelector(`[data-issue-items="${CSS.escape(id)}"]`); const drawRows = (query = "") => { const filtered = rows.filter((row) => rowMatches(row, query)); list.innerHTML = filtered.length ? filtered.map((row) => `<div class="issue-item"><div class="issue-item-copy"><strong>${esc(row.label)}</strong><small>${esc(row.source || "Fișier")}${row.row ? ` · rândul ${esc(row.row)}` : ""}${row.field ? ` · ${esc(row.field)}` : ""}</small></div></div>`).join("") : `<p class="empty-value">Nu există rânduri pentru această căutare.</p>`; const more = filtered.length < rows.length ? `<p class="issue-more">Se afișează ${filtered.length} din ${rows.length} rânduri afectate.</p>` : ""; list.insertAdjacentHTML("afterend", more); };
  drawRows(); article.querySelector("[data-issue-search]")?.addEventListener("input", (event) => { article.querySelector(".issue-more")?.remove(); drawRows(event.target.value); });
  article.querySelector("[data-confirm-group]")?.addEventListener("click", () => { group.items.forEach((item) => state.acknowledged.add(warningKey(item))); announce(`Avertismentul ${title} a fost confirmat.`); renderReview(); sync(); $("#review-title")?.focus(); });
  return article;
}

function renderReview() {
  const root = $("#issue-list"); if (!root || !payload) return;
  const blockers = payload.blockers || []; const pending = payload.warnings.filter((item) => !item.reviewed); const totalRows = (payload.schemas || []).reduce((sum, schema) => sum + Math.max(0, schema.rows.length - schema.headerRow), 0);
  const overview = $("#review-overview");
  if (overview) overview.innerHTML = `<div class="review-state ${blockers.length || pending.length ? "needs-attention" : "all-clear"}"><strong>${blockers.length || pending.length ? "Mai este ceva de judecat" : "Totul este pregătit"}</strong><span>${blockers.length} blocaje · ${pending.length} avertismente · ${totalRows} rânduri citite</span></div><div class="review-stats"><div class="review-stat"><strong>${blockers.length}</strong><span>blocaje</span></div><div class="review-stat"><strong>${pending.length}</strong><span>avertismente de confirmat</span></div><div class="review-stat"><strong>${payload.participantCounts?.included ?? 0}</strong><span>participanți incluși</span></div><div class="review-stat"><strong>${payload.participantCounts?.excludedUnrated ?? 0}</strong><span>fără scor, excluși</span></div></div>`;
  root.replaceChildren();
  groupedIssues(blockers, "blocker").forEach((group) => root.append(renderIssueGroup(group)));
  groupedIssues(pending, "warning").forEach((group) => root.append(renderIssueGroup(group)));
  if (!root.children.length) root.innerHTML = `<div class="review-state all-clear"><strong>Nu mai există blocaje sau avertismente de confirmat.</strong><span>Poți alege structura raportului.</span></div>`;
  const confirmAll = $("#confirm-all"); if (confirmAll) { confirmAll.hidden = blockers.length > 0 || pending.length === 0; confirmAll.disabled = blockers.length > 0 || pending.length === 0; }
  const note = $("#review-action-note"); if (note) note.textContent = blockers.length ? "Rezolvă blocajele din export și încarcă fișierele corectate." : pending.length ? "Confirmă fiecare grup sau folosește Confirmă toate avertismentele." : "Totul este confirmat; poți continua.";
}

function outlineLabel(slide) { const family = { cover: "Copertă", "how-to-read": "Cum se citește", methodology: "Metodologie", "executive-summary": "Rezumat executiv", population: "Distribuția rezultatelor", ranking: "Medii pe competențe", zone: "Regiuni", "appendix-divider": "Anexă · rezultate individuale", "participant-mean": "Rezultate individuale", "participant-comparison": "Comparație pe participanți", "divider-observations": "Analiza observațiilor", "key-findings": "Constatări cheie", "competency-participants": "Grafice pe competențe", "divider-behaviors": "Comportamente cheie", behavior: "Comportamente cheie", conclusions: "Concluzii și recomandări", close: "Mulțumim" }; return family[slide.family] || slide.title || "Secțiune"; }
function groupOutlineSlides(plan) { const groups = []; const keys = [...new Set(plan.map((slide) => slide.groupKey || ""))]; for (const key of keys) { const slides = plan.filter((slide) => (slide.groupKey || "") === key); if (!slides.length) continue; const main = slides.filter((slide) => slide.deliverable !== "appendix" && !["conclusions", "close"].includes(slide.family)); const annex = slides.filter((slide) => slide.deliverable === "appendix"); const ending = slides.filter((slide) => ["conclusions", "close"].includes(slide.family)); if (main.length) groups.push({ title: key ? `Raport principal · grup ${key}` : "Raport principal · întregul proiect", slides: main }); if (annex.length) groups.push({ title: key ? `Anexă · grup ${key}` : "Anexă", slides: annex }); if (ending.length) groups.push({ title: key ? `Concluzii · grup ${key}` : "Concluzii și recomandări", slides: ending }); } return groups; }
function renderStructure() {
  const root = $("#structure-summary"); if (!root || !payload?.readiness) return;
  const plan = reportPlan(payload, { scope: "whole" }); const groups = groupOutlineSlides(plan); const annex = payload.metadata.annex; const mainCount = plan.filter((slide) => slide.deliverable !== "appendix" && !["conclusions", "close"].includes(slide.family)).length; const appendixCount = plan.filter((slide) => slide.deliverable === "appendix").length; const endCount = plan.filter((slide) => ["conclusions", "close"].includes(slide.family)).length;
  $("#slide-total") && ($("#slide-total").textContent = `${plan.length} slide-uri`);
  root.parentElement.querySelectorAll(".structure-summary-note").forEach((node) => node.remove());
  root.innerHTML = groups.map((group) => `<section class="outline-group"><h4>${esc(group.title)}<small>${group.slides.length} slide-uri · întregul proiect înaintea grupurilor</small></h4><span class="outline-count">${group.slides.length}</span><div class="outline-items">${group.slides.slice(0, 9).map((slide) => `<div class="outline-item"><span>${esc(outlineLabel(slide))}</span><span>${slide.number}</span></div>`).join("")}${group.slides.length > 9 ? `<div class="outline-item"><span>și alte secțiuni</span><span>${group.slides.length - 9}</span></div>` : ""}</div></section>`).join("");
  root.insertAdjacentHTML("beforebegin", `<p class="field-where structure-summary-note">${plan.length} slide-uri: ${mainCount} principal · ${appendixCount} anexă · ${endCount} concluzii și încheiere. ${annex === "separate" ? "Anexa se descarcă separat, cu ambele nume de fișier afișate." : annex === "none" ? "Fără anexă: rezultatele individuale, analiza observațiilor și comportamentele cheie nu apar în livrare." : "Anexa rămâne la finalul raportului principal."}</p>`);
}

function setReason(input, reason) { const label = input?.closest("label"); if (!label) return; let node = label.querySelector(".control-reason"); if (!reason) { node?.remove(); return; } if (!node) { node = document.createElement("small"); node.className = "control-reason"; label.append(node); } node.textContent = reason; input.setAttribute("aria-describedby", "split-help"); }
function guardControls() {
  if (!payload) return;
  const split = $("#split-groups"); if (split) { const unavailable = !payload.codeReadiness?.splitAvailable; split.disabled = unavailable; if (unavailable) split.checked = false; const reason = unavailable ? "Împărțirea este indisponibilă: sunt necesare cel puțin două grupuri CODE." : ""; $("#split-help") && ($("#split-help").textContent = reason || `Dezactivat implicit. ${payload.codeReadiness.blank ? `${payload.codeReadiness.blank} participanți fără grup rămân numai în vederea întregului proiect.` : "Participanții fără grup rămân numai în vederea întregului proiect."}`); setReason(split, reason); }
  const previewButton = $("#preview-trigger"); if (previewButton) previewButton.disabled = !payload.readiness;
  const separate = payload.metadata.annex === "separate"; $("#pptx-whole") && ($("#pptx-whole").hidden = separate); $("#pptx-main") && ($("#pptx-main").hidden = !separate); $("#pptx-appendix") && ($("#pptx-appendix").hidden = !separate);
}

function sync() {
  document.body.dataset.workflowStep = String(state.step);
  const allowed = files.length ? payload?.readiness ? 4 : 2 : 1;
  $$(`[data-section]`).forEach((section) => { const number = Number(section.dataset.section); const unlocked = number <= allowed; section.classList.toggle("is-locked", !unlocked); section.classList.toggle("is-unlocked", unlocked && number > 1); section.classList.toggle("is-current", number === state.step); section.setAttribute("aria-disabled", String(!unlocked)); const body = section.querySelector(`[data-body="${number}"]`); if (body) body.hidden = !unlocked; });
  $$(`#workflow-steps [data-step]`).forEach((link) => { const number = Number(link.dataset.step); const unlocked = number <= allowed; link.classList.toggle("active", number === state.step); link.classList.toggle("done", number < state.step); link.classList.toggle("locked", !unlocked); link.setAttribute("aria-disabled", String(!unlocked)); });
  $("#to-step-2") && ($("#to-step-2").disabled = !files.length);
  $("#to-step-3") && ($("#to-step-3").disabled = !ready(2));
  $("#to-step-4") && ($("#to-step-4").disabled = !ready(3));
  const enabled = ready(3) && !state.busy; ["#xlsx", "#csv-template", "#bundle", "#pptx-whole", "#pptx-main", "#pptx-appendix"].forEach((selector) => { const button = $(selector); if (button) button.disabled = !enabled; });
  guardControls();
  const chooseNote = $("#choose-note"); if (chooseNote) chooseNote.textContent = payload?.metadata?.annex === "none" ? "Fără anexă a fost aleasă; rezultatele individuale și analiza observațiilor dispar din livrare." : `${reportPlan && payload?.readiness ? reportPlan(payload, { scope: "whole" }).length : "—"} slide-uri după alegerile curente.`;
}

function renderDownload() { const time = $("#receipt-time"); if (time && state.receipts.length) time.textContent = state.receipts.at(-1).time; const receipt = $("#receipt"); if (receipt && state.receipts.length) receipt.innerHTML = state.receipts.map((item) => `<div class="receipt-entry"><strong>${esc(item.name)}</strong><span>${esc(item.kind)} · ${esc(item.detail)} · ${esc(item.time)}</span></div>`).join(""); }
function addReceipt(name, kind, detail) { const time = new Intl.DateTimeFormat("ro-RO", { dateStyle: "short", timeStyle: "short" }).format(new Date()); state.receipts.push({ name, kind, detail, time }); renderDownload(); announce(`${name} a fost pregătit pentru descărcare.`); }

function render() {
  const current = metadata(); if (files.length && !state.projectName) state.projectName = current.projectName || deriveProjectName(); if (!state.reportDate) state.reportDate = current.reportDate || today();
  payload = buildPayload(window.XLSX, files, metadata(), state.corrections, { acknowledgedWarningIds: [...state.acknowledged] });
  Object.assign(payload.metadata, metadata());
  const warningIds = new Set(payload.warnings.map(warningKey)); state.acknowledged = new Set([...state.acknowledged].filter((id) => warningIds.has(id))); payload.warnings = payload.warnings.map((item) => ({ ...item, reviewed: state.acknowledged.has(warningKey(item)) })); payload.warningReviews = payload.warnings; payload.readiness = payload.blockers.length === 0 && payload.warnings.every((item) => item.reviewed);
  renderFileCards(); renderFound(); renderReview(); renderStructure(); renderDownload(); sync();
}
function recompute() { render(); }

async function readSources(event) { const incoming = await Promise.all([...event.target.files].map(async (file) => ({ name: file.name, bytes: await file.arrayBuffer() }))); files = mergeSelectedFiles(files, incoming); state.acknowledged.clear(); state.corrections = { values: {} }; state.projectName = state.projectName || deriveProjectName(); state.reportDate = state.reportDate || today(); event.target.value = ""; render(); announce("Fișierele au fost citite. Verifică datele găsite și mergi la verificare."); $("#review-title")?.focus(); }
function activateStep(number) { const allowed = files.length ? payload?.readiness ? 4 : 2 : 1; if (number > allowed) return; state.step = number; sync(); document.querySelector(`#step-${["upload", "review", "choose", "download"][number - 1]} h2`)?.focus(); if (number > 1) document.querySelector(`#step-${["upload", "review", "choose", "download"][number - 1]}`)?.scrollIntoView({ block: "start" }); }

function expectedNames(kind) { const name = slug(payload?.metadata?.projectName); return { whole: `raport-trend-${name}.pptx`, main: `raport-trend-principal-${name}.pptx`, appendix: `raport-trend-anexa-${name}.pptx`, bundle: `pachet-bhb-${name}.zip`, xlsx: `audit-raport-grup-${name}.xlsx`, "csv-template": "evaluation-sheet-template.csv" }; }
async function createDownload(kind) {
  if (state.busy || !ready(3)) return; state.busy = true; sync(); const names = expectedNames(kind); const status = $("#download-status"); try {
    if (kind === "xlsx") { const bytes = XLSX.write(createAuditWorkbook(XLSX, payload), { type: "array", compression: true, bookType: "xlsx" }); download(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), names.xlsx); addReceipt(names.xlsx, "fișier de lucru", "audit calculat"); }
    else if (kind === "csv-template") { const bytes = XLSX.write(createEvaluationSheetTemplate(XLSX, payload), { type: "string", bookType: "csv" }); download(new Blob([bytes], { type: "text/csv;charset=utf-8" }), names["csv-template"]); addReceipt(names["csv-template"], "fișier de lucru", "șablon declinații"); }
    else if (kind === "bundle") { await downloadBundle(payload, names.bundle); addReceipt(names.bundle, "livrabil client", "Pachet BHB"); }
    else { const scope = kind === "main" ? "main" : kind === "appendix" ? "appendix" : "whole"; await downloadPptx(payload, names[kind] || names.whole, { scope }); const count = reportPlan(payload, { scope }).length; addReceipt(names[kind] || names.whole, "livrabil client", `${count} slide-uri`); }
    if (status) status.textContent = "Livrabilul a fost pregătit. Chitanța de mai jos a fost actualizată.";
  } catch (error) { if (status) status.textContent = `Descărcarea nu a pornit: ${error.message || error}`; } finally { state.busy = false; sync(); }
}

function resetSession() { files = []; payload = null; state.step = 1; state.acknowledged.clear(); state.corrections = { values: {} }; state.projectName = ""; state.reportDate = ""; state.receipts = []; state.previewOpen = false; $("#sources").value = ""; $("#reset-confirm").hidden = true; render(); announce("Sesiunea a fost curățată. Fișierele sursă nu au fost schimbate."); $("#upload-title")?.focus(); }

$("#sources")?.addEventListener("change", readSources);
$("#choose-files")?.addEventListener("click", () => $("#sources")?.click());
$("#drop-zone")?.addEventListener("click", (event) => { if (event.target.closest("button")) return; $("#sources")?.click(); });
$("#drop-zone")?.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); $("#sources")?.click(); } });
$("#drop-zone")?.addEventListener("dragover", (event) => { event.preventDefault(); $("#drop-zone").classList.add("is-dragging"); });
$("#drop-zone")?.addEventListener("dragleave", () => $("#drop-zone").classList.remove("is-dragging"));
$("#drop-zone")?.addEventListener("drop", (event) => { event.preventDefault(); $("#drop-zone").classList.remove("is-dragging"); const input = $("#sources"); const transfer = event.dataTransfer; if (input && transfer?.files?.length) { const dt = new DataTransfer(); [...transfer.files].forEach((file) => dt.items.add(file)); input.files = dt.files; input.dispatchEvent(new Event("change", { bubbles: true })); } });
$("#to-step-2")?.addEventListener("click", () => activateStep(2)); $("#to-step-3")?.addEventListener("click", () => activateStep(3)); $("#to-step-4")?.addEventListener("click", () => activateStep(4));
$("#confirm-all")?.addEventListener("click", () => { payload.warnings.forEach((item) => state.acknowledged.add(warningKey(item))); announce("Toate avertismentele au fost confirmate."); render(); $("#review-title")?.focus(); });
$$(`#workflow-steps [data-step]`).forEach((link) => link.addEventListener("click", (event) => { event.preventDefault(); activateStep(Number(link.dataset.step)); }));
$$(`[data-section]`).forEach((section) => section.addEventListener("click", (event) => { const anchor = event.target.closest("a[data-step]"); if (anchor) { event.preventDefault(); activateStep(Number(anchor.dataset.step)); } }));
$("#split-groups")?.addEventListener("change", recompute); $("#benchmark-low")?.addEventListener("change", recompute); $("#benchmark-high")?.addEventListener("change", recompute); $$(`input[name="annex"]`).forEach((input) => input.addEventListener("change", () => { $$(`.option`).forEach((option) => option.classList.toggle("selected", option.querySelector("input")?.checked)); recompute(); }));
$("#methodology-text")?.addEventListener("change", recompute); $("#evaluators")?.addEventListener("change", recompute); $("#days")?.addEventListener("change", recompute); $("#conclusions")?.addEventListener("change", recompute); $("#conclusions-destination")?.addEventListener("change", recompute);
$("#preview-trigger")?.addEventListener("click", () => { state.previewOpen = !state.previewOpen; const root = $("#preview"); if (!root) return; root.hidden = !state.previewOpen; $("#preview-trigger").textContent = state.previewOpen ? "Ascunde structura slide-urilor" : "Vezi structura slide-urilor"; if (state.previewOpen && payload?.readiness) { mountPreview(root, payload, { scope: "whole" }); root.querySelector(".preview-stage")?.focus(); } });
[["#xlsx", "xlsx"], ["#csv-template", "csv-template"], ["#csv-template-import", "csv-template"], ["#bundle", "bundle"], ["#pptx-whole", "whole"], ["#pptx-main", "main"], ["#pptx-appendix", "appendix"]].forEach(([selector, kind]) => $(selector)?.addEventListener("click", () => createDownload(kind)));
$("#reset")?.addEventListener("click", () => { $("#reset-confirm").hidden = false; $("#reset-confirm-yes")?.focus(); }); $("#reset-confirm-yes")?.addEventListener("click", resetSession); $("#reset-cancel")?.addEventListener("click", () => { $("#reset-confirm").hidden = true; $("#reset")?.focus(); });

render(); window.__grfBooted?.();
