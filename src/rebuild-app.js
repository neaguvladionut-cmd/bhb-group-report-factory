import { renderIssueGroup } from "./rebuild-issues.js";

function groupedIssues(items, severity) {
  const groups = new Map();
  for (const item of items) { const key = item.code || "other"; if (!groups.has(key)) groups.set(key, { key, severity, items: [] }); groups.get(key).items.push(item); }
  return [...groups.values()];
}

if (!window.XLSX || !window.JSZip) throw new Error("Lipsesc bibliotecile locale necesare");

let files = [];
let payload = null;
let activeDownloadUrl = "";
const state = { step: 1, acknowledged: new Set(), corrections: { values: {} }, methodologyProposals: {}, completedSteps: new Set(), busy: false, receipts: [], projectName: "", reportDate: "", previewOpen: false };
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
  const value = (proposalKey, ...ids) => Object.hasOwn(state.methodologyProposals, proposalKey) ? state.methodologyProposals[proposalKey] : ids.map((id) => $(`#${id}`)?.value).find((item) => text(item)) || "";
  return {
    projectName: $("#project-name")?.value || state.projectName,
    clientName: $("#client-name")?.value || "",
    reportDate: $("#report-date")?.value || state.reportDate || today(),
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
    evaluators: value("evaluators", "proposal-evaluators", "evaluators"),
    days: value("days", "proposal-days", "days"),
    teamSize: value("teamSize", "proposal-team-size", "team-size"),
    evaluationPeriod: value("evaluationPeriod", "proposal-period", "evaluation-period"),
    populationByRole: value("populationByRole", "proposal-population", "population-role"),
    location: value("location", "proposal-location", "location"),
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
function ready(step) { if (step === 1) return files.length > 0 && requiredFilesPresent() && Boolean(text(metadata().projectName)); if (step === 2) return Boolean(payload?.readiness && warningComplete()); if (step === 3) return Boolean(payload?.readiness && warningComplete()); return false; }
function invalidate() { recompute(); }

function deriveProjectName() {
  const first = files[0]?.name || "proiect-nou";
  return first.replace(/\.[^.]+$/u, "").replace(/(?:sinteza|summary|detaliat|detail|export|raport|ac)[_-]*/giu, " ").replace(/[._-]+/gu, " ").replace(/\s+/gu, " ").trim() || "Proiect Trend";
}

function friendlyKind(kind) { return { "ac-summary-1-5": "Export de sinteză", "ac-detailed-0-2": "Export detaliat", "devplan-descriptors": "Declinații", unsupported: "Fișier nerecunoscut" }[kind] || "Fișier"; }
function sourceDates() { return payload?.methodology?.dates || []; }
function periodLabel() { const dates = sourceDates(); if (!dates.length) return "Nu apare în export"; if (dates.length === 1) return dates[0]; return `${dates[0]} – ${dates.at(-1)}`; }
function setProposal(input) { const key = input?.dataset.proposalKey; if (key) state.methodologyProposals[key] = input.value; }
function addFieldListeners() { $$(`[data-derived-field]`).forEach((input) => { if (input.dataset.listenerAttached) return; input.dataset.listenerAttached = "true"; input.addEventListener("input", () => { setProposal(input); recompute(); }); input.addEventListener("change", () => { setProposal(input); recompute(); }); }); }
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
  if (!payload || !files.length) { $("#found-panel")?.setAttribute("hidden", ""); $("#derived-fields")?.replaceChildren(); return; }
  $("#found-panel")?.removeAttribute("hidden");
  const method = payload.methodology || {};
  const roles = (method.populationByRole || []).slice().sort((a, b) => b.count - a.count || a.role.localeCompare(b.role, "ro"));
  const population = roles.length ? `${payload.participantCounts?.included ?? 0} participanți (${roles.map((item) => `${item.count} ${item.role}`).join(", ")})` : `${payload.participantCounts?.included ?? 0} participanți`;
  const proposalChip = (value) => value ? "din export · de verificat" : "";
  const stats = [
    [payload.participantCounts?.included ?? 0, "participanți incluși"],
    [payload.competencies?.length ?? 0, "competențe"],
    [payload.behaviorAggregates?.length ?? 0, "comportamente"],
    [method.evaluatorNames?.length || "", "evaluatori distincți"],
    [method.commonTeamSize || "", "consultanți / participant"],
    [method.dates?.length || "", "zile de evaluare"]
  ];
  const statsNode = $("#findings-stats");
  if (statsNode) statsNode.innerHTML = stats.map(([value, label]) => `<div class="stat"><strong>${esc(value)}</strong><span>${esc(label)}</span></div>`).join("");
  const factsNode = $("#source-facts");
  if (statsNode) statsNode.innerHTML = stats.map(([value, label]) => { const missing = value === "" || value === null || value === undefined; return "<div class=\"stat\"><strong>" + (missing ? "—" : esc(value)) + "</strong><span>" + esc(label) + "</span><small class=\"stat-caption\">" + (missing ? "Nu apare în export." : "Propunere din export.") + "</small></div>"; }).join("");
  if (factsNode) factsNode.innerHTML = (payload.schemas || []).map((schema) => `<div class="source-fact"><strong>${esc(friendlyKind(schema.kind))}</strong><span>${esc(schema.sourceName)} · ${Math.max(0, schema.rows.length - schema.headerRow)} rânduri</span></div>`).join("");
  const defaults = {
    evaluators: (method.evaluatorNames || []).join(", "),
    teamSize: method.commonTeamSize || "",
    days: method.dates?.length || "",
    evaluationPeriod: periodLabel() === "Nu apare în export" ? "" : periodLabel(),
    populationByRole: population,
    location: (method.locations || []).join(", ")
  };
  Object.entries(defaults).forEach(([key, value]) => { if (!Object.hasOwn(state.methodologyProposals, key)) state.methodologyProposals[key] = String(value); });
  const values = metadata();
  const examples = {
    "project-name": "Proiect Delta",
    "client-name": "Compania Delta",
    "report-date": today(),
    program: "Centru de Dezvoltare",
    "proposal-evaluators": "Ana Popescu, Mihai Ionescu",
    "proposal-team-size": "3 consultanți",
    "proposal-days": "2",
    "proposal-period": "01.10.2026 – 02.10.2026",
    "proposal-population": "20 participanți (12 manageri, 8 specialiști)",
    "proposal-location": "București"
  };
  const fields = [
    ["project-name", "Nume proiect", "Numele proiectului apare pe copertă.", values.projectName || state.projectName, "Apare în: copertă și numele fișierelor.", "din numele fișierului", "text"],
    ["client-name", "Client", "Numele clientului rămâne în metodologia raportului și în numele fișierelor.", values.clientName, "Apare în: metodologia raportului și numele fișierelor.", "de confirmat de consultant", "text"],
    ["report-date", "Data raportului", "Data propusă pentru livrare.", values.reportDate || today(), "Apare în: copertă și chitanță.", "propunere", "text"],
    ["program", "Program", "Denumirea programului din livrare.", values.program || "Centru de Dezvoltare", "Apare în: copertă și metodologia raportului.", "standard Trend", "text"],
    ["proposal-evaluators", "Evaluatori", "Numele distincte găsite în exporturile proiectului.", values.evaluators, "Apare în: metodologia raportului.", proposalChip(method.evaluatorNames?.length), "text", "evaluators"],
    ["proposal-team-size", "Echipa unui participant", "Cea mai frecventă echipă de evaluatori pentru un participant.", values.teamSize, "Apare în: metodologia raportului.", proposalChip(method.commonTeamSize), "text", "teamSize"],
    ["proposal-days", "Zile de evaluare", "Numărul de date distincte găsite în export.", values.days, "Apare în: metodologia raportului.", proposalChip(method.dates?.length), "text", "days"],
    ["proposal-period", "Perioada evaluării", "Prima și ultima dată găsite în export.", values.evaluationPeriod, "Apare în: metodologia raportului.", proposalChip(method.dates?.length), "text", "evaluationPeriod"],
    ["proposal-population", "Participanți și roluri", "Rolurile sunt propuse din informațiile despre rol din export.", values.populationByRole, "Apare în: metodologia raportului.", proposalChip(roles.length), "text", "populationByRole"],
    ["proposal-location", "Locația evaluării", "Locația este propusă din informațiile despre locația evaluării din export.", values.location, "Apare în: metodologia raportului.", proposalChip(method.locations?.length), "text", "location"]
  ];
  const derived = $("#derived-fields");
  if (derived) derived.innerHTML = fields.map(([id, label, help, value, where, source, type, proposalKey]) => { const helpId = `${id}-help`; const exampleId = `${id}-example`; const whereId = `${id}-where`; return `<div class="derived-field"><label for="${id}">${esc(label)}</label><p id="${helpId}" class="field-help">${esc(help)}</p><p id="${exampleId}" class="field-help">Exemplu: ${esc(examples[id])}.</p><input id="${id}" data-derived-field="true" data-proposal-key="${esc(proposalKey || "")}" type="${type}" value="${esc(value)}" placeholder="${esc(examples[id])}" aria-describedby="${helpId} ${exampleId} ${whereId}">${source ? `<span class="source-chip">${esc(source)}</span>` : ""}<span id="${whereId}" class="field-where">${esc(where)}</span></div>`; }).join("");
  addFieldListeners();
  [["evaluators", "evaluators"], ["days", "days"], ["team-size", "teamSize"], ["evaluation-period", "evaluationPeriod"], ["population-role", "populationByRole"], ["location", "location"]].forEach(([id, key]) => { const input = $(`#${id}`); if (input) input.value = state.methodologyProposals[key] ?? ""; });
}

function renderReview() {
  const root = $("#issue-list"); if (!root || !payload) return;
  const blockers = payload.blockers || []; const pending = payload.warnings.filter((item) => !item.reviewed); const totalRows = (payload.schemas || []).reduce((sum, schema) => sum + Math.max(0, schema.rows.length - schema.headerRow), 0);
  const pendingGroups = groupedIssues(pending, "warning").map((group) => group.key);
  const overview = $("#review-overview");
  if (overview) overview.innerHTML = `<div class="review-state ${blockers.length || pending.length ? "needs-attention" : "all-clear"}"><strong>${blockers.length || pending.length ? "Mai este ceva de judecat" : "Totul este pregătit"}</strong><span>${blockers.length} blocaje · ${pending.length} avertismente · ${totalRows} rânduri citite</span></div><div class="review-stats"><div class="review-stat"><strong>${blockers.length}</strong><span>blocaje</span></div><div class="review-stat"><strong>${pending.length}</strong><span>avertismente de confirmat</span></div><div class="review-stat"><strong>${payload.participantCounts?.included ?? 0}</strong><span>participanți citiți</span></div><div class="review-stat"><strong>${payload.participantCounts?.excludedUnrated ?? 0}</strong><span>fără scor, excluși</span></div></div>`;
  root.replaceChildren();
  const onConfirm = ({ group, title }) => { group.items.forEach((item) => state.acknowledged.add(warningKey(item))); announce(`Avertismentul ${title} a fost confirmat.`); render(); const nextKey = pendingGroups[pendingGroups.indexOf(group.key) + 1]; const next = nextKey ? [...document.querySelectorAll("[data-confirm-group]")].find((button) => button.dataset.confirmGroup === nextKey) : null; (next || $("#to-step-3"))?.focus(); };
  groupedIssues(blockers, "blocker").forEach((group) => root.append(renderIssueGroup(group, { payload })));
  groupedIssues(pending, "warning").forEach((group) => root.append(renderIssueGroup(group, { payload, onConfirm })));
  if (!root.children.length) root.innerHTML = `<div class="review-state all-clear"><strong>Nu mai există blocaje sau avertismente de confirmat.</strong><span>Poți alege structura raportului.</span></div>`;
  const confirmAll = $("#confirm-all"); if (confirmAll) { confirmAll.hidden = blockers.length > 0 || pending.length === 0; confirmAll.disabled = blockers.length > 0 || pending.length === 0; }
  const note = $("#review-action-note"); if (note) note.textContent = blockers.length ? "Rezolvă blocajele din export și încarcă fișierele corectate." : pending.length ? "Confirmă fiecare grup sau folosește Confirmă toate avertismentele." : "Totul este confirmat; poți continua.";
}

function outlineLabel(slide) { const family = { cover: "Copertă", "how-to-read": "Cum se citește", methodology: "Metodologie", "executive-summary": "Rezumat executiv", population: "Distribuția rezultatelor", ranking: "Medii pe competențe", zone: "Regiuni", "appendix-divider": "Anexă · rezultate individuale", "participant-mean": "Rezultate individuale", "participant-comparison": "Comparație pe participanți", "divider-observations": "Analiza observațiilor", "key-findings": "Constatări cheie", "competency-participants": "Grafice pe competențe", "divider-behaviors": "Comportamente cheie", behavior: "Comportamente cheie", conclusions: "Concluzii și recomandări", close: "Mulțumim" }; return family[slide.family] || slide.title || "Secțiune"; }
function groupOutlineSlides(plan) { const groups = []; const keys = [...new Set(plan.map((slide) => slide.groupKey || ""))]; for (const key of keys) { const slides = plan.filter((slide) => (slide.groupKey || "") === key); if (!slides.length) continue; const main = slides.filter((slide) => slide.deliverable !== "appendix" && !["conclusions", "close"].includes(slide.family)); const annex = slides.filter((slide) => slide.deliverable === "appendix"); const ending = slides.filter((slide) => ["conclusions", "close"].includes(slide.family)); if (main.length) groups.push({ title: key ? `Raport principal · grup ${key}` : "Raport principal · întregul proiect", slides: main }); if (annex.length) groups.push({ title: key ? `Anexă · grup ${key}` : "Anexă", slides: annex }); if (ending.length) groups.push({ title: key ? `Concluzii · grup ${key}` : "Concluzii și recomandări", slides: ending }); } return groups; }
function renderStructure() {
  const root = $("#structure-summary"); if (!root || !payload?.readiness) return;
  const plan = reportPlan(payload, { scope: "whole" }); const groups = groupOutlineSlides(plan); const annex = payload.metadata.annex; const mainPlan = reportPlan(payload, { scope: "main" }); const appendixPlan = annex === "separate" ? reportPlan(payload, { scope: "appendix" }) : []; const deliveredTotal = annex === "separate" ? mainPlan.length + appendixPlan.length : plan.length; const mainCount = mainPlan.filter((slide) => slide.deliverable !== "appendix" && !["conclusions", "close"].includes(slide.family)).length; const appendixCount = annex === "separate" ? appendixPlan.length : plan.filter((slide) => slide.deliverable === "appendix").length; const endCount = mainPlan.filter((slide) => ["conclusions", "close"].includes(slide.family)).length;
  const slideLabel = (count) => `${count} ${count === 1 ? "slide" : "slide-uri"}`;
  $("#slide-total") && ($("#slide-total").textContent = slideLabel(deliveredTotal));
  root.parentElement.querySelectorAll(".structure-summary-note").forEach((node) => node.remove());
  root.innerHTML = groups.map((group) => `<section class="outline-group"><h4>${esc(group.title)}<small>${slideLabel(group.slides.length)}</small></h4><span class="outline-count">${group.slides.length}</span><div class="outline-items">${group.slides.slice(0, 9).map((slide) => `<div class="outline-item"><span>${esc(outlineLabel(slide))}</span><span>${slide.number}</span></div>`).join("")}${group.slides.length > 9 ? `<div class="outline-item"><span>și alte secțiuni</span><span>${group.slides.length - 9}</span></div>` : ""}</div></section>`).join("");
  root.insertAdjacentHTML("beforebegin", `<p class="field-where structure-summary-note">${slideLabel(deliveredTotal)} în livrare: ${mainCount} principal · ${appendixCount} anexă · ${endCount} concluzii și încheiere. ${annex === "separate" ? "Anexa se descarcă separat, cu ambele nume reale afișate." : annex === "none" ? "Fără anexă: rezultatele individuale, analiza observațiilor și comportamentele cheie nu apar în livrare." : "Anexa rămâne înaintea concluziilor în raportul principal."}</p>`);
}

function codeGroupSummary() {
  const groups = (payload?.groups || []).map((group) => `${group.name || group.code} · ${group.records?.length ?? 0}`);
  if (payload?.codeReadiness?.blank) groups.push(`fără grup · ${payload.codeReadiness.blank}`);
  return groups.length ? `Grupuri disponibile: ${groups.join(" · ")}.` : "Nu există grupuri în export.";
}
function setReason(input, reason) { const label = input?.closest("label"); if (!label) return; let node = label.querySelector(".control-reason"); if (!reason) { node?.remove(); return; } if (!node) { node = document.createElement("small"); node.className = "control-reason"; label.append(node); } node.textContent = reason; input.setAttribute("aria-describedby", "split-help"); }
function guardControls() {
  if (!payload) return;
  const split = $("#split-groups"); if (split) { const unavailable = !payload.codeReadiness?.splitAvailable; split.disabled = unavailable; if (unavailable) split.checked = false; const reason = unavailable ? "Împărțirea este indisponibilă: sunt necesare cel puțin două grupuri." : ""; $("#split-help") && ($("#split-help").textContent = `${codeGroupSummary()} ${reason || "Dezactivat implicit. Participanții fără grup rămân numai în vederea întregului proiect."}`); setReason(split, reason); }
  const previewButton = $("#preview-trigger"); if (previewButton) previewButton.disabled = !payload.readiness;
  const separate = payload.metadata.annex === "separate"; const names = expectedNames("whole"); $("#pptx-whole") && ($("#pptx-whole").hidden = separate, $("#pptx-whole").textContent = `Descarcă ${names.whole}`); $("#pptx-main") && ($("#pptx-main").hidden = !separate, $("#pptx-main").textContent = `Descarcă ${names.main}`); $("#pptx-appendix") && ($("#pptx-appendix").hidden = !separate, $("#pptx-appendix").textContent = `Descarcă ${names.appendix}`); const annexNames = $("#annex-file-names"); if (annexNames) annexNames.textContent = `Descarcă ${names.main} și ${names.appendix}.`;
}

function sync() {
  document.body.dataset.workflowStep = String(state.step);
  const allowed = ready(3) ? 4 : ready(2) ? 3 : ready(1) ? 2 : 1;
  if (state.step > allowed) state.step = allowed;
  $$(`[data-section]`).forEach((section) => { const number = Number(section.dataset.section); const unlocked = number <= allowed; section.classList.toggle("is-locked", !unlocked); section.classList.toggle("is-unlocked", unlocked && number > 1); section.classList.toggle("is-current", number === state.step); section.setAttribute("aria-disabled", String(!unlocked)); const body = section.querySelector(`[data-body="${number}"]`); if (body) body.hidden = !unlocked; });
  $$(`[data-section]`).forEach((section) => { section.hidden = Number(section.dataset.section) !== state.step; });
  $$(`#workflow-steps [data-step]`).forEach((link) => { const number = Number(link.dataset.step); const unlocked = number <= allowed; link.classList.toggle("active", number === state.step); link.classList.toggle("done", number < state.step); link.classList.toggle("locked", !unlocked); link.setAttribute("aria-disabled", String(!unlocked)); link.setAttribute("aria-current", number === state.step ? "step" : "false"); });
  const lockReasons = { 2: "Încarcă mai întâi exporturile pentru a debloca verificarea.", 3: "Confirmă datele și avertismentele pentru a debloca alegerile.", 4: "Alege structura raportului pentru a debloca descărcările." };
  $$("[data-step]").forEach((link) => { const number = Number(link.dataset.step); const unlocked = number <= allowed; link.classList.toggle("done", state.completedSteps.has(number)); if (!unlocked) { link.setAttribute("title", lockReasons[number]); link.setAttribute("aria-label", (link.querySelector("strong")?.textContent || "Pas") + ". " + lockReasons[number]); } else { link.removeAttribute("title"); link.removeAttribute("aria-label"); } });
  $("#to-step-2") && ($("#to-step-2").disabled = !ready(1));
  $("#to-step-3") && ($("#to-step-3").disabled = !ready(2));
  $("#to-step-4") && ($("#to-step-4").disabled = !ready(3));
  const enabled = ready(3) && !state.busy; ["#xlsx", "#csv-template", "#bundle", "#pptx-whole", "#pptx-main", "#pptx-appendix"].forEach((selector) => { const button = $(selector); if (button) button.disabled = !enabled; });
  guardControls();
  const chooseNote = $("#choose-note"); if (chooseNote) { const count = reportPlan && payload?.readiness ? reportPlan(payload, { scope: "whole" }).length : null; chooseNote.textContent = payload?.metadata?.annex === "none" ? "Fără anexă a fost aleasă; rezultatele individuale și analiza observațiilor dispar din livrare." : `${count === null ? "—" : count} ${count === 1 ? "slide" : "slide-uri"} după alegerile curente.`; }
}

function renderDownload() { const time = $("#receipt-time"); const receipt = $("#receipt"); if (time) time.textContent = state.receipts.length ? state.receipts.at(-1).time : "Încă nu ai descărcat un fișier"; if (receipt) receipt.innerHTML = state.receipts.length ? state.receipts.map((item) => `<div class="receipt-entry"><strong>${esc(item.name)}</strong><span>${esc(item.kind)} · ${esc(item.detail)} · ${esc(item.time)}</span></div>`).join("") : "<p>După prima descărcare vei vedea aici numele exact, tipul livrării și numărul de slide-uri.</p>"; }
function addReceipt(name, kind, detail) { const time = new Intl.DateTimeFormat("ro-RO", { dateStyle: "short", timeStyle: "short" }).format(new Date()); const normalizedDetail = detail.replace(/(\d+) slide-uri?/u, (_, count) => `${count} ${count === "1" ? "slide" : "slide-uri"}`); state.receipts.push({ name, kind, detail: normalizedDetail, time }); renderDownload(); announce(`${name} a fost pregătit pentru descărcare.`); }

function render() {
  const active = document.activeElement;
  const focus = active?.id ? { id: active.id, start: typeof active.selectionStart === "number" ? active.selectionStart : null, end: typeof active.selectionEnd === "number" ? active.selectionEnd : null } : null;
  const current = metadata(); if (files.length && !state.projectName) state.projectName = current.projectName || deriveProjectName(); if (!state.reportDate) state.reportDate = current.reportDate || today();
  payload = buildPayload(window.XLSX, files, metadata(), state.corrections, { acknowledgedWarningIds: [...state.acknowledged] });
  Object.assign(payload.metadata, metadata());
  const warningIds = new Set(payload.warnings.map(warningKey)); state.acknowledged = new Set([...state.acknowledged].filter((id) => warningIds.has(id))); payload.warnings = payload.warnings.map((item) => ({ ...item, reviewed: state.acknowledged.has(warningKey(item)) })); payload.warningReviews = payload.warnings; payload.readiness = payload.blockers.length === 0 && payload.warnings.every((item) => item.reviewed);
  renderFileCards(); renderFound(); renderDrawerInstructions(); Object.assign(payload.metadata, metadata()); renderReview(); renderStructure(); renderDownload(); sync(); refreshDeliveryCounts();
  if (focus) { const target = document.getElementById(focus.id); if (target) { target.focus({ preventScroll: true }); if (focus.start !== null && typeof target.setSelectionRange === "function") target.setSelectionRange(focus.start, focus.end); } }
}
function recompute() { render(); }

async function readSources(event) { const incoming = await Promise.all([...event.target.files].map(async (file) => ({ name: file.name, bytes: await file.arrayBuffer() }))); files = mergeSelectedFiles(files, incoming); state.acknowledged.clear(); state.corrections = { values: {} }; state.methodologyProposals = {}; state.completedSteps.clear(); state.projectName = state.projectName || deriveProjectName(); state.reportDate = state.reportDate || today(); state.step = 1; event.target.value = ""; render(); announce("Fișierele au fost citite. Verifică ce am găsit și apoi mergi la verificare."); $("#found-title")?.focus(); }
function activateStep(number) { const allowed = ready(3) ? 4 : ready(2) ? 3 : ready(1) ? 2 : 1; if (number > allowed) return; state.step = number; sync(); document.querySelector(`#step-${["upload", "review", "choose", "download"][number - 1]} h2`)?.focus(); document.querySelector(`#step-${["upload", "review", "choose", "download"][number - 1]}`)?.scrollIntoView({ block: "start" }); }

function expectedNames(kind) { const name = slug(payload?.metadata?.projectName); return { whole: `raport-trend-${name}.pptx`, main: `raport-trend-principal-${name}.pptx`, appendix: `raport-trend-anexa-${name}.pptx`, bundle: `pachet-bhb-${name}.zip`, xlsx: `audit-raport-grup-${name}.xlsx`, "csv-template": `sablon-declinatii-${name}.csv` }; }
async function createDownload(kind) {
  if (state.busy || !ready(3)) return; const trigger = document.activeElement; state.busy = true; sync(); const names = expectedNames(kind); const status = $("#download-status"); try {
    if (kind === "xlsx") { const bytes = XLSX.write(createAuditWorkbook(XLSX, payload), { type: "array", compression: true, bookType: "xlsx" }); download(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), names.xlsx); addReceipt(names.xlsx, "fișier de lucru", "audit calculat"); }
    else if (kind === "csv-template") { const bytes = XLSX.write(createEvaluationSheetTemplate(XLSX, payload), { type: "string", bookType: "csv" }); download(new Blob([bytes], { type: "text/csv;charset=utf-8" }), names["csv-template"]); addReceipt(names["csv-template"], "fișier de lucru", "șablon declinații"); }
    else if (kind === "bundle") { await downloadBundle(payload, names.bundle); addReceipt(names.bundle, "livrabil client", "Pachet BHB"); }
    else { const scope = kind === "main" ? "main" : kind === "appendix" ? "appendix" : "whole"; await downloadPptx(payload, names[kind] || names.whole, { scope }); const count = reportPlan(payload, { scope }).length; addReceipt(names[kind] || names.whole, "livrabil client", `${count} slide-uri`); }
    if (status) status.textContent = "Livrabilul a fost pregătit. Chitanța de mai jos a fost actualizată.";
  } catch (error) { if (status) status.textContent = `Descărcarea nu a pornit: ${error.message || error}`; } finally { state.busy = false; sync(); trigger?.focus?.(); }
}

function renderDrawerInstructions() {
  const examples = { evaluators: "Ana Popescu, Mihai Ionescu", days: "2", "team-size": "3 consultanți", "evaluation-period": "01.10.2026 – 02.10.2026", "population-role": "20 participanți (12 manageri, 8 specialiști)", location: "București" };
  Object.entries(examples).forEach(([id, example]) => { const input = $("#" + id); if (input) input.placeholder = "De completat · Exemplu: " + example; });
}

function refreshDeliveryCounts() {
  if (!payload?.readiness) return;
  const annex = payload.metadata.annex || "end";
  const whole = reportPlan(payload, { scope: "whole" });
  const main = reportPlan(payload, { scope: "main" });
  const appendix = reportPlan(payload, { scope: "appendix" });
  const appendixCount = annex === "separate" ? appendix.length : whole.filter((slide) => slide.deliverable === "appendix").length;
  const principalCount = main.length;
  const total = annex === "separate" ? main.length + appendix.length : whole.length;
  const label = (count) => String(count) + " " + (count === 1 ? "slide" : "slide-uri");
  const note = annex === "separate"
    ? label(total) + " în 2 fișiere: " + label(principalCount) + " principal · " + label(appendixCount) + " anexă."
    : label(total) + " într-un singur fișier: " + label(principalCount) + " principal · " + label(appendixCount) + " anexă.";
  if ($("#slide-total")) $("#slide-total").textContent = label(total);
  $("#structure-summary")?.parentElement.querySelectorAll(".structure-summary-note").forEach((node) => { node.textContent = note; });
  if ($("#choose-note")) $("#choose-note").textContent = annex === "none" ? label(total) + " într-un singur fișier; fără anexă." : note + " După alegerile curente.";
}

function resetSession() {
  files = []; payload = null; state.step = 1; state.acknowledged.clear(); state.corrections = { values: {} }; state.methodologyProposals = {}; state.completedSteps.clear(); state.projectName = ""; state.reportDate = ""; state.receipts = []; state.previewOpen = false;
  $("#sources").value = ""; $("#reset-confirm").hidden = true; $("#download-fallback").hidden = true; $("#preview") && ($("#preview").innerHTML = "");
  ["benchmark-low", "benchmark-high"].forEach((id, index) => { const input = $(`#${id}`); if (input) input.value = index ? "3.5" : "2.75"; });
  $$(`input[name="annex"]`).forEach((input) => { input.checked = input.value === "end"; }); $("#split-groups") && ($("#split-groups").checked = false);
  ["evaluators", "days", "team-size", "evaluation-period", "population-role", "location"].forEach((id) => { const input = $(`#${id}`); if (input) { input.value = ""; delete input.dataset.userEdited; } });
  render(); announce("Sesiunea a fost curățată. Fișierele sursă nu au fost schimbate."); $("#upload-title")?.focus();
}

$("#sources")?.addEventListener("change", readSources);
$("#choose-files")?.addEventListener("click", () => $("#sources")?.click());
$("#drop-zone")?.addEventListener("dragover", (event) => { event.preventDefault(); $("#drop-zone").classList.add("is-dragging"); });
$("#drop-zone")?.addEventListener("dragleave", () => $("#drop-zone").classList.remove("is-dragging"));
$("#drop-zone")?.addEventListener("drop", (event) => { event.preventDefault(); $("#drop-zone").classList.remove("is-dragging"); const input = $("#sources"); const transfer = event.dataTransfer; if (input && transfer?.files?.length) { const dt = new DataTransfer(); [...transfer.files].forEach((file) => dt.items.add(file)); input.files = dt.files; input.dispatchEvent(new Event("change", { bubbles: true })); } });
$("#to-step-2")?.addEventListener("click", () => { state.completedSteps.add(1); activateStep(2); }); $("#to-step-3")?.addEventListener("click", () => { state.completedSteps.add(2); activateStep(3); }); $("#to-step-4")?.addEventListener("click", () => { state.completedSteps.add(3); activateStep(4); });
$$(`[data-back-step]`).forEach((button) => button.addEventListener("click", () => activateStep(Number(button.dataset.backStep))));
$("#confirm-all")?.addEventListener("click", () => { payload.warnings.forEach((item) => state.acknowledged.add(warningKey(item))); announce("Toate avertismentele au fost confirmate."); render(); $("#to-step-3")?.focus(); });
$$(`#workflow-steps [data-step]`).forEach((link) => link.addEventListener("click", (event) => { event.preventDefault(); activateStep(Number(link.dataset.step)); }));
$$(`[data-section]`).forEach((section) => section.addEventListener("click", (event) => { const anchor = event.target.closest("a[data-step]"); if (anchor) { event.preventDefault(); activateStep(Number(anchor.dataset.step)); } }));
$("#split-groups")?.addEventListener("change", recompute); $("#benchmark-low")?.addEventListener("change", recompute); $("#benchmark-high")?.addEventListener("change", recompute); $$(`input[name="annex"]`).forEach((input) => input.addEventListener("change", () => { $$(`.option`).forEach((option) => option.classList.toggle("selected", option.querySelector("input")?.checked)); recompute(); }));
$("#methodology-text")?.addEventListener("change", recompute); ["evaluators", "days", "team-size", "evaluation-period", "population-role", "location"].forEach((id) => { const input = $(`#${id}`); input?.addEventListener("input", () => { setProposal(input); recompute(); }); input?.addEventListener("change", () => { setProposal(input); recompute(); }); }); $("#conclusions")?.addEventListener("change", recompute); $("#conclusions-destination")?.addEventListener("change", recompute);
$("#preview-trigger")?.addEventListener("click", () => { state.previewOpen = !state.previewOpen; const root = $("#preview"); if (!root) return; root.hidden = !state.previewOpen; $("#preview-trigger").textContent = state.previewOpen ? "Ascunde structura slide-urilor" : "Vezi structura slide-urilor"; if (state.previewOpen && payload?.readiness) { mountPreview(root, payload, { scope: "whole" }); root.querySelector(".preview-stage")?.focus(); } });
[["#xlsx", "xlsx"], ["#csv-template", "csv-template"], ["#csv-template-import", "csv-template"], ["#bundle", "bundle"], ["#pptx-whole", "whole"], ["#pptx-main", "main"], ["#pptx-appendix", "appendix"]].forEach(([selector, kind]) => $(selector)?.addEventListener("click", () => createDownload(kind)));
$("#reset")?.addEventListener("click", () => { $("#reset-confirm").hidden = false; $("#reset-cancel")?.focus(); }); $("#reset-confirm-yes")?.addEventListener("click", resetSession); $("#reset-cancel")?.addEventListener("click", () => { $("#reset-confirm").hidden = true; $("#reset")?.focus(); });

render(); window.__grfBooted?.();
