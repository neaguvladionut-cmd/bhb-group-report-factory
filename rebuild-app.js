if (!window.XLSX || !window.JSZip) throw new Error("Lipsesc bibliotecile locale necesare");

let files = [];
let payload = null;
let activeDownloadUrl = "";
const state = { step: 1, acknowledged: new Set(), corrections: { values: {} }, busy: false };
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const warningKey = (item) => item.id || `${item.code}:${item.rowNumber || ""}:${item.message}`;
const text = (value) => String(value ?? "").trim();

function metadata() {
  const groupNames = Object.fromEntries($$(`[data-group-name]`).map((input) => [input.dataset.groupName, input.value]));
  const slideToggles = Object.fromEntries($$(`[data-slide-toggle]`).map((input) => [input.dataset.slideToggle, input.checked]));
  const groupConclusions = {};
  $$(`[data-group-conclusion]`).forEach((input) => { (groupConclusions[input.dataset.groupConclusion] ||= {})[input.dataset.field] = input.value; });
  return { projectName: $("#project-name")?.value || "", clientName: $("#client-name")?.value || "", reportDate: $("#report-date")?.value || "", context: $("#report-context")?.value || "", exercises: $("#exercise-list")?.value || "", otherInstruments: $("#other-instruments")?.value || "", conclusions: $("#conclusions")?.value || "", executiveConclusions: $("#conclusions")?.value || "", program: $("#program")?.value || "", methodologyText: $("#methodology-text")?.value ?? "", evaluators: $("#evaluators")?.value || "", days: $("#days")?.value || "", exerciseCount: $("#exercise-count")?.value || "", conclusionsStrengths: $("#conclusions-strengths")?.value || "", conclusionsDevelopment: $("#conclusions-development")?.value || "", conclusionsInterventions: $("#conclusions-interventions")?.value || "", benchmarkLow: $("#benchmark-low")?.value ?? "2.75", benchmarkHigh: $("#benchmark-high")?.value ?? "3.5", annex: $("#annex-setting")?.value || "end", splitGroups: Boolean($("#split-groups")?.checked), groupNames, groupConclusions, slideToggles };
}

function download(blob, name) {
  if (activeDownloadUrl) URL.revokeObjectURL(activeDownloadUrl);
  activeDownloadUrl = URL.createObjectURL(blob);
  const link = $("#download-fallback");
  if (!link) return;
  link.href = activeDownloadUrl;
  link.download = name;
  link.textContent = `Dacă descărcarea nu a pornit, apasă aici pentru ${name}.`;
  link.hidden = false;
  link.click();
}
window.__grfDownload = download;

function warningComplete() { return payload && payload.warnings.every((item) => item.reviewed); }
function ready(step) { if (step === 1) return files.length > 0; if (step >= 2) return Boolean(payload?.readiness && warningComplete()); return false; }
function invalidate() { if (state.step > 1) state.step = 1; render(); }

function renderReview() {
  const root = $("#issue-list");
  const overview = $("#review-overview");
  if (!root || !payload) return;
  const blockers = payload.blockers.length;
  const warnings = payload.warnings.filter((item) => !item.reviewed).length;
  if (overview) overview.innerHTML = `<div class="review-state ${blockers || warnings ? "needs-attention" : "all-clear"}"><strong>${blockers || warnings ? "Necesită atenție" : "Totul este pregătit"}</strong><span>${blockers} blocaje · ${warnings} avertismente de confirmat</span></div>`;
  root.replaceChildren();
  [...payload.blockers, ...payload.warnings.filter((item) => !item.reviewed)].forEach((item) => {
    const row = document.createElement("article");
    row.className = `issue-group ${item.severity}`;
    const copy = document.createElement("p");
    copy.textContent = item.message;
    row.append(copy);
    if (item.severity === "warning") {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "btn-ghost btn-sm";
      button.textContent = "Confirmă avertismentul";
      button.addEventListener("click", () => { state.acknowledged.add(warningKey(item)); render(); });
      row.append(button);
    }
    root.append(row);
  });
  if (!root.children.length) root.innerHTML = "<div class=\"review-empty\"><strong>Totul este pregătit.</strong><span>Poți continua.</span></div>";
}

function renderStructure() {
  const root = $("#structure-summary");
  if (!root || !payload) return;
  const plan = reportPlan(payload, { scope: "whole" });
  const main = plan.filter((slide) => slide.deliverable !== "appendix");
  const appendix = plan.filter((slide) => slide.deliverable === "appendix");
  root.innerHTML = `<div class="structure-card"><strong>Trend · raport principal</strong><span>${main.length} slide-uri</span><small>Întregul proiect primul; grupurile CODE urmează doar când activezi împărțirea.</small></div><div class="structure-card"><strong>Anexă</strong><span>${appendix.length ? `${appendix.length} slide-uri` : "dezactivată"}</span><small>${payload.metadata.annex === "separate" ? "Se descarcă separat." : payload.metadata.annex === "none" ? "Nu se generează." : "Se include la final."}</small></div>`;
  if (payload.metadata.annex === "none") {
    const note = document.createElement("p"); note.className = "structure-card methodology-warning"; note.setAttribute("role", "status");
    note.textContent = "Fără anexă: constatările cheie pe competențe și „Comportamente cheie” se află în anexă și nu vor fi generate.";
    root.append(note);
  }
  for (const competency of unrankedCompetencies(payload)) {
    const note = document.createElement("p"); note.className = "structure-card methodology-warning"; note.setAttribute("role", "status");
    note.textContent = `Competența ${competency} are un singur comportament evaluat: fără clasament (fără „Comportamente cheie” și fără casetele de abilități pe slide-ul de constatări).`;
    root.append(note);
  }
  const missingMethodology = methodologyColumns(payload).missingLabels || [];
  if (missingMethodology.length) {
    const warning = document.createElement("p");
    warning.className = "structure-card methodology-warning";
    warning.setAttribute("role", "status");
    warning.textContent = `Câmpuri de metodologie necompletate: ${missingMethodology.join(", ")}. Rândurile lor nu apar în raport până nu le completezi.`;
    root.append(warning);
  }
  if (payload.groups.length >= 2) {
    const groupSection = document.createElement("section");
    groupSection.className = "structure-card group-settings";
    groupSection.innerHTML = `<strong>Nume afișat pentru grupuri</strong><small>Valorile CODE sunt implicite și pot fi schimbate fără a schimba datele.</small>`;
    payload.groups.forEach((group) => { const label = document.createElement("label"); label.textContent = group.code; const input = document.createElement("input"); input.dataset.groupName = group.code; input.value = payload.metadata.groupNames?.[group.code] || group.code; input.addEventListener("input", invalidate); label.append(input); groupSection.append(label); });
    root.append(groupSection);
    if (payload.metadata.splitGroups) {
      const fields = [["strengths", "Abilități cheie"], ["development", "Arii de dezvoltare"], ["interventions", "Posibile intervenții"]];
      const conclusions = document.createElement("section");
      conclusions.className = "structure-card group-settings";
      conclusions.innerHTML = `<strong>Concluzii și recomandări pe grupuri</strong><small>Fiecare grup are propriul slide de concluzii; câmpurile goale rămân de completat în PowerPoint.</small>`;
      payload.groups.forEach((group) => fields.forEach(([field, title]) => { const label = document.createElement("label"); label.textContent = `${payload.metadata.groupNames?.[group.code] || group.code} – ${title}`; const area = document.createElement("textarea"); area.rows = 3; area.dataset.groupConclusion = group.code; area.dataset.field = field; area.value = payload.metadata.groupConclusions?.[group.code]?.[field] || ""; area.addEventListener("change", invalidate); label.append(area); conclusions.append(label); }));
      root.append(conclusions);
    }
  }
}

function sync() {
  document.body.dataset.workflowStep = String(state.step);
  $$(`[data-panel]`).forEach((panel) => { panel.hidden = Number(panel.dataset.panel) !== state.step; });
  $$(`#workflow-steps [data-step]`).forEach((button) => { const step = Number(button.dataset.step); button.disabled = step > state.step && !ready(step - 1); button.className = `${step === state.step ? "active " : ""}${step < state.step ? "done" : ""}`; });
  $("#to-step-2") && ($("#to-step-2").disabled = !ready(1));
  $("#to-step-3") && ($("#to-step-3").disabled = !ready(2));
  $("#to-step-4") && ($("#to-step-4").disabled = !ready(3));
  const canExport = ready(4) && !state.busy;
  if ($("#csv-template-import")) $("#csv-template-import").disabled = state.busy;
  ["#xlsx", "#csv-template", "#bundle", "#pptx-whole", "#pptx-main", "#pptx-appendix"].forEach((selector) => { if ($(selector)) $(selector).disabled = !canExport; });
  if ($("#pptx-whole")) $("#pptx-whole").hidden = payload?.metadata.annex === "separate";
  if ($("#pptx-appendix")) $("#pptx-appendix").hidden = payload?.metadata.annex !== "separate";
  renderStructure();
  const stage = $("#stage-summary");
  if (stage) stage.textContent = ["Încarcă exporturile standard și completează numai faptele confirmate.", "Rezolvă blocajele și confirmă avertismentele.", "Alege secțiunile, benchmarkul și anexa.", "Descarcă livrabilele Trend, bundle-ul BHB și auditul."][state.step - 1];
}

function render() {
  payload = buildPayload(window.XLSX, files, metadata(), state.corrections, { acknowledgedWarningIds: [...state.acknowledged] });
  Object.assign(payload.metadata, metadata());
  const warningIds = new Set(payload.warnings.map(warningKey));
  state.acknowledged = new Set([...state.acknowledged].filter((id) => warningIds.has(id)));
  const filesNode = $("#files");
  if (filesNode) filesNode.textContent = files.length ? files.map((file) => file.name).join(" · ") : "Niciun fișier selectat";
  const metrics = $("#metrics");
  if (metrics) metrics.replaceChildren(...payload.calculations.map((item) => { const node = document.createElement("article"); node.innerHTML = `<strong>${item.mean?.toFixed(2) ?? "—"}</strong><span>${item.competency}</span><small>min ${item.min?.toFixed(2) ?? "—"} · mediană ${item.median?.toFixed(2) ?? "—"} · max ${item.max?.toFixed(2) ?? "—"} · n ${item.n}</small>`; return node; }));
  const preview = $("#preview");
  if (preview) { if (payload.readiness) mountPreview(preview, payload, { scope: "whole" }); else preview.innerHTML = "<p class=\"preview-empty\">Încarcă exporturile complete pentru a vedea structura raportului.</p>"; }
  renderReview();
  guardControls();
  sync();
}
// F25: a control without data is unchecked, disabled and says why.
function guardControls() {
  if (!payload) return;
  const zone = $(`[data-slide-toggle="zone"]`);
  const imported = files.length > 0;
  if (zone) { const missing = imported && !payload.regionReadiness.available; if (missing) zone.checked = false; zone.disabled = missing; zone.title = missing ? payload.regionReadiness.disabledReason : ""; setReason(zone, missing ? payload.regionReadiness.disabledReason : ""); }
  const split = $("#split-groups");
  if (split) { const missing = imported && !payload.codeReadiness.splitAvailable; if (missing) split.checked = false; split.disabled = missing; const reason = missing ? "Exportul are mai puțin de două valori CODE; împărțirea pe grupuri nu este disponibilă." : ""; split.title = reason; setReason(split, reason); }
}
function setReason(input, reason) {
  const label = input.closest("label"); if (!label) return;
  let note = label.querySelector(".control-reason");
  if (!reason) { note?.remove(); return; }
  if (!note) { note = document.createElement("small"); note.className = "control-reason"; label.append(note); }
  note.textContent = reason;
}

async function readSources(event) { files = mergeSelectedFiles(files, await Promise.all([...event.target.files].map(async (file) => ({ name: file.name, bytes: await file.arrayBuffer() })))); state.acknowledged.clear(); state.corrections = { values: {} }; event.target.value = ""; invalidate(); }
async function createDownload(kind) {
  if (state.busy || (kind !== "csv-template" && !ready(4))) return;
  state.busy = true;
  sync();
  const slug = payload.metadata.projectName.replace(/[^a-z0-9]+/giu, "-") || "raport-grup";
  try {
    if (kind === "xlsx") { const bytes = XLSX.write(createAuditWorkbook(XLSX, payload), { type: "array", compression: true, bookType: "xlsx" }); download(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `audit-raport-grup-${slug}.xlsx`); }
    else if (kind === "csv-template") { const bytes = XLSX.write(createEvaluationSheetTemplate(XLSX, payload), { type: "string", bookType: "csv" }); download(new Blob([bytes], { type: "text/csv;charset=utf-8" }), "evaluation-sheet-template.csv"); }
    else if (kind === "bundle") await downloadBundle(payload, `bhb-bundle-${slug}.zip`);
    else { const scope = kind === "main" ? "main" : kind === "appendix" ? "appendix" : "whole"; await downloadPptx(payload, `${scope === "appendix" ? "anexa" : scope === "main" ? "raport-principal" : "raport-trend"}-${slug}.pptx`, { scope }); }
    $("#actions-help") && ($("#actions-help").textContent = "Livrabilul a fost pregătit pentru descărcare.");
  } catch (error) { $("#actions-help") && ($("#actions-help").textContent = `Descărcarea nu a pornit: ${error.message || error}`); }
  state.busy = false;
  sync();
}

$("#sources")?.addEventListener("change", readSources);
["#project-name", "#client-name", "#report-date", "#report-context", "#program", "#evaluators", "#days", "#exercise-count", "#exercise-list", "#other-instruments", "#methodology-text", "#conclusions", "#conclusions-strengths", "#conclusions-development", "#conclusions-interventions", "#benchmark-low", "#benchmark-high", "#annex-setting", "#split-groups"].forEach((selector) => $(selector)?.addEventListener("input", invalidate));
$$(`[data-slide-toggle]`).forEach((input) => input.addEventListener("change", invalidate));
$("#to-step-2")?.addEventListener("click", () => { state.step = 2; render(); });
$("#to-step-3")?.addEventListener("click", () => { state.step = 3; render(); });
$("#to-step-4")?.addEventListener("click", () => { state.step = 4; render(); });
$$(`[data-step]`).forEach((button) => button.addEventListener("click", () => { const step = Number(button.dataset.step); if (step <= state.step || ready(step - 1)) { state.step = step; render(); } }));
$("#reset")?.addEventListener("click", () => { files = []; state.step = 1; state.acknowledged.clear(); state.corrections = { values: {} }; $("#sources").value = ""; render(); });
[["#xlsx", "xlsx"], ["#csv-template", "csv-template"], ["#csv-template-import", "csv-template"], ["#bundle", "bundle"], ["#pptx-whole", "whole"], ["#pptx-main", "main"], ["#pptx-appendix", "appendix"]].forEach(([selector, kind]) => $(selector)?.addEventListener("click", () => createDownload(kind)));
render();
window.__grfBooted?.();
