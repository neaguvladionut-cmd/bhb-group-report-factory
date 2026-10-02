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

const text = (value) => String(value ?? "").trim();
const warningKey = (item) => item.id || `${item.code}:${item.rowNumber || ""}:${item.message}`;
const esc = (value) => String(value ?? "").replace(/[&<>"']/gu, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);

function sourceRowCount(payload, sourceName) {
  const schema = payload?.schemas?.find((item) => item.sourceName === sourceName);
  return schema ? Math.max(1, schema.rows.length - schema.headerRow) : 0;
}

function likelyCause(payload, group) {
  const bySource = new Map();
  group.items.forEach((item) => {
    const source = item.sourceName || "fișierul importat";
    if (!bySource.has(source)) bySource.set(source, new Set());
    if (item.rowNumber) bySource.get(source).add(item.rowNumber);
  });
  for (const [source, rows] of bySource) if (rows.size / sourceRowCount(payload, source) >= .8) return `Probabil fișierul este greșit sau este alt tip de export: ${source}. Înlocuiește-l cu exportul potrivit.`;
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

function rowMatches(row, query) {
  const normalized = text(query).toLocaleLowerCase("ro");
  return !normalized || [row.source, row.row, row.label, row.field].join(" ").toLocaleLowerCase("ro").includes(normalized);
}

function rowMarkup(row) {
  return `<div class="issue-item"><div class="issue-item-copy"><strong>${esc(row.label)}</strong><small>${esc(row.source || "Fișier")}${row.row ? ` · rândul ${esc(row.row)}` : ""}${row.field ? ` · ${esc(row.field)}` : ""}</small></div></div>`;
}

export function renderIssueGroup(group, { documentRef = document, payload = null, onConfirm = () => {}, rowLimit = 8 } = {}) {
  const title = group.severity === "blocker" ? (issueTitles[group.key] || "Blocaj de verificare") : (warningTitles[group.key] || "Avertisment");
  const explanation = group.severity === "blocker" ? (issueExplanations[group.key] || group.items[0]?.message || "Verificarea nu a putut fi încheiată.") : (warningExplanations[group.key] || group.items[0]?.message || "Apare o diferență care cere confirmare.");
  const fix = group.severity === "blocker" ? (issueFixes[group.key] || "Corectează exportul și încarcă fișierele din nou.") : (warningFixes[group.key] || "Verifică exportul și confirmă dacă situația este intenționată.");
  const rows = affectedRows(group);
  const cause = likelyCause(payload, group);
  const id = `issue-${group.severity}-${group.key}`;
  const article = documentRef.createElement("article");
  article.className = `issue-group ${group.severity}`;
  article.dataset.issueGroup = id;
  article.dataset.issueExpanded = "false";
  article.innerHTML = `<div class="issue-group-head"><div class="issue-group-title"><small>${group.severity === "blocker" ? "BLOCAJ" : "AVERTISMENT"}</small><h3>${esc(title)}</h3><span class="issue-count" data-issue-count>${rows.length}</span></div><div class="issue-group-actions">${group.severity === "warning" ? `<button type="button" class="button button-secondary button-small" data-confirm-group="${esc(group.key)}">Confirmă acest grup</button>` : ""}</div></div><p class="issue-group-summary">${esc(explanation)}</p><p class="issue-fix">Cum repari: ${esc(fix)}</p>${cause ? `<p class="issue-fix">${esc(cause)}</p>` : ""}<div class="issue-tools"><label>Caută în rândurile afectate<input type="search" data-issue-search="${esc(id)}" placeholder="nume, fișier sau rând" autocomplete="off"></label><span class="source-chip">${esc(new Set(rows.map((row) => row.source).filter(Boolean)).size ? [...new Set(rows.map((row) => row.source).filter(Boolean))].join(" · ") : "exportul importat")}</span></div><div class="issue-panel"><button type="button" class="button button-secondary button-small issue-toggle" data-issue-toggle aria-expanded="false" aria-controls="${esc(id)}-items">Arată rândurile afectate</button><div class="issue-items" id="${esc(id)}-items" data-issue-items hidden></div><button type="button" class="button button-secondary button-small issue-reveal" data-issue-reveal hidden></button></div>`;

  const list = article.querySelector("[data-issue-items]");
  const count = article.querySelector("[data-issue-count]");
  const toggle = article.querySelector("[data-issue-toggle]");
  const reveal = article.querySelector("[data-issue-reveal]");
  const search = article.querySelector("[data-issue-search]");
  let expanded = false;
  let showAll = false;

  const drawRows = () => {
    const filtered = rows.filter((row) => rowMatches(row, search?.value || ""));
    const query = text(search?.value);
    if (count) count.textContent = query ? `${filtered.length} din ${rows.length}` : String(rows.length);
    if (!expanded) {
      list.replaceChildren();
      list.hidden = true;
    } else {
      const visible = showAll ? filtered : filtered.slice(0, rowLimit);
      list.hidden = false;
      list.innerHTML = visible.length ? visible.map(rowMarkup).join("") : `<p class="empty-value">Nu există rânduri pentru această căutare.</p>`;
    }
    const canReveal = expanded && filtered.length > rowLimit;
    if (reveal) {
      reveal.hidden = !canReveal;
      reveal.textContent = showAll ? "Arată mai puține" : `Arată toate (${filtered.length})`;
    }
  };

  toggle?.addEventListener("click", () => {
    expanded = !expanded;
    showAll = false;
    article.dataset.issueExpanded = String(expanded);
    toggle.setAttribute("aria-expanded", String(expanded));
    toggle.textContent = expanded ? "Ascunde rândurile" : "Arată rândurile afectate";
    drawRows();
  });
  reveal?.addEventListener("click", () => {
    showAll = !showAll;
    drawRows();
  });
  search?.addEventListener("input", () => {
    showAll = false;
    drawRows();
  });
  article.querySelector("[data-confirm-group]")?.addEventListener("click", () => onConfirm({ group, title }));
  drawRows();
  return article;
}
