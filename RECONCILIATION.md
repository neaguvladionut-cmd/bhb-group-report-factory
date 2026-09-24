# Group Report Factory reconciliation gate

Status: RESOLVED — pending Inspector

## First task

Resolve the differences between the existing GRF design branch and the incoming GRF implementation packet before doing any other work.

- Destination baseline: the existing GRF design branch at the branch point for this reconciliation.
- Incoming material: the GRF-specific source, tests, build changes and safe assets from the platform-side work packet.
- Do not copy governance notes, project records, human-document work, client exports, participant data, screenshots or unreviewed reference material into this public repository.
- Preserve the source-to-deploy contract: source changes must be built into the deploy surface and independently checked.

## Branch gate

Until the reconciliation and privacy review are resolved:

- no feature development, refactoring, deployment or merge to `main`;
- no publication of unreviewed assets or generated Office material;
- no client, participant, project or real-person identity may enter this repository;
- no claim that the GRF output is accepted or production-ready;
- every conflict disposition must identify the winning behavior, the preserved behavior, and the verification that proves the result.

The fresh Engineer session that resolves this branch must update this file with the conflict disposition, privacy checks and source-to-deploy verification. A separate Inspector session must review the resolved diff before any merge or branch cleanup.

---

# Resolution (2026-09-24, Claude / Engineer)

The branch gate above still applies until a separate Inspector session has reviewed this record and
the branch diff. Nothing here merges to `main`, deploys, or claims that the GRF output is accepted.
The live page (this repository's root files) is unchanged: `git diff origin/main HEAD` over every file
on `main` is empty.

## Scope and method

- **Incoming platform source:** `bhb-platform` `origin/master` (`4b9a5aa`), folder
  `40-standalone/group-report-factory/`, unchanged there since `37f9a7e`. 105 files.
- **Destination:** this branch. Its base `7b7165e` is `feature/design-alignment-v2` (`a6f216d`) plus
  this file. The design branch begins with `849cb48`, which imported the platform folder and applied
  the redesign in the same commit. There is no import-only commit.
- **Method:** a git blob SHA comparison of every path. Then every differing file was diffed with both
  sides formatted by Prettier 3.8.1, so the dense one-line sources diff line by line. Every removed or
  changed platform behavior is classified as *re-implemented*, *dropped by the redesign* (commit cited)
  or *lost*.

## 1. File-by-file result (105 platform files)

| Result | Count | Files |
|---|---|---|
| Byte-identical (same blob) | 86 | `src/core.js`, `deploy/core.js`, `tests/core.test.mjs`, `package.json`; all 35 images and 6 vendor/licence files in `src/assets/`, plus their 41 `deploy/assets/` copies |
| Differ: redesign work | 15 | `README.md`, `tools/build.mjs`, `src/{app.js,index.html,pptx.js,report-plan.js,styles.css}`, `tests/{launch,pptx,preview}.test.mjs`, and the five built `deploy/` counterparts |
| Platform-only (removed on purpose) | 4 | `src/preview.js`, `src/cover-preview.css`, `deploy/preview.js`, `deploy/cover-preview.css` |

**Branch-only files.** The redesign's PDF-first additions: `src/pdf.js`, `tests/pdf.test.mjs`,
`src/assets/fonts/*` (Noto Sans + OFL), and pdf-lib, PDF.js, fontkit, their licences and
`vendor-manifest.json` in `src/assets/vendor/`, each with its `deploy/` copy. This reconciliation adds
`tests/report-plan.test.mjs` and `tests/download-recovery.test.mjs` (see *Lost and restored*). The root
live bundle (`index.html`, `app.js`, `core.js`, `pptx.js`, `preview.js`, `report-plan.js`, `styles.css`,
`cover-preview.css`, `assets/**`) is identical to `main`.

### Dispositions

| Area | Winning behavior (branch) | Preserved platform behavior | Proving check |
|---|---|---|---|
| Wizard (`app.js`, `index.html`) | Three steps, Încarcă / Revizuiește (seven sub-steps) / Generează (`849cb48`) | Warning-acknowledgement gate (`stepReady`, `warningComplete`), review filter and search, grouped issues with "Confirmă grupul" and "Înlocuiește fișierul", in-context correction editor, review summary stats, structure cards, download-error message, shared download helper plus fallback link, boot fallback banner, `#reset` (now also `#reset-session`), same-name file replacement (in core) | `launch.test.mjs` ("warning acknowledgement…", "substep ARIA…"), `core.test.mjs` "sequential file selections accumulate and same-name reselection replaces", restored `download-recovery.test.mjs` |
| Step 4 "Exportă" | Folded into step 3 "Generează" (`849cb48`) | All four platform downloads (XLSX, whole, main and appendix PPTX), plus the new PDF | `launch.test.mjs`, `pdf.test.mjs` |
| HTML slide preview (`preview.js`, `cover-preview.css`) | The PDF is the review artifact: `pdf.js`, a PDF.js canvas and an open-in-new-window fallback (`72655b1`; `f022bea` removed the fake preview chrome first) | The report inventory is still `reportPlan`, and the PDF page count equals it | `pdf.test.mjs` "PDF bytes are real… match the whole report plan", `preview.test.mjs` "all section toggles preserve reportPlan…" |
| "Indicatori calculați" metrics panel (mean/min/median/max/n per competency) | Removed with the old report surface in `849cb48` (its message does not name it) | The same figures remain in the XLSX audit "Calcule" sheet and in the range and ranking slides | `core.test.mjs` "calculations include the median", "audit workbook preserves calculation…". **Design question: Finding 1** |
| Section controls (`report-plan.js`, `pptx.js`) | Eight section toggles, on by default, applied to the plan, the PDF and both PPTX renderers (`849cb48`, `ea0087e`) | With every toggle on, the platform slide families and order are unchanged | `pptx.test.mjs` "both PPTX renderers omit every disabled family…", restored `report-plan.test.mjs` |
| BHB cover | Wrapped, auto-sized title and metadata block with `Client:`, `Data raportului:` and `Context:` lines (`bd345da`, `a10809b`) | Project-or-client title, client shown only when it differs, report date | `pptx.test.mjs` (long-cover and `Client:` assertions) |
| BHB behavior slides | Dark background with white logo and text; the red "Arii prioritare" heading becomes white (`bd345da`) | Same two columns and the same selected behaviors | `pptx.test.mjs`, restored "behavior insights select deterministic…" |
| TREND cover | Project-first title plus a `Client:` line; the "Raport de grup" label becomes `context · date` (`427d655`) | Title slot and TREND frame | `pptx.test.mjs` TREND delivery test renders it. The new label text is not asserted. |
| `reportPlan` participant sort | `records.slice().sort` with no in-place mutation (`849cb48`) | Same ordering | restored "report plan paginates long participants…" |
| `tools/build.mjs` | Bundles `pdf.js` instead of `preview.js` and drops `preview.js` and `cover-preview.css` from `deploy/` (`72655b1`) | Same classic-script, `file://` bundle model | §4 |
| Copy removed with the old markup | Shorter boot-fallback text (`72655b1`), shorter privacy line, "Sesiune nouă" instead of "Șterge sesiunea" (`849cb48`) | Boot fallback with runtime message; the in-memory-only statement | **Copy: Finding 3** |
| CSS | Redesigned stylesheet. Every removed selector is presentational (hero, metrics, old steps, old warning feed). | `body[data-booted="true"] #boot-fallback`, `[hidden]` rules | `launch.test.mjs` boot test |

### Lost and restored (separate commit)

No platform *product* fix was lost. Every removed code path was either re-implemented or belonged to
a surface the redesign removed. What was lost was **regression proof for code that is still live**. When
`72655b1` deleted `tests/preview.test.mjs` and rewrote `tests/launch.test.mjs`, these platform tests went
with them. They are restored from `37f9a7e`:

- `tests/report-plan.test.mjs`, 4 tests, verbatim: metadata and optional-conclusions visibility;
  compact methodology that never lists dates or locations; long-participant pagination that keeps
  full behavior sentences; deterministic top score-2 and score-0 behavior insights.
- `tests/download-recovery.test.mjs`, 1 test: the fallback link, `reportDownloadError`, the shared
  `window.__grfDownload`, and the PPTX use of that helper. One regex now accepts spaces around `=`,
  matching the redesign's source style.

One platform test is still without an equivalent. The behavior exists in code, but the test needs a
new DOM harness, so it was not restored (Finding 4). It is the end-to-end test "deployed replacement,
reset and reload paths stay fresh and escape hostile workbook headers". A code read shows the branch
writes workbook-derived text only through `textContent`, and every `innerHTML` interpolation is numeric
or constant. The PDF path has its own hostile-text test.

## 2. Romanian copy from `main` `dd56952`

`dd56952` changed only the published root `index.html`, 4 strings. It is an ancestor of this branch
(`849cb48` builds on it).

| `dd56952` change | On this branch (`src/index.html` → `deploy/index.html`) |
|---|---|
| Hero `h1` "Din exporturile AC, în…" → "Transformă exporturile AC în livrabile de grup verificabile." | The redesigned markup has no hero `h1`. Not carried over (Finding 2). |
| Legend "Renderer" → "Motor de randare" | Superseded by the legend "Stilul prezentării". No English label remains. |
| "Automatizează template-urile TREND" → "…șabloanele TREND" | Present in source and deploy (since `849cb48`) |
| Delivery captions → "Raportul principal și anexa în același fișier" / "…livrate separat" | Superseded by rewritten captions in the same register: "Raportul principal este urmat de anexă în aceeași prezentare." / "Primești separat raportul principal și anexa cu rezultatele detaliate." |

None of the pre-`dd56952` strings ("Renderer", "template-urile", "Raport principal și anexă…") remain
in `src/` or `deploy/`.

## 3. Privacy review

Nothing was removed. The dispositions below were first proposals for the Architect or Vlad. The Disposition column now records Vlad's ruling, with the original proposal in parentheses where it adds context. The coverage check below confirms that every image file is named in the table (GRF-INS-01).

**Ruled by Vlad, 2026-09-24:** every image and slide in this table, including the photo, the Trend
template slides and the example content on slides 5, 7, 8, 12 and 21, comes from the BHB/Trend library
and is cleared for use and public hosting. **All rows: keep.** Recorded by the Architect.

| Asset (each file is in `src/assets/images/`, with byte-identical copies in `deploy/assets/images/` and the live root `assets/images/`) | What it shows | Origin | Risk | Disposition |
|---|---|---|---|---|
| `bhb-observer-photo.jpg` (1080×1080, black and white) | A smiling woman in a blazer holding a tablet in an open office, with three blurred people behind her. The face is recognisable. | Added to `bhb-platform` in `788ca71` (2026-08-31, "Increment 3 — Apply group report reference styling") and unchanged since. In this repository since `1d863d1` and `849cb48`. XMP: made in Photoshop 24.3 (Mac) on 2023-05-28, before the project began, like the other BHB brand-deck assets. No stock-agency or licence metadata. | A recognisable face of an unidentified person, licence unknown at review time. Not participant data. | **Keep, cleared by Vlad 2026-09-24** (proposed: Vlad to confirm it is licensed BHB brand or stock imagery; confirmed by the ruling that it comes from the BHB/Trend library) |
| `trend-template-slide-1.png`, `trend-template-slide-3.png`, `trend-template-slide-11.png`, `trend-template-slide-18.png`, `trend-template-slide-20.png` | Trend template photos: a meeting room with several faces (1, which also carries "2024 © www.trendconsult.eu"), a building (3), a man in a teal shirt (11, 18), a woman at a desk (20) | Trend group-report template (`Template - raport de grup - RO.pptx`) | Faces in stock-style photography from Trend's template | **Keep, cleared by Vlad 2026-09-24** (proposed: keep if Trend's template licence covers public hosting) |
| `trend-template-slide-5.png`, `trend-template-slide-7.png`, `trend-template-slide-8.png`, `trend-template-slide-12.png`, `trend-template-slide-21.png` | Filled-in example content: 35 participant averages without names, 17/51/31 % banding, competency means (Colaborare și asertivitate 3.42, People Management 2.92, Gestionarea schimbării 2.83), a sales-competency distribution, behavior lists, and a "Concluzii și recomandări" page | Same Trend template | Looked like aggregate results from a real engagement baked into the template. No names and no client name visible. | **Keep, cleared by Vlad 2026-09-24** (proposed: accept as anonymised template residue or replace with blanked frames; the renderer tests already require these regions to be cleared in output) |
| `trend-template-slide-2.png`, `trend-template-slide-4.png`, `trend-template-slide-9.png`, `trend-template-slide-13.png`, `trend-template-slide-19.png` | Static frames (GRF-INS-01, inspected 2026-09-24). **2**: methodology page, "Privire de ansamblu asupra proiectului – Metodologie", with `#` placeholders ("# participanți Client", "# consultanți Trend", "# zi de evaluare… sediul CUI") and generic method text naming BHB Profiler. **4**: empty median range chart, axis 1–5, one category `C1`, BHB/Trend legend and benchmark note. **9**: participant bar chart with placeholder series `C1` ×5 and a single participant `P1`. **13**: competency distribution chart, one bar `P1` with sample labels (4.25, 3.75, 2.75, 2.50, 2.25). **19**: empty "Comportamente cheie" table with the headers "Abilități cheie" and "Abilități de dezvoltat". | Same Trend template | None. The Inspector's finding is **confirmed**: placeholders only (`#`, `C1`, `P1`), with no client, participant or real-person data. | **Keep, cleared by Vlad 2026-09-24** |
| `trend-template-slide-22.png`, `trend-template-slide-1.png` | Trend's Bucharest address, office e-mail, QR code (22); trendconsult.eu (1) | Trend template | Trend company data (allowed) | **Keep, cleared by Vlad 2026-09-24** |
| `bhb-closing-template.png`, closing text in `pptx.js` | BHB name, `office@bh-bar.biz`, phone `+40 742 123 456`, placeholder street | BHB brand | BHB data (allowed) | **Keep, cleared by Vlad 2026-09-24** (proposed: Vlad to confirm the phone number is a placeholder) |
| `bhb-360-profiler-report-illustration.webp`, `bhb-competency-profiler-report-illustration.webp` | Line-art cartoon figures with a sample chart | BHB brand | None | **Keep, cleared by Vlad 2026-09-24** |
| `bhb-behavior-background.png`, `bhb-behavior-divider.png`, `bhb-content-white.png`, `bhb-cover-template.png`, `bhb-distribution-divider.png`, `bhb-intro-divider.png`, `bhb-logo-full-white.png`, `bhb-logo-on-white.png`, `bhb-logo.svg`, `bhb-observations-background.png`, `bhb-photo-divider.png`, `bhb-range-legend-template.png`, `bhb-report-background.png`, `bhb-section-gradient.png`, `bhb-symbol.svg` | Gradients and backgrounds, logos and symbol, an empty white content frame, the cover template ("Lorem Ipsum"), and the range legend text | BHB brand | None | **Keep, cleared by Vlad 2026-09-24** |
| Test fixtures (`tests/*.mjs`) | "D210 Ana", "D210 Bogdan", "Ana Popescu", "Bogdan Ionescu", "Participant N", "Client Test/PDF/fixture", "Trend fixture", "Consultant 1/2", "București", generic competency and behavior text | Synthetic, built inside the tests. The repository holds no `.xlsx`, `.csv`, `.pptx` or `.pdf` files. | Common first-name and surname pairs, not tied to anyone | **Keep, cleared by Vlad 2026-09-24** |
| UI placeholders | "ex. AC Leadership 2026", "ex. Numele clientului", and similar | Source | None | **Keep, cleared by Vlad 2026-09-24** |

## 4. Source → deploy and tests (2026-09-24)

- Running `npm run build` and then `git status --short` and `git diff --stat` produces **no output**,
  both before and after the reconciliation commits. The build reproduces every committed `deploy/`
  file byte for byte.
- `npm test` on Node 22.22.2, the container default. Before the reconciliation commits, without
  poppler: 30 tests, 28 pass, 2 fail (`pptx.test.mjs` counts as one entry because it fails to load).
  After them, with the 5 restored tests and with poppler installed: 35 tests, 32 pass, 3 fail. The
  restored tests all pass. The third failure is the `pdftotext` finding below. The two baseline
  failures are:
  - `tests/pptx.test.mjs` does not load. It imports PptxGenJS, JSZip and pngjs from
    `/Users/vladneagu/.cache/codex-runtimes/...`. This is the known Mac-only test (bhb-platform #31),
    recorded here and not fixed.
  - "PDF preserves exact Romanian diacritics in embedded text" fails with
    `Promise.try is not a function`, thrown by the vendored PDF.js that the test uses to extract text.
    It fails on Node 24.21 too (`toHex`) and **passes on Node 25.9**. This is a runtime requirement of
    the test, not a product defect.
- Extra evidence, on Node 25.9 with the Mac-only imports redirected to npm `pptxgenjs@4.0.1`, `jszip`
  and `pngjs`, and with poppler installed:
  - The four PPTX tests load. Three pass. "TREND rendered routes clear baked template pixels" cannot
    run here because the container's LibreOffice cannot convert files.
  - With `pdftotext` present, the normally skipped branch of "PDF follows selected-section omission …
    hostile text safely" runs and **fails**. The extracted text is `Titlu script alert(1) /script`,
    and the test expects `Titlu script alert(1) script`. Its raw-byte assertion that no `<script>`
    reaches the PDF passes. The fault is the test's expected string, not the product. The test was
    not changed because it is the redesign's own (Finding 5).

## Findings for the Architect / Vlad (not decided here)

1. `849cb48` dropped the "Indicatori calculați" per-competency metrics panel. Its figures survive in
   the audit and in the PDF/PPTX. Should the redesign show them before generation?
2. The redesigned markup has no slot for the `dd56952` hero line.
3. The redesign dropped some guidance copy: "Corecțiile se aplică doar sesiunii locale. Fișierele
   originale nu sunt modificate…", "Poți selecta exporturile împreună sau pe rând; un fișier cu același
   nume îl înlocuiește…", and "…sau intervenție în aplicațiile legacy" from the privacy line. Also,
   `README.md` still names the button „Șterge sesiunea”, but the UI now labels it "Sesiune nouă".
4. The end-to-end replacement/reset/reload/hostile-DOM regression test has no equivalent (§1).
5. The hostile-text expectation in `tests/pdf.test.mjs` (`/script` vs `script`) fails wherever
   `pdftotext` is installed.
6. The diacritics test needs Node 25 or later. This belongs in #31 or the README.
7. ~~The privacy dispositions in §3 need Vlad's decision.~~ Resolved: Vlad cleared all assets, 2026-09-24 (see §3).
