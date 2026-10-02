import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import { resolve } from "node:path";
import vm from "node:vm";
import test from "node:test";

const root = resolve(new URL("..", import.meta.url).pathname);

test("deploy is the source-owned one-page Romanian builder with the approved four-step spine", async () => {
  const source = await readFile(resolve(root, "src/rebuild-index.html"), "utf8");
  const deploy = await readFile(resolve(root, "deploy/index.html"), "utf8");
  const app = await readFile(resolve(root, "deploy/app.js"), "utf8");
  const css = await readFile(resolve(root, "deploy/styles.css"), "utf8");
  assert.equal(deploy, await readFile(resolve(root, "src/rebuild-index.html"), "utf8"));
  assert.match(deploy, /id="split-groups"/u);
  assert.match(deploy, /name="annex"/u);
  assert.match(deploy, /șablon-declinații\.csv/u);
  assert.doesNotMatch(deploy, /Context copertă|evaluation-sheet-template\.csv|Împărțire pe grupuri CODE/u);
  assert.match(deploy, /id="bundle"/u);
  assert.match(deploy, /id="download-fallback"/u);
  assert.equal((deploy.match(/class="workflow-section/g) || []).length, 4);
  assert.equal((deploy.match(/class="section-guide"/g) || []).length, 4);
  assert.equal((deploy.match(/data-back-step/g) || []).length, 3);
  assert.doesNotMatch(deploy, /Context copertă/u);
  assert.ok((deploy.match(/class="field-context"/gu) || []).length >= 6);
  assert.match(deploy, /Ce faci aici/u);
  assert.match(deploy, /Ce urmează/u);
  assert.match(deploy, /Apare în/u);
  assert.match(deploy, /Pachet BHB \(ZIP\)/u);
  const visibleCopy = deploy.replace(/<[^>]+>/gu, " ");
  assert.doesNotMatch(visibleCopy, /\b(?:schema|upload|API)\b/iu);
  assert.doesNotMatch(visibleCopy, /\bCODE\b|evaluation-sheet-template/iu);
  assert.match(app, /downloadBundle/u);
  assert.match(app, /downloadPptx/u);
  assert.match(app, /groupedIssues\(blockers, "blocker"\)/u);
  assert.match(app, /data-issue-search/u);
  assert.match(app, /Confirmă toate avertismentele/u);
  assert.match(app, /data-back-step/u);
  assert.match(app, /section\.hidden = Number\(section\.dataset\.section\) !== state\.step/u);
  assert.match(app, /link\.classList\.toggle\("done", number < state\.step\)/u);
  assert.match(app, /aria-current/u);
  assert.match(app, /state\.step = 1/u);
  assert.match(app, /trigger\?\.focus/u);
  assert.match(app, /sablon-declinatii-\$\{name\}\.csv/u);
  assert.match(app, /projectName: \$\("#project-name"\)/u);
  assert.match(app, /names\.whole/u);
  assert.doesNotMatch(app, /PptxGenJS/u);
  assert.match(css, /\[hidden\]\{display:none!important\}/u);
  assert.match(css, /\.issue-panel\{/u);
  assert.match(css, /\.issue-toggle/u);
  assert.doesNotMatch(css, /\.issue-items\{[^}]*overflow/u);
  assert.match(css, /\.workflow-rail\{/u);
  assert.match(source, /id="step-upload"/u);
  assert.match(source, /id="step-download"/u);
  const scripts = [...deploy.matchAll(/<script(?:(?:\s+src="([^"]+)")?)>([\s\S]*?)<\/script>/gu)];
  assert.equal(scripts.length, 5);
  for (const [, src] of scripts) if (src) await access(resolve(root, "deploy", src.split("?")[0]));
});

test("drawer methodology fields share the page-one proposal keys", async () => {
  const source = await readFile(resolve(root, "src/rebuild-index.html"), "utf8");
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  for (const [id, key] of [["evaluators", "evaluators"], ["days", "days"], ["team-size", "teamSize"], ["evaluation-period", "evaluationPeriod"], ["population-role", "populationByRole"], ["location", "location"]]) {
    assert.match(source, new RegExp(`id="${id}" data-proposal-key="${key}"`, "u"));
    assert.match(app, new RegExp(`\["${id}", "${key}"\]`, "u"));
  }
  assert.match(app, /setProposal\(input\); recompute\(\)/u);
});

test("adding a source preserves edited proposals and corrections for the next render", async () => {
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  assert.match(app, /methodologyProposalEdits: new Set\(\)/u);
  assert.match(app, /state\.methodologyProposalEdits\.add\(key\)/u);
  assert.doesNotMatch(app, /files = mergeSelectedFiles\(files, incoming\); state\.acknowledged\.clear\(\); state\.corrections = \{ values: \{\} \}; state\.methodologyProposals = \{\};/u);
  assert.match(app, /if \(!state\.methodologyProposalEdits\.has\(key\)\) state\.methodologyProposals\[key\] = String\(value\)/u);
});

test("page-two review footer reports blockers, warning groups and confirmed state", async () => {
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  assert.match(app, /Rezolvă blocajele înainte de a merge mai departe\./u);
  assert.match(app, /Mai sunt \$\{pendingGroups\.length\} grupuri de confirmat\./u);
  assert.match(app, /Totul este confirmat; poți continua\./u);
});

test("separate-annex outline is built from the two delivered plans", async () => {
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  assert.match(app, /\(annex === "separate" \? \[mainPlan, appendixPlan\] : \[plan\]\)\.flatMap\(groupOutlineSlides\)/u);
  assert.match(app, /const deliveredTotal = annex === "separate" \? mainPlan\.length \+ appendixPlan\.length/u);
});

test("empty findings tiles are omitted and count tiles have no export-proposal caption", async () => {
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  assert.match(app, /\]\.filter\(\(\[value\]\) => value !== "" && value !== null && value !== undefined\)/u);
  assert.match(app, /stats\.map\(\(\[value, label\]\) => `<div class="stat">/u);
  assert.doesNotMatch(app, /stat-caption|Propunere din export\./u);
});

test("confirming the final warning focuses the first blocker when blockers remain", async () => {
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  const issues = await readFile(resolve(root, "src/rebuild-issues.js"), "utf8");
  assert.match(app, /else if \(blockers\.length\) \{ const firstBlocker = document\.querySelector\("\.issue-group\.blocker"\); firstBlocker\?\.focus\(\{ preventScroll: true \}\); \}/u);
  assert.match(issues, /article\.tabIndex = -1/u);
});

test("manual project edits are labelled, empty page-one fields are actionable, and downloads use the native path", async () => {
  const app = await readFile(resolve(root, "src/rebuild-app.js"), "utf8");
  const template = await readFile(resolve(root, "src/template-pptx.js"), "utf8");
  const styles = await readFile(resolve(root, "src/styles.css"), "utf8");
  assert.match(app, /text\(values\.projectName\) !== text\(state\.projectName\) \? "completat de tine"/u);
  assert.match(app, /const placeholder = text\(value\) \? examples\[id\] : `De completat · Exemplu: \$\{examples\[id\]\}`/u);
  assert.doesNotMatch(app, /window\.__grfDownload/u);
  assert.doesNotMatch(template, /window\.__grfDownload/u);
  assert.doesNotMatch(styles, /stat-caption/u);
});

test("source-owned build emits the new normalized payload and package-preserving route", async () => {
  const core = await readFile(resolve(root, "src/rebuild-core.js"), "utf8");
  const plan = await readFile(resolve(root, "src/rebuild-report-plan.js"), "utf8");
  const template = await readFile(resolve(root, "src/template-pptx.js"), "utf8");
  assert.match(core, /CODE/u);
  assert.match(core, /region-missing/u);
  assert.match(core, /objective_text_score_0/u);
  assert.match(plan, /rankBehaviors/u);
  assert.match(plan, /splitGroups/u);
  assert.match(template, /ppt\/embeddings/u);
  assert.match(template, /generateBundle/u);
  assert.doesNotMatch(template, /PptxGenJS/u);
});

test("built deploy bundle initializes without throwing in a DOM-like vm", async () => {
  const bundle = await readFile(resolve(root, "deploy/app.js"), "utf8");
  const document = { body: { dataset: {} }, querySelector: () => null, querySelectorAll: () => [] };
  const window = { XLSX: {}, JSZip: {}, __grfBooted() {} };
  assert.doesNotThrow(() => vm.runInNewContext(bundle, { Blob, Intl, Map, Set, URL, document, window }));
  assert.equal(window.__grfDownload, undefined);
});
