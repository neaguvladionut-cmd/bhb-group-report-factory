import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import vm from "node:vm";
import test from "node:test";
import { buildPayload } from "../src/rebuild-core.js";
import { reportPlan } from "../src/rebuild-report-plan.js";
import { createMetaFixture } from "./fixtures/grf-ux-meta-fixture.mjs";

const root = resolve(new URL("..", import.meta.url).pathname);
const sourceHtml = await readFile(resolve(root, "src/rebuild-index.html"), "utf8");
const bundle = await readFile(resolve(root, "deploy/app.js"), "utf8");
const vendor = await readFile(resolve(root, "src/assets/vendor/xlsx.full.min.js"), "utf8");

class MiniElement {
  constructor(tagName, ownerDocument) {
    this.tagName = tagName.toUpperCase(); this.ownerDocument = ownerDocument; this.parentElement = null;
    this.children = []; this.attributes = new Map(); this.listeners = new Map(); this.dataset = {};
    this.hidden = false; this.disabled = false; this.checked = false; this.value = ""; this._text = "";
  }
  set className(value) { this.setAttribute("class", value); }
  get className() { return this.getAttribute("class") || ""; }
  get classList() {
    return {
      add: (...names) => this.className = [...new Set([...this.className.split(/\s+/u).filter(Boolean), ...names])].join(" "),
      remove: (...names) => this.className = this.className.split(/\s+/u).filter((name) => name && !names.includes(name)).join(" "),
      toggle: (name, force) => { const present = this.className.split(/\s+/u).includes(name); const next = force === undefined ? !present : Boolean(force); if (next && !present) this.className = `${this.className} ${name}`.trim(); if (!next && present) this.className = this.className.split(/\s+/u).filter((item) => item && item !== name).join(" "); return next; },
      contains: (name) => this.className.split(/\s+/u).includes(name)
    };
  }
  set innerHTML(value) { this.replaceChildren(); parseMarkup(this, value, this.ownerDocument); }
  get innerHTML() { return this.children.map((child) => child.outerHTML()).join("") || this._text; }
  set textContent(value) { this.replaceChildren(); this._text = String(value ?? ""); }
  get textContent() { return this._text + this.children.map((child) => child.textContent).join(""); }
  setAttribute(name, value) {
    const stringValue = String(value); this.attributes.set(name, stringValue);
    if (name === "hidden") this.hidden = true;
    if (name === "disabled") this.disabled = true;
    if (name === "checked") this.checked = true;
    if (name === "value") this.value = stringValue;
    if (name.startsWith("data-")) this.dataset[name.slice(5).replace(/-([a-z])/gu, (_, letter) => letter.toUpperCase())] = stringValue;
  }
  removeAttribute(name) { this.attributes.delete(name); if (name === "hidden") this.hidden = false; if (name === "disabled") this.disabled = false; if (name === "checked") this.checked = false; }
  getAttribute(name) { return this.attributes.get(name) ?? null; }
  append(...nodes) { for (const node of nodes) { if (!node) continue; node.parentElement = this; this.children.push(node); } }
  replaceChildren(...nodes) { this.children.forEach((node) => { node.parentElement = null; }); this.children = []; this._text = ""; this.append(...nodes); }
  remove() { if (!this.parentElement) return; this.parentElement.children = this.parentElement.children.filter((child) => child !== this); this.parentElement = null; }
  addEventListener(type, listener) { if (!this.listeners.has(type)) this.listeners.set(type, []); this.listeners.get(type).push(listener); }
  dispatchEvent(event) { for (const listener of this.listeners.get(event.type) || []) listener.call(this, { ...event, target: this }); return true; }
  click() { this.dispatchEvent({ type: "click" }); }
  focus() { if (this.ownerDocument) this.ownerDocument.activeElement = this; }
  scrollIntoView() {}
  closest(selector) { let node = this; while (node) { if (matchesSimple(node, selector)) return node; node = node.parentElement; } return null; }
  querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
  querySelectorAll(selector) {
    const parts = selector.trim().split(/\s+/u); const found = [];
    const visit = (node) => { for (const child of node.children) { if (matchesSelector(child, parts)) found.push(child); visit(child); } };
    visit(this); return found;
  }
  outerHTML() { return `<${this.tagName.toLowerCase()}>${this.textContent}</${this.tagName.toLowerCase()}>`; }
  insertAdjacentHTML(position, markup) {
    if (position !== "beforebegin" || !this.parentElement) return;
    const fragment = new MiniElement("fragment", this.ownerDocument); parseMarkup(fragment, markup, this.ownerDocument);
    const index = this.parentElement.children.indexOf(this); this.parentElement.children.splice(index, 0, ...fragment.children); fragment.children.forEach((child) => { child.parentElement = this.parentElement; });
  }
}

class MiniDocument {
  constructor(markup) { this.body = new MiniElement("body", this); this.activeElement = this.body; this.body.innerHTML = markup.match(/<body[^>]*>([\s\S]*?)<\/body>/iu)?.[1] || markup; }
  createElement(tagName) { return new MiniElement(tagName, this); }
  querySelector(selector) { return this.body.querySelector(selector); }
  querySelectorAll(selector) { return this.body.querySelectorAll(selector); }
}

function matchesSimple(node, selector) {
  let source = selector; const checked = source.includes(":checked"); source = source.replace(":checked", "");
  const tag = source.match(/^([a-z][a-z0-9-]*)/iu)?.[1]; if (tag && node.tagName.toLowerCase() !== tag.toLowerCase()) return false;
  const id = source.match(/#([a-z0-9_-]+)/iu)?.[1]; if (id && node.getAttribute("id") !== id) return false;
  for (const className of source.matchAll(/\.([a-z0-9_-]+)/giu)) if (!node.className.split(/\s+/u).includes(className[1])) return false;
  for (const attribute of source.matchAll(/\[([a-z0-9:-]+)(?:=["']?([^\]"']+)["']?)?\]/giu)) {
    const actual = node.getAttribute(attribute[1]); if (actual === null || (attribute[2] !== undefined && actual !== attribute[2])) return false;
  }
  return !checked || node.checked === true;
}

function matchesSelector(node, parts) {
  if (!matchesSimple(node, parts.at(-1))) return false;
  let current = node.parentElement;
  for (let index = parts.length - 2; index >= 0; index -= 1) {
    while (current && !matchesSimple(current, parts[index])) current = current.parentElement;
    if (!current) return false; current = current.parentElement;
  }
  return true;
}

function parseMarkup(rootNode, markup, documentRef) {
  const stack = [rootNode]; const voidTags = new Set(["input", "img", "br", "hr", "meta", "link"]);
  for (const token of String(markup).match(/<[^>]+>|[^<]+/gu) || []) {
    if (token.startsWith("</")) { if (stack.length > 1) stack.pop(); continue; }
    if (!token.startsWith("<") || token.startsWith("<!--") || token.startsWith("<!DOCTYPE")) { stack.at(-1)._text += token; continue; }
    const opening = token.match(/^<([a-z0-9-]+)([^>]*)>/iu); if (!opening) continue;
    const element = new MiniElement(opening[1], documentRef);
    for (const attribute of opening[2].matchAll(/([a-z0-9:-]+)(?:=(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/giu)) element.setAttribute(attribute[1], attribute[2] ?? attribute[3] ?? attribute[4] ?? "");
    stack.at(-1).append(element); if (!voidTags.has(opening[1].toLowerCase()) && !token.endsWith("/>") && opening[1].toLowerCase() !== "script") stack.push(element);
  }
}

function loadXlsx() { const sandbox = { exports: {}, module: { exports: {} }, Buffer, process }; vm.runInNewContext(vendor, sandbox); return sandbox.exports; }
const XLSX = loadXlsx();
const workbook = (rows) => { const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1"); return XLSX.write(book, { type: "buffer", bookType: "xlsx" }); };

async function bootApp() {
  const document = new MiniDocument(sourceHtml);
  const window = { XLSX, JSZip: {}, __GRF_TEST__: true, document, __grfBooted() {} };
  vm.runInNewContext(bundle, { window, document, Blob, Date, Intl, Map, Set, URL, ArrayBuffer, Uint8Array, TextEncoder, TextDecoder, console });
  return { document, window };
}

const metaFixture = createMetaFixture(XLSX, "d");
const validFiles = [{ name: "summary.xlsx", bytes: metaFixture.summary }, { name: "detail.xlsx", bytes: metaFixture.detailed }];
const warningSummary = workbook([["CODE", "name", "cod cp", "Leadership"], ["North", "Synthetic Ana", "A-1", 4], ["", "Synthetic Bogdan", "A-2", 3]]);
const warningDetail = workbook([["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", "Leadership"], ["", "", "", "", "Subcompetente", "L"], ["", "", "", "", "behavior", "B1"], ["North", "Synthetic Ana", "Nord", "A-1", "", 2], ["", "Synthetic Bogdan", "Sud", "A-2", "", ""]]);
const warningFiles = [{ name: "summary.xlsx", bytes: warningSummary }, { name: "detail.xlsx", bytes: warningDetail }];
const singularSummary = workbook([["CODE", "name", "cod cp", "Leadership"], ["North", "Synthetic Ana", "A-1", 4], ["South", "Synthetic Bogdan", "A-2", 3]]);
const singularDetail = workbook([["CODE", "name the person evaluated", "regiune", "cod ac", "Competente", "Leadership"], ["", "", "", "", "Subcompetente", "L"], ["", "", "", "", "behavior", "B1"], ["North", "Synthetic Ana", "Nord", "A-1", "", 2], ["South", "Synthetic Bogdan", "Sud", "A-2", "", ""]]);
const singularFiles = [{ name: "summary.xlsx", bytes: singularSummary }, { name: "detail.xlsx", bytes: singularDetail }];
const invalidSummary = workbook([["CODE", "name", "cod cp", "Leadership"], ["North", "Synthetic Ana", "A-1", 9]]);
const blockerFiles = [{ name: "summary.xlsx", bytes: invalidSummary }, { name: "detail.xlsx", bytes: metaFixture.detailed }];
const makePayload = (files, metadata = {}, acknowledgedWarningIds = []) => buildPayload(XLSX, files, { projectName: "Synthetic GRF-UX", ...metadata }, {}, { acknowledgedWarningIds });

test("R2 footer renders blocker, singular/plural warning-group, and all-confirmed states", async () => {
  const { document, window } = await bootApp();
  const blocker = makePayload(blockerFiles);
  window.__grfTest.renderReview(blocker);
  assert.equal(document.querySelector("#review-footer-note").textContent, "Rezolvă blocajele înainte de a merge mai departe.");

  const pending = makePayload(warningFiles);
  window.__grfTest.renderReview(pending);
  assert.equal(document.querySelector("#review-footer-note").textContent, "Mai sunt 2 grupuri de confirmat.");
  const singular = makePayload(singularFiles);
  window.__grfTest.renderReview(singular);
  assert.equal(document.querySelector("#review-footer-note").textContent, "Mai este 1 grup de confirmat.");

  const allDone = makePayload(singularFiles, {}, singular.warnings.map((item) => item.id));
  window.__grfTest.renderReview(allDone);
  assert.equal(document.querySelector("#review-footer-note").textContent, "Avertismentele sunt confirmate; poți alege structura.");
});

test("N7 rendered outline assigns every generated slide to exactly one block for end, separate, and none", async () => {
  const { document, window } = await bootApp();
  for (const annex of ["end", "separate", "none"]) {
    const initial = makePayload(validFiles, { annex });
    const payload = makePayload(validFiles, { annex }, initial.warnings.map((item) => item.id));
    window.__grfTest.renderStructure(payload);
    const headerTotal = Number.parseInt(document.querySelector("#slide-total").textContent, 10);
    const blockTotal = document.querySelectorAll("#structure-summary .outline-count").reduce((sum, node) => sum + Number(node.textContent), 0);
    const generatedTotal = annex === "separate" ? reportPlan(payload, { scope: "main" }).length + reportPlan(payload, { scope: "appendix" }).length : reportPlan(payload, { scope: "whole" }).length;
    assert.equal(blockTotal, headerTotal, `${annex}: outline blocks must add to the header total`);
    assert.equal(headerTotal, generatedTotal, `${annex}: header total must match generated slides`);
  }
});

test("N3 rendered derived-field chip changes after edit and returns after Sesiune nouă", async () => {
  const { document, window } = await bootApp();
  const payload = makePayload(validFiles);
  window.__grfTest.renderFound(payload, validFiles);
  const chip = () => document.querySelector("#proposal-evaluators").parentElement.querySelector(".source-chip").textContent;
  assert.equal(chip(), "din export · de verificat");
  const input = document.querySelector("#proposal-evaluators"); input.value = "Consultant corectat"; input.dispatchEvent({ type: "input" });
  assert.equal(chip(), "corectat de tine");
  document.querySelector("#reset").click(); document.querySelector("#reset-confirm-yes").click();
  window.__grfTest.renderFound(makePayload(validFiles), validFiles);
  assert.equal(chip(), "din export · de verificat");
});
