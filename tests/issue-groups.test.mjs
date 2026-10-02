import assert from "node:assert/strict";
import test from "node:test";
import { renderIssueGroup } from "../src/rebuild-issues.js";

class MiniElement {
  constructor(tagName) {
    this.tagName = tagName.toUpperCase(); this.children = []; this.attributes = new Map(); this.listeners = new Map();
    this.dataset = {}; this.hidden = false; this.value = ""; this._text = "";
  }
  set className(value) { this.setAttribute("class", value); }
  get className() { return this.getAttribute("class") || ""; }
  set innerHTML(value) { this.replaceChildren(); parseMarkup(this, value); }
  get innerHTML() { return this.textContent; }
  set textContent(value) { this.replaceChildren(); this._text = String(value ?? ""); }
  get textContent() { return this._text + this.children.map((child) => child.textContent).join(""); }
  setAttribute(name, value) {
    const stringValue = String(value); this.attributes.set(name, stringValue); if (name === "hidden") this.hidden = true;
    if (name.startsWith("data-")) this.dataset[name.slice(5).replace(/-([a-z])/gu, (_, letter) => letter.toUpperCase())] = stringValue;
  }
  getAttribute(name) { return this.attributes.get(name) ?? null; }
  append(...nodes) { this.children.push(...nodes); }
  replaceChildren(...nodes) { this.children = nodes; this._text = ""; }
  addEventListener(type, listener) { if (!this.listeners.has(type)) this.listeners.set(type, []); this.listeners.get(type).push(listener); }
  dispatchEvent(event) { for (const listener of this.listeners.get(event.type) || []) listener.call(this, { ...event, target: this }); }
  click() { this.dispatchEvent({ type: "click" }); }
  querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
  querySelectorAll(selector) {
    const matches = (node) => selector.startsWith(".")
      ? node.className.split(/\s+/u).includes(selector.slice(1))
      : selector.startsWith("[") && selector.endsWith("]") ? node.attributes.has(selector.slice(1, -1).split("=")[0]) : false;
    const found = []; const visit = (node) => { for (const child of node.children) { if (matches(child)) found.push(child); visit(child); } };
    visit(this); return found;
  }
}

class MiniDocument { createElement(tagName) { return new MiniElement(tagName); } }

function parseMarkup(root, markup) {
  const stack = [root]; const voidTags = new Set(["input"]);
  for (const token of markup.match(/<[^>]+>|[^<]+/gu) || []) {
    if (token.startsWith("</")) { stack.pop(); continue; }
    if (!token.startsWith("<") || token.startsWith("<!--")) { stack.at(-1)._text += token; continue; }
    const opening = token.match(/^<([a-z0-9-]+)([^>]*)>/iu); if (!opening) continue;
    const element = new MiniElement(opening[1]);
    for (const attribute of opening[2].matchAll(/([a-z0-9-]+)(?:="([^"]*)")?/giu)) element.setAttribute(attribute[1], attribute[2] ?? "");
    stack.at(-1).append(element); if (!voidTags.has(opening[1].toLowerCase()) && !token.endsWith("/>") ) stack.push(element);
  }
}

function bigFixture() {
  const people = Array.from({ length: 200 }, (_, index) => ({
    name: `Participant test ${String(index + 1).padStart(3, "0")}`,
    region: index % 5 === 0 ? "" : ["București", "Cluj-Napoca", "Iași"][index % 3],
    code: index % 3 === 0 ? "" : index % 2 ? "NORD" : "SUD"
  }));
  const invalid = people.slice(0, 160).map((person, index) => ({
    id: `summary-score-${index + 1}`,
    code: "summary-score",
    sourceName: "sinteza-200.xlsx",
    rowNumber: index + 2,
    message: `${person.name} · scor 6.5`
  }));
  return {
    people,
    payload: { schemas: [{ sourceName: "sinteza-200.xlsx", rows: Array.from({ length: 201 }, () => []), headerRow: 1 }] },
    group: { key: "summary-score", severity: "blocker", items: invalid }
  };
}

test("200-participant issue card is collapsed and caps expanded rows at eight", () => {
  const fixture = bigFixture();
  const card = renderIssueGroup(fixture.group, { documentRef: new MiniDocument(), payload: fixture.payload });

  assert.equal(fixture.people.length, 200);
  assert.equal(card.dataset.issueExpanded, "false");
  assert.equal(card.querySelectorAll(".issue-item").length, 0);
  assert.equal(card.querySelector("[data-issue-count]").textContent, "160");

  card.querySelector("[data-issue-toggle]").click();
  assert.equal(card.dataset.issueExpanded, "true");
  assert.equal(card.querySelectorAll(".issue-item").length, 8);
  assert.equal(card.querySelector("[data-issue-reveal]").textContent, "Arată încă 25 (din 160)");
});

test("issue search narrows the visible rows and then reveals the rest in pages", () => {
  const fixture = bigFixture();
  const card = renderIssueGroup(fixture.group, { documentRef: new MiniDocument(), payload: fixture.payload });
  card.querySelector("[data-issue-toggle]").click();

  const search = card.querySelector("[data-issue-search]");
  search.value = "Participant test 012";
  search.dispatchEvent({ type: "input" });
  assert.equal(card.querySelectorAll(".issue-item").length, 1);
  assert.equal(card.querySelector("[data-issue-count]").textContent, "1 din 160");

  search.value = "";
  search.dispatchEvent({ type: "input" });
  const reveal = card.querySelector("[data-issue-reveal]");
  reveal.click();
  assert.equal(card.querySelectorAll(".issue-item").length, 25);
  assert.equal(reveal.textContent, "Arată încă 25 (din 160)");

  reveal.click();
  assert.equal(card.querySelectorAll(".issue-item").length, 25);
  assert.equal(reveal.textContent, "Arată încă 25 (din 160)");

  for (let page = 0; page < 6; page += 1) reveal.click();
  assert.equal(card.querySelectorAll(".issue-item").length, 8);
  assert.equal(reveal.textContent, "Arată încă 25 (din 160)");
});

test("the served app bundle defines every issue helper it calls", async () => {
  const { readFile } = await import("node:fs/promises");
  const bundle = await readFile(new URL("../deploy/app.js", import.meta.url), "utf8");
  for (const name of ["groupedIssues", "renderIssueGroup"]) {
    if (/\b/.test(name) && bundle.includes(`${name}(`)) assert.match(bundle, new RegExp(`function ${name}\\b|${name}[,}]`), `${name} is called but never defined in deploy/app.js`);
  }
});

test("the served bundle embeds the Poppins font so the BHB bundle works from a local file", async () => {
  const { readFile } = await import("node:fs/promises");
  const bundle = await readFile(new URL("../deploy/app.js", import.meta.url), "utf8");
  assert.match(bundle, /window\.__GRF_FONT_BASE64__="[A-Za-z0-9+/=]{1000,}"/u);
});
