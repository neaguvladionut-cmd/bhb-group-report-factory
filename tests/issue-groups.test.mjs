import assert from "node:assert/strict";
import { JSDOM } from "jsdom";
import test from "node:test";
import { renderIssueGroup } from "../src/rebuild-issues.js";

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
  const dom = new JSDOM("<!doctype html><body></body>");
  const card = renderIssueGroup(fixture.group, { documentRef: dom.window.document, payload: fixture.payload });
  dom.window.document.body.append(card);

  assert.equal(fixture.people.length, 200);
  assert.equal(card.dataset.issueExpanded, "false");
  assert.equal(card.querySelectorAll(".issue-item").length, 0);
  assert.equal(card.querySelector("[data-issue-count]").textContent, "160");

  card.querySelector("[data-issue-toggle]").click();
  assert.equal(card.dataset.issueExpanded, "true");
  assert.equal(card.querySelectorAll(".issue-item").length, 8);
  assert.equal(card.querySelector("[data-issue-reveal]").textContent, "Arată toate (160)");
});

test("issue search narrows the visible rows and count, then Arată toate reveals the rest", () => {
  const fixture = bigFixture();
  const dom = new JSDOM("<!doctype html><body></body>");
  const card = renderIssueGroup(fixture.group, { documentRef: dom.window.document, payload: fixture.payload });
  dom.window.document.body.append(card);
  card.querySelector("[data-issue-toggle]").click();

  const search = card.querySelector("[data-issue-search]");
  search.value = "Participant test 012";
  search.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
  assert.equal(card.querySelectorAll(".issue-item").length, 1);
  assert.equal(card.querySelector("[data-issue-count]").textContent, "1 din 160");

  search.value = "";
  search.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
  const reveal = card.querySelector("[data-issue-reveal]");
  reveal.click();
  assert.equal(card.querySelectorAll(".issue-item").length, 160);
  assert.equal(reveal.textContent, "Arată mai puține");

  reveal.click();
  assert.equal(card.querySelectorAll(".issue-item").length, 8);
  assert.equal(reveal.textContent, "Arată toate (160)");
});
