// Run: node tests/rules.test.js
const assert = require("assert");
const R = require("../src/rules.js");

let passed = 0;
function t(name, fn) { fn(); passed++; }
const rules = (list) => list.map((i) => i.rule);

t("em dash flagged and fixable", () => {
  const i = R.checkText("Revenue grew — margins fell", {});
  assert.deepStrictEqual(rules(i), ["em-dash"]);
  assert.strictEqual(i[0].fix, "dash");
});
t("table em dash uses table fix", () => {
  assert.strictEqual(R.checkText("a—b", { inTable: true })[0].fix, "table-dash");
});
t("spaced en dash flagged, range en dash not", () => {
  assert.deepStrictEqual(rules(R.checkText("Cost – high", {})), ["en-dash"]);
  assert.deepStrictEqual(rules(R.checkText("2020–2024", {})), []);
});
t("placeholders flagged, NEEDED is info", () => {
  const i = R.checkText("[Insert challenge] and [NEEDED: client revenue, FY2025]", {});
  assert.deepStrictEqual(rules(i), ["placeholder", "needed"]);
  assert.strictEqual(i[1].severity, "info");
  assert.strictEqual(i[1].needed, "[NEEDED: client revenue, FY2025]");
});
t("example text xx flagged", () => {
  assert.deepStrictEqual(rules(R.checkText("Source: xx", {})), ["example-text"]);
  assert.deepStrictEqual(rules(R.checkText("Rate: xxx/day", {})), ["example-text"]);
  assert.deepStrictEqual(rules(R.checkText("Box exceeds plan", {})), []);
});
t("typo flagged", () => {
  assert.deepStrictEqual(rules(R.checkText("Delive-rables", {})), ["typo"]);
});
t("fonts", () => {
  assert.deepStrictEqual(rules(R.checkFont({ name: "Lora", size: 10 })), []);
  assert.deepStrictEqual(rules(R.checkFont({ name: "+mn-lt", size: 10 })), []);
  assert.deepStrictEqual(rules(R.checkFont({ name: "Arial", size: 10 })), ["font"]);
  assert.deepStrictEqual(rules(R.checkFont({ name: null, size: 10 })), ["font-mixed"]);
  assert.deepStrictEqual(rules(R.checkFont({ name: "Lora", size: 7 })), ["small-text"]);
  assert.deepStrictEqual(rules(R.checkFont({ name: undefined, size: undefined })), []);
});
t("title rules", () => {
  assert.deepStrictEqual(rules(R.checkTitle("Revenue grew 12% while costs held flat", {})), []);
  assert.deepStrictEqual(rules(R.checkTitle("REVENUE GREW STRONGLY", {})), ["title-caps"]);
  assert.strictEqual(R.checkTitle("Revenue grew", { allCaps: true })[0].fix, "allcaps");
  assert.deepStrictEqual(rules(R.checkTitle("Revenue Grew Faster Than Costs In Every Region", {})), ["title-case"]);
  assert.deepStrictEqual(rules(R.checkTitle("GNA and C.A. Hull agree on three Claude pilots", {})), []);
  assert.deepStrictEqual(rules(R.checkTitle("Costs fell.", {})), ["title-period"]);
  assert.deepStrictEqual(rules(R.checkTitle("", {})), ["title-empty"]);
});
t("source detection", () => {
  assert.strictEqual(R.hasSource(["Body", "Source: Acme data"]), true);
  assert.strictEqual(R.hasSource(["Source: xx"]), false);
  assert.strictEqual(R.hasSource(["Source: "]), false);
  assert.strictEqual(R.hasSource(["Footnote: 1. abc\nSource: GNA analysis"]), true);
  assert.strictEqual(R.hasSource(["No source here"]), false);
});
t("dash edits reproduce fixDashes", () => {
  const samples = ["a — b", "a—b — c", "x – y and 2020–2024", "end —", "— start"];
  for (const s of samples) {
    let out = s;
    for (const e of R.dashEdits(s)) out = out.slice(0, e.start) + e.text + out.slice(e.start + e.length);
    assert.strictEqual(out, R.fixDashes(s).replace(/, ,/g, ","), s);
    assert.ok(!/—/.test(out));
  }
});
t("typo and period edits", () => {
  const apply = (s, edits) => edits.reduce((o, e) => o.slice(0, e.start) + e.text + o.slice(e.start + e.length), s);
  assert.strictEqual(apply("Two Delive-rables and Deliver-ables", R.typoEdits("Two Delive-rables and Deliver-ables")), "Two Deliverables and Deliverables");
  assert.strictEqual(apply("Costs fell.", R.titlePeriodEdits("Costs fell.")), "Costs fell");
});

console.log(`rules: ${passed} tests passed`);
