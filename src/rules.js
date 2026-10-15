/*
 * GNA house-style rules, as pure functions (no Office.js), so they can be unit
 * tested in Node and reused by the style checker in the task pane.
 *
 * Each rule returns issues shaped like:
 *   { rule, severity: "error" | "warning" | "info", message, excerpt, fix }
 * where `fix` is a fix kind the checker knows how to apply, or null.
 */
(function (root) {
  "use strict";

  var EM_DASH = /—/;
  var SPACED_EN_DASH = / – /;
  var PLACEHOLDER = /\[(?!NEEDED:)[^\[\]\r\n]{1,160}\]/g;
  var NEEDED = /\[NEEDED:[^\]]*\]/g;
  var EXAMPLE = /\bx{2,}\b|lorem ipsum|click to edit/i;
  var TYPOS = /Delive-rables|Deliver-ables/;
  var SOURCE_LINE = /^\s*Sources?\s*:/im;

  // Small words that stay lowercase in Title Case, used by the sentence-case heuristic
  var MINOR = { a: 1, an: 1, and: 1, as: 1, at: 1, but: 1, by: 1, for: 1, from: 1, in: 1, into: 1, of: 1, on: 1, or: 1, the: 1, to: 1, vs: 1, via: 1, with: 1 };

  function excerptAround(text, index, len) {
    var start = Math.max(0, index - 24);
    var end = Math.min(text.length, index + (len || 1) + 24);
    var s = text.slice(start, end).replace(/[\r\n\v\u000b]+/g, " ");
    return (start > 0 ? "…" : "") + s + (end < text.length ? "…" : "");
  }

  function isLora(name) {
    if (!name) return false;
    var n = String(name).toLowerCase();
    return n === "lora" || n.indexOf("lora ") === 0 || n === "+mn-lt" || n === "+mj-lt";
  }

  /** Rules that apply to any piece of text (shape or table cell). */
  function checkText(text, ctx) {
    ctx = ctx || {};
    var issues = [];
    if (!text || !text.trim()) return issues;
    var m;

    if ((m = EM_DASH.exec(text))) {
      issues.push({ rule: "em-dash", severity: "error", message: "Em dash. House style uses a comma, colon or full stop.",
        excerpt: excerptAround(text, m.index, 1), fix: ctx.inTable ? "table-dash" : "dash" });
    }
    if ((m = SPACED_EN_DASH.exec(text))) {
      issues.push({ rule: "en-dash", severity: "warning", message: "En dash used as a dash. Use a comma, colon or full stop.",
        excerpt: excerptAround(text, m.index, 3), fix: ctx.inTable ? "table-dash" : "dash" });
    }
    PLACEHOLDER.lastIndex = 0;
    var ph = [];
    while ((m = PLACEHOLDER.exec(text)) && ph.length < 50) ph.push(m[0]);
    if (ph.length) {
      issues.push({ rule: "placeholder", severity: "error",
        message: ph.length === 1 ? "Template placeholder not filled in." : ph.length + " template placeholders not filled in.",
        excerpt: ph.slice(0, 3).join("  "), fix: null });
    }
    if ((m = EXAMPLE.exec(text))) {
      issues.push({ rule: "example-text", severity: "error", message: "Template example text left in (e.g. \"xx\").",
        excerpt: excerptAround(text, m.index, m[0].length), fix: null });
    }
    if ((m = TYPOS.exec(text))) {
      issues.push({ rule: "typo", severity: "error", message: "Template typo: should be \"Deliverables\".",
        excerpt: m[0], fix: ctx.inTable ? null : "typo" });
    }
    NEEDED.lastIndex = 0;
    while ((m = NEEDED.exec(text))) {
      issues.push({ rule: "needed", severity: "info", message: "Open [NEEDED] flag.", excerpt: m[0], fix: null, needed: m[0] });
    }
    return issues;
  }

  /** Font rules for a shape's whole text range. `font` = {name, size, allCaps}. */
  function checkFont(font, ctx) {
    ctx = ctx || {};
    var issues = [];
    if (!font) return issues;
    if (font.name === null) {
      issues.push({ rule: "font-mixed", severity: "warning", message: "Mixed fonts in one text box. GNA uses Lora only.", excerpt: "", fix: "font" });
    } else if (font.name !== undefined && !isLora(font.name)) {
      issues.push({ rule: "font", severity: "error", message: "Font is " + font.name + ". GNA uses Lora only.", excerpt: "", fix: "font" });
    }
    if (typeof font.size === "number" && font.size > 0 && font.size < 8) {
      issues.push({ rule: "small-text", severity: "warning", message: "Text is " + font.size + " pt. Template minimum is 8 pt (sources) and 9 pt (labels).", excerpt: "", fix: null });
    }
    return issues;
  }

  function looksTitleCase(text) {
    var words = text.replace(/[’'][a-z]+/g, "").split(/\s+/).filter(function (w) { return /^[A-Za-z]/.test(w); });
    if (words.length < 4) return false;
    var rest = words.slice(1);
    var counted = 0, capped = 0;
    for (var i = 0; i < rest.length; i++) {
      var w = rest[i].replace(/[^A-Za-z]/g, "");
      if (!w || MINOR[w.toLowerCase()]) continue;
      if (/^[A-Z]{2,}$/.test(w)) continue; // acronyms (GNA, KPI)
      counted++;
      if (/^[A-Z]/.test(w)) capped++;
    }
    return counted >= 3 && capped / counted >= 0.75;
  }

  /** Title rules. `font` may include allCaps. */
  function checkTitle(text, font) {
    var issues = [];
    var t = (text || "").replace(/[\r\n\v\u000b]+/g, " ").trim();
    if (!t) {
      issues.push({ rule: "title-empty", severity: "error", message: "Slide has no title.", excerpt: "", fix: null });
      return issues;
    }
    var letters = t.replace(/[^A-Za-z]/g, "");
    if ((font && font.allCaps === true) || (letters.length >= 6 && letters === letters.toUpperCase())) {
      issues.push({ rule: "title-caps", severity: "error", message: "Title in all caps. Use sentence case.", excerpt: t.slice(0, 80),
        fix: font && font.allCaps === true ? "allcaps" : null });
    } else if (looksTitleCase(t)) {
      issues.push({ rule: "title-case", severity: "warning", message: "Title looks like Title Case. Use sentence case: capitalize the first word and proper nouns only.", excerpt: t.slice(0, 80), fix: null });
    }
    if (/[^.]\.\s*$/.test(t) && !/\b(etc|Inc|Ltd|Co)\.\s*$/.test(t)) {
      issues.push({ rule: "title-period", severity: "warning", message: "Title ends with a full stop.", excerpt: t.slice(-40), fix: "title-period" });
    }
    if (t.length > 120) {
      issues.push({ rule: "title-long", severity: "warning", message: "Title is " + t.length + " characters and may run past two lines.", excerpt: t.slice(0, 60) + "…", fix: null });
    }
    return issues;
  }

  /** Does any text on the slide count as a filled-in source line? */
  function hasSource(texts) {
    for (var i = 0; i < texts.length; i++) {
      var t = texts[i] || "";
      if (SOURCE_LINE.test(t)) {
        var after = t.replace(SOURCE_LINE, "").trim();
        if (after && !/^x{2,}$/i.test(after) && after !== "…") return true;
      }
    }
    return false;
  }

  /** Text replacements used by the safe fixes. Returns new string. */
  function fixDashes(text) {
    return text
      .replace(/ ?— ?/g, ", ")
      .replace(/ – /g, ", ")
      .replace(/, ,/g, ",");
  }

  /**
   * List the edits that turn `text` into fixDashes(text), so the checker can
   * apply them with TextRange.getSubstring (which keeps run formatting).
   * Edits are returned last-to-first so indexes stay valid while applying.
   */
  function dashEdits(text) {
    var edits = [];
    var re = / ?— ?| – /g, m;
    while ((m = re.exec(text))) edits.push({ start: m.index, length: m[0].length, text: ", " });
    return edits.reverse();
  }

  function typoEdits(text) {
    var edits = [];
    var re = /Delive-rables|Deliver-ables/g, m;
    while ((m = re.exec(text))) edits.push({ start: m.index, length: m[0].length, text: "Deliverables" });
    return edits.reverse();
  }

  function titlePeriodEdits(text) {
    var m = /\.(\s*)$/.exec(text);
    return m ? [{ start: m.index, length: 1, text: "" }] : [];
  }

  var api = {
    checkText: checkText,
    checkFont: checkFont,
    checkTitle: checkTitle,
    looksTitleCase: looksTitleCase,
    hasSource: hasSource,
    isLora: isLora,
    fixDashes: fixDashes,
    dashEdits: dashEdits,
    typoEdits: typoEdits,
    titlePeriodEdits: titlePeriodEdits,
    SOURCE_LINE: SOURCE_LINE
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.GNA_RULES = api;
})(typeof window !== "undefined" ? window : this);
