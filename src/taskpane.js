/* GNA Toolkit task pane UI. */
(function () {
  "use strict";

  var CAT = window.GNA_CATALOG;
  var GNA = window.GNA;
  var $ = function (sel, root) { return (root || document).querySelector(sel); };
  var $$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };

  function el(tag, attrs, children) {
    var e = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      if (k === "class") e.className = attrs[k];
      else if (k === "text") e.textContent = attrs[k];
      else if (k === "html") e.innerHTML = attrs[k];
      else if (k.indexOf("on") === 0) e.addEventListener(k.slice(2), attrs[k]);
      else e.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(function (c) { if (c) e.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return e;
  }

  // ------------------------------------------------------------ toast
  var toastTimer = null;
  function toast(msg, isError) {
    var t = $("#toast");
    t.textContent = msg;
    t.className = isError ? "err" : "";
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, isError ? 7000 : 3200);
  }

  async function run(button, fn, okMsg) {
    if (button) button.disabled = true;
    try {
      var result = await fn();
      if (okMsg) toast(typeof okMsg === "function" ? okMsg(result) : okMsg);
      return result;
    } catch (e) {
      console.error(e);
      toast(GNA.friendly(e), true);
    } finally {
      if (button) button.disabled = false;
    }
  }

  // ------------------------------------------------------------ tabs
  function showTab(name) {
    $$(".tabs button").forEach(function (b) { b.setAttribute("aria-selected", b.dataset.tab === name ? "true" : "false"); });
    $$(".tab").forEach(function (s) { s.hidden = s.id !== "tab-" + name; });
    GNA.store.set("tab", name);
  }

  // ------------------------------------------------------------ library
  function slideName(n) { return (CAT.slides[String(n)] || {}).name || "Slide " + n; }

  function card(n, onInsert) {
    var c = el("button", { class: "card", title: "Insert template slide " + n + ": " + slideName(n), "data-n": n }, [
      el("img", { src: GNA.thumbPath(n), alt: "", loading: "lazy" }),
      el("div", { class: "meta" }, [el("span", { class: "num", text: String(n) }), el("span", { class: "name", text: slideName(n) })])
    ]);
    c.addEventListener("click", function () { onInsert(n, c); });
    return c;
  }

  async function insertSlide(n, cardEl) {
    if (cardEl.classList.contains("busy")) return;
    cardEl.classList.add("busy");
    try {
      var res = await GNA.insertTemplateSlide(n);
      var recent = GNA.store.get("recent", []).filter(function (x) { return x !== n; });
      recent.unshift(n);
      GNA.store.set("recent", recent.slice(0, 8));
      renderRecent();
      toast("Inserted " + slideName(n).toLowerCase() + " (" + n + ")" + (res.formatting === "KeepSourceFormatting" ? ", template formatting kept" : ""));
    } catch (e) {
      console.error(e);
      toast(GNA.friendly(e), true);
    } finally {
      cardEl.classList.remove("busy");
    }
  }

  function renderRecent() {
    var recent = GNA.store.get("recent", []).filter(function (n) { return CAT.slides[String(n)]; });
    var row = $("#lib-recent-row");
    row.innerHTML = "";
    recent.forEach(function (n) { row.appendChild(card(n, insertSlide)); });
    $("#lib-recent").hidden = !recent.length || !!$("#lib-search").value.trim() || $("#lib-cat").value !== "all";
  }

  function renderLibrary() {
    var q = $("#lib-search").value.trim().toLowerCase();
    var catId = $("#lib-cat").value;
    var wrap = $("#lib-groups");
    wrap.innerHTML = "";
    var shown = 0;
    CAT.categories.forEach(function (cat) {
      if (catId !== "all" && cat.id !== catId) return;
      var matches = cat.slides.filter(function (n) {
        if (!q) return true;
        var s = CAT.slides[String(n)] || {};
        var hay = (n + " " + s.name + " " + (s.keys || "") + " " + cat.name).toLowerCase();
        return q.split(/\s+/).every(function (w) { return hay.indexOf(w) >= 0; });
      });
      if (!matches.length) return;
      shown += matches.length;
      var grid = el("div", { class: "grid" });
      matches.forEach(function (n) { grid.appendChild(card(n, insertSlide)); });
      wrap.appendChild(el("div", { class: "group" }, [
        el("h3", {}, [cat.name, el("span", { text: matches.length + (matches.length === 1 ? " slide" : " slides") })]),
        grid
      ]));
    });
    if (!shown) wrap.appendChild(el("p", { class: "empty", text: "No slides match “" + q + "”." }));
    renderRecent();
  }

  function initLibrary() {
    var sel = $("#lib-cat");
    sel.appendChild(el("option", { value: "all", text: "All categories" }));
    CAT.categories.forEach(function (c) { sel.appendChild(el("option", { value: c.id, text: c.name })); });
    $("#lib-search").addEventListener("input", renderLibrary);
    sel.addEventListener("change", renderLibrary);
    var fmt = $("#fmt");
    fmt.value = GNA.store.get("formatting", "auto");
    fmt.addEventListener("change", function () { GNA.store.set("formatting", fmt.value); });
    renderLibrary();
  }

  // ------------------------------------------------------------ starters
  function initStarters() {
    var list = $("#starter-list");
    CAT.starters.forEach(function (st) {
      var items = st.slides.map(function (n) { return el("li", {}, [el("b", { text: String(n) }), " " + slideName(n)]); });
      var openBtn = el("button", { class: "primary", text: "Open as new deck" });
      var insBtn = el("button", { text: "Insert into this deck" });
      openBtn.addEventListener("click", function () {
        run(openBtn, function () { return GNA.openStarterAsNewDeck(st.id); }, st.name + " deck opened in a new window.");
      });
      insBtn.addEventListener("click", function () {
        run(insBtn, function () { return GNA.insertStarter(st.id); }, function (r) { return "Inserted " + r.inserted + (r.inserted === 1 ? " slide." : " slides."); });
      });
      list.appendChild(el("div", { class: "starter" }, [
        el("h3", { text: st.name + " (" + st.slides.length + " slides)" }),
        el("ol", {}, items),
        st.note ? el("p", { class: "small", text: st.note }) : null,
        el("div", { class: "btn-row" }, [openBtn, insBtn])
      ]));
    });
  }

  // ------------------------------------------------------------ tools
  function initTools() {
    var row = $("#sticker-row");
    GNA.STICKERS.forEach(function (text) {
      var b = el("button", { text: text });
      b.addEventListener("click", function () {
        run(b, function () { return GNA.setSticker(text); }, function (n) { return "“" + text + "” on " + n + (n === 1 ? " slide." : " slides."); });
      });
      row.appendChild(b);
    });
    $("#sticker-remove").addEventListener("click", function (e) {
      run(e.target, function () { return GNA.setSticker(null); }, function (n) { return n ? "Removed " + n + (n === 1 ? " sticker." : " stickers.") : "No stickers on the selected slides."; });
    });

    var client = $("#src-client");
    client.value = GNA.store.get("client", "");
    client.addEventListener("change", function () { GNA.store.set("client", client.value.trim()); });
    client.addEventListener("focus", function () { $('input[name="src"][value="client"]').checked = true; });
    $("#src-custom").addEventListener("focus", function () { $('input[name="src"][value="custom"]').checked = true; });
    $("#src-apply").addEventListener("click", function (e) {
      var mode = $('input[name="src"]:checked').value;
      var text;
      if (mode === "client") {
        var name = client.value.trim();
        if (!name) { toast("Type the client name first.", true); client.focus(); return; }
        text = name + " data";
      } else if (mode === "gna") {
        text = "GNA analysis";
      } else {
        text = $("#src-custom").value.trim();
        if (!text) { toast("Type the source first.", true); $("#src-custom").focus(); return; }
      }
      run(e.target, function () { return GNA.setSource(text); }, function (n) { return "Source line set on " + n + (n === 1 ? " slide." : " slides."); });
    });

    $("#sec-apply").addEventListener("click", function (e) {
      var t = $("#sec-text").value.trim();
      if (!t) { toast("Type the section name first.", true); return; }
      run(e.target, function () { return GNA.setSectionTag(t); }, function (n) { return "Section tag set on " + n + (n === 1 ? " slide." : " slides."); });
    });

    $("#needed-insert").addEventListener("click", function (e) {
      run(e.target, function () { return GNA.insertNeeded($("#needed-text").value); }, function (l) { return "Inserted " + l; });
    });
    $("#needed-wrap").addEventListener("click", function (e) {
      run(e.target, function () { return GNA.wrapSelectionAsNeeded(); }, "Wrapped. Now highlight it yellow (Home > Text Highlight Color).");
    });

    $$(".dot-btn").forEach(function (b) {
      b.addEventListener("click", function () { run(b, function () { return GNA.addStatusDot(b.dataset.dot); }); });
    });

    var hrow = $("#harvey-row");
    [0, 25, 50, 75, 100].forEach(function (p) {
      var b = el("button", { class: "harvey-btn", title: p + "%", html: GNA.harveySvg(p, 36) + "<span>" + p + "</span>" });
      b.addEventListener("click", function () { run(b, function () { return GNA.insertHarvey(p); }); });
      hrow.appendChild(b);
    });
  }

  // ------------------------------------------------------------ checker
  var lastScan = null;

  function sevFilter() {
    var on = {};
    $$("#check-filters input").forEach(function (i) { on[i.dataset.sev] = i.checked; });
    return on;
  }

  function renderScan() {
    var scan = lastScan;
    var s = scan.stats;
    var sum = $("#check-summary");
    sum.innerHTML = "";
    if (!s.error && !s.warning) sum.appendChild(el("span", { class: "pill ok", text: "No errors or warnings in " + scan.slideCount + " slides" }));
    if (s.error) sum.appendChild(el("span", { class: "pill error", text: s.error + (s.error === 1 ? " error" : " errors") }));
    if (s.warning) sum.appendChild(el("span", { class: "pill warning", text: s.warning + (s.warning === 1 ? " warning" : " warnings") }));
    if (s.info) sum.appendChild(el("span", { class: "pill info", text: s.info + " info" }));
    sum.hidden = false;
    $("#check-filters").hidden = !scan.issues.length;

    var pending = scan.issues.filter(function (i) { return i.fix && i.fix !== "add-source" && !i.fixed; }).length;
    var fixBtn = $("#check-fix");
    fixBtn.hidden = !pending;
    fixBtn.textContent = "Fix safe issues (" + pending + ")";

    var on = sevFilter();
    var bySlide = {};
    scan.issues.forEach(function (i) {
      if (!on[i.severity] || i.rule === "needed") return;
      (bySlide[i.slideNumber] = bySlide[i.slideNumber] || []).push(i);
    });
    var res = $("#check-results");
    res.innerHTML = "";
    Object.keys(bySlide).map(Number).sort(function (a, b) { return a - b; }).forEach(function (num) {
      var list = bySlide[num];
      var goBtn = num ? el("button", { text: "Go to" }) : null;
      if (goBtn) goBtn.addEventListener("click", function () { run(goBtn, function () { return GNA.goToSlide(list[0].slideId); }); });
      var block = el("div", { class: "slide-block" }, [el("header", {}, [el("span", { text: num ? "Slide " + num : "Whole deck" }), goBtn])]);
      var order = { error: 0, warning: 1, info: 2 };
      list.sort(function (a, b) { return order[a.severity] - order[b.severity]; }).forEach(function (issue) {
        var fixBtn2 = null;
        if (issue.fix && !issue.fixed) {
          fixBtn2 = el("button", { text: issue.fix === "add-source" ? "Add" : "Fix" });
          fixBtn2.addEventListener("click", function () {
            run(fixBtn2, async function () {
              var changed = await GNA.fixIssue(issue);
              if (changed) { issue.fixed = true; renderScan(); }
              return changed;
            }, function (changed) { return changed ? "Fixed." : "Nothing to change (already fixed?)."; });
          });
        }
        block.appendChild(el("div", { class: "issue" + (issue.fixed ? " fixed" : "") }, [
          el("span", { class: "sev " + issue.severity, title: issue.severity }),
          el("div", { class: "body" }, [el("div", { text: issue.message }), issue.excerpt ? el("div", { class: "ex", text: issue.excerpt }) : null]),
          fixBtn2
        ]));
      });
      res.appendChild(block);
    });

    var nl = $("#needed-list");
    nl.innerHTML = "";
    scan.needed.forEach(function (n) { nl.appendChild(el("li", { text: "Slide " + n.slideNumber + ": " + n.text })); });
    $("#check-needed").hidden = !scan.needed.length;
  }

  async function runCheck(btn) {
    var prog = $("#check-progress");
    prog.hidden = false;
    prog.textContent = "Reading slides…";
    await run(btn, async function () {
      lastScan = await GNA.scanDeck(function (i, n) { prog.textContent = "Checking slide " + i + " of " + n + "…"; });
      renderScan();
    });
    prog.hidden = true;
  }

  function initCheck() {
    $("#check-run").addEventListener("click", function (e) { runCheck(e.target); });
    $("#check-fix").addEventListener("click", function (e) {
      var btn = e.target;
      run(btn, async function () {
        var todo = lastScan.issues.filter(function (i) { return i.fix && i.fix !== "add-source" && !i.fixed; });
        var done = 0, failed = 0;
        for (var k = 0; k < todo.length; k++) {
          btn.textContent = "Fixing " + (k + 1) + " of " + todo.length + "…";
          try { if (await GNA.fixIssue(todo[k])) done++; todo[k].fixed = true; } catch (err) { failed++; console.warn(err); }
        }
        await runCheck($("#check-run"));
        return { done: done, failed: failed };
      }, function (r) { return "Fixed " + r.done + (r.failed ? ", " + r.failed + " need a manual fix" : "") + ". Re-checked the deck."; });
    });
    $$("#check-filters input").forEach(function (i) { i.addEventListener("change", function () { if (lastScan) renderScan(); }); });
    $("#needed-copy").addEventListener("click", function () {
      var text = lastScan.needed.map(function (n) { return "Slide " + n.slideNumber + ": " + n.text; }).join("\n");
      try {
        navigator.clipboard.writeText(text).then(function () { toast("Copied."); }, function () { fallbackCopy(text); });
      } catch (e) { fallbackCopy(text); }
    });
  }

  function fallbackCopy(text) {
    var ta = el("textarea", {});
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); toast("Copied."); } catch (e) { toast("Couldn't copy. Select the list and copy it.", true); }
    ta.remove();
  }

  // ------------------------------------------------------------ boot
  function boot() {
    $$(".tabs button").forEach(function (b) { b.addEventListener("click", function () { showTab(b.dataset.tab); }); });
    var params = new URLSearchParams(location.search);
    var tab = params.get("tab") || GNA.store.get("tab", "library");
    if (!$("#tab-" + tab)) tab = "library";
    showTab(tab);
    initLibrary();
    initStarters();
    initTools();
    initCheck();
    $("#version").textContent = "GNA Toolkit, catalog version " + CAT.version + ".";
  }

  Office.onReady(function () {
    if (!GNA.supports("1.2")) {
      toast("This PowerPoint version is too old for the toolkit. Update Microsoft 365.", true);
    }
    boot();
  });
})();
