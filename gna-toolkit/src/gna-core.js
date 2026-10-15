/*
 * GNA Toolkit core: everything that talks to PowerPoint.
 * Shared by the task pane (taskpane.html) and the ribbon commands (commands.html).
 * Requires office.js, rules.js and catalog.js to be loaded first.
 */
(function () {
  "use strict";

  var GNA = (window.GNA = window.GNA || {});
  var R = window.GNA_RULES;

  // ---------------------------------------------------------------- constants
  GNA.COLORS = {
    teal: "#025043", midTeal: "#5B9188", deepTeal: "#013A31", teal4: "#3D7A6F",
    lightTeal: "#B3CBC7", grey: "#8C8C8C", lightGrey: "#F2F2F2", text: "#262626",
    stickerGrey: "#A6A6A6", harveyGrey: "#BFBFBF",
    // Status colors used in the template (green and red from slides 35 and 64).
    // Amber is not defined in the template: Office standard amber is used.
    green: "#00B050", amber: "#FFC000", red: "#E5546C"
  };

  // Slide geometry (template is 16:9, 960 x 540 pt) and template element positions in points
  GNA.GEOMETRY = {
    slideW: 960, slideH: 540,
    sticker: { left: 708.0, top: 101.6, width: 208.3, height: 19.9 },     // "For discussion" (template slides 14, 56)
    source: { left: 36.8, top: 503.2, width: 574.5, height: 28.75 },      // "Source:" line (footer, bottom left)
    sectionTag: { left: 600.8, top: 7.9, width: 315.5, height: 9.7 }      // on-page tracker (top right)
  };

  GNA.STICKERS = ["For discussion", "Preliminary", "Draft"];
  GNA.NAMES = { sticker: "GNA sticker", source: "GNA source", needed: "GNA needed flag", rag: "GNA status dot", section: "GNA section tag" };

  var TEXT_TYPES = { GeometricShape: 1, TextBox: 1, Placeholder: 1, Callout: 1, Freeform: 1 };

  // ---------------------------------------------------------------- helpers
  GNA.supports = function (version) {
    try {
      return Office.context.requirements.isSetSupported("PowerPointApi", version);
    } catch (e) {
      return false;
    }
  };

  GNA.assetUrl = function (path) {
    return new URL(path, document.baseURI).href;
  };

  GNA.store = {
    get: function (k, d) {
      try { var v = localStorage.getItem("gna." + k); return v === null ? d : JSON.parse(v); } catch (e) { return d; }
    },
    set: function (k, v) {
      try { localStorage.setItem("gna." + k, JSON.stringify(v)); } catch (e) { /* storage unavailable */ }
    }
  };

  function arrayBufferToBase64(buf) {
    var bytes = new Uint8Array(buf), chunk = 0x8000, parts = [];
    for (var i = 0; i < bytes.length; i += chunk) {
      parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
    }
    return btoa(parts.join(""));
  }

  var cache = {};
  GNA.fetchBase64 = async function (path) {
    if (cache[path]) return cache[path];
    var res = await fetch(GNA.assetUrl(path), { cache: "no-cache" });
    if (!res.ok) throw new Error("Could not download " + path + " (" + res.status + ")");
    var b64 = arrayBufferToBase64(await res.arrayBuffer());
    cache[path] = b64;
    return b64;
  };

  GNA.slidePath = function (n) { return "assets/slides/s" + (n < 10 ? "0" : "") + n + ".pptx"; };
  GNA.starterPath = function (id) { return "assets/starters/" + id + ".pptx"; };
  GNA.thumbPath = function (n) { return "assets/thumbs/s" + (n < 10 ? "0" : "") + n + ".jpg"; };

  function friendly(err) {
    var msg = (err && (err.message || err.toString())) || "Unknown error";
    if (/InvalidArgument|GeneralException/.test(msg) && err.debugInfo && err.debugInfo.errorLocation) {
      msg += " (" + err.debugInfo.errorLocation + ")";
    }
    return msg;
  }
  GNA.friendly = friendly;

  /** Selected slides, in deck order. Falls back to the active slide. */
  async function selectedSlides(ctx) {
    if (!GNA.supports("1.5")) throw new Error("This version of PowerPoint can't read the slide selection. Update Microsoft 365.");
    var sel = ctx.presentation.getSelectedSlides();
    sel.load("items/id");
    var all = ctx.presentation.slides;
    all.load("items/id");
    await ctx.sync();
    var ids = {};
    sel.items.forEach(function (s) { ids[s.id] = true; });
    var ordered = all.items.filter(function (s) { return ids[s.id]; });
    if (!ordered.length) throw new Error("Select a slide first (click it in the thumbnail pane).");
    return ordered;
  }
  GNA.selectedSlides = selectedSlides;

  // ---------------------------------------------------------------- inserting slides
  /** "auto": use the destination theme when the open deck already has the GNA master. */
  async function resolveFormatting(ctx) {
    var mode = GNA.store.get("formatting", "auto");
    if (mode === "keep") return "KeepSourceFormatting";
    if (mode === "destination") return "UseDestinationTheme";
    if (!GNA.supports("1.3")) return "KeepSourceFormatting";
    try {
      var masters = ctx.presentation.slideMasters;
      masters.load("items/id");
      await ctx.sync();
      var layoutSets = masters.items.map(function (m) { var l = m.layouts; l.load("items/name"); return l; });
      await ctx.sync();
      var isGna = layoutSets.some(function (ls) {
        var names = ls.items.map(function (l) { return l.name; });
        return names.indexOf("Default") >= 0 && names.indexOf("Frontpage dark") >= 0;
      });
      return isGna ? "UseDestinationTheme" : "KeepSourceFormatting";
    } catch (e) {
      return "KeepSourceFormatting";
    }
  }

  /**
   * Insert a .pptx (by asset path) after the last selected slide, or at the end.
   * Returns { inserted, formatting }.
   */
  GNA.insertFile = async function (path) {
    if (!GNA.supports("1.2")) throw new Error("This version of PowerPoint can't insert slides from add-ins. Update Microsoft 365.");
    var b64 = await GNA.fetchBase64(path);
    return PowerPoint.run(async function (ctx) {
      var pres = ctx.presentation;
      var slides = pres.slides;
      slides.load("items/id");
      var sel = null;
      if (GNA.supports("1.5")) { sel = pres.getSelectedSlides(); sel.load("items/id"); }
      await ctx.sync();

      var before = {};
      slides.items.forEach(function (s) { before[s.id] = true; });
      var target;
      if (sel && sel.items.length) {
        var selIds = {};
        sel.items.forEach(function (s) { selIds[s.id] = true; });
        for (var i = slides.items.length - 1; i >= 0; i--) {
          if (selIds[slides.items[i].id]) { target = slides.items[i].id; break; }
        }
      } else if (slides.items.length) {
        target = slides.items[slides.items.length - 1].id;
      }

      var formatting = await resolveFormatting(ctx);
      var opts = { formatting: formatting };
      if (target) opts.targetSlideId = target;
      pres.insertSlidesFromBase64(b64, opts);
      await ctx.sync();

      slides.load("items/id");
      await ctx.sync();
      var added = slides.items.filter(function (s) { return !before[s.id]; }).map(function (s) { return s.id; });
      if (added.length && GNA.supports("1.5")) {
        try { pres.setSelectedSlides([added[0]]); await ctx.sync(); } catch (e) { /* selection is a nicety */ }
      }
      return { inserted: added.length, formatting: formatting };
    });
  };

  GNA.insertTemplateSlide = function (n) { return GNA.insertFile(GNA.slidePath(n)); };

  GNA.openStarterAsNewDeck = async function (id) {
    if (!PowerPoint.createPresentation) throw new Error("This version of PowerPoint can't open new decks from add-ins. Use 'Insert into this deck' instead.");
    var b64 = await GNA.fetchBase64(GNA.starterPath(id));
    await PowerPoint.createPresentation(b64);
  };

  GNA.insertStarter = function (id) { return GNA.insertFile(GNA.starterPath(id)); };

  // ---------------------------------------------------------------- quick tools
  function needs14() {
    if (!GNA.supports("1.4")) throw new Error("This version of PowerPoint can't add shapes from add-ins. Update Microsoft 365.");
  }

  /** Load shapes (name, text) for a list of slides. Text only for text-capable shapes. */
  async function loadShapes(ctx, slides) {
    var colls = slides.map(function (s) {
      var c = s.shapes;
      c.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
      return c;
    });
    await ctx.sync();
    var withText = [];
    colls.forEach(function (c) {
      c.items.forEach(function (sh) {
        if (TEXT_TYPES[sh.type]) { sh.textFrame.textRange.load("text"); withText.push(sh); }
      });
    });
    try {
      await ctx.sync();
    } catch (e) {
      // One odd shape can fail the batch; retry one by one and skip failures.
      for (var i = 0; i < withText.length; i++) {
        try { withText[i].textFrame.textRange.load("text"); await ctx.sync(); } catch (e2) { withText[i]._noText = true; }
      }
    }
    return colls.map(function (c) { return c.items; });
  }

  function shapeText(sh) {
    try { return sh._noText ? "" : (sh.textFrame.textRange.text || ""); } catch (e) { return ""; }
  }

  function styleText(shape, opts) {
    var f = shape.textFrame.textRange.font;
    f.name = "Lora";
    if (opts.size) f.size = opts.size;
    if (opts.color) f.color = opts.color;
    if (opts.italic !== undefined) f.italic = opts.italic;
    if (opts.bold !== undefined) f.bold = opts.bold;
    if (opts.align) shape.textFrame.textRange.paragraphFormat.horizontalAlignment = opts.align;
    if (opts.zeroMargins) {
      shape.textFrame.leftMargin = 0; shape.textFrame.rightMargin = 0;
      shape.textFrame.topMargin = 0; shape.textFrame.bottomMargin = 0;
    }
    if (opts.valign) shape.textFrame.verticalAlignment = opts.valign;
    shape.textFrame.autoSizeSetting = "AutoSizeNone";
    shape.textFrame.wordWrap = true;
  }

  /** Add or replace the sticker on every selected slide. text = null removes it. */
  GNA.setSticker = async function (text) {
    needs14();
    return PowerPoint.run(async function (ctx) {
      var slides = await selectedSlides(ctx);
      var shapeLists = await loadShapes(ctx, slides);
      var g = GNA.GEOMETRY.sticker, count = 0;
      slides.forEach(function (slide, i) {
        var existing = shapeLists[i].filter(function (sh) {
          return sh.name === GNA.NAMES.sticker ||
            (GNA.STICKERS.concat(["Initial view"]).indexOf(shapeText(sh).trim()) >= 0 && sh.top > 80 && sh.top < 130 && sh.left > 600);
        });
        if (text === null) {
          existing.forEach(function (sh) { sh.delete(); count++; });
          return;
        }
        if (existing.length) {
          existing[0].textFrame.textRange.text = text;
          existing[0].name = GNA.NAMES.sticker;
          existing.slice(1).forEach(function (sh) { sh.delete(); });
        } else {
          var sh = slide.shapes.addTextBox(text, { left: g.left, top: g.top, width: g.width, height: g.height });
          sh.name = GNA.NAMES.sticker;
          styleText(sh, { size: 14, italic: true, color: GNA.COLORS.stickerGrey, align: "Right", zeroMargins: true });
        }
        count++;
      });
      await ctx.sync();
      return count;
    });
  };

  /** Replace the slide's source line text, or add one at the template position. */
  function setSourceOnSlide(slide, shapes, full) {
    var g = GNA.GEOMETRY.source;
    var existing = shapes.filter(function (sh) {
      return sh.name === GNA.NAMES.source || (R.SOURCE_LINE.test(shapeText(sh)) && sh.top > 440);
    })[0];
    if (existing) {
      existing.textFrame.textRange.text = full;
      existing.textFrame.textRange.font.name = "Lora";
      existing.textFrame.textRange.font.size = 9;
    } else {
      var sh = slide.shapes.addTextBox(full, { left: g.left, top: g.top, width: g.width, height: g.height });
      sh.name = GNA.NAMES.source;
      styleText(sh, { size: 9, color: GNA.COLORS.text, align: "Left", valign: "Middle" });
    }
  }

  /** Set the source line on every selected slide (adds one where missing). */
  GNA.setSource = async function (sourceText) {
    needs14();
    var full = /^\s*sources?\s*:/i.test(sourceText) ? sourceText.trim() : "Source: " + sourceText.trim();
    return PowerPoint.run(async function (ctx) {
      var slides = await selectedSlides(ctx);
      var shapeLists = await loadShapes(ctx, slides);
      slides.forEach(function (slide, i) { setSourceOnSlide(slide, shapeLists[i], full); });
      await ctx.sync();
      return slides.length;
    });
  };

  /** Set the section tag (top right) on every selected slide. */
  GNA.setSectionTag = async function (text) {
    needs14();
    return PowerPoint.run(async function (ctx) {
      var slides = await selectedSlides(ctx);
      var shapeLists = await loadShapes(ctx, slides);
      var g = GNA.GEOMETRY.sectionTag;
      slides.forEach(function (slide, i) {
        var tag = shapeLists[i].filter(function (sh) {
          return sh.name === GNA.NAMES.section || (TEXT_TYPES[sh.type] && sh.top < 25 && sh.left > 450 && sh.width > 100 && sh.height < 30);
        })[0];
        if (tag) {
          tag.textFrame.textRange.text = text;
        } else {
          var sh = slide.shapes.addTextBox(text, { left: g.left, top: g.top, width: g.width, height: g.height });
          sh.name = GNA.NAMES.section;
          styleText(sh, { size: 8, color: GNA.COLORS.text, align: "Right", zeroMargins: true, valign: "Middle" });
        }
      });
      await ctx.sync();
      return slides.length;
    });
  };

  /** Insert a yellow [NEEDED: ...] box on the active slide. */
  GNA.insertNeeded = async function (what) {
    needs14();
    var label = "[NEEDED: " + ((what || "").trim() || "describe what is missing") + "]";
    return PowerPoint.run(async function (ctx) {
      var slide = (await selectedSlides(ctx))[0];
      var sh = slide.shapes.addTextBox(label, { left: 330, top: 255, width: 300, height: 26 });
      sh.name = GNA.NAMES.needed;
      sh.fill.setSolidColor("#FFFF00");
      styleText(sh, { size: 10, bold: true, color: GNA.COLORS.text, align: "Center", valign: "Middle" });
      sh.load("id");
      await ctx.sync();
      if (GNA.supports("1.5")) { try { slide.setSelectedShapes([sh.id]); await ctx.sync(); } catch (e) { /* ignore */ } }
      return label;
    });
  };

  /** Wrap the selected text as [NEEDED: selection]. */
  GNA.wrapSelectionAsNeeded = async function () {
    if (!GNA.supports("1.5")) throw new Error("This version of PowerPoint can't read selected text. Update Microsoft 365.");
    return PowerPoint.run(async function (ctx) {
      var tr = ctx.presentation.getSelectedTextRangeOrNullObject();
      tr.load("text");
      await ctx.sync();
      if (tr.isNullObject || !tr.text) throw new Error("Select some text inside a text box first.");
      var inner = tr.text.replace(/^\[|\]$/g, "").replace(/^NEEDED:\s*/i, "");
      tr.text = "[NEEDED: " + inner + "]";
      await ctx.sync();
      return tr.text;
    });
  };

  /** Add a status dot. Next to the selected shape if there is one, else mid-slide. */
  GNA.addStatusDot = async function (colorKey) {
    needs14();
    var color = GNA.COLORS[colorKey] || GNA.COLORS.grey;
    return PowerPoint.run(async function (ctx) {
      var slide = (await selectedSlides(ctx))[0];
      var left = 474, top = 264, size = 12;
      if (GNA.supports("1.5")) {
        var shapes = ctx.presentation.getSelectedShapes();
        shapes.load("items/left,items/top,items/width,items/height");
        await ctx.sync();
        if (shapes.items.length) {
          var s = shapes.items[shapes.items.length - 1];
          left = Math.min(GNA.GEOMETRY.slideW - size - 4, s.left + s.width + 6);
          top = s.top + s.height / 2 - size / 2;
        }
      }
      var dot = slide.shapes.addGeometricShape("Ellipse", { left: left, top: top, width: size, height: size });
      dot.name = GNA.NAMES.rag + " (" + colorKey + ")";
      dot.fill.setSolidColor(color);
      dot.lineFormat.visible = false;
      dot.load("id");
      await ctx.sync();
      if (GNA.supports("1.5")) { try { slide.setSelectedShapes([dot.id]); await ctx.sync(); } catch (e) { /* ignore */ } }
    });
  };

  /** Harvey ball as SVG (0, 25, 50, 75, 100). */
  GNA.harveySvg = function (pct, px) {
    px = px || 36;
    var c = px / 2, r = px / 2 - 1, teal = GNA.COLORS.teal, grey = GNA.COLORS.harveyGrey;
    var body;
    if (pct >= 100) body = '<circle cx="' + c + '" cy="' + c + '" r="' + r + '" fill="' + teal + '"/>';
    else if (pct <= 0) body = '<circle cx="' + c + '" cy="' + c + '" r="' + r + '" fill="' + grey + '"/>';
    else {
      var a = (pct / 100) * 2 * Math.PI;
      var x = c + r * Math.sin(a), y = c - r * Math.cos(a);
      var large = pct > 50 ? 1 : 0;
      body = '<circle cx="' + c + '" cy="' + c + '" r="' + r + '" fill="' + grey + '"/>' +
        '<path d="M' + c + "," + c + " L" + c + "," + (c - r) + " A" + r + "," + r + " 0 " + large + ",1 " +
        x.toFixed(3) + "," + y.toFixed(3) + ' Z" fill="' + teal + '"/>';
    }
    return '<svg xmlns="http://www.w3.org/2000/svg" width="' + px + '" height="' + px + '" viewBox="0 0 ' + px + " " + px + '">' + body + "</svg>";
  };

  function svgToPngBase64(svg, px) {
    return new Promise(function (resolve, reject) {
      var img = new Image();
      img.onload = function () {
        var cv = document.createElement("canvas");
        cv.width = px; cv.height = px;
        cv.getContext("2d").drawImage(img, 0, 0, px, px);
        resolve(cv.toDataURL("image/png").split(",")[1]);
      };
      img.onerror = reject;
      img.src = "data:image/svg+xml;base64," + btoa(svg);
    });
  }

  /** Insert a Harvey ball on the current slide (as a picture). */
  GNA.insertHarvey = async function (pct) {
    var sizePt = 18;
    var opts = { imageLeft: 471, imageTop: 261, imageWidth: sizePt, imageHeight: sizePt };
    var svgOk = false;
    try { svgOk = Office.context.requirements.isSetSupported("ImageCoercion", "1.2"); } catch (e) { /* no */ }
    var data, type;
    if (svgOk) { data = GNA.harveySvg(pct, 36); type = Office.CoercionType.XmlSvg; }
    else { data = await svgToPngBase64(GNA.harveySvg(pct, 144), 144); type = Office.CoercionType.Image; }
    opts.coercionType = type;
    return new Promise(function (resolve, reject) {
      Office.context.document.setSelectedDataAsync(data, opts, function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res.error ? res.error.message : "Could not insert the Harvey ball."));
      });
    });
  };

  // ---------------------------------------------------------------- style checker
  /**
   * Scan the whole deck. Returns { slides, issues, needed, stats }.
   * Each issue: { key, slideId, slideNumber, path:[shapeIds], cell:[r,c]|null, rule, severity, message, excerpt, fix }
   */
  GNA.scanDeck = async function (onProgress) {
    if (!GNA.supports("1.4")) throw new Error("The style checker needs a newer version of PowerPoint (PowerPointApi 1.4). Update Microsoft 365.");
    var has18 = GNA.supports("1.8");
    var cat = window.GNA_CATALOG || { noSourceSlides: [] };
    return PowerPoint.run(async function (ctx) {
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      var masters = null;
      if (GNA.supports("1.3")) { masters = ctx.presentation.slideMasters; masters.load("items/id,items/name"); }
      await ctx.sync();

      var issues = [], needed = [], seq = 0;
      function add(slideInfo, path, cell, list) {
        list.forEach(function (it) {
          it.key = "i" + (seq++);
          it.slideId = slideInfo.id; it.slideNumber = slideInfo.number;
          it.path = path; it.cell = cell;
          if (it.rule === "needed") needed.push({ slideNumber: slideInfo.number, text: it.needed });
          issues.push(it);
        });
      }

      for (var si = 0; si < slides.items.length; si++) {
        var slide = slides.items[si];
        var info = { id: slide.id, number: si + 1 };
        if (onProgress) onProgress(si + 1, slides.items.length);

        // Slide meta + top-level shapes
        var tag = null, layout = null;
        if (GNA.supports("1.3")) {
          tag = slide.tags.getItemOrNullObject("GNA_TEMPLATE_ID"); tag.load("value");
          layout = slide.layout; layout.load("name");
        }
        var shapes = slide.shapes;
        shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
        try { await ctx.sync(); } catch (e) { add(info, [], null, [{ rule: "scan", severity: "info", message: "Couldn't read this slide: " + friendly(e), excerpt: "", fix: null }]); continue; }

        var templateId = tag && !tag.isNullObject ? parseInt(tag.value, 10) : null;
        var layoutName = layout ? layout.name : "";
        var exemptFromSource = (templateId && cat.noSourceSlides.indexOf(templateId) >= 0) ||
          /^(\d+_)?(Frontpage( dark)?|Section breaker|Top left title|Mid left title|End|Agenda|Blank dark)$/i.test(layoutName);

        if (templateId && templateId <= 3) {
          add(info, [], null, [{ rule: "style-guide", severity: "error", message: "Template style guide slide (" + templateId + ") is in the deck. Remove it before sending.", excerpt: "", fix: null }]);
        }

        // Collect shapes, descending into groups (1.8)
        var flat = []; // {shape, path}
        var groupQueue = [];
        shapes.items.forEach(function (sh) { flat.push({ shape: sh, path: [sh.id] }); if (sh.type === "Group" && has18) groupQueue.push({ shape: sh, path: [sh.id] }); });
        var depth = 0;
        while (groupQueue.length && depth < 3) {
          var batch = groupQueue.splice(0);
          var colls = batch.map(function (g) { var c = g.shape.group.shapes; c.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height"); return c; });
          try {
            await ctx.sync();
          } catch (e) {
            // Load groups one by one so one bad group doesn't hide the others
            for (var gb = 0; gb < batch.length; gb++) {
              try { colls[gb].load("items/id,items/name,items/type,items/left,items/top,items/width,items/height"); await ctx.sync(); colls[gb]._ok = true; } catch (e2) { colls[gb]._bad = true; }
            }
          }
          colls.forEach(function (c, gi) {
            if (c._bad) return;
            c.items.forEach(function (child) {
              var p = batch[gi].path.concat([child.id]);
              flat.push({ shape: child, path: p });
              if (child.type === "Group") groupQueue.push({ shape: child, path: p });
            });
          });
          depth++;
        }

        // Load text + fonts for text shapes, values for tables
        var textItems = flat.filter(function (f) { return TEXT_TYPES[f.shape.type]; });
        var tableItems = has18 ? flat.filter(function (f) { return f.shape.type === "Table"; }) : [];
        function queueLoads(item) {
          var tr = item.shape.textFrame.textRange;
          tr.load("text");
          tr.font.load(has18 ? "name,size,allCaps" : "name,size");
          if (has18 && item.shape.type === "Placeholder") item.shape.placeholderFormat.load("type");
        }
        textItems.forEach(queueLoads);
        tableItems.forEach(function (t) { t.table = t.shape.getTable(); t.table.load("values"); });
        try {
          await ctx.sync();
        } catch (e) {
          for (var k = 0; k < textItems.length; k++) {
            try { queueLoads(textItems[k]); await ctx.sync(); } catch (e2) { textItems[k].failed = true; }
          }
          for (var t2 = 0; t2 < tableItems.length; t2++) {
            try { tableItems[t2].table.load("values"); await ctx.sync(); } catch (e3) { tableItems[t2].failed = true; }
          }
        }

        var allTexts = [], titleFound = false;
        textItems.forEach(function (item) {
          if (item.failed) return;
          var sh = item.shape, text, font;
          try { text = sh.textFrame.textRange.text || ""; font = sh.textFrame.textRange.font; } catch (e) { return; }
          allTexts.push(text);
          var phType = null;
          try { phType = has18 && sh.type === "Placeholder" ? sh.placeholderFormat.type : null; } catch (e) { /* not loaded */ }
          var isTitle = phType === "Title" || phType === "CenterTitle" || (!phType && /^Title\b/i.test(sh.name) && sh.top < 120);
          if (!text.trim()) {
            if (isTitle && !titleFound) { titleFound = true; add(info, item.path, null, R.checkTitle("", {})); }
            return;
          }
          add(info, item.path, null, R.checkText(text, {}));
          var f = { name: safeGet(font, "name"), size: safeGet(font, "size"), allCaps: has18 ? safeGet(font, "allCaps") : undefined };
          add(info, item.path, null, R.checkFont(f, {}));
          if (isTitle && !titleFound) {
            titleFound = true;
            add(info, item.path, null, R.checkTitle(text, { allCaps: f.allCaps }));
          }
        });
        tableItems.forEach(function (t) {
          if (t.failed || !t.table.values) return;
          t.table.values.forEach(function (row, r) {
            row.forEach(function (cellText, c) {
              allTexts.push(cellText);
              add(info, t.path, [r, c], R.checkText(cellText, { inTable: true }));
            });
          });
        });

        if (!exemptFromSource && !R.hasSource(allTexts)) {
          add(info, [], null, [{ rule: "source", severity: "error", message: "No source line. Add \"Source: [Client] data\" or \"Source: GNA analysis\".", excerpt: "", fix: "add-source" }]);
        }
        if (!exemptFromSource) {
          var tracker = textItems.filter(function (it) { var s = it.shape; return s.top < 25 && s.left > 450 && s.width > 100; })[0];
          var trackerText = "";
          try { trackerText = tracker ? (tracker.shape.textFrame.textRange.text || "").trim() : ""; } catch (e) { /* ignore */ }
          if (!trackerText || /^section marker$/i.test(trackerText)) {
            add(info, [], null, [{ rule: "section-tag", severity: "info", message: "Section tag (top right) is empty. Use the agenda section name.", excerpt: "", fix: null }]);
          }
        }
      }

      if (masters && masters.items.length > 1) {
        issues.push({ key: "i" + (seq++), slideId: null, slideNumber: 0, path: [], cell: null, rule: "masters", severity: "info",
          message: "Deck has " + masters.items.length + " slide masters. Extra masters usually come from pasted slides and bloat the file.", excerpt: masters.items.map(function (m) { return m.name; }).join(", "), fix: null });
      }

      var stats = { error: 0, warning: 0, info: 0, fixable: 0 };
      issues.forEach(function (i) { stats[i.severity]++; if (i.fix && i.fix !== "add-source") stats.fixable++; });
      return { slideCount: slides.items.length, issues: issues, needed: needed, stats: stats };
    });
  };

  function safeGet(obj, prop) {
    try { return obj[prop]; } catch (e) { return undefined; }
  }

  /** Resolve a shape from a path of ids (top-level id, then group child ids). */
  function shapeFromPath(slide, path) {
    var sh = slide.shapes.getItem(path[0]);
    for (var i = 1; i < path.length; i++) sh = sh.group.shapes.getItem(path[i]);
    return sh;
  }

  async function applyEdits(ctx, shape, editsFn) {
    var tr = shape.textFrame.textRange;
    tr.load("text");
    await ctx.sync();
    var edits = editsFn(tr.text || "");
    edits.forEach(function (e) { tr.getSubstring(e.start, e.length).text = e.text; });
    await ctx.sync();
    return edits.length;
  }

  /** Apply the safe fix for one issue. Returns true if something changed. */
  GNA.fixIssue = async function (issue) {
    if (!issue.fix) return false;
    if (issue.fix === "add-source") {
      return PowerPoint.run(async function (ctx) {
        var slide = ctx.presentation.slides.getItem(issue.slideId);
        var shapes = (await loadShapes(ctx, [slide]))[0];
        setSourceOnSlide(slide, shapes, "Source: [NEEDED: source]");
        await ctx.sync();
        return true;
      });
    }
    return PowerPoint.run(async function (ctx) {
      var slide = ctx.presentation.slides.getItem(issue.slideId);
      if (issue.fix === "table-dash") {
        if (!GNA.supports("1.9")) throw new Error("Fixing text inside tables needs a newer PowerPoint. Fix this one by hand.");
        var table = shapeFromPath(slide, issue.path).getTable();
        var cell = table.getCellOrNullObject(issue.cell[0], issue.cell[1]);
        cell.load("textRuns");
        await ctx.sync();
        if (cell.isNullObject) return false;
        var runs = cell.textRuns;
        if (runs.length <= 1) {
          cell.text = R.fixDashes(runs.map(function (x) { return x.text; }).join(""));
        } else {
          cell.textRuns = runs.map(function (run) {
            var font = {};
            Object.keys(run.font || {}).forEach(function (k) { if (run.font[k] !== null && run.font[k] !== undefined) font[k] = run.font[k]; });
            return { text: R.fixDashes(run.text), font: font };
          });
        }
        await ctx.sync();
        return true;
      }
      var shape = shapeFromPath(slide, issue.path);
      if (issue.fix === "font") { shape.textFrame.textRange.font.name = "Lora"; await ctx.sync(); return true; }
      if (issue.fix === "allcaps") { shape.textFrame.textRange.font.allCaps = false; await ctx.sync(); return true; }
      if (issue.fix === "dash") return (await applyEdits(ctx, shape, R.dashEdits)) > 0;
      if (issue.fix === "typo") return (await applyEdits(ctx, shape, R.typoEdits)) > 0;
      if (issue.fix === "title-period") return (await applyEdits(ctx, shape, R.titlePeriodEdits)) > 0;
      return false;
    });
  };

  GNA.goToSlide = async function (slideId) {
    if (!slideId || !GNA.supports("1.5")) return;
    await PowerPoint.run(async function (ctx) { ctx.presentation.setSelectedSlides([slideId]); await ctx.sync(); });
  };
})();
