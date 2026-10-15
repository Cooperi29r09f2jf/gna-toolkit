/*
 * A small in-memory stand-in for Office.js / PowerPoint, used only by tests/ui_test.py
 * to exercise the task pane in a normal browser. It mimics the proxy API shapes the
 * toolkit uses (load/sync, collections, getItem, textRange.getSubstring...), not
 * PowerPoint's behavior in every detail.
 */
(function () {
  var nextId = 1000;
  var calls = (window.__calls = []);

  function textShape(name, text, o) {
    o = o || {};
    return { id: String(nextId++), name: name, type: o.type || "TextBox", left: o.left || 50, top: o.top || 150, width: o.width || 300, height: o.height || 40,
      text: text, font: { name: o.font || "Lora", size: o.size || 12, allCaps: !!o.allCaps }, ph: o.ph || null };
  }

  // ---------------------------------------------------------------- the fake deck
  var deck = (window.__deck = {
    masters: [{ id: "m1", name: "GNA", layouts: ["Frontpage", "Default", "Frontpage dark", "Agenda"] }],
    selected: [],
    slides: [
      { id: "256#", tags: { GNA_TEMPLATE_ID: "4" }, layout: "Frontpage dark",
        shapes: [textShape("Title 1", "Claude rollout proposal for C.A. Hull", { type: "Placeholder", ph: "CenterTitle", top: 200 })] },
      { id: "257#", tags: { GNA_TEMPLATE_ID: "16" }, layout: "Default",
        shapes: [
          textShape("Title 1", "Revenue Grew Faster Than Costs In Every Region.", { type: "Placeholder", ph: "Title", top: 22, left: 44, width: 873, size: 25 }),
          textShape("Text Placeholder 8", "", { type: "Placeholder", ph: "Body", top: 8, left: 601, width: 316, height: 10, size: 8 }),
          textShape("TextBox 26", "Dispatch is manual — crews wait for calls", { font: "Arial" }),
          textShape("TextBox 27", "[Insert challenge]", {}),
          textShape("Google Shape;507;p14", "Source: xx", { top: 503, left: 37, width: 574, height: 29, size: 9 })
        ] },
      { id: "258#", tags: { GNA_TEMPLATE_ID: "35" }, layout: "Default",
        shapes: [
          textShape("Title 1", "Three of seven KPIs are behind plan", { type: "Placeholder", ph: "Title", top: 22, left: 44, width: 873, size: 25 }),
          { id: String(nextId++), name: "Table 4", type: "Table", left: 40, top: 114, width: 873, height: 300,
            table: [["KPI", "Status"], ["Close rate — Q3", "[NEEDED: Q3 close rate]"]] },
          { id: String(nextId++), name: "Group 9", type: "Group", left: 40, top: 420, width: 300, height: 40,
            children: [textShape("TextBox 3", "Owner — Clay", { top: 420 })] }
        ] }
    ]
  });
  deck.selected = ["257#"];

  // ---------------------------------------------------------------- proxies
  function loadable(o) { o.load = function () { return o; }; return o; }

  function substring(model, start, len) {
    return {
      set text(v) { model.text = model.text.slice(0, start) + v + model.text.slice(start + (len === undefined ? model.text.length : len)); },
      get text() { return model.text.substr(start, len); }
    };
  }

  function shapeProxy(model, slideModel, parentList) {
    var fontProxy = loadable({});
    ["name", "size", "allCaps", "bold", "italic", "color"].forEach(function (k) {
      Object.defineProperty(fontProxy, k, { get: function () { return model.font[k]; }, set: function (v) { model.font[k] = v; } });
    });
    var tr = loadable({ font: fontProxy, paragraphFormat: {},
      getSubstring: function (s, l) { return substring(model, s, l); } });
    Object.defineProperty(tr, "text", { get: function () { return model.text; }, set: function (v) { model.text = v; } });
    var p = {
      get id() { return model.id; }, get type() { return model.type; },
      get left() { return model.left; }, get top() { return model.top; }, get width() { return model.width; }, get height() { return model.height; },
      get name() { return model.name; }, set name(v) { model.name = v; },
      textFrame: { textRange: tr },
      fill: { setSolidColor: function (c) { model.fill = c; } },
      lineFormat: {},
      placeholderFormat: loadable({ get type() { return model.ph; } }),
      delete: function () { parentList.splice(parentList.indexOf(model), 1); },
      getTable: function () {
        return loadable({
          get values() { return model.table; },
          getCellOrNullObject: function (r, c) {
            return loadable({ isNullObject: false,
              get text() { return model.table[r][c]; },
              set text(v) { model.table[r][c] = v; },
              get textRuns() { return [{ text: model.table[r][c], font: {} }]; },
              set textRuns(runs) { model.table[r][c] = runs.map(function (x) { return x.text; }).join(""); } });
          }
        });
      }
    };
    if (model.type === "Group") p.group = { shapes: shapesProxy(model.children, slideModel) };
    if (model.type !== "Table") {
      Object.defineProperty(p.textFrame, "leftMargin", { set: function () {} });
    }
    return p;
  }

  function shapesProxy(list, slideModel) {
    var c = loadable({
      get items() { return list.map(function (m) { return shapeProxy(m, slideModel, list); }); },
      getItem: function (id) { var m = list.filter(function (x) { return x.id === id; })[0]; if (!m) throw new Error("ItemNotFound " + id); return shapeProxy(m, slideModel, list); },
      addTextBox: function (text, o) {
        var m = textShape("TextBox", text, { left: o.left, top: o.top, width: o.width, height: o.height, font: "Lora", size: 18 });
        list.push(m); calls.push(["addTextBox", slideModel.id, text]);
        return shapeProxy(m, slideModel, list);
      },
      addGeometricShape: function (type, o) {
        var m = { id: String(nextId++), name: type, type: "GeometricShape", left: o.left, top: o.top, width: o.width, height: o.height, text: "", font: { name: "Lora", size: 12 } };
        list.push(m); calls.push(["addGeometricShape", slideModel.id, type]);
        return shapeProxy(m, slideModel, list);
      }
    });
    return c;
  }

  function slideProxy(model) {
    return {
      id: model.id,
      tags: { getItemOrNullObject: function (k) { return loadable({ isNullObject: !(k in model.tags), value: model.tags[k] }); } },
      layout: loadable({ name: model.layout }),
      shapes: shapesProxy(model.shapes, model),
      setSelectedShapes: function (ids) { calls.push(["setSelectedShapes", ids]); }
    };
  }

  function slidesProxy() {
    return loadable({
      get items() { return deck.slides.map(slideProxy); },
      getItem: function (id) { return slideProxy(deck.slides.filter(function (s) { return s.id === id; })[0]); }
    });
  }

  var presentation = {
    get slides() { return slidesProxy(); },
    get slideMasters() {
      return loadable({ items: deck.masters.map(function (m) {
        return { id: m.id, name: m.name, layouts: loadable({ items: m.layouts.map(function (n) { return { name: n }; }) }) };
      }) });
    },
    getSelectedSlides: function () { return loadable({ items: deck.slides.filter(function (s) { return deck.selected.indexOf(s.id) >= 0; }).map(slideProxy) }); },
    setSelectedSlides: function (ids) { deck.selected = ids.slice(); calls.push(["setSelectedSlides", ids]); },
    getSelectedShapes: function () { return loadable({ items: [] }); },
    getSelectedTextRangeOrNullObject: function () { return loadable({ isNullObject: true }); },
    insertSlidesFromBase64: function (b64, opts) {
      calls.push(["insertSlidesFromBase64", b64.length, opts]);
      var idx = opts.targetSlideId ? deck.slides.findIndex(function (s) { return s.id === opts.targetSlideId; }) + 1 : 0;
      deck.slides.splice(idx, 0, { id: (nextId++) + "#", tags: { GNA_TEMPLATE_ID: "x" }, layout: "Default", shapes: [] });
    }
  };

  window.PowerPoint = {
    run: function (fn) { return Promise.resolve(fn({ presentation: presentation, sync: function () { return Promise.resolve(); } })); },
    createPresentation: function (b64) { calls.push(["createPresentation", b64.length]); return Promise.resolve(); }
  };
  window.Office = {
    onReady: function (cb) { setTimeout(function () { cb({ host: "PowerPoint" }); }, 0); return Promise.resolve(); },
    context: {
      requirements: { isSetSupported: function (name, v) { return name === "ImageCoercion" || parseFloat(v) <= 1.9; } },
      document: { setSelectedDataAsync: function (data, opts, cb) { calls.push(["setSelectedDataAsync", opts.coercionType, data.slice(0, 40)]); cb({ status: "succeeded" }); } }
    },
    CoercionType: { XmlSvg: "xmlSvg", Image: "image" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    actions: { associate: function () {} }
  };
})();
