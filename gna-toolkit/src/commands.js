/*
 * Ribbon buttons that act immediately (no task pane).
 * Function names must match <FunctionName> values in manifest.xml.
 * Errors are swallowed: ribbon commands have no UI in PowerPoint, so the
 * task pane's Tools tab is the place to see error messages.
 */
(function () {
  "use strict";

  function wrap(fn) {
    return function (event) {
      Promise.resolve()
        .then(fn)
        .catch(function (e) { console.error("GNA Toolkit command failed:", e); })
        .then(function () { event.completed(); });
    };
  }

  var handlers = {
    stickerDiscussion: wrap(function () { return window.GNA.setSticker("For discussion"); }),
    stickerPreliminary: wrap(function () { return window.GNA.setSticker("Preliminary"); }),
    stickerDraft: wrap(function () { return window.GNA.setSticker("Draft"); }),
    stickerRemove: wrap(function () { return window.GNA.setSticker(null); }),
    sourceGna: wrap(function () { return window.GNA.setSource("GNA analysis"); }),
    sourceNeeded: wrap(function () { return window.GNA.setSource("[NEEDED: source]"); }),
    insertNeeded: wrap(function () { return window.GNA.insertNeeded(""); })
  };

  Office.onReady(function () {
    Object.keys(handlers).forEach(function (name) {
      if (Office.actions && Office.actions.associate) Office.actions.associate(name, handlers[name]);
      else window[name] = handlers[name];
    });
  });
  // Older hosts look the functions up on the global object.
  Object.keys(handlers).forEach(function (name) { window[name] = handlers[name]; });
})();
