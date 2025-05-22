// === Ensure Overridden Labeled Frames on All Pages ===
(function() {
  var doc = app.activeDocument;

  for (var p = 0; p < doc.pages.length; p++) {
    ensureOverriddenLabeledFrame(doc.pages[p], 'BibleBody');
  }

  for (var p = 0; p < doc.pages.length; p++) {
    ensureOverriddenLabeledFrame(doc.pages[p], 'BibleHeader');
  }

  // === Utility: Frame override logic ===
  function findLabeledFrame(page, label) {
    var items = page.allPageItems;
    for (var i = 0; i < items.length; i++) {
      var tf = items[i];
      if (tf.constructor.name === 'TextFrame' && tf.label === label) return tf;
    }
    return null;
  }
  function getMasterPageForDocPage(docPage) {
    var master = docPage.appliedMaster;
    if (!master) return null;
    if (master.pages.length === 1) return master.pages[0];
    var isLeft = (docPage.side && docPage.side === PageSideOptions.LEFT_HAND) || (docPage.documentOffset % 2 === 0);
    return master.pages[isLeft ? 0 : 1];
  }
  function ensureOverriddenLabeledFrame(page, label) {
    var local = findLabeledFrame(page, label);
    if (local) return local;
    var masterPage = getMasterPageForDocPage(page);
    if (!masterPage) return null;
    for (var i = 0; i < masterPage.textFrames.length; i++) {
      var tf = masterPage.textFrames[i];
      if (tf.label === label) {
        try {
          return tf.override(page);
        } catch (e) {
          continue;
        }
      }
    }
    return null;
  }
})();



