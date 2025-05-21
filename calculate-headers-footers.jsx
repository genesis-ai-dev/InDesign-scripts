// calculate-headers-footers.jsx
// Efficiently populate running headers and footers for all pages in the current InDesign document.
// Assumes frames are labeled 'BibleHeader', 'BibleFooter', and 'BibleBody' on each page.
// Place this script in the same directory as script.jsx.

(function () {
    var doc = app.activeDocument;
    var verseNumStyle = doc.characterStyles.itemByName('VerseNum');
    var bookStyle = doc.paragraphStyles.itemByName('BookTitle');

    // --- Pass 1: Build a map of which verses and book titles appear on which pages ---
    // We'll cache for each page: firstVerse, lastVerse, bookName
    var pageInfo = {};

    // Build a map from each story to its BibleBody frames by page
    var storyToFrames = {};
    for (var p = 0; p < doc.pages.length; p++) {
        var page = doc.pages[p];
        var items = page.allPageItems;
        for (var i = 0; i < items.length; i++) {
            var tf = items[i];
            if (tf.constructor && tf.constructor.name === 'TextFrame' && tf.label === 'BibleBody') {
                var story = tf.parentStory;
                if (!storyToFrames[story.id]) storyToFrames[story.id] = {};
                storyToFrames[story.id][page.documentOffset] = tf;
            }
        }
    }

    // For each story, do a single pass to find verse numbers and book titles by character index
    for (var storyId in storyToFrames) {
        var story = null;
        // Find the story object by id
        for (var s = 0; s < doc.stories.length; s++) {
            if (doc.stories[s].id == storyId) {
                story = doc.stories[s];
                break;
            }
        }
        if (!story) continue;

        // Build a list of verse numbers and book titles with their character indices and page numbers
        var verseEntries = [];
        var bookEntries = [];
        for (var i = 0; i < story.characters.length; i++) {
            var ch = story.characters[i];
            var page = ch.parentTextFrames.length > 0 ? ch.parentTextFrames[0].parentPage : null;
            if (!page) continue;
            var pageNum = page.documentOffset;
            if (ch.appliedCharacterStyle == verseNumStyle) {
                verseEntries.push({
                    index: i,
                    page: pageNum,
                    value: ch.contents
                });
            }
        }
        for (var j = 0; j < story.paragraphs.length; j++) {
            var para = story.paragraphs[j];
            var page = para.characters.length > 0 && para.characters[0].parentTextFrames.length > 0 ? para.characters[0].parentTextFrames[0].parentPage : null;
            if (!page) continue;
            var pageNum = page.documentOffset;
            if (para.appliedParagraphStyle == bookStyle) {
                bookEntries.push({
                    index: j,
                    page: pageNum,
                    value: String(para.contents).replace(/\r/g, '').replace(/^\s+|\s+$/g, '')
                });
            }
        }

        // For each page, cache first/last verse and last book title up to that page
        var lastBook = '';
        var pageVerses = {};
        for (var p = 0; p < doc.pages.length; p++) {
            var firstVerse = '', lastVerse = '';
            // Find first and last verse on this page
            for (var v = 0; v < verseEntries.length; v++) {
                if (verseEntries[v].page == p) {
                    if (!firstVerse) firstVerse = verseEntries[v].value;
                    lastVerse = verseEntries[v].value;
                }
            }
            // Find the last book title up to this page
            for (var b = 0; b < bookEntries.length; b++) {
                if (bookEntries[b].page <= p) {
                    lastBook = bookEntries[b].value;
                }
            }
            pageInfo[p] = pageInfo[p] || {};
            pageInfo[p].firstVerse = firstVerse;
            pageInfo[p].lastVerse = lastVerse;
            pageInfo[p].bookName = lastBook;
        }
    }

    // --- Pass 2: Fill headers and footers using the cached info ---
    for (var p = 0; p < doc.pages.length; p++) {
        var page = doc.pages[p];
        // Find header/footer frames
        var headerFrame = null, footerFrame = null;
        var items = page.allPageItems;
        for (var i = 0; i < items.length; i++) {
            var tf = items[i];
            if (tf.constructor && tf.constructor.name === 'TextFrame') {
                if (tf.label === 'BibleHeader') headerFrame = tf;
                if (tf.label === 'BibleFooter') footerFrame = tf;
            }
        }
        // Fill header
        if (headerFrame) {
            var info = pageInfo[p] || {};
            var headerText = '';
            if (p % 2 === 0) { // left page
                headerText = (info.bookName || '') + (info.firstVerse ? ' ' + info.firstVerse : '');
            } else { // right page
                headerText = (info.lastVerse ? info.lastVerse + ' ' : '') + (info.bookName || '');
            }
            headerFrame.contents = headerText;
            headerFrame.texts[0].justification = Justification.CENTER_ALIGN;
            headerFrame.texts[0].pointSize = 9;
        }
        // Fill footer
        if (footerFrame) {
            footerFrame.contents = SpecialCharacters.AUTO_PAGE_NUMBER;
            footerFrame.texts[0].justification = Justification.CENTER_ALIGN;
            footerFrame.texts[0].pointSize = 9;
        }
    }
})(); 