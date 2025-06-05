// === Bible Header and Footer Styling Script ===
// This script creates and applies proper styles to all BibleHeader and BibleFooter frames

// === Debug logging setup ===
var debugLog = [];
function log(message) {
    debugLog.push(message);
    // Uncomment if using ExtendScript Toolkit
    // $.writeln(message);
}

// === Prepare document ===
var doc = app.activeDocument;

// === Style creators ===
function defineParagraphStyle(doc, name, settings) {
    var style;
    // Remove existing style if it exists (except [Basic Paragraph])
    try {
        style = doc.paragraphStyles.itemByName(name);
        if (style.isValid && name !== '[Basic Paragraph]') {
            style.remove();
        }
    } catch (e) { }
    style = doc.paragraphStyles.add({ name: name });
    for (var key in settings) {
        try {
            style[key] = settings[key];
        } catch (e) {
            log("Could not set property " + key + " on style " + name + ": " + e);
        }
    }
    return style;
}

function defineCharacterStyle(doc, name, settings) {
    var style;
    // Remove existing style if it exists (except [None])
    try {
        style = doc.characterStyles.itemByName(name);
        if (style.isValid && name !== '[None]') {
            style.remove();
        }
    } catch (e) { }
    style = doc.characterStyles.add({ name: name });
    for (var key in settings) {
        try {
            style[key] = settings[key];
        } catch (e) {
            log("Could not set property " + key + " on character style " + name + ": " + e);
        }
    }
    return style;
}

// === Define Header and Footer Styles ===

// Header paragraph styles
var headerLeftStyle = defineParagraphStyle(doc, 'BibleHeader-Left', {
    appliedFont: 'Times New Roman',
    pointSize: 9,
    leading: 11,
    justification: Justification.LEFT_ALIGN,
    spaceBefore: 0,
    spaceAfter: 0,
    alignToBaseline: false,
    ruleBelow: true,
    ruleBelowWeight: 0.5,
    ruleBelowOffset: 0.25,
    ruleBelowLeftIndent: 0,
    ruleBelowRightIndent: 0,
    ruleBelowWidth: RuleWidth.COLUMN_WIDTH
});

var headerRightStyle = defineParagraphStyle(doc, 'BibleHeader-Right', {
    appliedFont: 'Times New Roman',
    pointSize: 9,
    leading: 11,
    justification: Justification.RIGHT_ALIGN,
    spaceBefore: 0,
    spaceAfter: 0,
    alignToBaseline: false,
    ruleBelow: true,
    ruleBelowWeight: 0.5,
    ruleBelowOffset: 2,
    ruleBelowLeftIndent: 0,
    ruleBelowRightIndent: 0,
    ruleBelowWidth: RuleWidth.COLUMN_WIDTH
});

// Footer paragraph styles  
var footerLeftStyle = defineParagraphStyle(doc, 'BibleFooter-Left', {
    appliedFont: 'Times New Roman',
    pointSize: 9,
    leading: 11,
    justification: Justification.LEFT_ALIGN,
    spaceBefore: 0,
    spaceAfter: 0,
    alignToBaseline: false
});

var footerRightStyle = defineParagraphStyle(doc, 'BibleFooter-Right', {
    appliedFont: 'Times New Roman',
    pointSize: 9,
    leading: 11,
    justification: Justification.RIGHT_ALIGN,
    spaceBefore: 0,
    spaceAfter: 0,
    alignToBaseline: false
});

// Character style for emphasis (optional)
var headerCharStyle = defineCharacterStyle(doc, 'BibleHeaderChar', {
    appliedFont: 'Times New Roman',
    fontStyle: 'Regular',
    pointSize: 9
});

// Try to set rule color to black
try {
    var blackColor = doc.colors.itemByName("Black");
    headerLeftStyle.ruleBelowColor = blackColor;
    headerRightStyle.ruleBelowColor = blackColor;
} catch (e) {
    try {
        var blackSwatch = doc.swatches.itemByName("Black");
        headerLeftStyle.ruleBelowColor = blackSwatch;
        headerRightStyle.ruleBelowColor = blackSwatch;
    } catch (e2) {
        log("Could not set rule color to black: " + e2);
    }
}

// === Utility: Frame finding functions ===
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
    var isLeft = (docPage.side && docPage.side === PageSideOptions.LEFT_HAND);
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

// === Apply styles to all headers and footers ===
var styledHeaderCount = 0;
var styledFooterCount = 0;

for (var p = 0; p < doc.pages.length; p++) {
    var page = doc.pages[p];
    var isLeft = (page.side && page.side === PageSideOptions.LEFT_HAND);

    log("=== Processing page " + (p + 1) + " (isLeft: " + isLeft + ") ===");

    // Style BibleHeader frame
    var headerFrame = findLabeledFrame(page, 'BibleHeader');
    if (!headerFrame) {
        // Try to override from master
        headerFrame = ensureOverriddenLabeledFrame(page, 'BibleHeader');
    }

    if (headerFrame) {
        log("Found header frame on page " + (p + 1));

        // Check if frame has content
        if (headerFrame.contents && headerFrame.contents.length > 0) {
            log("Header content: '" + headerFrame.contents + "'");

            // Apply appropriate style based on page side
            var headerStyle = isLeft ? headerLeftStyle : headerRightStyle;

            try {
                // Apply paragraph style to all paragraphs in the frame
                for (var i = 0; i < headerFrame.paragraphs.length; i++) {
                    headerFrame.paragraphs[i].appliedParagraphStyle = headerStyle;

                    // Apply character style to all characters
                    for (var j = 0; j < headerFrame.paragraphs[i].characters.length; j++) {
                        headerFrame.paragraphs[i].characters[j].appliedCharacterStyle = headerCharStyle;
                    }
                }

                styledHeaderCount++;
                log("Applied " + headerStyle.name + " to header on page " + (p + 1));

            } catch (e) {
                log("Error styling header on page " + (p + 1) + ": " + e);
            }
        } else {
            log("Header frame is empty on page " + (p + 1));
        }
    } else {
        log("No header frame found on page " + (p + 1));
    }

    // Style BibleFooter frame
    var footerFrame = findLabeledFrame(page, 'BibleFooter');
    if (!footerFrame) {
        // Try to override from master
        footerFrame = ensureOverriddenLabeledFrame(page, 'BibleFooter');
    }

    if (footerFrame) {
        log("Found footer frame on page " + (p + 1));

        // Check if frame has content
        if (footerFrame.contents && footerFrame.contents.length > 0) {
            log("Footer content: '" + footerFrame.contents + "'");

            // Apply appropriate style based on page side
            var footerStyle = isLeft ? footerLeftStyle : footerRightStyle;

            try {
                // Apply paragraph style to all paragraphs in the frame
                for (var i = 0; i < footerFrame.paragraphs.length; i++) {
                    footerFrame.paragraphs[i].appliedParagraphStyle = footerStyle;

                    // Apply character style to all characters
                    for (var j = 0; j < footerFrame.paragraphs[i].characters.length; j++) {
                        footerFrame.paragraphs[i].characters[j].appliedCharacterStyle = headerCharStyle;
                    }
                }

                styledFooterCount++;
                log("Applied " + footerStyle.name + " to footer on page " + (p + 1));

            } catch (e) {
                log("Error styling footer on page " + (p + 1) + ": " + e);
            }
        } else {
            log("Footer frame is empty on page " + (p + 1));
        }
    } else {
        log("No footer frame found on page " + (p + 1));
    }
}

// Force document recomposition to apply all changes
try {
    app.activeDocument.recompose();
    log("Document recomposed successfully");
} catch (e) {
    log("Error during document recomposition: " + e);
}

// Summary
log("=== STYLING COMPLETE ===");
log("Styled headers: " + styledHeaderCount);
log("Styled footers: " + styledFooterCount);
log("Total pages processed: " + doc.pages.length);

// Display the debug log
// alert("Header and footer styling complete!\n\nStyled " + styledHeaderCount + " headers and " + styledFooterCount + " footers.\n\n" + debugLog.join("\n"));