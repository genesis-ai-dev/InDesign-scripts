// === Bible Header Recalculation Script ===
// This script recalculates running headers for Bible content based on verse ranges

// === Debug logging setup ===
var debugLog = [];
function log(message) {
    debugLog.push(message);
    // Uncomment if using ExtendScript Toolkit
    // $.writeln(message);
}

// === Prepare document ===
var doc = app.activeDocument;

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

// === Header Styling Function ===
function styleHeaderFrame(page, headerFrame) {
    try {
        var isLeft = (page.side && page.side === PageSideOptions.LEFT_HAND);
        log('=== Styling header on page ' + page.name + ' (isLeft: ' + isLeft + ') ===');
        log('Header frame has ' + headerFrame.texts.length + ' texts');
        log('Header frame has ' + headerFrame.paragraphs.length + ' paragraphs');

        // Check if frame has content
        if (headerFrame.contents.length === 0) {
            log('Header frame is empty on page ' + page.name);
            return;
        }

        log('Header frame content: "' + headerFrame.contents + '"');

        // Apply styling to all text in the frame
        if (headerFrame.paragraphs.length > 0) {
            for (var i = 0; i < headerFrame.paragraphs.length; i++) {
                var paragraph = headerFrame.paragraphs[i];
                log('Processing paragraph ' + i + ': "' + paragraph.contents + '"');

                // Apply alignment
                if (isLeft) {
                    paragraph.justification = Justification.LEFT_ALIGN;
                    log('Applied LEFT alignment to paragraph ' + i + ' on page ' + page.name);
                } else {
                    paragraph.justification = Justification.RIGHT_ALIGN;
                    log('Applied RIGHT alignment to paragraph ' + i + ' on page ' + page.name);
                }

                // Add horizontal rule below the paragraph
                try {
                    paragraph.ruleBelow = true;
                    paragraph.ruleBelowWeight = 0.5; // Use numeric value instead of string

                    // Try to get black color, fallback if not found
                    try {
                        paragraph.ruleBelowColor = doc.colors.itemByName("Black");
                    } catch (colorError) {
                        try {
                            paragraph.ruleBelowColor = doc.swatches.itemByName("Black");
                        } catch (swatchError) {
                            log('Could not find Black color, using default');
                        }
                    }

                    paragraph.ruleBelowOffset = 2; // Use numeric value instead of string
                    paragraph.ruleBelowLeftIndent = 0;
                    paragraph.ruleBelowRightIndent = 0;
                    paragraph.ruleBelowWidth = RuleWidth.COLUMN_WIDTH;

                    log('Applied horizontal rule below paragraph ' + i + ' on page ' + page.name);
                } catch (ruleError) {
                    log('Error applying rule to paragraph ' + i + ': ' + ruleError);
                }
            }
        } else {
            // If no paragraphs, try to style the entire text frame
            log('No paragraphs found, styling entire text frame');
            try {
                if (isLeft) {
                    headerFrame.texts[0].justification = Justification.LEFT_ALIGN;
                    log('Applied LEFT alignment to entire text frame on page ' + page.name);
                } else {
                    headerFrame.texts[0].justification = Justification.RIGHT_ALIGN;
                    log('Applied RIGHT alignment to entire text frame on page ' + page.name);
                }
            } catch (textError) {
                log('Error styling text frame: ' + textError);
            }
        }

        // Force redraw to make changes visible
        try {
            app.activeDocument.recompose();
            log('Recomposed document to show changes');
        } catch (recomposeError) {
            log('Could not recompose document: ' + recomposeError);
        }

    } catch (e) {
        log('Error styling header on page ' + page.name + ': ' + e);
    }
}

// === Function to format verse reference for header ===
function formatVerseReference(ref) {
    if (!ref) return '';
    if (ref.book && ref.chapter && ref.verse) {
        return ref.book + ' ' + ref.chapter + ':' + ref.verse;
    } else if (ref.book && ref.chapter) {
        return ref.book + ' ' + ref.chapter;
    } else if (ref.book) {
        return ref.book;
    }
    return '';
}

// === Build page-to-verse map by walking through content sequentially ===
var pageVerseMap = {}; // pageIndex -> {verses: [{book, chapter, verse}], firstVerse, lastVerse}
var currentBook = '';
var currentChapter = '';

log("=== Building page-to-verse map by walking through content sequentially ===");

// Walk through all stories and paragraphs sequentially
for (var s = 0; s < doc.stories.length; s++) {
    var story = doc.stories[s];

    for (var p = 0; p < story.paragraphs.length; p++) {
        var para = story.paragraphs[p];
        var styleName = para.appliedParagraphStyle.name;
        var content = String(para.contents);

        // Skip empty paragraphs
        if (!content || content === '' || content === '\r' || content === '\n' || content === '\r\n') {
            continue;
        }

        // Check for book title
        if (styleName === 'BookTitle') {
            var bookName = content.replace(/\r$/, '').replace(/\n$/, '');
            currentBook = bookName;
            log("Sequential analysis: Set book to " + currentBook);
            continue;
        }

        // Check for chapter start
        if (styleName.indexOf('VerseText') === 0) {
            // Check if this paragraph has drop cap formatting (which indicates chapter start)
            // Chapter starts are identified by having dropCapCharacters > 0
            var isChapterStart = false;
            var chapterNumber = '';

            try {
                // Check if this paragraph style has drop cap settings
                var paraStyle = para.appliedParagraphStyle;
                if (paraStyle.dropCapCharacters > 0) {
                    // This is a chapter start - extract the chapter number from the beginning
                    var chapterMatch = content.match(/^(\d+)/);
                    if (chapterMatch) {
                        chapterNumber = chapterMatch[1];
                        isChapterStart = true;
                        log("Sequential analysis: Detected chapter start by drop cap: " + chapterNumber);
                    }
                }
            } catch (e) {
                // Fallback to text-based detection if drop cap check fails
                var chapterStartWords = /^\s*(In|Long|Na|Ol|Em|God|Jisas|Man|Woman|Nau|Taim|Wanpela|Mi|Bihaen|Na)/i;
                var chapterMatch = content.match(/^(\d+)\s+(.+)/);

                if (chapterMatch && chapterMatch[2].match(chapterStartWords)) {
                    chapterNumber = chapterMatch[1];
                    isChapterStart = true;
                    log("Sequential analysis: Detected chapter start by text pattern: " + chapterNumber);
                }
            }

            if (isChapterStart) {
                currentChapter = chapterNumber;
                log("Sequential analysis: Set chapter to " + currentChapter);

                // This paragraph contains chapter start (verse 1)
                // Find which page this paragraph is on
                try {
                    if (para.characters.length > 0) {
                        var paraFrame = para.characters[0].parentTextFrames[0];
                        if (paraFrame && paraFrame.parentPage) {
                            var pageIndex = paraFrame.parentPage.documentOffset;

                            if (!pageVerseMap[pageIndex]) {
                                pageVerseMap[pageIndex] = { verses: [], firstVerse: null, lastVerse: null };
                            }

                            var verseRef = {
                                book: currentBook,
                                chapter: currentChapter,
                                verse: '1'
                            };

                            pageVerseMap[pageIndex].verses.push(verseRef);
                            if (!pageVerseMap[pageIndex].firstVerse) {
                                pageVerseMap[pageIndex].firstVerse = verseRef;
                            }
                            pageVerseMap[pageIndex].lastVerse = verseRef;

                            log("Sequential analysis: Added " + formatVerseReference(verseRef) + " to page " + (pageIndex + 1));
                        }
                    }
                } catch (e) {
                    // Skip if can't determine page
                }
            }

            // Now look for individual verse numbers in this paragraph
            try {
                for (var i = 0; i < para.characters.length; i++) {
                    var character = para.characters[i];
                    var charContent = String(character.contents);

                    // Check if this character has superscript formatting (verse number)
                    if (character.appliedCharacterStyle && character.appliedCharacterStyle.name === 'VerseNum') {
                        // Collect the full verse number
                        var verseNum = charContent;
                        var j = i + 1;

                        // Collect consecutive superscript digits for multi-digit verse numbers
                        while (j < para.characters.length) {
                            var nextChar = para.characters[j];
                            if (nextChar.appliedCharacterStyle &&
                                nextChar.appliedCharacterStyle.name === 'VerseNum' &&
                                /\d/.test(String(nextChar.contents))) {
                                verseNum += String(nextChar.contents);
                                j++;
                            } else {
                                break;
                            }
                        }

                        if (/^\d+$/.test(verseNum)) {
                            log("Sequential analysis: Found superscript verse number: " + verseNum + " in chapter " + currentChapter);

                            // Find which page this character is on
                            try {
                                if (character.parentTextFrames.length > 0) {
                                    var charFrame = character.parentTextFrames[0];
                                    if (charFrame && charFrame.parentPage) {
                                        var pageIndex = charFrame.parentPage.documentOffset;

                                        if (!pageVerseMap[pageIndex]) {
                                            pageVerseMap[pageIndex] = { verses: [], firstVerse: null, lastVerse: null };
                                        }

                                        var verseRef = {
                                            book: currentBook,
                                            chapter: currentChapter,
                                            verse: verseNum
                                        };

                                        pageVerseMap[pageIndex].verses.push(verseRef);
                                        if (!pageVerseMap[pageIndex].firstVerse) {
                                            pageVerseMap[pageIndex].firstVerse = verseRef;
                                        }
                                        pageVerseMap[pageIndex].lastVerse = verseRef;

                                        log("Sequential analysis: Added " + formatVerseReference(verseRef) + " to page " + (pageIndex + 1));
                                    }
                                }
                            } catch (e) {
                                // Skip if can't determine page
                            }

                            i = j - 1; // Skip the characters we've already processed
                        }
                    }
                }
            } catch (e) {
                log("Error examining characters in paragraph: " + e);
            }
        }
    }
}

// Log the final page-to-verse map
for (var pageIndex in pageVerseMap) {
    var pageData = pageVerseMap[pageIndex];
    var verseRefs = [];
    for (var v = 0; v < pageData.verses.length; v++) {
        verseRefs.push(formatVerseReference(pageData.verses[v]));
    }
    log("Page " + (parseInt(pageIndex) + 1) + " verses: [" + verseRefs.join(", ") + "]");
    log("  First: " + formatVerseReference(pageData.firstVerse) + ", Last: " + formatVerseReference(pageData.lastVerse));
}

// === Process pages in pairs (spreads) for proper left/right header logic ===
for (var p = 0; p < doc.pages.length; p += 2) {
    var leftPage = doc.pages[p];
    var rightPage = (p + 1 < doc.pages.length) ? doc.pages[p + 1] : null;

    log("=== Processing spread: pages " + (p + 1) + " and " + (rightPage ? (p + 2) : "none") + " ===");

    // Get verse ranges from our map
    var rightPageData = pageVerseMap[p] || { firstVerse: null, lastVerse: null };
    var leftPageData = rightPage ? (pageVerseMap[p + 1] || { firstVerse: null, lastVerse: null }) : { firstVerse: null, lastVerse: null };

    log("Left page first verse: " + formatVerseReference(rightPageData.firstVerse) + ", last verse: " + formatVerseReference(rightPageData.lastVerse));
    log("Right page first verse: " + formatVerseReference(leftPageData.firstVerse) + ", last verse: " + formatVerseReference(leftPageData.lastVerse));

    // Set left page header (last verse on LEFT page)
    var leftHeaderFrame = ensureOverriddenLabeledFrame(leftPage, 'BibleHeader');
    if (leftHeaderFrame) {
        if (rightPageData.lastVerse) {
            var leftHeaderText = formatVerseReference(rightPageData.lastVerse);
            leftHeaderFrame.contents = leftHeaderText;
            log("Set left header to: '" + leftHeaderText + "'");
            // Apply styling after setting content
            styleHeaderFrame(leftPage, leftHeaderFrame);
        } else {
            log("No last verse found for left header");
        }
    } else {
        log("No left header frame found");
    }

    // Set right page header (first verse on RIGHT page)
    if (rightPage) {
        var rightHeaderFrame = ensureOverriddenLabeledFrame(rightPage, 'BibleHeader');
        if (rightHeaderFrame) {
            if (leftPageData.firstVerse) {
                var rightHeaderText = formatVerseReference(leftPageData.firstVerse);
                rightHeaderFrame.contents = rightHeaderText;
                log("Set right header to: '" + rightHeaderText + "'");
                // Apply styling after setting content
                styleHeaderFrame(rightPage, rightHeaderFrame);
            } else {
                log("No first verse found for right header");
            }
        } else {
            log("No right header frame found");
        }
    }
}

// Force final recomposition
app.activeDocument.recompose();

// Display the debug log
// alert("Header recalculation complete!\n\n" + debugLog.join("\n"));