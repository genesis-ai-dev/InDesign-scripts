// === Bible Content Import Script ===
// After running this script, run calculate-headers-footers.jsx to populate running headers and footers.
// This script only inserts Bible content and links BibleBody frames.

// FIXMES:
// - new chapters aren't going to newlines and having the chapter number inserted
// - verse text style needs to be drop caps 3 lines on character style [none]
// - need to make sure verse number gets inserted; don't insert chapter number

// === JSON shim for ExtendScript ===
if (typeof JSON === 'undefined') {
  JSON = {};
  JSON.parse = function (s) { return eval('(' + s + ')'); };
}

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
  for (var key in settings) { style[key] = settings[key]; }
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
  for (var key in settings) { style[key] = settings[key]; }
  return style;
}

// === Load JSON file ===
var file = File.openDialog('Select Bible JSON file');
if (!file) exit();
file.open('r');
var jsonString = file.read();
file.close();

var data;
try {
  data = JSON.parse(jsonString);
} catch (e) {
  alert('Failed to parse JSON: ' + e);
  exit();
}
if (!data || !data.length) {
  alert('No data found in JSON.');
  exit();
}
alert('Loaded ' + data.length + ' entries');

// === Prepare document ===
var doc = app.activeDocument;

// Ensure [Basic Paragraph] does not align to baseline grid
try {
  var basicParagraphStyle = doc.paragraphStyles.itemByName('[Basic Paragraph]');
  if (basicParagraphStyle.isValid) {
    basicParagraphStyle.alignToBaseline = false;
  }
} catch (e) { }

// Target the first page and first BibleBody frame
var page = doc.pages[0];
if (!page) {
  alert('No pages found in document.');
  exit();
}

// === OVERRIDE FRAMES FIRST TO ENSURE BIBLEBODY EXISTS ===
// (function () {
//   var LABELS = ['BibleBody', 'BibleHeader'];
//   for (var l = 0; l < LABELS.length; l++) {
//     ensureOverriddenLabeledFrame(page, LABELS[l]);
//   }
// })();

// Override BibleHeader frames on all pages
for (var p = 0; p < doc.pages.length; p++) {
  ensureOverriddenLabeledFrame(doc.pages[p], 'BibleHeader');
}

// Find the first BibleBody frame on the first page
var textFrame = null;
var items = page.allPageItems;
for (var i = 0; i < items.length; i++) {
  var tf = items[i];
  if (tf.constructor && tf.constructor.name === 'TextFrame' && tf.label === 'BibleBody') {
    textFrame = tf;
    break;
  }
}

if (!textFrame) {
  alert('No BibleBody frame found on the first page. Make sure your master page has a frame labeled BibleBody.');
  exit();
}

// Set the text frame to 2 columns (if needed)
textFrame.textFramePreferences.textColumnCount = 2;

// Clear existing content
if (textFrame.parentStory) {
  textFrame.parentStory.contents = '';
  var story = textFrame.parentStory;
} else {
  alert('BibleBody frame is not connected to a story');
  exit();
}

// === Define Styles ===
// Shared style properties
var sharedStyleProps = {
  appliedFont: 'Times New Roman',
  justification: Justification.LEFT_JUSTIFIED,
  hyphenation: false,
  spaceAfter: 0.1875,
  alignToBaseline: false,
  dropCapLines: 2,
  leading: 13.5
};

var bookStyle = defineParagraphStyle(doc, 'BookTitle', {
  appliedFont: sharedStyleProps.appliedFont,
  pointSize: 18,
  leading: 16,
  justification: Justification.CENTER_ALIGN,
  spaceBefore: 0.3125,
  spaceAfter: 0,
  alignToBaseline: sharedStyleProps.alignToBaseline,
  capitalization: Capitalization.SMALL_CAPS,
  // spanColumn: 1634495520, // Make book titles span all columns -- can't figure out how to do this from the docs
});

var verseTextStyle = defineParagraphStyle(doc, 'VerseText', {
  appliedFont: sharedStyleProps.appliedFont,
  pointSize: 10,
  leading: sharedStyleProps.leading,
  justification: sharedStyleProps.justification,
  hyphenation: sharedStyleProps.hyphenation,
  spaceAfter: sharedStyleProps.spaceAfter,
  alignToBaseline: sharedStyleProps.alignToBaseline,
  dropCapCharacters: 1,
  dropCapLines: sharedStyleProps.dropCapLines,
});

var verseTextStyle2 = defineParagraphStyle(doc, 'VerseTextTwoDigitChapter', {
  appliedFont: sharedStyleProps.appliedFont,
  pointSize: 10,
  leading: sharedStyleProps.leading,
  justification: sharedStyleProps.justification,
  hyphenation: sharedStyleProps.hyphenation,
  spaceAfter: sharedStyleProps.spaceAfter,
  alignToBaseline: sharedStyleProps.alignToBaseline,
  dropCapCharacters: 2,
  dropCapLines: sharedStyleProps.dropCapLines,
});

var verseTextStyle3 = defineParagraphStyle(doc, 'VerseTextThreeDigitChapter', {
  appliedFont: sharedStyleProps.appliedFont,
  pointSize: 10,
  leading: sharedStyleProps.leading,
  justification: sharedStyleProps.justification,
  hyphenation: sharedStyleProps.hyphenation,
  spaceAfter: sharedStyleProps.spaceAfter,
  alignToBaseline: sharedStyleProps.alignToBaseline,
  dropCapCharacters: 3,
  dropCapLines: sharedStyleProps.dropCapLines,
});
var verseNumStyle = defineCharacterStyle(doc, 'VerseNum', {
  position: Position.SUPERSCRIPT,
  pointSize: 10,
  appliedFont: 'Times New Roman',
  fontStyle: 'Bold',
  baselineShift: 2,
});
// Add DropCap character style for multi-character drop caps
var dropCapStyle = defineCharacterStyle(doc, 'DropCap', {
  tracking: 20, // Adjust as needed for spacing
  pointSize: 10
});

// === Insert Bible content ===
var prevBook = '', prevChapter = '';
var malformedIdCount = 0;
var missingFieldCount = 0;
var malformedChapterVerseCount = 0;
var errorProcessingCount = 0;
var successCount = 0;

var debugLog = [];
function log(message) {
  debugLog.push(message);
  // Uncomment if using ExtendScript Toolkit
  // $.writeln(message);
}

for (var i = 0; i < data.length; i++) {
  var entry = data[i];
  if (!entry.id || !entry.translation) {
    missingFieldCount++;
    continue;
  }
  // Handle multi-word book names (e.g., "1 Corinthians", "3 John")
  var idParts = entry.id.split(' ');
  if (idParts.length < 2) {
    malformedIdCount++;
    continue;
  }
  // Check if first part is a number (e.g., "1", "3")
  var book = idParts[0];
  if (!isNaN(parseInt(book))) {
    // If it is, combine with next word for book name
    book = book + ' ' + idParts[1];
    // Remove the second word from idParts since we've used it
    idParts.splice(1, 1);
  }
  var chapterVerse = idParts[1].split(':');
  if (chapterVerse.length < 2) {
    malformedChapterVerseCount++;
    continue;
  }
  var chapter = chapterVerse[0];
  var verse = chapterVerse[1];
  var text = entry.translation;

  if (book !== prevBook) {
    // Ensure we are starting a new paragraph for the book title
    if (story.characters.length > 0 && story.characters[-1].contents !== '\r') {
      story.insertionPoints[-1].contents = '\r';
    }
    // Insert book title and ensure a paragraph break after
    story.insertionPoints[-1].contents = book + '\r';
    story.paragraphs[-1].appliedParagraphStyle = bookStyle;
    // Reset character style for the book title paragraph
    story.characters.itemByRange(story.paragraphs[-1].characters[0], story.paragraphs[-1].characters[-1]).appliedCharacterStyle = doc.characterStyles.itemByName('[None]');
    // Insert a space to start a new paragraph for the next content and to ensure book name styles are applied
    story.insertionPoints[-1].contents = ' ';
    // Set the style of the new (empty) paragraph to VerseText
    story.paragraphs[-1].appliedParagraphStyle = verseTextStyle;
    prevBook = book;
    prevChapter = '';
  }

  // Insert paragraph break before new chapter
  if (chapter !== prevChapter) {
    story.insertionPoints[-1].contents = '\r';
    // Insert chapter number as plain text (will become drop cap)
    story.insertionPoints[-1].contents = chapter + ' ';

    // Apply the verse text style first
    story.paragraphs[-1].appliedParagraphStyle = verseTextStyle;

    // Then directly set the dropCapCharacters property based on chapter length
    if (chapter.length === 2) {
      story.paragraphs[-1].appliedParagraphStyle = verseTextStyle2;
    } else if (chapter.length === 3) {
      story.paragraphs[-1].appliedParagraphStyle = verseTextStyle3;
    }

    // Apply DropCap style to the chapter number
    var para = story.paragraphs[-1];
    para.characters.itemByRange(0, chapter.length - 1).appliedCharacterStyle = dropCapStyle;
    // Add extra space before the drop cap paragraph to prevent collision with the line above
    para.spaceBefore = 10; // Adjust this value as needed
    

    prevChapter = chapter;
  }

  // Insert verse text with verse number for all verses
  try {
    // Get the current insertion point (where text will be added) and its index
    var ip = story.insertionPoints[-1];
    var verseStartIndex = ip.index;

    // Only add verse numbers for verses after the first verse of each chapter
    // (verse 1 doesn't get a number since the chapter number serves as the drop cap)
    if (verse !== '1' || chapter === prevChapter) {
      if (verse && verse.length > 0) {
        // Insert the verse number as plain text
        ip.contents = verse;
        // Apply verse number style to all characters in the verse number
        var verseLength = verse.length; // Store the actual length of the verse number
        if (verse !== '1') {
          story.characters.itemByRange(
            story.characters[verseStartIndex],
            story.characters[verseStartIndex + verseLength - 1]
          ).appliedCharacterStyle = verseNumStyle;
        } 
      }
    }

    // Then add the text with default character style
    ip = story.insertionPoints[-1];
    var textStartIndex = ip.index; // Store where the actual verse text begins
    ip.contents = text + ' ';

    // Apply verse text style to the paragraph first
    story.paragraphs[-1].appliedParagraphStyle = verseTextStyle;

    // Ensure all text after the verse number has no character style
    if (verse !== '1' || chapter === prevChapter) {
      if (verse && verse.length > 0) {
        // Clear any character styling from the verse text (not the verse number)
        story.characters.itemByRange(
          story.characters[textStartIndex],
          story.characters[-1]
        ).appliedCharacterStyle = doc.characterStyles.itemByName('[None]');

        // Re-apply verse number style to ensure it stays styled correctly
        if (verse !== '1') {
          story.characters.itemByRange(
            story.characters[verseStartIndex],
            story.characters[verseStartIndex + verse.length - 1]
          ).appliedCharacterStyle = verseNumStyle;
        } else {
          // remove the first character of the verse text
          story.characters.itemByRange(
            story.characters[verseStartIndex],
            story.characters[verseStartIndex + verse.length - 1]
          ).remove();

        }
      }
    } else {
      // For verse 1 of each chapter, ensure all text has no character style
      // The drop cap is handled by the paragraph style
      story.characters.itemByRange(
        story.characters[verseStartIndex],
        story.characters[-1]
      ).appliedCharacterStyle = doc.characterStyles.itemByName('[None]');
    }

    // Apply verse text style to the paragraph
    if (chapter.length === 2) {
      story.paragraphs[-1].appliedParagraphStyle = verseTextStyle2;
    } else if (chapter.length === 3) {
      story.paragraphs[-1].appliedParagraphStyle = verseTextStyle3;
    } else {
      story.paragraphs[-1].appliedParagraphStyle = verseTextStyle;
    }
    successCount++;
  } catch (e) {
    errorProcessingCount++;
    break;
  }
}

// === Link all BibleBody frames in order ===
var frames = [];
for (var p = 0; p < doc.pages.length; p++) {
  var page = doc.pages[p];
  var items = page.allPageItems;
  for (var i = 0; i < items.length; i++) {
    var tf = items[i];
    if (tf.constructor && tf.constructor.name === 'TextFrame' && tf.label === 'BibleBody') {
      frames.push(tf);
    }
  }
}
if (frames.length > 0 && frames[0] !== textFrame) {
  frames = frames.filter(function (frame) { return frame !== textFrame; });
  frames.unshift(textFrame);
}
for (var i = 0; i < frames.length - 1; i++) {
  if (frames[i].nextTextFrame !== frames[i + 1]) {
    frames[i].nextTextFrame = frames[i + 1];
  }
}

// Check for overflow and create new pages as needed
var lastFrame = frames[frames.length - 1] || textFrame;
var maxNewPages = 1000; // Safety limit to prevent infinite loops
var pagesAdded = 0;

// Keep adding pages until no more overflow or we hit our safety limit
while (pagesAdded < maxNewPages) {
  // Force recomposition to correctly detect overflow
  app.activeDocument.recompose();
  
  // Check if we still have overflow
  if (!lastFrame.overflows) {
    log("All content fits - no more pages needed");
    break;
  }
  
  log("Content overflow detected - creating new spread");
  
  // Create two new pages (a spread) at the end of the document
  var leftPage = doc.pages.add(LocationOptions.AFTER, doc.pages[-1]);
  var rightPage = doc.pages.add(LocationOptions.AFTER, leftPage);
  pagesAdded += 2;
  
  // Override BibleBody frame on left page
  var leftFrame = ensureOverriddenLabeledFrame(leftPage, 'BibleBody');
  if (!leftFrame) {
    log("Failed to create BibleBody frame on left page");
    break;
  }
  
  // Override BibleHeader frame on left page
  ensureOverriddenLabeledFrame(leftPage, 'BibleHeader');
  
  // Connect the previous last frame to the new left frame
  lastFrame.nextTextFrame = leftFrame;
  
  // Override BibleBody frame on right page
  var rightFrame = ensureOverriddenLabeledFrame(rightPage, 'BibleBody');
  if (!rightFrame) {
    log("Failed to create BibleBody frame on right page");
    break;
  }
  
  // Override BibleHeader frame on right page
  ensureOverriddenLabeledFrame(rightPage, 'BibleHeader');
  
  // Override BibleFooter frame on right page
  ensureOverriddenLabeledFrame(rightPage, 'BibleFooter');
  
  // Connect left frame to right frame
  leftFrame.nextTextFrame = rightFrame;
  
  // Update lastFrame for next iteration
  lastFrame = rightFrame;
  
  // Add the new frames to our frames array
  frames.push(leftFrame);
  frames.push(rightFrame);
}

if (pagesAdded >= maxNewPages) {
  log("WARNING: Reached maximum page limit. Content may still be overflowing.");
}

// === Add running headers with proper verse ranges ===
// Walk through all content sequentially to build a page-to-verse map with proper context
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
                pageVerseMap[pageIndex] = {verses: [], firstVerse: null, lastVerse: null};
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
                      pageVerseMap[pageIndex] = {verses: [], firstVerse: null, lastVerse: null};
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

// Process pages in pairs (spreads) for proper left/right header logic
for (var p = 0; p < doc.pages.length; p += 2) {
  var leftPage = doc.pages[p];
  var rightPage = (p + 1 < doc.pages.length) ? doc.pages[p + 1] : null;
  
  log("=== Processing spread: pages " + (p + 1) + " and " + (rightPage ? (p + 2) : "none") + " ===");
  
  // Get verse ranges from our map
  var leftPageData = pageVerseMap[p] || {firstVerse: null, lastVerse: null};
  var rightPageData = rightPage ? (pageVerseMap[p + 1] || {firstVerse: null, lastVerse: null}) : {firstVerse: null, lastVerse: null};
  
  // Determine the spread's verse range
  var spreadFirstVerse = leftPageData.firstVerse;
  var spreadLastVerse = rightPageData.lastVerse || leftPageData.lastVerse;
  
  log("Spread first verse: " + formatVerseReference(spreadFirstVerse));
  log("Spread last verse: " + formatVerseReference(spreadLastVerse));
  
  // Set left page header (first verse of spread)
  var leftHeaderFrame = ensureOverriddenLabeledFrame(leftPage, 'BibleHeader');
  if (leftHeaderFrame) {
    if (spreadFirstVerse) {
      var leftHeaderText = formatVerseReference(spreadFirstVerse);
      leftHeaderFrame.contents = leftHeaderText;
      log("Set left header to: '" + leftHeaderText + "'");
    } else {
      log("No first verse found for left header");
    }
  } else {
    log("No left header frame found");
  }
  
  // Set right page header (last verse of spread)
  if (rightPage) {
    var rightHeaderFrame = ensureOverriddenLabeledFrame(rightPage, 'BibleHeader');
    if (rightHeaderFrame) {
      if (spreadLastVerse) {
        var rightHeaderText = formatVerseReference(spreadLastVerse);
        rightHeaderFrame.contents = rightHeaderText;
        log("Set right header to: '" + rightHeaderText + "'");
      } else {
        log("No last verse found for right header");
      }
    } else {
      log("No right header frame found");
    }
  }
}

// === Add page numbers to BibleFooter frames ===
for (var p = 0; p < doc.pages.length; p++) {
  var page = doc.pages[p];
  
  // Ensure the BibleFooter frame is overridden
  var footerFrame = ensureOverriddenLabeledFrame(page, 'BibleFooter');
  
  if (!footerFrame) {
    log("No BibleFooter frame found on page " + (p+1));
    continue;
  }

  // Add the page number to the footer
  footerFrame.contents = String(p + 1); // Adding 1 since pages are 0-indexed
}

// At the end of your script, display the log
alert(debugLog.join("\n"));
