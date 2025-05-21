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
(function () {
  var LABELS = ['BibleBody'];
  for (var l = 0; l < LABELS.length; l++) {
    ensureOverriddenLabeledFrame(page, LABELS[l]);
  }
})();

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
var bookStyle = defineParagraphStyle(doc, 'BookTitle', {
  appliedFont: 'Times New Roman',
  pointSize: 18,
  leading: 16,
  justification: Justification.CENTER_ALIGN,
  spaceBefore: 1,
  spaceAfter: 0.25,
  alignToBaseline: false,
  capitalization: Capitalization.SMALL_CAPS,
  // spanColumn: 1634495520, // Make book titles span all columns -- can't figure out how to do this from the docs
});
var verseTextStyle = defineParagraphStyle(doc, 'VerseText', {
  appliedFont: 'Times New Roman',
  pointSize: 10,
  leading: 12,
  justification: Justification.FULLY_JUSTIFIED,
  hyphenation: false,
  spaceAfter: 0,
  alignToBaseline: false,
  dropCapCharacters: 1,
  dropCapLines: 2,
});
var verseTextStyle2 = defineParagraphStyle(doc, 'VerseTextTwoDigitChapter', {
  appliedFont: 'Times New Roman',
  pointSize: 10,
  leading: 12,
  justification: Justification.FULLY_JUSTIFIED,
  hyphenation: false,
  spaceAfter: 0,
  alignToBaseline: false,
  dropCapCharacters: 2,
  dropCapLines: 2,
});
var verseTextStyle3 = defineParagraphStyle(doc, 'VerseTextThreeDigitChapter', {
  appliedFont: 'Times New Roman',
  pointSize: 10,
  leading: 12,
  justification: Justification.FULLY_JUSTIFIED,
  hyphenation: false,
  spaceAfter: 0,
  alignToBaseline: false,
  dropCapCharacters: 3,
  dropCapLines: 2,
});
var verseNumStyle = defineCharacterStyle(doc, 'VerseNum', {
  position: Position.SUPERSCRIPT,
  pointSize: 6.5,
  appliedFont: 'Times New Roman',
  baselineShift: 2,
});
// Add DropCap character style for multi-character drop caps
var dropCapStyle = defineCharacterStyle(doc, 'DropCap', {
  tracking: 50, // Adjust as needed for spacing
  pointSize: 7
});

// === Insert Bible content ===
var prevBook = '', prevChapter = '';
var malformedIdCount = 0;
var missingFieldCount = 0;
var malformedChapterVerseCount = 0;
var errorProcessingCount = 0;
var successCount = 0;

for (var i = 0; i < data.length; i++) {
  var entry = data[i];
  if (!entry.id || !entry.translation) {
    // alert('Missing id or translation at entry ' + i);
    missingFieldCount++;
    continue;
  }
  // Handle multi-word book names (e.g., "1 Corinthians", "3 John")
  var idParts = entry.id.split(' ');
  if (idParts.length < 2) {
    // alert('Malformed id at entry ' + i + ': ' + entry.id);
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
    // Insert a paragraph break to start a new paragraph for the next content
    story.insertionPoints[-1].contents = '\r';
    // Set the style of the new (empty) paragraph to VerseText
    story.paragraphs[-1].appliedParagraphStyle = verseTextStyle;
    prevBook = book;
    prevChapter = '';
  }

  // Insert paragraph break before new chapter
  if (chapter !== prevChapter) {
    story.insertionPoints[-1].contents = '\r';
    // Insert chapter number as plain text (will become drop cap)
    story.insertionPoints[-1].contents = chapter;

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
        story.characters.itemByRange(
          story.characters[verseStartIndex],
          story.characters[verseStartIndex + verseLength - 1]
        ).appliedCharacterStyle = verseNumStyle;
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
        story.characters.itemByRange(
          story.characters[verseStartIndex],
          story.characters[verseStartIndex + verse.length - 1]
        ).appliedCharacterStyle = verseNumStyle;
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
