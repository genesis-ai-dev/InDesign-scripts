var doc = app.activeDocument;

for (var i = 0; i < doc.pages.length; i++) {
    var page = doc.pages[i];
    var frames = page.textFrames;

    for (var j = 0; j < frames.length; j++) {
        var frame = frames[j];

        // Check if the frame has the script label "BibleFooter"
        if (frame.label === "BibleFooter") {
            frame.contents = ""; // clear any existing content

            // Insert current page number marker
            frame.insertionPoints[0].contents = SpecialCharacters.AUTO_PAGE_NUMBER;

            // Apply paragraph style
            frame.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName('BibleFooter-Left');

            // Apply character style to all characters
            for (var k = 0; k < frame.paragraphs[0].characters.length; k++) {
                frame.paragraphs[0].characters[k].appliedCharacterStyle = doc.characterStyles.itemByName('BibleHeaderChar');
            }
        }
    }
}