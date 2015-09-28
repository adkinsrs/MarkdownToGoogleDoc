/*
Usage: 
  Adding this script to your doc: 
    - Tools > Script Manager > New
    - Select "Blank Project", then paste this code in and save.
  Running the script:
    - Select some text in your Google Document
    - Click the new Convert Markdown menu, and click Convert from Markdown
    
Note:
    - Google App Scripts has limited regex support so I had to be a bit hacky
    - One problem I've had is combining bold and italic emphasis if the italic portion is at the end of the bold portion (** bold and *italic***)
        - My recommentation is to designate one with asterisks and the other with underscores
*/

// Create a menu item on opening document
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Convert Markdown')
    .addItem('Convert from Markdown', 'convertFromMarkdown')
    .addToUi();
}

// Converts a selection of text from Markdown to standard Google Document format
function convertFromMarkdown() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {    
      processElement(elements[i]);
    }
  } else {
    throw 'Please select some text.';
  }  
}

// Process current element
function processElement(elt) {
  //TODO:  Handle partial elements (like in the Google Docs Apps Script examples)
  var element = elt.getElement();
  // Only valid elements that can be edited as text; skip images and
  // other non-text elements.
  if (element.editAsText()) {
    var elementText = element.asText();
    processText(elementText);
  }
}

// Parse the provided text and check for Markdown syntax
function processText(text) {
  // Found it easier to process asterisk and underscore tags seperate instead of combining into one regex
  handleBoldAsterisk(text);
  handleBoldUnderscore(text);
  handleItalicAsterisk(text);
  handleItalicUnderscore(text);
  handleStrikethrough(text);
}

function handleBoldAsterisk(text) {   
  var coords = handleEnclosedText("\\*{2}", text);
  // Only go further if new tags within Text element were found
  if (coords[0]) {
    text.setBold(coords[0], coords[1]-1, true);     
    // Recursively call command until all tag pairs have been processed
    handleBoldAsterisk(text);   
  }
}

function handleBoldUnderscore(text) {
  var coords = handleEnclosedText("_{2}", text);
  // Only go further if new tags within Text element were found
  if (coords[0]) {
    text.setBold(coords[0], coords[1]-1, true);     
    // Recursively call command until all tag pairs have been processed
    handleBoldAsterisk(text);   
  }
}

function handleItalicAsterisk(text) {     
  var coords = handleEnclosedText("\\*{1}", text);
  // Only go further if new tags within Text element were found
  if (coords[0]) {
    text.setItalic(coords[0], coords[1]-1, true);      
    // Recursively call command until all tag pairs have been processed
      handleItalicUnderscore(text);
  }
}

function handleItalicUnderscore(text) {     
  var coords = handleEnclosedText("_{1}", text);
  // Only go further if new tags within Text element were found
  if (coords[0]) {
    text.setItalic(coords[0], coords[1]-1, true);      
    // Recursively call command until all tag pairs have been processed
      handleItalicUnderscore(text);
  }
}

function handleStrikethrough(text) {     
  var coords = handleEnclosedText("~{2}", text);
  // Only go further if new tags within Text element were found
  if (coords[0]) {
    text.setSuperscript(coords[0], coords[1]-1, true);     
    // Recursively call command until all tag pairs have been processed
    handleStrikethrough(text);   
  }
}

// Handle Markdown symbols that enclose a formatted item, such as **item**
function handleEnclosedText(regex, text) {
  var coords = [];  
  var first_elt = text.findText(regex);
  // If first tag doesn't exist, don't bother with rest
  if (! first_elt) {
    return coords;
  } else {
    // Get coords of first elt
    var first_start = first_elt.getStartOffset();
    var first_end = first_elt.getEndOffsetInclusive();
    text.deleteText(first_start, first_end);
    var second_elt = text.findText(regex, first_elt);
    // Same deal with first tag, since a pair is needed
    if (! second_elt) {
      // If no second elt, then just restore original text
      text.setText(orig_text);  // TODO:  Change, since setting text removes stylizing
      return coords;
    } else {
      var second_start = second_elt.getStartOffset();
      var second_end = second_elt.getEndOffsetInclusive();
      text.deleteText(second_start, second_end);
      // first_start is now start location of enclosed text
      // second_start will be the substring ending coord
      coords.push(first_start, second_start);
    }
  }
  return coords;
}
