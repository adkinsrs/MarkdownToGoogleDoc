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
    var elements = selection.getSelectedElements();
    for (var i = 0; i < elements.length; i++) {    
      processElement(elements[i]);
    }
  } else {
    throw 'Please select some text.';
  }  
}

// Process current element
function processElement(elt) {
  var text;
  //TODO:  Handle partial elements (like in the Google Docs Apps Script examples)
  var element = elt.getElement();
  // Only valid elements that can be edited as text; skip images and
  // other non-text elements.
  if (element.editAsText) {
    var elementText = element.asText();
    processText(elementText);
  }
}

// Parse the provided text and check for Markdown syntax
function processText(text) {
  // Found it easier to process asterisk and underscore tags seperate instead of combining into one regex
  handleBoldAsterisk(text);
  //handleBoldUnderscore(text);
  //handleItalicAsterisk(text);
  handleItalicUnderscore(text);
  //handleStrikethrough(text);
}

function handleBoldAsterisk(text) {
  var coords = handleEnclosedText("\\*{2}", text);
  // Only go further if new tags within Text element were found
  if (coords[0] !== '') {
    var end_string = formNewTextString(coords, text);
    // Recursively call command until all tag pairs have been processed
    if (end_string.length > 0) {
      text.appendText(end_string);
      handleBoldAsterisk(text);
    }
    // Bold after the recursive call is bolded
    text.setBold(coords[0], coords[1]-1, true);    
  }
}

function handleBoldUnderscore(text) {
  var coords = handleEnclosedText("_{2}", text);
  // Only go further if new tags within Text element were found
  if (coords[0] !== '') {
    var end_string = formNewTextString(coords, text);
    // Recursively call command until all tag pairs have been processed
    if (end_string.length > 0) {
      text.appendText(end_string);
      handleBoldUnderscore(text);
    }
    // Bold after the recursive call is bolded
    text.setBold(coords[0], coords[1]-1, true);    
  }
}

function handleItalicAsterisk(text) {
  var coords = handleEnclosedText("\\*{1}", text);
  // Only go further if new tags within Text element were found
  if (coords[0] !== '') {
    var end_string = formNewTextString(coords, text);
    // Recursively call command until all tag pairs have been processed
    if (end_string.length > 0) {
      text.appendText(end_string);
      handleItalicAsterisk(text);
    }
    // Italicize after the recursive call is italicized
    text.setItalic(coords[0], coords[1]-1, true);    
  }
}

function handleItalicUnderscore(text) {
  var coords = handleEnclosedText("_{1}", text);
  // Only go further if new tags within Text element were found
  if (coords[0] !== '') {
    var end_string = formNewTextString(coords, text); 
    // Recursively call command until all tag pairs have been processed
    if (end_string.length > 0) {
      text.appendText(end_string);
      handleItalicUnderscore(text);
    }
    // Italicize after the recursive call is italicized
    text.setItalic(coords[0], coords[1]-1, true);    
  }
}

function handleStrikethrough(text) {
  var coords = handleEnclosedText("~{2}", text);
  // Only go further if new tags within Text element were found
  if (coords[0] !== '') {
    var end_string = formNewTextString(coords, text); 
    // Recursively call command until all tag pairs have been processed
    if (end_string.length > 0) {
      text.appendText(end_string);
      handleStrikethrough(text);
    }
    // Strikethrough after the recursive call is done
    text.setStrikethrough(coords[0], coords[1]-1, true);    
  }
}

// Handle Markdown symbols that enclose a formatted item, such as **item**
function handleEnclosedText(regex, text) {
  var coords = [];
  var orig_text = text.getText();
  var first_elt = text.findText(regex);
  // If first tag doesn't exist, don't bother with rest
  if (! first_elt) {
    coords.push('', '');
  } else {
    // Get coords of first elt
    var first_start = first_elt.getStartOffset();
    var first_end = first_elt.getEndOffsetInclusive();
    text.deleteText(first_start, first_end);
    var second_elt = text.findText(regex, first_elt);
    // Same deal with first tag, since a pair is needed
    if (! second_elt) {
      // If no second elt, then just restore original text
      text.setText(orig_text);
      coords.push('', '');
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

// Break into string pieces not containing the Markdown tags and reform
function formNewTextString(coords, text) {
  var beginning = text.getText().substring(0,coords[0]);
  var enclosed = text.getText().substring(coords[0], coords[1]);
  var end = text.getText().substring(coords[1]);  
  // Process end text in next recursive go-around
  var new_text = beginning + enclosed;
  text.setText(new_text);
  return end;
}
