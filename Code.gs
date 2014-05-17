/*
Copyright 2014-2015 Matt L.
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

   1. Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.

   2. Redistributions in binary form must reproduce the above
      copyright notice, this list of conditions and the following
      disclaimer in the documentation and/or other materials provided
      with the distribution.

THIS SOFTWARE IS PROVIDED BY THE AUTHORS ``AS IS'' AND ANY EXPRESS OR
IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE
OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

The views and conclusions contained in the software and documentation
are those of the authors and should not be interpreted as representing
official policies, either expressed or implied, of the authors.
*/

function onInstall() {
  onOpen();
}

/**
 * The onOpen function runs automatically when the Google Docs document is
 * opened. Use it to add custom menus to Google Docs that allow the user to run
 * custom scripts. For more information, please consult the following two
 * resources.
 *
 * Extending Google Docs developer guide:
 *     https://developers.google.com/apps-script/guides/docs
 *
 * Document service reference documentation:
 *     https://developers.google.com/apps-script/reference/document/
 */
function onOpen() {
  // Add a menu with some items, some separators, and a sub-menu.
  
  var menu = DocumentApp.getUi().createAddonMenu();
  
  menu.addItem('Encrypt', 'encryptSelectedText');
  menu.addItem('Decrypt', 'decryptSelectedText');
  
  menu.addToUi();
}

/**
 * Shows an input box in the Google Docs editor.
 */
function showPromptForPassphrase() {
  // Displays a dialog box with "OK" and "Cancel" buttons, as well as a text box
  // allowing the user to enter a response to a question.
  var result = DocumentApp.getUi().prompt('Input password:', DocumentApp.getUi().ButtonSet.OK_CANCEL);

  // Process the user's response:
  if (result.getSelectedButton() == DocumentApp.getUi().Button.OK) {
    // The user clicked the "OK" button.
    return result.getResponseText();
  } 
  
  return null;
}

function replaceText(text, replacement, startOffset, endOffsetInclusive) {
  if (startOffset < 0) {
    text.removeFromParent();
    text.setText(replacement);
  } else if (endOffsetInclusive > startOffset) {
    text.deleteText(startOffset, endOffsetInclusive);
    text.insertText(startOffset, replacement);
  }
}

function encryptOrDecryptText(isEncrypt) {
  // Try to get the current selection in the document. If this fails (e.g.,
  // because nothing is selected), show an alert and exit the function.
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    DocumentApp.getUi().alert('Cannot find a selection in the document.');
    return;
  }

  var password = showPromptForPassphrase();
  if (null != password) {
    var selectedElements = selection.getRangeElements();
    for (var i = 0; i < selectedElements.length; ++i) {
      var selectedElement = selectedElements[i];
      
      // Only modify elements that can be edited as text; skip images and other
      // non-text elements.
      var text = selectedElement.getElement().editAsText();
      
      // Change the background color of the selected part of the element, or the
      // full element if it's completely selected.
      
      var replacement, selectedTextString;
      var startOffset = -1;
      var endOffsetInclusive = -1;
      
      if (selectedElement.isPartial()) {
        startOffset = selectedElement.getStartOffset();
        endOffsetInclusive = selectedElement.getEndOffsetInclusive();
        selectedTextString = text.getText().substring(startOffset, endOffsetInclusive + 1);
        
      } else {
        selectedTextString = text.getText();  
      }
      
      if (isEncrypt) {
        replacement = sjcl.encrypt(password, selectedTextString);
      } else {
        replacement = sjcl.decrypt(password, selectedTextString);
      }
      
      replaceText(text, replacement, startOffset, endOffsetInclusive);
    }
  }
}

/**
 * Encrypt the selected text
 */
function encryptSelectedText() {
  encryptOrDecryptText(true /*isEncrypt*/);
}

/**
 * Encrypt the selected text
 */
function decryptSelectedText() {
  encryptOrDecryptText(false /*isEncrypt*/);
}


