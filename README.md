
## AI 4 Docs
An elegant, open-source Google Workspace Add-on script that seamlessly embeds Pollinations.ai directly into Google Docs and Google Sheets. It adds a premium custom top menu and a right-hand sidebar helper to read highlighted text, stream prompt context with active memory tracking, and inject responses directly back into your workflow files.
------------------------------
## 🔑 How to Get a Pollinations API Key
This add-on requires an API key for authentication. You can generate one for free in under a minute:

   1. Visit the Developer Console: Navigate to enter.pollinations.ai.
   2. Sign In: Authenticate using your GitHub account.
   3. Generate Key: Click on Create New API Key (or navigate to your Developer Settings tab).
   4. Choose the Right Key Type:
   * Secret Key (sk_...): Highly recommended for this app. This key is intended for secure backend execution (like Google Apps Script UrlFetchApp).
      * Publishable Key (pk_...): Designed only for client-side environments like browsers or mobile apps.
   5. Set Scope: Choose to "Allow all models" or restrict it to specific choices (e.g., openai).
   6. Save Key Safely: Copy the generated sk_... key instantly. It will only be shown to you once. Paste this directly into the apiKey variable in your code.gs file.

------------------------------
## 🚀 Key Features

* Dual-App Ecosystem: Operates instantly out of a single script file inside both Google Docs and Google Sheets.
* Contextual Chat Sidebar: Highlight text or structured cells, pull them into the scratchpad, and submit requests.
* Direct Canvas Infiltration: Tap a single button to insert generated answers exactly where your text blinking cursor or active cell layout is targeted.
* Rolling Chat Memory: Automatically holds a cache window of the last 10 conversational interaction turns safely saved inside Google's private user properties engine.
* Zero System Bloat: Standalone client layout built cleanly with Vanilla JS scripting, optimized CSS stylesheets, and no third-party framework overhead.

------------------------------
## 🛠️ Code Architecture
The implementation requires two structural layout assets added directly to your application container: code.gs and sidebar.html.
## 1. Backend Engine (code.gs)
Create a script file named code.gs and paste the following code:

/**
 * Helper function to retrieve the active UI instance for either Docs or Sheets.
 */function getGenericUi() {
  if (typeof DocumentApp !== 'undefined' && DocumentApp.getActiveDocument()) {
    return DocumentApp.getUi();
  } else if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
    return SpreadsheetApp.getUi();
  }
  return null;
}
/**
 * Helper function to handle toast notifications safely across environments.
 */function showToastNotification(message, title) {
  if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}
/**
 * Clears stored chat history from User Properties.
 */function clearChatHistory() {
  const ui = getGenericUi();
  PropertiesService.getUserProperties().deleteProperty('AI_CHAT_HISTORY');
  if (ui) ui.alert('System', 'Chat memory cleared successfully.', ui.ButtonSet.OK);
}
/**
 * Automatically creates a custom menu in the top bar when the Document or Sheet opens.
 * Adapts dynamically to either Google Docs or Google Sheets environment.
 */function onOpen() {
  const ui = getGenericUi();
  if (!ui) return;
  
  ui.createMenu('🤖 AI Copilot')
    .addItem('👉 Open AI Sidebar', 'showSidebar')
    .addItem('Ask Pollinations AI (Pop-up)', 'runAiFromTopBar')
    .addSeparator()
    .addItem('🧹 Clear Chat History', 'clearChatHistory')
    .addToUi();
}
/**
 * Renders the sidebar.html file directly within the application's right gutter panel.
 */function showSidebar() {
  const ui = getGenericUi();
  if (!ui) return;
  
  const htmlOutput = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('AI 4 Docs Copilot')
      .setWidth(300);
      
  if (typeof DocumentApp !== 'undefined' && DocumentApp.getActiveDocument()) {
    DocumentApp.getUi().showSidebar(htmlOutput);
  } else if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  }
}
/**
 * Sidebar Hook: Reads highlighted text from the user's document or active spreadsheet cell.
 * @return {string} The collected highlight string, or an empty string.
 */function getSelectedTextFromDoc() {
  try {
    const doc = DocumentApp.getActiveDocument();
    if (doc) {
      const selection = doc.getSelection();
      if (selection) {
        const elements = selection.getRangeElements();
        let text = [];
        for (let i = 0; i < elements.length; i++) {
          const element = elements[i].getElement();
          if (element.editAsText) {
            const textObj = element.editAsText();
            if (elements[i].isPartial()) {
              text.push(textObj.getText().substring(elements[i].getStartOffset(), elements[i].getEndOffsetInclusive() + 1));
            } else {
              text.push(textObj.getText());
            }
          }
        }
        return text.join('\n').trim();
      }
    }
  } catch (e) {}

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    if (sheet) {
      return sheet.getActiveCell().getValue().toString().trim();
    }
  } catch (e) {}
  
  return "";
}
/**
 * Sidebar Hook: Injects final text responses back down into the active file cursor coordinates.
 * @param {string} textToInsert - Raw string received from the Pollinations text payload.
 */function insertTextAtCursor(textToInsert) {
  if (!textToInsert) return;
  
  try {
    const doc = DocumentApp.getActiveDocument();
    if (doc) {
      const cursor = doc.getCursor();
      if (cursor) {
        cursor.insertText(textToInsert);
        return;
      }
      const selection = doc.getSelection();
      if (selection) {
        const elements = selection.getRangeElements();
        if (elements.length > 0 && elements.getElement().editAsText) {
          elements.getElement().editAsText().setText(textToInsert);
          return;
        }
      }
      doc.getBody().appendParagraph(textToInsert);
      return;
    }
  } catch (e) {}

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    if (sheet) {
      sheet.getActiveCell().setValue(textToInsert);
      return;
    }
  } catch (e) {}
}
/**
 * Triggered from the top bar menu. Handles user input, API execution, and results display.
 */function runAiFromTopBar() {
  const ui = getGenericUi();
  if (!ui) return;
  
  const response = ui.prompt('Ask Pollinations AI', 'Enter your prompt below:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const userPrompt = response.getResponseText();
    
    if (!userPrompt || !userPrompt.trim()) {
      ui.alert('Error', 'Prompt cannot be empty.', ui.ButtonSet.OK);
      return;
    }
    
    showToastNotification('Fetching response from AI...', 'Working...');
    const aiResult = callPollinatorAPI(userPrompt);
    
    showToastNotification('Done!', 'AI Response Received');
    ui.alert('AI Response', aiResult, ui.ButtonSet.OK);
  }
}
/**
 * Sends text to the Pollinations.ai API and maintains ongoing conversation context.
 */function callPollinatorAPI(promptText) {
  const apiKey = 'YOUR_POLLINATIONS_API_KEY'; 
  const url = 'https://pollinations.ai'; 

  const userProperties = PropertiesService.getUserProperties();
  let chatHistoryString = userProperties.getProperty('AI_CHAT_HISTORY');
  let messages = [];

  if (chatHistoryString) {
    try {
      messages = JSON.parse(chatHistoryString);
    } catch(e) {
      messages = [];
    }
  }

  messages.push({ "role": "user", "content": promptText });

  if (messages.length > 10) {
    messages = messages.slice(-10);
  }

  const payload = {
    "model": "openai", 
    "messages": messages, 
    "temperature": 0.7
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const json = JSON.parse(responseText);
      if (json.choices && json.choices && json.choices.message) {
        const aiResponse = json.choices.message.content;
        
        messages.push({ "role": "assistant", "content": aiResponse });
        userProperties.setProperty('AI_CHAT_HISTORY', JSON.stringify(messages));
        
        return aiResponse;
      }
      return "Error: Received unexpected JSON schema structure from API.";
    } else {
      let errorMsg = responseText;
      try {
        const errorJson = JSON.parse(responseText);
        if (errorJson.error) errorMsg = errorJson.error.message || JSON.stringify(errorJson.error);
      } catch(e) {}
      return "API Error (" + responseCode + "): " + errorMsg;
    }
  } catch(e) {
    return "Network or runtime execution error: " + e.toString();
  }
}

## 2. Frontend Interface (sidebar.html)
Create an HTML file named sidebar.html and paste the following code: [2] 

<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; padding: 16px; margin: 0; background-color: #fafafa; font-size: 14px; color: #202124; }
      h3 { margin-top: 0; color: #1a73e8; font-weight: 500; font-size: 16px; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px; }
      label { font-weight: bold; display: block; margin-bottom: 6px; font-size: 12px; color: #5f6368; text-transform: uppercase; }
      textarea { width: 100%; height: 110px; box-sizing: border-box; margin-bottom: 12px; padding: 10px; border: 1px solid #dadce0; border-radius: 4px; font-family: inherit; resize: vertical; font-size: 13px; }
      textarea:focus { border-color: #1a73e8; outline: none; }
      .btn-group { display: flex; flex-direction: column; gap: 8px; margin-bottom: 16px; }
      button { width: 100%; padding: 10px 16px; font-size: 13px; font-weight: 500; border-radius: 4px; border: 1px solid transparent; cursor: pointer; transition: background-color 0.2s; }
      .primary-btn { background-color: #1a73e8; color: white; }
      .primary-btn:hover { background-color: #1557b0; }
      .secondary-btn { background-color: #ffffff; color: #3c4043; border: 1px solid #dadce0; }
      .secondary-btn:hover { background-color: #f8f9fa; }
      .success-btn { background-color: #1e7e34; color: white; }
      .success-btn:hover { background-color: #155d23; }
      button:disabled { background-color: #f1f3f4 !important; color: #3c4043 !important; border-color: #e0e0e0 !important; cursor: not-allowed; opacity: 0.6; }
      #outputBox { display: none; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 4px; padding: 12px; margin-top: 12px; }
      #outputContent { white-space: pre-wrap; word-break: break-word; font-size: 13px; line-height: 1.5; color: #3c4043; max-height: 220px; overflow-y: auto; background: #f8f9fa; padding: 8px; border-radius: 4px; }
      .spinner { color: #5f6368; font-style: italic; font-size: 13px; display: inline-flex; align-items: center; gap: 6px; }
    </style>
  </head>
  <body>
    <h3>Pollinations Assistant</h3>
    
    <label for="prompt">Prompt / Text Input</label>
    <textarea id="prompt" placeholder="Write a summary, compose a draft, or rewrite selected text..."></textarea>
    
    <div class="btn-group">
      <button class="secondary-btn" id="btnGrab" onclick="fetchSelectedText()">✨ Import Highlighted Text</button>
      <button class="primary-btn" id="btnSubmit" onclick="submitToAI()">🚀 Generate Response</button>
      <button class="success-btn" id="btnInsert" onclick="insertToDoc()" disabled>📥 Insert at Cursor</button>
    </div>
    
    <div id="outputBox">
      <label>AI Output Response</label>
      <div id="outputContent"></div>
    </div>

    <script>
      let savedResponse = "";

      function fetchSelectedText() {
        const btnGrab = document.getElementById('btnGrab');
        btnGrab.disabled = true;
        
        google.script.run
          .withSuccessHandler(function(selectedText) {
            btnGrab.disabled = false;
            if (selectedText) {
              document.getElementById('prompt').value = selectedText;
            } else {
              alert('Please highlight a block of text in your document or cell first before running this command.');
            }
          })
          .withFailureHandler(function(err) {
            btnGrab.disabled = false;
            alert('Error pulling selection: ' + err.message);
          })
          .getSelectedTextFromDoc();
      }

      function submitToAI() {
        const promptInput = document.getElementById('prompt').value.trim();
        if (!promptInput) return alert('The prompt text window cannot be empty.');
        
        const btnSubmit = document.getElementById('btnSubmit');
        const btnInsert = document.getElementById('btnInsert');
        const outputBox = document.getElementById('outputBox');
        const outputContent = document.getElementById('outputContent');
        
        btnSubmit.disabled = true;
        btnSubmit.innerHTML = '<span class="spinner">Thinking...</span>';
        outputBox.style.display = 'block';
        outputContent.innerHTML = '<em>Connecting to Pollinations.ai instance...</em>';
        btnInsert.disabled = true;
        
        google.script.run
          .withSuccessHandler(function(apiResult) {
            savedResponse = apiResult;
            outputContent.innerText = apiResult;
            btnSubmit.disabled = false;
            btnSubmit.innerText = '🚀 Generate Response';
            
            if (!apiResult.startsWith("Error") && !apiResult.startsWith("API Error") && !apiResult.startsWith("Network")) {
              btnInsert.disabled = false;
            }
          })
          .withFailureHandler(function(err) {
            outputContent.innerText = 'Runtime Script Error: ' + err.message;
            btnSubmit.disabled = false;
            btnSubmit.innerText = '🚀 Generate Response';
          })
          .callPollinatorAPI(promptInput);
      }

      // Passes the output payload straight back to the cursor inside the canvas layout
      function insertToDoc() {
        if (!savedResponse) return;
        const btnInsert = document.getElementById('btnInsert');
        btnInsert.disabled = true;
        
        google.script.run
          .withSuccessHandler(function() {
            btnInsert.disabled = false;
          })
          .withFailureHandler(function(err) {
            btnInsert.disabled = false;
            alert('Failed to insert text: ' + err.message);
          })
          .insertTextAtCursor(savedResponse);
      }
    </script>
  </body>
</html>

------------------------------
