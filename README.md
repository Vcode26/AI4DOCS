## AI 4 Docs
An elegant, open-source Google Workspace Add-on style script that seamlessly embeds [Pollinations.ai](https://pollinations.ai/) directly into Google Docs and Google Sheets. It adds a premium custom top menu and a right-hand sidebar helper to read highlighted text, stream prompt context with active memory tracking, and inject responses directly back into your workflow files.
------------------------------
## 🔑 How to Get a Pollinations API Key
This add-on requires an API key for authentication. You can generate one for free in under a minute: [1] 

   1. Visit the Developer Console: Navigate to enter.pollinations.ai. [2, 3] 
   2. Sign In: Authenticate using your GitHub account. [4, 5] 
   3. Generate Key: Click on Create New API Key (or navigate to your Developer Settings tab). [4, 6] 
   4. Choose the Right Key Type:
   * Secret Key (sk_...): Highly recommended for this app. This key is intended for secure backend execution (like Google Apps Script UrlFetchApp).
      * Publishable Key (pk_...): Designed only for client-side environments like browsers or mobile apps. [2, 7, 8] 
   5. Set Scope: Choose to "Allow all models" or restrict it to specific choices (e.g., openai). [9] 
   6. Save Key Safely: Copy the generated sk_... key instantly. It will only be shown to you once. Paste this directly into the apiKey variable in your code.gs file. [6, 10] 

------------------------------
## 🚀 Key Features

* Dual-App Ecosystem: Operates instantly out of a single script file inside both Google Docs and Google Sheets.
* Contextual Chat Sidebar: Highlight text or structured cells, pull them into the scratchpad, and submit requests.
* Direct Canvas Infiltration: Tap a single button to insert generated answers exactly where your text blinking cursor or active cell layout is targeted.
* Rolling Chat Memory: Automatically holds a cache window of the last 10 conversational interaction turns safely saved inside Google's private user properties engine.
* Zero System Bloat: Standalone client layout built cleanly with vanillajs scripting, optimized CSS stylesheets, and no third-party framework overhead.

------------------------------
## 🛠️ Code Architecture
The implementation requires two structural layout assets added directly to your application container: code.gs and sidebar.html.
## 1. Backend Engine (code.gs)

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
  const url = '

https://pollinations.ai

'; 

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
        if (errorJson.error) errorMsg = errorJson.error.message;
      } catch(e) {}
      
      return "API Error (" + responseCode + "): " + errorMsg;
    }
  } catch (e) {
    return "Connection Exception: " + e.toString();
  }
}
/**
 * Resets the memory cache and starts a brand new conversation topic thread.
 */function clearChatHistory() {
  PropertiesService.getUserProperties().deleteProperty('AI_CHAT_HISTORY');
  const ui = getGenericUi();
  if (ui) ui.alert('Success', 'Conversation memory has been cleared.', ui.ButtonSet.OK);
}
/**
 * Helper: Safely fetches the active application UI context (Docs or Sheets).
 */function getGenericUi() {
  try {
    return DocumentApp.getUi();
  } catch (e) {
    try {
      return SpreadsheetApp.getUi();
    } catch (err) {
      return null;
    }
  }
}
/**
 * Helper: Displays a subtle toast pop-up notification if inside Google Sheets.
 */function showToastNotification(msg, title) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, 4);
  } catch(e) {
    // Fail silently if running inside Google Docs
  }
}

## 2. Interface Layer (sidebar.html)

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
      .status-text { color: #5f6368; font-style: italic; font-size: 13px; margin-top: 4px; display: block; text-align: center; }
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
    
    <div id="statusIndicator" class="status-text"></div>

    <div id="outputBox">
      <label>AI Output Response</label>
      <div id="outputContent"></div>
    </div>

    <script>
      let savedResponse = "";

      function fetchSelectedText() {
        document.getElementById('statusIndicator').innerText = "Reading selection...";
        google.script.run
          .withSuccessHandler(function(selectedText) {
            if (selectedText) {
              document.getElementById('prompt').value = selectedText;
              document.getElementById('statusIndicator').innerText = "Text imported!";
            } else {
              document.getElementById('statusIndicator').innerText = "No selection or cell found.";
            }
          })
          .withFailureHandler(function(err) {
            document.getElementById('statusIndicator').innerText = "Error: " + err.message;
          })
          .getSelectedTextFromDoc();
      }

      function submitToAI() {
        const promptText = document.getElementById('prompt').value;
        if (!promptText.trim()) {
          alert("Please type something or import highlighted text first!");
          return;
        }

        document.getElementById('btnSubmit').disabled = true;
        document.getElementById('btnGrab').disabled = true;
        document.getElementById('btnInsert').disabled = true;
        document.getElementById('statusIndicator').innerText = "AI is thinking...";
        document.getElementById('outputBox').style.display = 'none';

        google.script.run
          .withSuccessHandler(function(aiResponse) {
            savedResponse = aiResponse;
            document.getElementById('outputContent').innerText = aiResponse;
            document.getElementById('outputBox').style.display = 'block';
            document.getElementById('statusIndicator').innerText = "Generation complete!";
            
            document.getElementById('btnSubmit').disabled = false;
            document.getElementById('btnGrab').disabled = false;
            document.getElementById('btnInsert').disabled = false;
          })
          .withFailureHandler(function(err) {
            document.getElementById('outputContent').innerText = "Execution failed: " + err.message;
            document.getElementById('outputBox').style.display = 'block';
            document.getElementById('statusIndicator').innerText = "Error encountered.";
            document.getElementById('btnSubmit').disabled = false;
            document.getElementById('btnGrab').disabled = false;
          })
          .callPollinatorAPI(promptText);
      }

      function insertToDoc() {
        if (!savedResponse) return;
        document.getElementById('statusIndicator').innerText = "Inserting text...";
        
        google.script.run
          .withSuccessHandler(function() {
            document.getElementById('statusIndicator').innerText = "Inserted successfully!";
          })
          .withFailureHandler(function(err) {
            document.getElementById('statusIndicator').innerText = "Insertion error: " + err.message;
          })
          .insertTextAtCursor(savedResponse);
      }
    </script>
  </body>
</html>

------------------------------
## 💾 Installation & Setup

   1. Open your target Google Doc or Google Sheet.
   2. Click Extensions > Apps Script in the top navigation bar.
   3. In the left panel, rename the default file to code.gs and paste the backend code block inside.
   4. Click the + icon next to Files and add an HTML asset named sidebar. Paste the interface layer block inside it.
   5. Paste your sk_... key from enter.pollinations.ai directly over the 'YOUR_POLLINATIONS_API_KEY' placeholder inside code.gs.
   6. Click the disk Save icon at the top.
   7. Refresh your browser tab hosting the Doc or Sheet. Within 5 seconds, the 🤖 AI Copilot menu will render next to the "Help" menu option. [11] 

------------------------------
## 🔒 Security & OAuth Permissions
When executing this script for the first time, Google Workspace requires security scope access authorizations. The add-on strictly limits its operation requirements to these three native parameters:

* https://googleapis.com — Targets highlighted canvas blocks or text insertions under active focus.
* https://googleapis.com — Pulls data arrays from spreadsheet cell blocks and modifies values.
* https://googleapis.com — Routes queries securely over external network tunnels directly to the server endpoints.

------------------------------
Let me know if you would like me to add anything else to the documentation, such as a model list breakdown or an section explaining how to add an custom System Prompt/Persona!

[1] [https://pollinations.ai](https://pollinations.ai/play)
[2] [https://github.com](https://github.com/pollinations/pollinations/blob/main/APIDOCS.md)
[3] https://enter.pollinations.ai
[4] [https://gen.pollinations.ai](https://gen.pollinations.ai/docs)
[5] [https://www.youtube.com](https://www.youtube.com/watch?v=7J7YowrN4aY)
[6] [https://docs.pollination.solutions](https://docs.pollination.solutions/user-manual/developers/create-an-api-key)
[7] [https://github.com](https://github.com/isaacgounton/n8n-nodes-pollinations)
[8] [https://help.openai.com](https://help.openai.com/en/articles/4936850-where-do-i-find-my-openai-api-key)
[9] [https://github.com](https://github.com/pollinations/pollinations)
[10] [https://www.heyhelp.ai](https://www.heyhelp.ai/ai-providers-and-api-keys/)
[11] [https://pollinations.ai](https://pollinations.ai/apps)
