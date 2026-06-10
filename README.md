## AI 4 Docs
An elegant, open-source Google Workspace Add-on style script that seamlessly embeds [Pollinations.ai](https://pollinations.ai/) directly into Google Docs and Google Sheets. It adds a premium custom top menu and a right-hand sidebar helper to read highlighted text, stream prompt context with active memory tracking, and inject responses directly back into your workflow files.

---

## 🔑 How to Get a Pollinations API Key
This add-on requires an API key for authentication. You can generate one for free in under a minute:

1. **Visit the Developer Console:** Navigate to `enter.pollinations.ai`.
2. **Sign In:** Authenticate using your GitHub account.
3. **Generate Key:** Click on **Create New API Key** (or navigate to your Developer Settings tab).
4. **Choose the Right Key Type:**
   * **Secret Key (`sk_...`):** Highly recommended for this app. This key is intended for secure backend execution (like Google Apps Script `UrlFetchApp`).
   * **Publishable Key (`pk_...`):** Designed only for client-side environments like browsers or mobile apps.
5. **Set Scope:** Choose to "Allow all models" or restrict it to specific choices (e.g., `openai`).
6. **Save Key Safely:** Copy the generated `sk_...` key instantly. It will only be shown to you once. Paste this directly into the `apiKey` variable in your `code.gs` file.

---

## 🚀 Key Features

* **Dual-App Ecosystem:** Operates instantly out of a single script file inside both Google Docs and Google Sheets.
* **Contextual Chat Sidebar:** Highlight text or structured cells, pull them into the scratchpad, and submit requests.
* **Direct Canvas Infiltration:** Tap a single button to insert generated answers exactly where your text blinking cursor or active cell layout is targeted.
* **Rolling Chat Memory:** Automatically holds a cache window of the last 10 conversational interaction turns safely saved inside Google's private user properties engine.
* **Zero System Bloat:** Standalone client layout built cleanly with vanillajs scripting, optimized CSS stylesheets, and no third-party framework overhead.

---

## 🛠️ Code Architecture
The implementation requires two structural layout assets added directly to your application container: `code.gs` and `sidebar.html`.

### 1. Backend Engine (code.gs)
```javascript
/**
 * Helper function to retrieve the active UI instance for either Docs or Sheets.
 */
function getGenericUi() {
  if (typeof DocumentApp !== 'undefined' && DocumentApp.getActiveDocument()) {
    return DocumentApp.getUi();
  } else if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
    return SpreadsheetApp.getUi();
  }
  return null;
}

/**
 * Helper function to handle toast notifications safely across environments.
 */
function showToastNotification(message, title) {
  if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

/**
 * Clears stored chat history from User Properties.
 */
function clearChatHistory() {
  const ui = getGenericUi();
  PropertiesService.getUserProperties().deleteProperty('AI_CHAT_HISTORY');
  if (ui) ui.alert('System', 'Chat memory cleared successfully.', ui.ButtonSet.OK);
}

/**
 * Automatically creates a custom menu in the top bar when the Document or Sheet opens.
 * Adapts dynamically to either Google Docs or Google Sheets environment.
 */
function onOpen() {
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
 */
function showSidebar() {
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
 */
function getSelectedTextFromDoc() {
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
 */
function insertTextAtCursor(textToInsert) {
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
        if (elements.length > 0 && elements[0].getElement().editAsText) {
          elements[0].getElement().editAsText().setText(textToInsert);
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
 */
function runAiFromTopBar() {
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
 */
function callPollinatorAPI(promptText) {
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
      if (json.choices && json.choices[0] && json.choices[0].message) {
        const aiResponse = json.choices[0].message.content;
        
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
```
