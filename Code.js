/**
 * AI Co-Pilot & Data Pipeline - By Ansh Jerath
 * Features: Generate, Debug, Explain, Fetch (API), Push (Webhook) + Memory
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('✨ AI Bot')
      .addItem('Open AI Co-Pilot', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('AI Co-Pilot')
      .setWidth(320); 
  SpreadsheetApp.getUi().showSidebar(html);
}

// 1. generator logic
function generateFormulaGAS(prompt) {
  const cache = CacheService.getUserCache();
  const sheetContext = getSheetContext();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API Key missing! Check Script Properties.");

  let chatHistory = [];
  const cachedHistory = cache.get('CHAT_HISTORY');
  if (cachedHistory) chatHistory = JSON.parse(cachedHistory);

  chatHistory.push({ role: "user", parts: [{ text: prompt }] });
  if (chatHistory.length > 10) chatHistory = chatHistory.slice(chatHistory.length - 10);

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=${apiKey}`;
  const systemInstruction = `You are an expert Google Sheets AI co-pilot. 
  If the user explicitly asks for a formula, output ONLY the raw formula without backticks or explanations.
  If the user asks you to modify a previous formula, look at the conversation history and output the updated raw formula.
  
  CRITICAL CONTEXT:
  ${sheetContext}`;

  const payload = { system_instruction: { parts: { text: systemInstruction } }, contents: chatHistory };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  
  if (response.getResponseCode() !== 200) {
    chatHistory.pop();
    cache.put('CHAT_HISTORY', JSON.stringify(chatHistory), 21600);
    throw new Error(json.error?.message || "The AI engine encountered an error.");
  }

  let aiResponse = json.candidates[0].content.parts[0].text.trim();
  aiResponse = aiResponse.replace(/^```[a-z]*\n?/, '').replace(/\n?```$/, '');

  chatHistory.push({ role: "model", parts: [{ text: aiResponse }] });
  cache.put('CHAT_HISTORY', JSON.stringify(chatHistory), 21600);

  return aiResponse;
}

// debugger logic

function debugFormulaGAS(brokenFormula) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API Key missing!");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=${apiKey}`;
  const debugInstruction = `You are an expert Google Sheets debugger. 
  The user will provide a broken formula. Explain why it is broken in ONE sentence, then output the corrected formula on a new line.
  CRITICAL: Do not use markdown formatting or backticks around the formula.`;

  const payload = { system_instruction: { parts: { text: debugInstruction } }, contents: [{ parts: [{ text: brokenFormula }] }] };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) throw new Error(JSON.parse(response.getContentText()).error?.message || "Error.");

  let result = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim();
  return result.replace(/^```[a-z]*\n?/, '').replace(/\n?```$/, '');
}

//explainer logic

function explainFormulaGAS(formula) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API Key missing!");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=${apiKey}`;
  const explainInstruction = `You are an expert Google Sheets instructor. 
  The user will provide a complex formula. Break down exactly how it works step-by-step in plain English. 
  Use short bullet points for readability. Be concise, clear, and easy to understand for a beginner.`;

  const payload = { system_instruction: { parts: { text: explainInstruction } }, contents: [{ parts: [{ text: formula }] }] };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) throw new Error(JSON.parse(response.getContentText()).error?.message || "Error.");

  return JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim();
}

//fetch logic

function fetchDataGAS(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API Key missing!");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=${apiKey}`; 
  const routerInstruction = `You are a strict API routing assistant for a data pipeline. 
  The user will ask to fetch data (e.g., 'crypto prices', 'random users', 'weather').
  Find a free, public API without auth requirements that fulfills this request.
  Output ONLY a raw, minified JSON object with a single key "url" pointing to the endpoint. 
  Example: {"url": "https://api.coindesk.com/v1/bpi/currentprice.json"}. Do NOT use backticks or markdown.`;

  const payload = { system_instruction: { parts: { text: routerInstruction } }, contents: [{ parts: [{ text: prompt }] }] };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  
  const aiResponse = UrlFetchApp.fetch(url, options);
  if (aiResponse.getResponseCode() !== 200) throw new Error("Failed to route API request.");
  
  let cleanJson = JSON.parse(aiResponse.getContentText()).candidates[0].content.parts[0].text.trim();
  cleanJson = cleanJson.replace(/^```[a-z]*\n?/, '').replace(/\n?```$/, ''); 
  
  let apiConfig;
  try { apiConfig = JSON.parse(cleanJson); } 
  catch (e) { throw new Error("AI failed to return a valid API route. Try being more specific."); }

  let rawData;
  try {
      const dataResponse = UrlFetchApp.fetch(apiConfig.url, {muteHttpExceptions: true});
      rawData = JSON.parse(dataResponse.getContentText());
  } catch (e) { throw new Error(`Failed to fetch data from: ${apiConfig.url}`); }

  let dataArray = Array.isArray(rawData) ? rawData : null;
  if (!dataArray) {
      for (const key in rawData) {
          if (Array.isArray(rawData[key])) { dataArray = rawData[key]; break; }
      }
  }
  
  if (!dataArray) dataArray = [rawData];
  if (dataArray.length === 0) throw new Error("API returned empty data.");

  const headers = Object.keys(dataArray[0]).filter(k => typeof dataArray[0][k] !== 'object');
  const finalData = [headers];

  dataArray.forEach(obj => {
      const row = headers.map(h => obj[h] !== undefined && obj[h] !== null ? String(obj[h]) : "");
      finalData.push(row);
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startCell = sheet.getActiveCell();
  sheet.getRange(startCell.getRow(), startCell.getColumn(), finalData.length, finalData[0].length).setValues(finalData);

  return `✅ Data Pipeline Success!\nFetched ${finalData.length - 1} rows from:\n${apiConfig.url}`;
}

//push logic

function pushDataGAS(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API Key missing!");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const rowIndex = activeCell.getRow();
  const lastCol = sheet.getLastColumn();

  if (rowIndex === 1) throw new Error("You selected the header row. Please click on a row containing actual data to push.");
  if (lastCol === 0) throw new Error("No data found in this sheet.");

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowData = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  let payloadObj = {};
  headers.forEach((header, index) => {
      if (header) payloadObj[header] = rowData[index];
  });

  const urlParams = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`; 
  const extractorInstruction = `You are a URL extraction tool. The user will ask to send data to a specific webhook or API link. 
  Extract the URL. Output ONLY the raw URL starting with http:// or https://. Do not include quotes or backticks.`;

  const aiPayload = { system_instruction: { parts: { text: extractorInstruction } }, contents: [{ parts: [{ text: prompt }] }] };
  const aiOptions = { method: 'post', contentType: 'application/json', payload: JSON.stringify(aiPayload), muteHttpExceptions: true };
  
  const aiResponse = UrlFetchApp.fetch(urlParams, aiOptions);
  if (aiResponse.getResponseCode() !== 200) throw new Error("Failed to parse the target URL from your prompt.");
  
  let targetUrl = JSON.parse(aiResponse.getContentText()).candidates[0].content.parts[0].text.trim();
  if (!targetUrl.startsWith('http')) throw new Error("Could not find a valid HTTP/HTTPS URL in your request.");

  const pushOptions = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payloadObj), muteHttpExceptions: true };

  try {
      const pushResponse = UrlFetchApp.fetch(targetUrl, pushOptions);
      const responseCode = pushResponse.getResponseCode();
      if (responseCode >= 200 && responseCode < 300) {
          return `🚀 Sync Successful!\nRow ${rowIndex} pushed to webhook.\nStatus: ${responseCode}`;
      } else {
          throw new Error(`Server rejected the payload. Status: ${responseCode}`);
      }
  } catch (error) { throw new Error(`Failed to reach webhook: ${error.message}`); }
}

//utilities

function getSheetContext() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell().getA1Notation();
  const lastCol = sheet.getLastColumn();
  let headers = [];
  if (lastCol > 0) headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  let contextString = `Active cell: ${activeCell}.\n`;
  if (headers.length > 0) {
    contextString += "Headers:\n";
    headers.forEach((h, i) => { if (h) contextString += `- Col ${String.fromCharCode(65 + i)}: "${h}"\n`; });
  }
  return contextString;
}

function insertFormulaIntoCell(formula) { SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setFormula(formula); }
function getActiveCellFormula() { return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getFormula() || ""; }
function clearChatMemory() { CacheService.getUserCache().remove('CHAT_HISTORY'); }
