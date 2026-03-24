# 🚀 AI-Powered Google Sheets Data Pipeline 
Built by Ansh Jerath : Architecting cloud solutions, machine learning integrations, and robust data pipelines.

👉 **[Watch the 2-Minute Loom Demo Here](https://www.loom.com/share/b197d926990a454a9954602e23095320)** 👈

A serverless, natively integrated Google Workspace Add-on that transforms Google Sheets from a static grid into a dynamic, two-way data pipeline using LLM-based API routing.

## 🧠 The Architecture Problem
Building robust data connectors inside Google Workspace is notoriously difficult due to strict execution timeouts (6 minutes), UI sandboxing restrictions (IFRAMEs), and the complexities of flattening deeply nested JSON responses into a clean 2D grid. 

This project solves those friction points by using a decoupled architecture: an LLM acts as the intelligent data router, while Google Apps Script (V8) acts as the execution engine for secure, native spreadsheet manipulation.

## 🛠️ Core Data Pipeline Features

This add-on features 5 distinct modes, shifting from a simple spreadsheet assistant to a full-fledged data connector:

* **📡 Fetch (Data Ingestion & Flattening):** Uses natural language to dynamically identify public API endpoints.
  * Fetches the JSON payload via `UrlFetchApp`.
  * *Crucial Step:* Automatically traverses the payload, isolates the primary data array, flattens deeply nested objects into column headers, and injects the clean 2D matrix directly into the active sheet.
* **📤 Push (Two-Way Webhook Syncing):** Select any row of data, and the engine automatically maps the row values back to their respective column headers.
  * Formats a clean JSON payload.
  * Extracts the target webhook/API URL directly from the user's natural language prompt.
  * Fires a POST request to sync the data back to an external server/CRM.

## 🤖 Spreadsheet Intelligence

* **⚙️ Context-Aware Generation:** The script reads the spreadsheet's active headers and dimensions to provide highly specific formula recommendations. It retains conversational memory via Google's `CacheService` (avoiding the need for an external database).
* **🐛 Debugger:** A 1-click syntax fixer that parses `#ERROR!` and `#REF!` outputs, explains the failure, and safely auto-injects the corrected formula back into the active cell using `google.script.run`.
* **📖 Explainer:** Reverse-engineers complex, nested spreadsheet logic into human-readable, bulleted documentation for non-technical users.

## 🏗️ Tech Stack
* **Frontend:** Vanilla JS/HTML/CSS (Built specifically to bypass Google's IFRAME sandboxing restrictions and maintain a lightweight footprint).
* **Backend:** Google Apps Script (Node-like V8 runtime).
* **Intelligence:** Gemini 2.5 Pro (REST API integration).
* **State Management:** Google Apps Script `CacheService`.

## ⚙️ Local Development & Setup
If you want to run this locally or push it to your own Google Workspace environment, you can run this entire setup sequence in your terminal:

```bash
# 1. Clone this repository and move into the directory
git clone [https://github.com/darkhorse0204/sheets-ai-pipeline.git](https://github.com/darkhorse0204/sheets-ai-pipeline.git)
cd sheets-ai-pipeline

# 2. Install Google's Apps Script CLI tool
npm install -g @google/clasp

# 3. Authenticate with Google and push the code to a new project
clasp login
clasp create --type sheets --title "AI Pipeline"
clasp push

Final Step: Open the Apps Script editor in your browser, go to Project Settings, and add your Gemini API key under the Script Properties key GEMINI_API_KEY.
