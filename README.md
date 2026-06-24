<div align="center">
  
  # Exa Websets Sheets Sync

  [![Watch Video on X](https://img.shields.io/badge/Watch_Video_on_X-%23000000.svg?style=for-the-badge&logo=X&logoColor=white)](https://x.com/knokvik/status/2069697486906921120)

  **A Google Sheets Add-on that continuously syncs Exa Websets directly into a spreadsheet.**

<img width="2938" height="1652" alt="Screenshot 2026-06-24 at 12 33 35 PM" src="https://github.com/user-attachments/assets/b6e00706-c7cb-468b-b9d1-5a2c6e691411" />
<br>

  [Installation](#installation) • [How it Works](#how-it-works) • [Architecture](#architecture)
</div>

<br>

## 📌 The Problem
Currently, exporting data from Exa Websets into a spreadsheet requires manual CSV exports. There is no easy way to establish a continuous sync so that as new items are discovered and enriched by Exa, they automatically populate in a Google Sheet for your team's workflow.

## 🚀 The Solution
This project bridges the gap by providing a native **Google Sheets Add-on** that connects directly to the Exa Websets REST API. 

**Features:**
- **One-Click Sync**: Pulls all items from a Webset directly into Google Sheets.
- **Smart Deduplication**: Uses Webset item IDs to cleanly update existing rows with new enrichment data and append only newly discovered rows.
- **Auto-Refresh**: Sets up a background time-driven trigger to continuously poll the Webset and keep the sheet updated on a schedule.
- **Demo Mode**: Includes a fully-functional Demo Mode with 25 realistic mock AI startups to let users test the workflow without needing an Exa Pro API key.
- **Secure Architecture**: API keys are securely stored in Apps Script's `PropertiesService`, meaning they are never exposed in spreadsheet cells or hardcoded.

## 🛠️ Installation

### Option 1: Using Clasp (For Developers)
1. Clone this repository.
2. Install [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
3. Login to clasp: `clasp login`
4. Create a new Google Sheet and bind a script: `clasp create --type sheets`
5. Push the code: `clasp push`
6. Open your Google Sheet, and the "Websets Sync" menu will appear!

### Option 2: Copying a Template
*(Provide a link to a view-only Google Sheet that users can File > Make a Copy of)*

## 🎥 Demo
*(Link to your Loom video here)*

## 🏗️ Architecture (Phase 1)
This MVP uses **Google Apps Script + TypeScript** to directly orchestrate the synchronization. No external backend server is required.
```text
Google Sheet
  └─ Apps Script (clasp)
       ├─ Sidebar UI (HTML service)
       ├─ Sync Engine (UrlFetchApp -> Exa REST API)
       └─ Auto-Refresh Trigger (every 15 min)
```

## 🔐 Security & Config Handling
To ensure enterprise-grade security:
- The Exa API Key is never written to a sheet.
- It is stored securely in the user's isolated `PropertiesService.getUserProperties()`.
- The UI handles errors gracefully, ensuring bad keys or missing permissions are clearly communicated.

---
*Built as a proof-of-work to solve a core workflow pain point for Websets users.*
