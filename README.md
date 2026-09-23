# Exa for Google Sheets

Use Exa inside Google Sheets to research the web, generate tables, and fill missing data.

Install the add-on from the [Google Workspace Marketplace](https://workspace.google.com/marketplace/app/exa_ai/465545439521).

This repo contains the Google Apps Script code for the Exa AI Google Sheets add-on.

## What You Can Do

- Generate full research tables from one prompt
- Fill blank cells in an existing table
- Continue a table by adding new rows
- Use `=EXA(...)` formulas for quick one-cell answers
- Refresh many Exa formula cells at once
- Convert Exa formulas to normal values

## Install

The easiest way is to install the public add-on:

1. Open the [Exa AI add-on](https://workspace.google.com/marketplace/app/exa_ai/465545439521) in the Google Workspace Marketplace.
2. Click **Install**.
3. Open a Google Sheet.
4. Go to **Extensions -> Exa AI -> Open Sidebar**.
5. Add your Exa API key from the [Exa dashboard](https://dashboard.exa.ai/api-keys).

You are ready to use Exa in Google Sheets.

## Exa Agent

Exa Agent is the easiest way to use Exa in Sheets.

Open the sidebar, then go to **Exa Agent**.

### Generate Table

Choose **Generate table** from the Task menu when you want Exa to create a new table.

Example prompt:

```text
Find top 40 AI companies and return company name, website URL, CEO, founding date, headquarters, and a short description.
```

Exa researches the web and writes the table into your sheet.

By default, the table starts at your selected cell. You can choose another start cell in **More options**.

### Find More

Use **Find more** to extend an existing table without returning rows already listed.

1. Select the entire existing table, including its header row.
2. Choose **Find more** from the Task menu and enter a prompt describing the additional rows you want, such as `Find five more companies like these`.
3. Click **Find more**.

The selected rows are sent to Exa Agent as `input.exclusion`. The selected headers define the output columns. New, non-duplicate rows are appended directly below the selection; existing rows are never overwritten. If the append area contains data, the add-on stops instead of writing over it. **Find more** does not use the separate Columns or Write table at options.

### For Each Row

Use **For each row** to generate multiple destination rows from each source row, with a separate Agent run per source row.

1. Put source records in one sheet, with headers in row 1 and a unique identifier in the first column. For example, a `Company` column.
2. Create a different destination sheet with headers in row 1. Include a column with the same header as the source identifier, such as `Company`; the other destination headers define the fields to find.
3. Select the source data rows to process. You can select any column in those rows; the add-on reads the full row and skips the header row.
4. From the source sheet, open **Exa Agent**, choose **For each row** from the Task menu, enter one prompt, and choose the destination sheet.
5. Click **Run for each row** to start one paid Agent request per selected row with a value in the first column. The sidebar shows the request count and progress; keep it open until the batch completes.

The add-on sends each selected source row as `input.data`, with up to 50 Agent runs active at once. It limits simultaneous Sheets-to-Apps-Script calls to eight while starting, polling, and writing those runs. Existing destination rows with the same source identifier go into `input.exclusion`. Every result must repeat the source identifier, and non-duplicate results append to the destination. Each selection is limited to 50 source rows; select another batch for larger tables. Exa API rate limits may still slow or reject a large batch.

### Fill Cells

Use **Fill cells** when you already have a table and want Exa to fill missing data.

1. Select blank cells in your sheet.
2. Open **Exa Agent**.
3. Choose **Fill cells**.
4. Click **Fill selected cells**.

Exa uses the table around your selection as context.

You can also select blank rows under a table. Exa will try to continue the table with new rows that match the same columns.

For citation columns, the header controls how grounding is written:

- `Sources` or `Evidence URLs` returns semicolon-separated URLs.
- `Primary Source URL` returns one best URL.
- `Sources (Markdown)` returns titled Markdown links.
- `Source Titles` returns source titles.
- `Sources JSON` returns a JSON string with titles and URLs.

## Formula

Use `=EXA(...)` when you want one answer in one cell.

```text
=EXA("Return only the CEO name", A2)
```

The second value, like `A2`, is the context. You can drag the formula down a column to run it for many rows.

More examples:

```text
=EXA("Return only the company website URL", A2)
=EXA("Return only the headquarters", A2)
=EXA("Return only the LinkedIn URL", A2)
```

## Batch

Use **Batch** when you want to work with many Exa formula cells at once.

Batch can:

- refresh selected cells with Exa formulas
- convert selected Exa formulas into normal values

Convert formulas to values when you want to keep the current results and stop the formulas from running again.

## Advanced Functions

The add-on also includes lower-level custom functions:

| Function | What it does |
| --- | --- |
| `EXA(prompt, context)` | Simple one-cell enrichment |
| `EXA_SEARCH(query, ...)` | Search the web and return URLs |
| `EXA_ANSWER(prompt, ...)` | Ask Exa for an answer |
| `EXA_CONTENTS(url)` | Get page content from a URL |
| `EXA_FINDSIMILAR(url, ...)` | Find pages similar to a URL |

Most users should start with **Exa Agent** or `=EXA(...)`.

## Local Development

This project uses Google Apps Script and `clasp`.

### Requirements

- Node.js 14 or newer
- A Google account
- An Exa API key

### Setup

Install dependencies:

```bash
npm ci
```

Log in to Google:

```bash
npm run login
```

Create a new Apps Script project:

```bash
npm run create
```

Push the local files to Apps Script:

```bash
npm run push
```

Run tests:

```bash
npm test
```

## Manual Apps Script Setup

If you do not want to use `clasp`, you can copy the files manually.

1. Open a Google Sheet.
2. Go to **Extensions -> Apps Script**.
3. Copy `Code.gs` into the Apps Script `Code.gs` file.
4. Create an HTML file named `Sidebar.html`.
5. Copy `Sidebar.html` from this repo into that file.
6. Save the project.
7. Refresh the Google Sheet.
8. Go to **Extensions -> Exa AI -> Open Sidebar**.
9. Add your Exa API key.

For normal use, install the [Marketplace add-on](https://workspace.google.com/marketplace/app/exa_ai/465545439521) instead.

## Project Files

- `Code.gs` - Apps Script backend, custom functions, Exa API calls, and sheet writes
- `Sidebar.html` - sidebar UI for Settings, Exa Agent, and Batch
- `appsscript.json` - Apps Script manifest and OAuth scopes
- `__tests__/` - Jest tests
- `scripts/bump-version.js` - version bump helper

made with :heart: by exa
