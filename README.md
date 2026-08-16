# Ledger — Invoice Dashboard

A simple, self-hosted invoicing web app built on top of **Google Sheets** and **Google Apps Script**. No server, no database, no subscription fees — your data lives in a Google Sheet you own, and the entire UI is served to you as a free Google Apps Script web app.

## Overview

Ledger is a single-page invoice manager that lets you:

- Create professional, printable **A4 invoices** with a live preview as you type
- Auto-generate sequential invoice numbers (e.g. `INV-1001`, `INV-1002`, …)
- Add multiple line items, tax percentage, and discounts
- Track invoice status (`Unpaid` / `Paid`), with **Overdue** detected automatically
- Log **expenses** and see them next to your income
- View a **dashboard** with revenue, outstanding, overdue, and monthly/daily charts
- **Print** invoices or **download as PDF** directly from the browser
- Customize your business profile (name, tagline, email, address, currency symbol, footer) from a Settings page
- Optionally link to a **separate spreadsheet** for your invoice data

Everything is stored in plain Google Sheet tabs, which means you can inspect, edit, or export your data at any time.

## How it works

```
┌─────────────────────────┐        ┌──────────────────────────────┐
│   Browser (index.html)  │  ⇄     │  Google Apps Script (code.gs) │
│   React UI + print CSS  │  run   │  doGet / google.script.run    │
└─────────────────────────┘        └──────────────┬───────────────┘
                                                   │
                                        ┌──────────▼──────────┐
                                        │   Google Sheet       │
                                        │  Invoices / Expenses │
                                        │  Config / Settings   │
                                        └─────────────────────┘
```

There are two files:

| File | Purpose |
| ---- | ------- |
| `index.html` | The entire front end — React (UMD build), Babel standalone, hand-rolled SVG charts, all styling, and the printable A4 invoice document. |
| `code.js` | The Google Apps Script backend — serves the HTML, reads/writes the spreadsheet, computes totals, and generates invoice/expense numbers. |

The front end talks to the backend through `google.script.run`, the standard Apps Script RPC bridge. When you open the page outside of a deployed Apps Script web app (e.g. by double-clicking the HTML file), the app detects this and runs in **Preview mode** with mocked data so you can still explore the interface.

## Google Sheets implementation

The app uses up to four tabs inside your spreadsheet. All of them are created automatically on first use.

### `Invoices` tab

One row per saved invoice. Columns (in order):

| Column | Notes |
| ------ | ----- |
| `Invoice ID` | Auto-generated, e.g. `INV-1001` |
| `Date` | Invoice date |
| `Due Date` | Payment due date |
| `Customer Name` | Required |
| `Customer Email` | |
| `Customer Phone` | |
| `Customer Address` | |
| `Items JSON` | Line items as a JSON string, e.g. `[{"description":"Consulting","qty":1,"price":850}]` |
| `Subtotal` | Sum of `qty × rate`, always computed on the server |
| `Tax %` | Percent, e.g. `5` |
| `Tax Amount` | `Subtotal × Tax % / 100` |
| `Discount` | Flat amount subtracted from the total |
| `Total` | `Subtotal + Tax Amount − Discount` |
| `Status` | `Unpaid` or `Paid` |
| `Notes` | Terms / thank-you note shown on the document |
| `Created At` | Timestamp set when the row is appended |

The tab doesn't have to be named "Invoices" — the script auto-detects any tab whose header row looks like an invoice sheet (contains `Invoice ID` or `Customer`) or whose first column contains `INV-` ids. If nothing matches, it creates the `Invoices` tab for you.

### `Expenses` tab

One row per logged expense. Columns: `Expense ID` (`EXP-1001`, …), `Date`, `Category`, `Description`, `Amount`, `Notes`, `Created At`.

### `Config` tab (hidden)

Holds counters used for auto-numbering. It is created hidden and stores two key/value rows:

| Key | Value |
| --- | ----- |
| `LastInvoiceNumber` | Last used invoice number (e.g. `1000` → next is `INV-1001`) |
| `LastExpenseNumber` | Last used expense number |

Number generation is wrapped in `LockService.getScriptLock()` so two users can never receive the same invoice number.

### `Settings` tab

A simple two-column `Key` / `Value` sheet with the business profile:

| Key | Example |
| --- | ------- |
| `Company Name` | Ledger Studio |
| `Tagline` | Invoicing for modern business |
| `Currency Symbol` | `$` |
| `Company Email` | billing@ledger.studio |
| `Company Address` | 128 Commerce Street, Springfield |
| `Invoice Prefix` | `INV-` |
| `Footer` | Thank you for your business |
| `Spreadsheet ID` | *(optional)* URL or ID of a different spreadsheet to hold invoice data |

Any keys you don't set fall back to the defaults in `DEFAULT_SETTINGS` in `code.js`.

### Linking a separate data spreadsheet

By default all data lives in the spreadsheet the script is attached to. If you want your invoices in a *different* sheet, paste that sheet's URL (or ID) into the `Spreadsheet ID` setting. The script extracts the ID, opens it via `SpreadsheetApp.openById()`, and uses it for all reads/writes. If the linked sheet is inaccessible, it falls back to the attached spreadsheet.

## Google Apps Script implementation process

The backend in `code.gs` follows a simple request/response model. Every public function is callable from the browser via `google.script.run`.

### Core flow

1. **`doGet()`** — entry point. Serves `index.html` as a web app with a viewport meta tag.
2. **`getSpreadsheet()`** — resolves which spreadsheet to use (linked one from Settings, otherwise the attached spreadsheet).
3. **`getInvoiceSheet()`** — finds the `Invoices` tab (by name, header shape, or `INV-` pattern), otherwise creates it with frozen headers.
4. **`getNextInvoiceNumber()`** — locks, increments `LastInvoiceNumber` in `Config`, returns the next ID, unlocks.
5. **`saveInvoice(invoiceData)`** — validates a customer name and at least one line item, **recomputes** `subtotal`, `taxAmount` and `total` on the server (client math is never trusted), then appends a row.
6. **`getAllInvoices()` / `getAllExpenses()`** — read all rows back as objects, most recent first, serialized with `JSON.stringify`. (Returned as JSON strings on purpose — `google.script.run` silently drops arrays that contain nested arrays of objects.)
7. **`updateInvoiceStatus(id, status)`** — finds the row and flips its `Status` cell (used by "Mark as paid").
8. **`saveExpense()` / `deleteExpense()`** — append/remove expense rows.
9. **`getSettings()` / `saveSettings()`** — read/write the `Settings` tab.
10. **`getDashboardStats()`** — aggregates invoices/expenses into dashboard numbers and chart series (last 6 months revenue, last 7 days income vs expense), using `Utilities.formatDate()` with the script's time zone.

### Public API surface

| Function | Returns |
| -------- | ------- |
| `getAllInvoices()` | JSON string of all invoices |
| `getAllExpenses()` | JSON string of all expenses |
| `saveInvoice(data)` | `{ success, invoiceId, subtotal, taxAmount, total }` |
| `saveExpense(data)` | `{ success, expenseId, amount }` |
| `deleteExpense(id)` | `{ success }` |
| `updateInvoiceStatus(id, status)` | `{ success }` |
| `getNextInvoiceNumber()` | Next ID (consumes it) |
| `previewNextInvoiceNumber()` | Next ID (does **not** consume it) |
| `getSettings()` | Merged settings object |
| `saveSettings(obj)` | `{ success }` |
| `getDashboardStats()` | Dashboard numbers + chart series |
| `getDataInfo()` | Diagnostics about which sheet is being read |

## Setup & deployment

### 1. Create the spreadsheet

1. Go to https://sheets.new and create a blank spreadsheet. Name it something like "Invoice Data".
2. Leave it blank — the script creates all tabs automatically.

### 2. Add the Apps Script project

1. In the spreadsheet, go to **Extensions → Apps Script**.
2. Delete the default `myFunction()`.
3. Copy the full contents of `code.js` into the editor.
4. Add `index.html` as an HTML file: click **+** next to "Files" → **HTML**, name it `index`, and paste the contents of `index.html`. (The name must be `index` — `doGet()` serves exactly that file.)

### 3. Deploy as a web app

1. Click **Deploy → New deployment**.
2. Choose type **Web app**.
3. Under **Execute as**, select **Me**.
4. Under **Who has access**, choose:
   - **Anyone** — if you want a public URL, or
   - **Anyone with Google account** — if viewers need to sign in.
5. Click **Deploy** and authorize the requested permissions (read/write your spreadsheet, run a script, etc.).
6. Copy the **Web app URL** it gives you. That's your app.

### 4. Use it

Open the deployed URL in any browser. A yellow **Preview mode** banner means you're not on the deployed URL (e.g. you opened `index.html` locally) — in that case nothing is written to your sheet. Open the deployed URL and the banner disappears.

## Features in detail

- **Live A4 preview** — the invoice document is rendered as real HTML/CSS and scaled to fit the sidebar; it updates as you type.
- **Print / Save as PDF** — the `@media print` styles hide everything except the A4 sheet. From the invoice viewer, click **Print** or **Download PDF**, then choose "Save as PDF" in the print dialog.
- **Status stamps** — `Paid` (teal), `Unpaid` (amber), `Overdue` (brick) are rendered as rotated "rubber stamp" badges. Overdue is computed from `Status !== Paid && Due Date < today`.
- **Search & filter** — invoices filter by customer name/ID and by status; expenses filter by description/category/ID.
- **Hand-rolled charts** — the revenue and income-vs-expense charts are inline SVG components, no chart library needed.
- **Boot diagnostics** — if React, ReactDOM or Babel fail to load from the CDN, the page shows exactly why instead of a blank screen.
- **Currency & branding** — everything driven from the `Settings` tab; change `$` to `€` and the whole app, including printed documents, follows.

## Dependencies

All third-party libraries are loaded from CDNs by `index.html`:

- React 18.3.1 (UMD) from `cdnjs.cloudflare.com`
- ReactDOM 18.3.1 (UMD) from `cdnjs.cloudflare.com`
- Babel standalone 7 from `unpkg.com`

The Charts, icons, fonts, and all styling are self-contained. The only Google Fonts loaded are Fraunces, Inter, and IBM Plex Mono (with system-font fallbacks).

## Project structure

```
.
├── code.js        # Google Apps Script backend (all server functions)
├── index.html     # React single-page app + A4 invoice print styles
├── .hintrc        # webhint lint configuration
└── .vscode/       # editor settings (live preview path)
```

## Limitations

- **No authentication/authorization** — whoever has the deployed URL can read and write your sheet. Grant access accordingly, or keep the URL private.
- **No invoice deletion or editing** in the UI — invoices can only be marked paid. Delete rows directly in the sheet if needed.
- **No email sending** — invoices are printed/exported as PDF rather than emailed.
- Requires internet access to the CDNs and Google's servers.

## License

This is a small personal project. Free to use, modify, and adapt for your own invoicing needs.
