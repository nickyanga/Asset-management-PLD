# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Platform

This is a **Google Apps Script (GAS) web app**. There is no Node.js build system, no package manager, and no test runner. All code runs inside Google's V8 runtime on Google's servers.

## Deployment Commands

Requires [clasp](https://github.com/google/clasp) (Google's CLI for Apps Script).

```bash
clasp login                        # authenticate once
clasp create --type webapp         # first-time setup (generates .clasp.json)
clasp push                         # upload local files to Apps Script project
clasp deploy                       # create a new versioned deployment
clasp open                         # open the Apps Script editor in browser
```

There is no lint, build, or test step. Testing is done manually via the Apps Script editor (`clasp open`) or through the live web app URL.

## Architecture

### Two-file design
The entire app lives in two files:

- **`Code.gs`** — Server-side GAS. Handles all Sheets I/O, email sending, and transaction logic. Exposes functions callable from the frontend via `google.script.run`.
- **`index.html`** — Single-file SPA (all CSS in `<style>`, all JS in `<script>`, no CDN). Served by `doGet()` in Code.gs.

### How frontend ↔ backend communication works
The frontend uses `google.script.run` (GAS-specific browser API). **It only accepts a single argument per function call.** The `runServer()` wrapper in `index.html` handles this pattern:

```javascript
function runServer(fnName, params, onSuccess, onError) {
  google.script.run
    .withSuccessHandler(onSuccess)
    .withFailureHandler(onError || showErrorToast)
    [fnName](params);
}
```

Because of the single-argument constraint, multi-param server functions must use a wrapper object. For example, `updateAsset` takes `{ assetTag, updatedFields }` as one `params` object — not two separate arguments.

### Database: Google Sheets
`SPREADSHEET_ID` in `Code.gs` must be set to your sheet's ID (from the URL). Two sheets are required:

- **`Assets`** — headers in row 1: `AssetTag, SerialNumber, DeviceName, Model, Category, Condition, Status, Notes, DateAdded, LastModified`
- **`Transactions`** — headers in row 1: `TransactionID, AssetTag, DeviceName, Category, Condition, BorrowerName, BorrowerEmail, LoanDate, ReturnDate, Type, AdminNotes`

`AssetTag` is the primary key across both sheets. Open loans are identified by `Type === 'Loan'` with a blank `ReturnDate`.

### Race condition protection
`submitLoan()` and `submitReturn()` use `LockService.getScriptLock().waitLock(15000)` to prevent two simultaneous submissions from double-loaning the same asset.

### Return flow (important)
A return does two writes:
1. Fills `ReturnDate` on the **original Loan row** (makes it a closed loan)
2. Appends a **new `Return` row** for full audit history

### Frontend state
All data loads once per tab activation into `APP_STATE`. Filtering/search is client-side. Server is only called for mutations or initial tab data load.

## Required First-Time Setup

1. Create a Google Sheet with `Assets` and `Transactions` sheets (row 1 = headers as above)
2. Set `SPREADSHEET_ID` in `Code.gs`
3. `clasp login && clasp create --type webapp`
4. `clasp push`
5. In Apps Script: Deploy → New Deployment → Web app (Execute as: Me, Access: Anyone)
