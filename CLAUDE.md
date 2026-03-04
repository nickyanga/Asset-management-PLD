# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Platform

**Google Apps Script (GAS) web app.** No Node.js, no package manager, no test runner. All code runs in Google's V8 runtime.

## Deployment

Requires [clasp](https://github.com/google/clasp).

```bash
clasp push                         # upload local files to Apps Script project
clasp deploy                       # create a new versioned deployment
clasp open                         # open the Apps Script editor in browser
```

No lint, build, or test step. Test manually via the Apps Script editor or the live web app URL.

## Architecture

### Two-file design
- **`Code.gs`** — Server-side GAS. Sheets I/O, email, transaction logic. Exposes functions callable via `google.script.run`.
- **`index.html`** — Single-file SPA (CSS in `<style>`, JS in `<script>`, no CDN). Served by `doGet()`.

### Frontend ↔ backend communication
`google.script.run` only accepts **a single argument** per call. The `runServer()` wrapper handles this:

```javascript
function runServer(fnName, params, onSuccess, onError) {
  google.script.run
    .withSuccessHandler(onSuccess)
    .withFailureHandler(onError || showErrorToast)
    [fnName](params);
}
```

Multi-param server functions use a wrapper object: `updateAsset({ serialNumber, updatedFields })`.

### Database: Google Sheets
`SPREADSHEET_ID` in `Code.gs` points to the backing spreadsheet. Two sheets:

- **`Assets`** (cols A–G): `SerialNumber, Model, Condition, Status, Notes, DateAdded, LastModified`
- **`Transactions`** (cols A–I): `TransactionID, SerialNumber, Condition, BorrowerName, BorrowerEmail, LoanDate, ReturnDate, Type, AdminNotes`

`SerialNumber` is the primary key across both sheets. Open loans: `Type === 'Loan'` with blank `ReturnDate`.

### Key server functions
| Function | Purpose |
|---|---|
| `getAssetBySerial(sn)` | Lookup single asset by serial number |
| `getAvailableAssets()` | All assets with Status=Available |
| `addAsset({ serialNumber, model, condition, notes })` | Create new asset |
| `updateAsset({ serialNumber, updatedFields })` | Update asset fields |
| `submitLoan({ borrowerName, borrowerEmail, serialNumber, condition })` | Process loan (lock-guarded) |
| `submitReturn({ transactionId, serialNumber, borrowerName, borrowerEmail, adminNotes })` | Process return (lock-guarded) |
| `lookupLoanBySerial(sn)` | Find open loan by serial number |
| `lookupLoansByBorrower(query)` | Find open loans by name/email |

### Race condition protection
`submitLoan()` and `submitReturn()` use `LockService.getScriptLock().waitLock(15000)`.

### Return flow
1. Fills `ReturnDate` on the **original Loan row** (closes it)
2. Appends a **new `Return` row** for audit history

### Loan condition
The Loan form has an **editable Condition dropdown**. If the admin changes condition at loan time, `submitLoan` updates the asset's condition in the Assets sheet.

### Frontend state
All data loads once per tab activation into `APP_STATE`. Filtering/search is client-side. Server is only called for mutations or initial tab data load.

## First-Time Setup

1. `clasp login && clasp create --type webapp`
2. `clasp push`
3. In the Apps Script editor, run `createNewSheet()` — it creates the spreadsheet with both sheets and correct headers. Copy the logged ID into `SPREADSHEET_ID` in `Code.gs`.
4. `clasp push` again (to push the updated ID)
5. Deploy → New Deployment → Web app (Execute as: Me, Access: Anyone)
