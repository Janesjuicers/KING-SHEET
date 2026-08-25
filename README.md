# KING SHEET — Google Sheets matched-betting tracker

This repository is a container-bound Google Apps Script project. It builds the connected financial sheets **Bookie Health → Accounts and account dropdowns**, and **Settings → Promos/BTO/NonPromos → Monthly Metrics and EV Breakdown → Master PnL**, plus Audit.

## Deploy to Google Sheets

### Apps Script editor (simplest)

1. Create or open a Google Sheet.
2. Choose **Extensions → Apps Script**.
3. Create one script file for every `.gs` file in this repository and paste the matching contents. Replace the generated manifest by enabling **Project Settings → Show `appsscript.json` manifest file**, then paste `appsscript.json`.
4. Save, select `buildWorkbook`, and click **Run**. Review and grant the requested spreadsheet permission.
5. Return to the spreadsheet and reload it. Use **Matched Betting Tracker → Build/Rebuild Workbook** whenever structure needs refreshing.

### clasp

Install Google's `clasp`, sign in with `clasp login`, run `clasp create --type sheets --title "KING SHEET"` (or add the target script ID to `.clasp.json`), copy these files into that project, then run `clasp push`. Open the spreadsheet, run `buildWorkbook` once, and reload.

The build is idempotent: it reuses named sheets, replaces managed formulas/charts/validation/protections, seeds settings only when their starting cells are empty, and never clears entry sheets. Three clearly labelled Promo examples are added only when Promos has no dated rows.

## Manual-entry columns

* **Settings:** BTO methods are in I, non-promo types in J, results in K, and the Bookie Health bookmaker list is in V. Column W is a hidden, generated helper backing the `Account_ID_List` named range; do not maintain account IDs manually. Bookie Health statuses (including colours, order, and active flags) are in P:T. Promo rates are A:B, EV tier limits D:F, and the missing-BSP BTO default is E10.
* **Bookie Health:** up to 250 nonblank base bookmaker names come from Settings column V. Each of profiles 1–24 has adjacent Status and protected Account ID columns. An ID (for example `23BetNation`) exists only while its Status is nonblank. Selecting another active status appends it with `, ` without duplicates; clearing Status also clears the ID. Refreshing preserves selections by bookmaker name and updates the central list, Accounts, and all account dropdowns.
* **Accounts:** generated Account IDs are appended once; existing columns and account data are preserved and matched by Account ID.
* **Promos:** the visible columns are exactly Date, Account, Promo Type, Notes, Turnover, Odds, Result, Estimated EV, Cash Change, and Bonus Change. The last three are protected formulas; reporting derives actual PnL as Cash Change plus Bonus Change and groups directly from Date.
* **BTO:** columns are Date, Account, BTO Method, Notes, Turnover, Odds, optional BSP, Result, EV Taken, EV %, Bookie Change, Betfair Lay Stake, Lay Odds, Commission %, Betfair Change, Actual Return, and Actual BTO %. It has no Entry ID or Month helper; monthly reporting groups directly from Date.
* **NonPromos:** Date, Account, Non-Promo Type, Notes, Turnover, Odds, BSP, Result, Bookie Change, and Betfair Change remain manual fields. Expected EV, EV %, Actual PnL, POT %, and Month remain calculated; Entry ID has been removed.
* **EV Breakdown:** reporting period in B2; for Custom, From Date in B3 and To Date in B4.

Formula columns use spill formulas and warning-only protections. Warning-only protection prevents accidental edits without locking the owner out. Add data below the header; formulas and validations are provisioned for 1,999 entry rows and the builder safely expands short sheets.

## Checks and calculation assumptions

* `runBtoFormulaTests()` verifies the four required results: AUD 24.50, 19.22, 5.44, and 18.75. The final case reads the editable Settings rate rather than embedding 75% in the calculation.
* `runAudit()` creates live PASS/FAIL checks with a one-cent monetary tolerance, scans displayed values for spreadsheet errors, and checks Bookie Health names, multi-status values, mappings, and validation.
* Bonus Change is valued at 100%. Blank Bookie/Betfair changes become zero after Date and Turnover exist.
* BSP is optional for BTO (the Settings fallback applies), but Audit deliberately flags missing BSP for review. Non-promo expected EV remains blank without BSP.
* Australian dates and AUD formats are display formats. The project time zone is Australia/Sydney.
* Settings updates recalculate spill formulas and every downstream summary/dashboard because all summaries reference entry-level calculated columns.
