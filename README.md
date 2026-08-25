# KING SHEET — Google Sheets matched-betting tracker

This repository is a container-bound Google Apps Script project. It builds the connected financial sheets **Settings → Promos/BTO/NonPromos → Monthly Metrics and EV Breakdown → Master PnL**, plus the independent **Bookie Health** status matrix and Audit. Dashboard totals always read the three entry sheets directly; Bookie Health is not included in EV or PnL.

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

* **Settings:** all list/rate cells are editable. Add bookmaker accounts in column H, BTO methods in I, non-promo types in J, results in K, general account statuses in L, transaction types in M, and banks in N. Bookie Health statuses (including colours, order, and active flags) are in P:T, and its bookmaker list is in V. Promo rates are A:B, EV tier limits D:F, and the missing-BSP BTO default is E10.
* **Bookie Health:** bookmaker rows come from Settings column V. Profiles 1–20 begin in columns B:U. Selecting another active status appends it with `, `; repeats are ignored and clearing a cell leaves it blank. Use **Matched Betting Tracker → Refresh Bookie Health** after editing bookmakers or status settings.
* **Promos:** Date, Account, Promo Type, Notes, Turnover, Odds, Result, Cash Change, Bonus Change (B:H and K:L, excluding generated columns). Entry ID, Promo EV %, Estimated EV, Actual PnL, POT %, EV Tier, and Month are calculated.
* **BTO:** Date, Account, BTO Method, Notes, Turnover, Odds, optional BSP, Result, Bookie Change, Betfair Lay Stake, Lay Odds, Commission %, and Betfair Change (B:I and L:P). Entry ID, EV Taken, EV %, Actual Return, Actual BTO %, and Month are calculated.
* **NonPromos:** Date, Account, Non-Promo Type, Notes, Turnover, Odds, BSP, Result, Bookie Change, and Betfair Change (B:I and L:M). Entry ID, Expected EV, EV %, Actual PnL, POT %, and Month are calculated.
* **EV Breakdown:** reporting period in B2; for Custom, From Date in B3 and To Date in B4.

Formula columns use spill formulas and warning-only protections. Warning-only protection prevents accidental edits without locking the owner out. Add data below the header; formulas and validations are provisioned for 1,999 entry rows and the builder safely expands short sheets.

## Checks and calculation assumptions

* `runBtoFormulaTests()` verifies the four required results: AUD 24.50, 19.22, 5.44, and 18.75. The final case reads the editable Settings rate rather than embedding 75% in the calculation.
* `runAudit()` creates live PASS/FAIL checks with a one-cent monetary tolerance, scans displayed values for spreadsheet errors, and checks Bookie Health names, multi-status values, mappings, and validation.
* Bonus Change is valued at 100%. Blank Bookie/Betfair changes become zero after Date and Turnover exist.
* BSP is optional for BTO (the Settings fallback applies), but Audit deliberately flags missing BSP for review. Non-promo expected EV remains blank without BSP.
* Entry IDs are deterministic from sheet, date, and row. Moving a row changes its generated ID; no separate immutable-ID service is assumed.
* Australian dates and AUD formats are display formats. The project time zone is Australia/Sydney.
* Settings updates recalculate spill formulas and every downstream summary/dashboard because all summaries reference entry-level calculated columns.
