# Hertz-Cashflow

Weekly 13-week cash flow forecast for Hertz Philippines (browser app, no server).

## Use
1. Open `index.html` in **Google Chrome or Microsoft Edge**.
2. **Connect folder** (bottom of the sidebar) → choose the shared `FILES` folder (it holds `Cashflow Summary` and `Uploaded Files`).
3. **Setup → Cash balances**: enter Friday's bank balances (Beginning Balance of Week 1).
4. **Upload**: AR Aging, Cash Flow Report, Unit Disposals, Loan Payment Schedule, Unreleased Checks, Projected Cash Outflow.
5. Review / edit the Summary (or the uploaded files via **Uploaded files** in the sidebar), then **Save summary**. The saved file also holds a copy of each source sheet (AR Corporate, Cash Flow Report, Unit Disposals, Loan Payments, Unreleased Checks, Projected Payments) and a Bank Balances sheet; the Summary amounts are formulas that point to them.

Company data (FILES, source Excel files) is not stored in this repository.

## Several people at the same time (V3.6)
Everyone connects the same shared **FILES** folder on the LAN. Each part (AR Aging, Cash Flow Report, Unit Disposals, Loans, Unreleased Checks, Projected payments, bank balances, typed amounts) is saved separately in `FILES/Current`; **Save summary** combines everyone's latest work into `Cashflow Summary/Cashflow_Summary_LATEST.xlsx`, and **Finalize week** keeps a dated copy. No data is sent to GitHub — the browser reads and writes the LAN folder directly.

## V3.7
- **Unidentified Collection**: a Cash In row under AR Corporate, with its own upload slot under AR Aging. The file has a Date and an Amount column (actual collections); each amount goes into the week of its date, full amount (no EWT).
- **Reference Folder**: Save summary also updates the source file of the same name in `FILES/Reference Folder` (for every user), after keeping a "(before edit)" backup in Uploaded Files. If the file is open in Excel, the system asks you to close it; an update that could not be written is retried on the next Save summary.
