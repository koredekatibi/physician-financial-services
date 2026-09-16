# Physician Financial Services Analytics Portfolio

**Korede Katibi, MBA** — built for interview prep: **Kaiser Permanente, Physician Financial Services Consultant III (Req #1436508)**

This portfolio simulates a physician-group revenue cycle and reimbursement analytics workflow — claims data, contracted payer rates, variance detection, and financial modeling — using the exact Excel toolset requested for the role:

| Tool | Where it lives |
|---|---|
| Pivot Tables | `Claims_Data` is a native Excel Table (`tbl_Claims`); `Pivot_Summary` tab |
| VLOOKUP / XLOOKUP | `VLOOKUP_Analysis` tab (with a documented note on where VLOOKUP breaks) |
| INDEX/MATCH | `INDEX_MATCH_Analysis` tab (two-way lookup: CPT code + payer) |
| Complex IF Statements | `Complex_IF_Variance_Flags` tab (nested IF + AND logic) |
| Power Query | `EHR_Export` / `Billing_Export` / `Power_Query_Consolidated` tabs + `power_query/consolidate_claims.pq` |
| Data Validation | Dropdowns on `Claims_Data` (Department, Payer, Status, Denial Reason) |
| Financial Models | `Financial_Model` tab — RVU-based physician compensation model |
| Macros/VBA | `vba/*.bas` — three macros, documented in the `VBA_Macros` tab |
| Formula Auditing | `Formula_Auditing_Notes` tab |
| Root Cause Analysis | `Root_Cause_Dashboard` tab — denial breakdown, chart, and written root-cause narrative |

## Repo structure

```
├── excel/
│   └── Physician_Financial_Services_Portfolio.xlsx   ← the main workbook (open this first)
├── power_query/
│   └── consolidate_claims.pq                          ← M-code for the EHR + Billing merge
├── vba/
│   ├── RefreshAndFormatDashboard.bas
│   ├── HighlightVarianceOutliers.bas
│   └── ExportRootCauseSummaryToPDF.bas
├── docs/
│   └── root_cause_writeup.md                          ← standalone narrative, same content as the dashboard tab
└── README.md
```

## Scenario

A 220-claim illustrative sample across five departments (Nephrology, Cardiology, Internal Medicine,
Endocrinology, Family Medicine), five payer types (Medicare, Medicaid, Kaiser HMO, and two commercial
plans), and ten CPT codes — modeled after the kind of physician-services revenue cycle data referenced
in my experience at Bethsaida Nephrology and the variance/root-cause work I did at USAA.

## How to use each tab (suggested interview walkthrough order)

1. **Overview** — one-page map of the workbook.
2. **Claims_Data** — the source of truth. Try changing a dropdown value; note it's a real Excel Table, so it's already the source for a native PivotTable (`Insert > PivotTable`) and for the Power Query source tabs.
3. **Reimbursement_Rates** — the contracted rate table, including the `Payer_Rate_Map` helper table used by INDEX/MATCH.
4. **VLOOKUP_Analysis** → **INDEX_MATCH_Analysis** — walk through *why* INDEX/MATCH replaces VLOOKUP here: VLOOKUP's column index is hard-coded and breaks if a column moves; INDEX/MATCH resolves the payer's rate column dynamically by name.
5. **Complex_IF_Variance_Flags** — the nested-IF decision tree that turns raw variance into an actionable status, plus an AND-based priority flag.
6. **Pivot_Summary** — the native-table pivot source, plus a SUMIFS/COUNTIFS cross-tab as a transparent, always-live alternative.
7. **EHR_Export / Billing_Export / Power_Query_Consolidated** — the two "source systems," their merge, and the actual M-code in `power_query/consolidate_claims.pq` that would produce it live in Excel.
8. **Financial_Model** — change any yellow assumption cell and watch every month recalculate.
9. **Root_Cause_Dashboard** — the chart plus a written root-cause narrative (see `docs/root_cause_writeup.md`).
10. **Formula_Auditing_Notes** — the auditing discipline used while building the file (Trace Precedents/Dependents, Evaluate Formula, IFERROR guards).
11. **VBA_Macros** — macro descriptions; actual code in `/vba`.

## Honest technical notes

- **Power Query and VBA can't be authored as live, compiled objects outside Excel's own engine.** The M-code and VBA modules here are real, complete, and ready to paste/import — `power_query/consolidate_claims.pq` via *Data > Get Data > Launch Power Query Editor > Advanced Editor*, and each `.bas` file via *Alt+F11 > File > Import File*. The workbook itself shows what each would produce, built with live formulas so it's fully verifiable.
- **XLOOKUP**: the newer XLOOKUP/dynamic-array functions aren't included as live formulas in this file because they need a very recent Excel build to evaluate reliably; INDEX/MATCH is used throughout as the more portable, universally-supported equivalent — which is also the safer real-world choice when a workbook is shared across an organization on mixed Excel versions.
- The Financial_Model's RVU target is scaled to this workbook's 220-claim illustrative sample (not a full monthly production extract) — the formula structure is what would carry over unchanged to a real dataset.

## Publishing this to your own GitHub

```bash
cd physician-financial-services-portfolio
git init
git add .
git commit -m "Physician financial services analytics portfolio"
git branch -M main
git remote add origin https://github.com/<your-username>/physician-financial-services-portfolio.git
git push -u origin main
```
