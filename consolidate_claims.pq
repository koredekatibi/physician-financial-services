// Power Query (M) — Consolidate_Claims
// Paste into: Data > Get Data > Launch Power Query Editor > Home > Advanced Editor
// Assumes two source queries already exist in the workbook: "EHR_Export" and "Billing_Export"
// (each pointing at the matching sheet, or swap the Excel.CurrentWorkbook() sources below
// for your real EHR / billing system connectors in production.)

let
    // Step 1: Load source tables
    EHR = Excel.CurrentWorkbook(){[Name="EHR_Export"]}[Content],
    Billing = Excel.CurrentWorkbook(){[Name="Billing_Export"]}[Content],

    // Step 2: Enforce data types on each source before merging
    EHR_Typed = Table.TransformColumnTypes(EHR, {
        {"Claim_ID", Int64.Type},
        {"Physician", type text},
        {"Department", type text},
        {"Service_Date", type date},
        {"CPT_Code", type text}
    }),
    Billing_Typed = Table.TransformColumnTypes(Billing, {
        {"Claim_ID", Int64.Type},
        {"Payer", type text},
        {"Billed_Amount", type number},
        {"Paid_Amount", type number},
        {"Status", type text},
        {"Denial_Reason", type text},
        {"RVU", type number}
    }),

    // Step 3: Merge queries on Claim_ID (equivalent of the INDEX/MATCH join in the workbook)
    Merged = Table.NestedJoin(
        EHR_Typed, {"Claim_ID"},
        Billing_Typed, {"Claim_ID"},
        "BillingData", JoinKind.LeftOuter
    ),

    // Step 4: Expand the matched billing columns
    Expanded = Table.ExpandTableColumn(Merged, "BillingData",
        {"Payer", "Billed_Amount", "Paid_Amount", "Status", "Denial_Reason"},
        {"Payer", "Billed_Amount", "Paid_Amount", "Status", "Denial_Reason"}
    ),

    // Step 5: Remove claims with no billing match (data quality check — flags EHR/billing sync gaps)
    RemovedUnmatched = Table.SelectRows(Expanded, each [Payer] <> null),

    // Step 6: Reorder columns to match the target Power_Query_Consolidated layout
    Reordered = Table.ReorderColumns(RemovedUnmatched,
        {"Claim_ID", "Physician", "Department", "Service_Date", "CPT_Code",
         "Payer", "Billed_Amount", "Paid_Amount", "Status", "Denial_Reason"}
    ),

    // Step 7: Sort by Claim_ID for consistent, diffable output
    Sorted = Table.Sort(Reordered, {{"Claim_ID", Order.Ascending}})
in
    Sorted

/* Notes:
   - Table.SelectRows in Step 5 is the data-quality check: any Claim_ID present in the EHR
     export but missing from the billing export means a claim was seen clinically but never
     billed — that's a revenue-leakage signal worth its own root-cause tracker.
   - In production, replace the two Excel.CurrentWorkbook() sources with the EHR and billing
     system's native connectors (SQL Server, ODBC, or a REST API via Web.Contents), then set
     the query to refresh on a schedule.
*/
