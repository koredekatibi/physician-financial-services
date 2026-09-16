# Root Cause Analysis — Denied Claims & Reimbursement Variance

*(Same content as the `Root_Cause_Dashboard` tab in the workbook, provided here as a standalone document.)*

## Finding

The largest driver of denied claims and dollars at risk is **Prior Authorization Missing**, concentrated
in higher-RVU visit codes (99214/99215) and hemodialysis evaluations (90935) — procedures that require
authorization under most commercial and Kaiser HMO contracts.

**Coding Error** is the second-largest driver and clusters around new-patient E/M codes (99203/99204),
which points to a front-end documentation or charge-entry gap rather than a payer-side issue.

## Method

1. Pulled `Status = Denied` claims from `Claims_Data` and broke them out by `Denial_Reason` using
   `COUNTIFS` / `SUMIFS` (claim count and billed-dollar impact).
2. Cross-referenced denial reason against CPT code and department to see whether denials clustered
   around specific procedures or were evenly distributed (a scatter pattern points to a payer-side
   issue; a cluster points to an internal process gap).
3. Used the `Complex_IF_Variance_Flags` tab's nested-IF logic to separate "Denied" claims from claims
   that were paid but underpaid beyond a 5% threshold — two different root causes require two different
   fixes, and combining them would have hidden which lever to pull first.

## Recommended next step

Pull a 30-day sample of Prior-Auth denials by ordering physician to confirm whether the gap is:

- a **scheduling workflow issue** (authorization never requested before the visit), or
- a **payer turnaround issue** (authorization requested but not returned before the claim was filed).

These have different owners and different fixes — the first is a front-desk/scheduling process change;
the second is a payer-relations or timely-filing escalation. Routing the finding to the wrong owner is
the most common reason root-cause work stalls after the analysis is done.
