# Root Cause Analysis — Denied Claims & Reimbursement Variance

*(Same content as the `Root_Cause_Dashboard` tab in the workbook, provided here as a standalone document.)*

## Finding

The largest **dollar exposure** among denied claims comes from **Duplicate Claim** submissions ($763.56)
and **Timely Filing Limit Exceeded** ($634.87). By **claim count**, Duplicate Claim and Eligibility Not
Verified are tied as the most frequent denial reasons (4 claims each) — Coding Error and Prior
Authorization Missing were comparatively rare in this sample.

## Method

1. Pulled `Status = Denied` claims from `Claims_Data` and broke them out by `Denial_Reason` using
   `COUNTIFS` / `SUMIFS` (claim count and billed-dollar impact).
2. Compared count-based ranking against dollar-based ranking side by side — they don't always agree,
   and the gap itself is informative. A high-count, low-dollar reason (Coding Error, at $187 total) is
   a lower-priority fix than a lower-count, high-dollar reason (Duplicate Claim, at $763 total from 4
   claims — nearly $191 per claim).
3. Used the `Complex_IF_Variance_Flags` tab's nested-IF logic to keep "Denied" claims separate from
   claims that were paid but underpaid beyond a 5% threshold — two different root causes require two
   different fixes, and combining them would have hidden which lever to pull first.

## What the pattern suggests

- **Duplicate Claim** denials typically point to a front-end billing or system-integration issue — the
  same claim submitted twice, often after a resubmission workflow error (e.g., a claim resubmitted after
  a status-check timeout without confirming the original hadn't already posted).
- **Eligibility Not Verified** points to a front-desk intake gap — coverage wasn't confirmed before the
  visit occurred, rather than a payer-side denial.

## Recommended next step

Audit the claims-submission log for duplicate-claim patterns by billing staff member or system
integration point, to confirm whether it's a manual resubmission habit or an automated double-post from
a system integration. Separately, confirm whether eligibility checks are being run before check-in or
only at time of billing — these are two different owners and two different fixes, and routing the
finding to the wrong owner is the most common reason root-cause work stalls after the analysis is done.
