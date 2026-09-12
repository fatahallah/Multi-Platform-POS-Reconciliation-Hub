# Multi-Platform Delivery Sales Reconciliation & Performance Analytics

An Excel-based reconciliation system that matches point-of-sale records against multiple delivery platform reports to catch missing orders, verify payouts, and surface commission costs — built as a fast, formula-driven alternative to manual cross-checking.

![Executive Summary](images/Executive Summary.png)

## Business Problem

Restaurants selling through multiple delivery apps (Jahez, Hungerstation, ToYou, Ninja, Chefz) end up with sales data scattered across their POS system and each platform's own settlement reports. Manually cross-checking these to confirm every order was paid out correctly — and by how much commission was deducted — is slow and error-prone. This project builds a reconciliation workbook that automates that matching and flags real discrepancies in seconds.

## Data

- **Source:** Simulated POS and delivery-platform data generated to reflect realistic reconciliation scenarios (a portfolio sample, not data from an actual business)
- **Scope:** 220 orders logged in the internal POS system, matched against 217 corresponding records reported by 5 delivery platforms
- **Files:** `Foodics_POS_Sales_Report.csv`, `Delivery_Platforms_Combined_Report.csv`, and the consolidated `Multi_Platform_Delivery_Sales_Reconciliation.xlsx` workbook

## Data Preparation

Each POS order carries a `Platform_Reference_ID` linking it to the matching record in the delivery platform report, and each platform record carries a `POS_Order_Ref` back to the POS order — this shared key is what makes automated matching possible. Three orders were deliberately left out of the platform report to test whether the reconciliation logic correctly flags orders as missing rather than silently ignoring them.

## Analysis

The reconciliation logic separates two very different questions that are easy to conflate:

1. **Did the order value reported by the platform match what the POS recorded?** — compares POS Net Sales against the platform's *gross* order value (before any deductions). Any gap here is a genuine error worth investigating.
2. **How much was deducted in commission and VAT?** — a separate calculation comparing gross order value to the actual net payout, reported on its own rather than mislabeled as a "discrepancy."

Keeping these separate matters: comparing POS sales directly to the *net* payout would flag every single order as a mismatch, since commission is always deducted — that's expected behavior, not an error. Splitting the two turns the audit into a signal worth acting on.

```excel
=IFERROR(INDEX(Delivery_Platform_Report!$E:$E, MATCH(B2, Delivery_Platform_Report!$A:$A, 0)), 0)
```

## Dashboard

The **Executive Summary** sheet surfaces four headline metrics (Total POS Net Sales, Total Expected Payout, Total Discrepancy, Matched Orders Rate), a per-platform performance table, and two clean charts — order volume and sales by platform — kept as separate visuals rather than a single combined chart, since order counts and sales values aren't meaningfully comparable on the same axis.

![Reconciliation Audit](images/Reconciliation_Audit.png)

The **Reconciliation_Audit** sheet lists every order with its match status (Matched / Missing in Platform), the commission and VAT deducted, and an audit note explaining what action each flagged record needs.

## Key Insights

- Of 220 POS orders, **217 matched cleanly (98.6%)** against platform records
- The 3 unmatched orders were correctly flagged as **"Missing in Platform"** — orders the POS recorded but that never appeared in a delivery platform's settlement report, representing uncollected payouts worth following up on
- Total discrepancy across all matched orders came to just **524.86 SAR**, confirming the reconciliation logic isn't over-flagging normal commission deductions as errors
- Order volume and average order value both varied meaningfully by platform, useful context for platform-level commission negotiations

## Business Impact

Instead of manually checking hundreds of orders across five separate platform reports, a restaurant owner can drop in updated CSVs and get a clear, automatically-flagged list of exactly which orders need investigation — separating real problems (missing payouts) from expected costs (commission).

## Recommendations

- Follow up directly with the platform(s) responsible for the flagged missing orders to recover uncollected payouts
- Track the Matched Orders Rate over time as a simple health metric — a sudden drop signals a settlement issue worth escalating quickly
- Extend the same reconciliation key structure to any additional delivery platforms the business adds later

## Tools Used

Excel · Power Query · INDEX/MATCH · Data Validation & Auditing

---

*Sample data was generated for portfolio demonstration purposes.*
