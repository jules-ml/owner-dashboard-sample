#!/usr/bin/env python3
"""Owner dashboard for the Lantern Key Pool & Air sample.

SAMPLE: synthetic data for a fictional business. Not a real company and not real results.
Run from the repository root:

    streamlit run app.py

Loads tables from data/clean/. If those files are missing, the cleaning step runs on startup.
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

ROOT = Path(__file__).resolve().parent
CLEAN = ROOT / "data" / "clean"

st.set_page_config(
    page_title="Sample owner dashboard — Lantern Key Pool & Air",
    layout="wide",
    initial_sidebar_state="collapsed",
)


REQUIRED_CLEAN = ("kpis.json", "data_quality_log.md", "invoices.csv", "customers.csv")


def ensure_clean_data() -> None:
    """Load prebuilt clean tables, or run the cleaning pipeline if they are missing."""
    if all((CLEAN / name).exists() for name in REQUIRED_CLEAN):
        return
    import clean_data

    clean_data.main()


ensure_clean_data()


@st.cache_data
def load():
    kpis = json.loads((CLEAN / "kpis.json").read_text(encoding="utf-8"))
    quality_md = (CLEAN / "data_quality_log.md").read_text(encoding="utf-8")
    invoices = pd.read_csv(CLEAN / "invoices.csv")
    flagged = invoices[invoices["status"].isin(["Flagged", "Overpaid"])].copy()
    customers = pd.read_csv(CLEAN / "customers.csv")
    return kpis, quality_md, flagged, customers


kpis, quality_md, flagged, customers = load()
t = kpis["totals"]
q = kpis["quality"]
monthly = pd.DataFrame(kpis["monthly"])
lines = pd.DataFrame(kpis["service_lines"])
aging = pd.DataFrame(kpis["aging"])
chase = pd.DataFrame(kpis["chase_first"])
cash = pd.DataFrame(kpis["cash_monthly"])

money0 = lambda x: f"${x:,.0f}"
money2 = lambda x: f"${x:,.2f}"
pct = lambda x: f"{x*100:.1f}%" if x is not None and pd.notna(x) else "—"

st.title("Owner dashboard")
st.error(
    "SAMPLE: synthetic data for a fictional business — Lantern Key Pool & Air "
    "(Delray Beach / Palm Beach County, FL area). Not a real company, not real customers, "
    "and not real results."
)
st.caption(
    f"As of {kpis['as_of']} (America/New_York). Source window: Jan 6–Nov 20, 2026. "
    "Revenue is pre-tax invoice totals. Gross profit is revenue minus direct job labor and materials. "
    "Cash out treats those job costs as paid in the job month, plus overhead expenses."
)

tabs = st.tabs(
    ["Overview", "Profit by service line", "Accounts receivable", "Cash in vs out", "Data quality"]
)

# ---------------------------------------------------------------------------
# Overview
# ---------------------------------------------------------------------------
with tabs[0]:
    st.write(
        "On this sample book, the shop is profitable on an accrual basis but cash is behind, "
        "mostly because of past-due invoices."
    )
    c = st.columns(4)
    c[0].metric("Revenue", money0(t["revenue"]), help="Pre-tax invoices with a parseable date and amount.")
    c[1].metric("Gross profit", money0(t["gross_profit"]), pct(t["margin_pct"]))
    c[2].metric("Open AR", money0(t["open_ar"]))
    c[3].metric("Net cash", money0(t["net_cash"]), help="Payments received minus job costs and overhead.")

    fig = go.Figure()
    fig.add_bar(name="Revenue", x=monthly["label"], y=monthly["revenue"], marker_color="#1f77b4")
    fig.add_scatter(
        name="Gross profit",
        x=monthly["label"],
        y=monthly["gross_profit"],
        mode="lines+markers",
        line=dict(color="#2ca02c", width=3),
    )
    fig.update_layout(
        title="Revenue and gross profit by month",
        yaxis_tickprefix="$",
        yaxis_tickformat=",.0f",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        height=420,
        margin=dict(t=60, b=40, l=40, r=20),
        bargap=0.25,
    )
    st.plotly_chart(fig, width="stretch")

    mt = monthly.copy()
    mt["margin"] = mt["margin_pct"].map(pct)
    st.dataframe(
        mt.rename(columns={
            "label": "Month",
            "revenue": "Revenue",
            "direct_cost": "Direct cost",
            "gross_profit": "Gross profit",
            "margin": "Margin",
        })[["Month", "Revenue", "Direct cost", "Gross profit", "Margin"]],
        width="stretch",
        hide_index=True,
        column_config={
            "Revenue": st.column_config.NumberColumn(format="$%.2f"),
            "Direct cost": st.column_config.NumberColumn(format="$%.2f"),
            "Gross profit": st.column_config.NumberColumn(format="$%.2f"),
        },
    )

# ---------------------------------------------------------------------------
# Service lines
# ---------------------------------------------------------------------------
with tabs[1]:
    st.write(
        "Weekly pool service has the strongest margin in this sample. Equipment installs "
        "bring in the most revenue but the thinnest margin."
    )
    lt = lines.sort_values("gross_profit", ascending=False).copy()
    fig = go.Figure(
        go.Bar(
            x=lt["gross_profit"],
            y=lt["service_line"],
            orientation="h",
            text=lt["gross_profit"].map(money0),
            textposition="outside",
            marker_color="#1f77b4",
        )
    )
    fig.update_layout(
        title="Gross profit by service line",
        xaxis_tickprefix="$",
        xaxis_tickformat=",.0f",
        height=360,
        margin=dict(t=60, b=40, l=20, r=80),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig, width="stretch")

    show = lt.copy()
    show["margin"] = show["margin_pct"].map(pct)
    st.dataframe(
        show.rename(columns={
            "service_line": "Service line",
            "revenue": "Revenue",
            "direct_cost": "Direct cost",
            "gross_profit": "Gross profit",
            "margin": "Margin",
            "invoice_count": "Invoices",
        })[["Service line", "Revenue", "Direct cost", "Gross profit", "Margin", "Invoices"]],
        width="stretch",
        hide_index=True,
        column_config={
            "Revenue": st.column_config.NumberColumn(format="$%.2f"),
            "Direct cost": st.column_config.NumberColumn(format="$%.2f"),
            "Gross profit": st.column_config.NumberColumn(format="$%.2f"),
        },
    )

# ---------------------------------------------------------------------------
# AR
# ---------------------------------------------------------------------------
with tabs[2]:
    past_due = sum(r["balance"] for r in kpis["aging"] if r["bucket"] != "Current")
    st.write(
        f"Open accounts receivable is {money2(t['open_ar'])}, of which {money2(past_due)} is past due. "
        "The list below is who to call first, ranked by past-due balance and age."
    )
    a, b = st.columns(2)
    with a:
        fig = go.Figure(
            go.Bar(
                x=aging["bucket"],
                y=aging["balance"],
                text=aging["balance"].map(money0),
                textposition="outside",
                marker_color=["#7f7f7f", "#1f77b4", "#1f77b4", "#ff7f0e", "#d62728"],
            )
        )
        fig.update_layout(
            title="AR aging",
            yaxis_tickprefix="$",
            yaxis_tickformat=",.0f",
            height=360,
            margin=dict(t=60, b=40, l=40, r=20),
        )
        st.plotly_chart(fig, width="stretch")
    with b:
        st.dataframe(
            aging.rename(columns={"bucket": "Bucket", "balance": "Balance", "invoice_count": "Invoices"}),
            width="stretch",
            hide_index=True,
            column_config={"Balance": st.column_config.NumberColumn(format="$%.2f")},
        )

    st.subheader("Chase first")
    ct = chase.copy()
    st.dataframe(
        ct.rename(columns={
            "customer_name": "Customer",
            "balance": "Past due",
            "total_open_ar": "Total open",
            "days_past_due": "Oldest (days)",
            "aging_bucket": "Bucket",
            "past_due_count": "Invoices",
            "reason": "Why chase",
        })[["Customer", "Past due", "Total open", "Oldest (days)", "Bucket", "Invoices", "Why chase"]],
        width="stretch",
        hide_index=True,
        column_config={
            "Past due": st.column_config.NumberColumn(format="$%.2f"),
            "Total open": st.column_config.NumberColumn(format="$%.2f"),
        },
    )

# ---------------------------------------------------------------------------
# Cash
# ---------------------------------------------------------------------------
with tabs[3]:
    st.write(
        "Cash in is payments received in the month. Cash out is direct job costs (treated as paid "
        "in the job month) plus overhead such as rent, insurance, fuel, and the truck lease. "
        "That lag versus revenue is the accounts-receivable gap."
    )
    fig = go.Figure()
    fig.add_bar(name="Cash in", x=cash["label"], y=cash["cash_in"], marker_color="#1f77b4")
    fig.add_bar(name="Cash out", x=cash["label"], y=cash["cash_out"], marker_color="#ff7f0e")
    fig.update_layout(
        title="Cash in vs cash out",
        barmode="group",
        yaxis_tickprefix="$",
        yaxis_tickformat=",.0f",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        height=420,
        margin=dict(t=60, b=40, l=40, r=20),
    )
    st.plotly_chart(fig, width="stretch")

    ct = cash.copy()
    st.dataframe(
        ct.rename(columns={
            "label": "Month",
            "cash_in": "Cash in",
            "job_costs": "Job costs",
            "overhead": "Overhead",
            "cash_out": "Cash out",
            "net_cash": "Net cash",
        })[["Month", "Cash in", "Job costs", "Overhead", "Cash out", "Net cash"]],
        width="stretch",
        hide_index=True,
        column_config={
            "Cash in": st.column_config.NumberColumn(format="$%.2f"),
            "Job costs": st.column_config.NumberColumn(format="$%.2f"),
            "Overhead": st.column_config.NumberColumn(format="$%.2f"),
            "Cash out": st.column_config.NumberColumn(format="$%.2f"),
            "Net cash": st.column_config.NumberColumn(format="$%.2f"),
        },
    )

# ---------------------------------------------------------------------------
# Data quality
# ---------------------------------------------------------------------------
with tabs[4]:
    st.write(
        "Three messy exports went in: a combined invoices-and-payments CSV, a jobs-and-expenses "
        "workbook with headers not on row 1, and a CRM lead export. The Python pipeline standardized "
        "dates and currency, merged duplicate customer names, dropped exact duplicate rows, and "
        "flagged anything it could not fix with confidence."
    )
    c = st.columns(4)
    c[0].metric("Invoice rows in", f"{q['raw_invoice_rows']}", help="Raw rows in invoices_payments.csv, including preamble-skipped header.")
    c[1].metric("Clean invoices", f"{q['clean_invoices']}")
    c[2].metric("Customers after merge", f"{q['customers_after_merge']}", f"{q['customer_variants_in']} name variants in")
    c[3].metric("Rows flagged for review", f"{q['rows_flagged']}")

    st.subheader("What the cleaning step did")
    before_after = pd.DataFrame([
        ["Raw invoice-export rows", q["raw_invoice_rows"], ""],
        ["Exact duplicate invoice rows dropped", q["duplicate_invoice_rows_dropped"], "Kept the first copy"],
        ["Clean invoices written", q["clean_invoices"], "Payment-split rows collapsed to one invoice"],
        ["Clean payments written", q["clean_payments"], ""],
        ["Customer name variants in the invoice export", q["customer_variants_in"], ""],
        ["Customers after merge", q["customers_after_merge"], f"{q['customer_groups_merged']} groups had more than one name spelling"],
        ["Dates parsed (invoice, due, and payment fields)", q["dates_standardized"], "1 typo year (2027→2026) was corrected"],
        ["Currency amounts parsed", q["amounts_parsed"], "US and European formats"],
        ["Service lines standardized", q["service_lines_standardized"], "Mapped aliases like 'pool clean' and 'A/C Repair'"],
        ["Unparseable dates flagged", q["unparseable_dates"], "Left out of monthly charts"],
        ["Unparseable amounts flagged", q["unparseable_amounts"], "Left out of revenue and AR totals"],
        ["Jobs with no matching customer", q["orphan_jobs"], "Not assigned to a customer"],
        ["Duplicate expense rows dropped", q["duplicate_expense_rows_dropped"], ""],
        ["Clean jobs written", q["clean_jobs"], "Includes 2 unmatched jobs and 1 cancelled job"],
        ["Clean expenses written", q["clean_expenses"], ""],
        ["Clean leads written", q["clean_leads"], ""],
    ], columns=["Step", "Count", "Note"])
    before_after["Count"] = before_after["Count"].astype(str)
    st.dataframe(before_after, width="stretch", hide_index=True)

    st.subheader("Example: name variants merged")
    show_c = customers[customers["variant_count"] >= 4][
        ["customer_name", "variant_count", "name_variants"]
    ].rename(columns={
        "customer_name": "Canonical name",
        "variant_count": "Variants",
        "name_variants": "Spellings found in the export",
    })
    st.dataframe(show_c, width="stretch", hide_index=True)

    if not flagged.empty:
        st.subheader("Invoices flagged for review")
        st.dataframe(
            flagged[["invoice_id", "customer_name", "invoice_date", "amount", "amount_paid", "balance", "status", "notes"]].rename(columns={
                "invoice_id": "Invoice",
                "customer_name": "Customer",
                "invoice_date": "Date",
                "amount": "Amount",
                "amount_paid": "Paid",
                "balance": "Balance",
                "status": "Status",
                "notes": "Notes",
            }),
            width="stretch",
            hide_index=True,
        )

    with st.expander("Full data-quality log"):
        st.markdown(quality_md)

st.divider()
st.caption(
    "Portfolio sample for Giulio Piccinonna. SAMPLE: synthetic data for a fictional business. "
    "Not a real company and not real results."
)
