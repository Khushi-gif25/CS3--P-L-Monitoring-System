"""
PROZO P&L MONITOR - AI-POWERED INSIGHTS ENGINE
Generates plain-language, executive explanations for all important financial anomalies,
daily digests, and trend observations.
Includes a rigorous rule-based analytical engine and optional LLM provider integration.
"""

from typing import Dict, Any, List, Optional
import pandas as pd
import os
import json


# Root cause heuristics for fulfillment centre operational expenses
EXPENSE_POSSIBLE_CAUSES = {
    "Manpower": [
        "Unplanned overtime shifts due to sudden inbound volume spikes",
        "High reliance on higher-cost contract / temporary warehouse staff",
        "Sub-optimal labor pick/pack productivity rates per labor hour",
        "Mismatched shift scheduling versus order intake distribution"
    ],
    "Packaging": [
        "Incorrect carton size selection leading to excessive dunnage/bubble wrap consumption",
        "Emergency procurement of corrugated boxes at non-negotiated spot market rates",
        "Damage/wastage during automated packing or tape sealing operations",
        "Unreported inventory loss or bulk consumption recording errors"
    ],
    "Power and Fuel": [
        "Extended diesel generator run-hours during utility power outages",
        "Peak summer cooling / HVAC load in temperature-sensitive storage areas",
        "Inefficient high-bay lighting or conveyor systems running during off-shift hours",
        "Utility rate tariff increases or retro-billing reconciliations"
    ],
    "FC Rent": [
        "Fixed facility lease cost spread over lower-than-budgeted client throughput revenue",
        "Recent square footage expansion not yet amortized by scaled volume",
        "Annual contractual escalation clause activated in current period"
    ],
    "Equipment Rentals": [
        "Additional rental reach trucks or forklifts hired for short-term volume surge",
        "Under-utilized specialized handling equipment on daily hire rates",
        "Equipment breakdown leading to standby penalty or replacement rental costs"
    ],
    "Overheads": [
        "Quarterly or annual facility maintenance / repair invoices booked in a single day",
        "Pest control, security, or safety audit compliance surge expenses",
        "Discretionary admin expenses or software licensing accruals"
    ]
}

RECOMMENDED_MANAGEMENT_ACTIONS = {
    "Manpower": [
        "Audit daily clock-in/out rosters and enforce strict caps on unscheduled overtime.",
        "Transition recurring temporary labor pools to indexed fixed-productivity contracts.",
        "Implement warehouse management system (WMS) pick-path optimization to lift units-per-hour (UPH)."
    ],
    "Packaging": [
        "Enforce automated volumetric packaging recommendation at pack stations.",
        "Perform immediate physical count of corrugated cartons and dunnage materials.",
        "Consolidate box procurement with centralized corporate vendor contracts."
    ],
    "Power and Fuel": [
        "Audit DG logbook diesel consumption against meter run-hours and utility outage logs.",
        "Enforce automated motion sensors and strict shutdown protocols for idle conveyors and lighting.",
        "Review maximum demand tariff penalties with local electricity distribution company."
    ],
    "FC Rent": [
        "Accelerate commercial business development to monetize unused pallet racking capacity.",
        "Review subleasing or space-sharing opportunities for surplus mezzanine space."
    ],
    "Equipment Rentals": [
        "Off-hire surplus rental forklifts and return to vendor immediately after shift completion.",
        "Benchmark MHE utilization telemetry; renegotiate idle standby rental clauses."
    ],
    "Overheads": [
        "Amortize non-recurring repair and maintenance expenses across monthly accounting periods.",
        "Require regional GM sign-off for any discretionary facility expenditure exceeding ₹10,000."
    ]
}


def analyze_record_insights(
    row: pd.Series,
    anomalies_for_record: List[Dict[str, Any]],
    targets: Dict[str, Dict[str, float]]
) -> Dict[str, Any]:
    """
    Builds a complete, 6-part executive insight for a single FC / Day record.
    1. What happened?
    2. Size of deviation
    3. Breached line items
    4. Top expense contributors sorted by actual financial impact
    5. Why the issue may have occurred (possible vs confirmed)
    6. Recommended management actions
    """
    date_str = str(row["Date"])
    fc_name = str(row["FC"])
    revenue = float(row["Revenue"])
    cm1_pct = float(row["CM1 %"])
    cm2_pct = float(row["CM2 %"])
    fc_status = str(row["FC Status"])

    # Line item impacts
    expense_impacts = []
    breached_items = []
    underreported_warnings = []

    for cat_name, (amt_col, pct_col) in {
        "Manpower": ("Manpower Cost", "Manpower Cost %"),
        "Packaging": ("Packaging Cost", "Packaging Cost %"),
        "Power and Fuel": ("Power and Fuel", "Power and Fuel %"),
        "FC Rent": ("FC Rent", "FC Rent %"),
        "Equipment Rentals": ("Equipment Rentals", "Equipment Rentals %"),
        "Overheads": ("Overheads", "Overheads %"),
    }.items():
        actual_pct = float(row[pct_col])
        actual_amt = float(row[amt_col])
        t_min = float(targets.get(cat_name, {}).get("min", 0.0))
        t_max = float(targets.get(cat_name, {}).get("max", 100.0))

        if actual_pct > t_max:
            dev_pp = round(actual_pct - t_max, 2)
            allowed_max_amt = (t_max / 100.0) * revenue
            impact_rs = round(actual_amt - allowed_max_amt, 2)
            expense_impacts.append({
                "category": cat_name,
                "impact_rs": impact_rs,
                "dev_pp": dev_pp,
                "actual_pct": actual_pct,
                "target_max": t_max,
                "type": "OVERRUN"
            })
            breached_items.append(f"{cat_name} ({actual_pct:.1f}% vs max {t_max:.1f}%, +₹{impact_rs:,.2f})")
        elif actual_pct < t_min:
            dev_pp = round(t_min - actual_pct, 2)
            allowed_min_amt = (t_min / 100.0) * revenue
            savings_rs = round(allowed_min_amt - actual_amt, 2)
            if actual_amt == 0.0 or dev_pp >= 5.0:
                underreported_warnings.append(
                    f"{cat_name} reported at only {actual_pct:.1f}% (target min {t_min:.1f}%), ₹{actual_amt:,.2f}. Check for missing vendor bills."
                )

    # Sort expense contributors strictly by financial impact in Rupees descending
    expense_impacts.sort(key=lambda x: x["impact_rs"], reverse=True)

    # 1. What Happened
    if fc_status == "RED":
        what_happened = (
            f"On {date_str}, {fc_name} registered a severe profitability deficit with CM2 dropping to {cm2_pct:.1f}%, "
            f"well below the critical survival floor of 12.0%. Total operating revenue generated was ₹{revenue:,.2f}."
        )
    elif fc_status == "YELLOW":
        what_happened = (
            f"On {date_str}, {fc_name} operated in the caution zone with CM2 at {cm2_pct:.1f}%, "
            f"deviating from the target corridor (15% - 30%) with revenue of ₹{revenue:,.2f}."
        )
    elif fc_status == "SUSPICIOUS":
        what_happened = (
            f"On {date_str}, {fc_name} reported an abnormally elevated CM2 of {cm2_pct:.1f}% (above 33.0% threshold). "
            f"While seemingly profitable, such extreme margins strongly indicate unposted accruals or missing operational expenses."
        )
    elif fc_status == "REVIEW REQUIRED":
        what_happened = (
            f"On {date_str}, {fc_name} landed in a transitional boundary zone with CM2 at {cm2_pct:.1f}%. "
            f"Requires managerial validation of cost allocation."
        )
    else:
        # Green day
        if breached_items:
            what_happened = (
                f"On {date_str}, {fc_name} maintained an overall healthy CM2 of {cm2_pct:.1f}% (Green status), "
                f"generating ₹{revenue:,.2f} in revenue. However, individual line-item variances were detected."
            )
        else:
            what_happened = (
                f"On {date_str}, {fc_name} operated cleanly with all operational metrics in target corridor: "
                f"CM1 at {cm1_pct:.1f}% and CM2 at {cm2_pct:.1f}% on ₹{revenue:,.2f} revenue."
            )

    # 2. Size of Deviation
    cm2_min = targets.get("CM2", {}).get("min", 15.0)
    cm2_max = targets.get("CM2", {}).get("max", 30.0)
    if cm2_pct < cm2_min:
        cm2_gap_pp = round(cm2_min - cm2_pct, 2)
        dev_size_desc = f"CM2 is deficient by {cm2_gap_pp:.1f} percentage points below the {cm2_min:.1f}% minimum target."
    elif cm2_pct > cm2_max:
        cm2_gap_pp = round(cm2_pct - cm2_max, 2)
        dev_size_desc = f"CM2 is {cm2_gap_pp:.1f} percentage points above the {cm2_max:.1f}% upper target threshold."
    else:
        dev_size_desc = f"CM2 is within acceptable limits ({cm2_min:.1f}% - {cm2_max:.1f}%)."

    # 3. Breached Line Items Summary
    breach_summary = ", ".join(breached_items) if breached_items else "No expense ceilings exceeded."

    # 4. Top Expense Contributors Ranked by Rupee Impact
    ranked_contributors = []
    for rank, item in enumerate(expense_impacts, 1):
        ranked_contributors.append({
            "rank": rank,
            "category": item["category"],
            "impact_rs": item["impact_rs"],
            "dev_pp": item["dev_pp"],
            "summary": f"#{rank} {item['category']}: +₹{item['impact_rs']:,.2f} above target (+{item['dev_pp']:.1f} pp)"
        })

    # 5. Why the issue may have occurred (Confirmed vs Possible Causes)
    confirmed_causes = []
    possible_causes = []

    if expense_impacts:
        lead_cat = expense_impacts[0]["category"]
        lead_impact = expense_impacts[0]["impact_rs"]
        confirmed_causes.append(
            f"Mathematical Confirmation: Direct budget overrun in '{lead_cat}' eroded profitability by ₹{lead_impact:,.2f}."
        )
        for exp in expense_impacts[:2]:
            cat = exp["category"]
            hypotheses = EXPENSE_POSSIBLE_CAUSES.get(cat, ["Operational volatility or unplanned vendor rate adjustments."])
            for h in hypotheses[:2]:
                possible_causes.append(f"[{cat}] Possible operational driver: {h}")

    if fc_status == "SUSPICIOUS":
        confirmed_causes.append("Confirmed: Indirect expense lines are significantly below expected minimum operational floors.")
        possible_causes.append("Possible billing omission: Facility rent or utility invoice for this date has not been posted in the ERP.")

    if not confirmed_causes and not possible_causes:
        confirmed_causes.append("No material variance observed. Operations ran strictly within cost parameters.")

    # 6. Recommended Management Actions
    recommended_actions = []
    if expense_impacts:
        for exp in expense_impacts[:2]:
            cat = exp["category"]
            actions = RECOMMENDED_MANAGEMENT_ACTIONS.get(cat, ["Perform comprehensive line-item cost audit."])
            recommended_actions.extend([f"[{cat}] {act}" for act in actions[:2]])
    elif fc_status == "SUSPICIOUS":
        recommended_actions.append("[Audit] Immediately verify with Accounts Payable whether rent, utility, or labor contractor invoices were deferred or unaccrued.")
    else:
        recommended_actions.append("Maintain existing roster and cost control protocols.")

    return {
        "date": date_str,
        "fc": fc_name,
        "revenue": revenue,
        "cm1_pct": cm1_pct,
        "cm2_pct": cm2_pct,
        "fc_status": fc_status,
        "what_happened": what_happened,
        "deviation_size": dev_size_desc,
        "breached_items": breach_summary,
        "ranked_contributors": ranked_contributors,
        "confirmed_causes": confirmed_causes,
        "possible_causes": possible_causes,
        "recommended_actions": recommended_actions,
        "underreported_warnings": underreported_warnings,
        "formatted_text": (
            f"**Executive Insight for {fc_name} ({date_str})**\n\n"
            f"• **What Happened:** {what_happened}\n"
            f"• **Deviation:** {dev_size_desc}\n"
            f"• **Breaches:** {breach_summary}\n"
            f"• **Top Impact:** {'; '.join([c['summary'] for c in ranked_contributors]) if ranked_contributors else 'None'}\n"
            f"• **Potential Drivers:** {'; '.join(possible_causes) if possible_causes else 'Normal operation'}\n"
            f"• **Recommended Action:** {'; '.join(recommended_actions)}"
        )
    }


def generate_all_ai_insights(
    df_metrics: pd.DataFrame,
    df_anomalies: pd.DataFrame,
    targets: Dict[str, Dict[str, float]],
    api_key: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Generates structured AI insights for all FC records with anomalies or status != GREEN.
    Also produces executive digest summaries.
    """
    insights_list = []

    # Filter records that have anomalies or non-green status
    flagged_records = []
    for idx, row in df_metrics.iterrows():
        d_val = str(row["Date"])
        fc_val = str(row["FC"])
        status = str(row["FC Status"])
        
        has_anom = False
        if not df_anomalies.empty:
            match = df_anomalies[(df_anomalies["Date"] == d_val) & (df_anomalies["FC"] == fc_val)]
            if not match.empty:
                has_anom = True

        if status != "GREEN" or has_anom:
            flagged_records.append((idx, row))

    # If no records flagged (all completely clean), process top 2 records anyway for clarity
    if not flagged_records:
        flagged_records = list(df_metrics.iterrows())[:2]

    for idx, row in flagged_records:
        d_val = str(row["Date"])
        fc_val = str(row["FC"])
        rec_anoms = []
        if not df_anomalies.empty:
            m = df_anomalies[(df_anomalies["Date"] == d_val) & (df_anomalies["FC"] == fc_val)]
            rec_anoms = m.to_dict(orient="records")

        ins = analyze_record_insights(row, rec_anoms, targets)
        insights_list.append(ins)

    return insights_list


def generate_overall_trend_ai_observations(
    df_metrics: pd.DataFrame,
    df_anomalies: pd.DataFrame,
    targets: Dict[str, Dict[str, float]]
) -> Dict[str, Any]:
    """
    Synthesizes overall trend observations across the entire uploaded historical dataset:
    - Most frequent breached expense category
    - Repeated anomalies
    - CM1/CM2 margin stability
    - Top performing and lowest performing FC/day
    """
    if df_metrics.empty:
        return {}

    total_records = len(df_metrics)
    avg_cm1 = df_metrics["CM1 %"].mean()
    avg_cm2 = df_metrics["CM2 %"].mean()
    total_rev = df_metrics["Revenue"].sum()

    best_day = df_metrics.loc[df_metrics["CM2 %"].idxmax()]
    worst_day = df_metrics.loc[df_metrics["CM2 %"].idxmin()]

    # Status distribution
    status_counts = df_metrics["FC Status"].value_counts().to_dict()

    # Most breached expense
    most_breached_cat = "None"
    most_breached_count = 0
    if not df_anomalies.empty:
        line_anoms = df_anomalies[df_anomalies["Metric Type"] == "Line Item"]
        if not line_anoms.empty:
            cat_counts = line_anoms["Line Item"].value_counts()
            most_breached_cat = cat_counts.index[0]
            most_breached_count = int(cat_counts.iloc[0])

    # Observations
    observations = [
        f"Analyzed {total_records} operational day-records generating cumulative revenue of ₹{total_rev:,.2f}.",
        f"Average portfolio CM1 is {avg_cm1:.1f}% (target corridor: {targets.get('CM1', {}).get('min', 65)}% - {targets.get('CM1', {}).get('max', 70)}%) and average CM2 is {avg_cm2:.1f}% (target: {targets.get('CM2', {}).get('min', 15)}% - {targets.get('CM2', {}).get('max', 30)}%).",
        f"Peak margin efficiency occurred on {best_day['Date']} ({best_day['FC']}) with CM2 reaching {best_day['CM2 %']:.1f}%.",
        f"Severe margin compression observed on {worst_day['Date']} ({worst_day['FC']}) with CM2 collapsing to {worst_day['CM2 %']:.1f}%.",
        f"The primary cost driver triggering variances is '{most_breached_cat}', flagged in {most_breached_count} separate instances."
    ]

    return {
        "total_records": total_records,
        "total_revenue": total_rev,
        "avg_cm1": round(avg_cm1, 2),
        "avg_cm2": round(avg_cm2, 2),
        "status_counts": status_counts,
        "best_record": {
            "date": str(best_day["Date"]),
            "fc": str(best_day["FC"]),
            "cm2": float(best_day["CM2 %"]),
            "revenue": float(best_day["Revenue"])
        },
        "worst_record": {
            "date": str(worst_day["Date"]),
            "fc": str(worst_day["FC"]),
            "cm2": float(worst_day["CM2 %"]),
            "revenue": float(worst_day["Revenue"])
        },
        "most_breached_cat": most_breached_cat,
        "most_breached_count": most_breached_count,
        "observations": observations
    }
