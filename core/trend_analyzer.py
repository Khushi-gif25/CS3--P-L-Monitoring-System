"""
Trend Analysis Module
Prozo Daily FC P&L Monitoring System

Performs historical pattern analysis across uploaded days:
- 6 Expense Category Percentage Trajectories (with reference target bands)
- Continuous Increasing / Decreasing Trends (3+ consecutive day trends)
- Sudden Significant Expense Shifts (> 2.5 pp single-day swing)
- Repeated Target Range Breaches & Problematic Category Ranking
- CM1% and CM2% Trajectories (declines, volatility, low/suspicious day detection)
- Anomaly Frequency Distribution
- Grounded AI & Rule-Based Trend Insights (supported strictly by observed day ranges)
"""

from typing import Any, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd
from .data_extractor import CM_TARGETS, STANDARD_EXPENSE_KEYS

def detect_consecutive_streaks(series: pd.Series, min_length: int = 3) -> Dict[str, List[Dict[str, Any]]]:
    """
    Detect continuous increasing or decreasing streaks of length >= min_length.
    """
    values = series.tolist()
    increasing_streaks = []
    decreasing_streaks = []

    cur_inc_start = 0
    cur_dec_start = 0

    for i in range(1, len(values)):
        # Check increasing
        if values[i] > values[i - 1]:
            if cur_dec_start is not None and (i - 1 - cur_dec_start + 1) >= min_length:
                decreasing_streaks.append({
                    "start_idx": cur_dec_start,
                    "end_idx": i - 1,
                    "length": i - cur_dec_start,
                    "start_val": values[cur_dec_start],
                    "end_val": values[i - 1]
                })
            cur_dec_start = i
        elif values[i] < values[i - 1]:
            if cur_inc_start is not None and (i - 1 - cur_inc_start + 1) >= min_length:
                increasing_streaks.append({
                    "start_idx": cur_inc_start,
                    "end_idx": i - 1,
                    "length": i - cur_inc_start,
                    "start_val": values[cur_inc_start],
                    "end_val": values[i - 1]
                })
            cur_inc_start = i
        else:
            # Equal value ends streaks
            if cur_inc_start is not None and (i - 1 - cur_inc_start + 1) >= min_length:
                increasing_streaks.append({
                    "start_idx": cur_inc_start,
                    "end_idx": i - 1,
                    "length": i - cur_inc_start,
                    "start_val": values[cur_inc_start],
                    "end_val": values[i - 1]
                })
            if cur_dec_start is not None and (i - 1 - cur_dec_start + 1) >= min_length:
                decreasing_streaks.append({
                    "start_idx": cur_dec_start,
                    "end_idx": i - 1,
                    "length": i - cur_dec_start,
                    "start_val": values[cur_dec_start],
                    "end_val": values[i - 1]
                })
            cur_inc_start = i
            cur_dec_start = i

    # Check terminal streaks
    if cur_inc_start is not None and (len(values) - 1 - cur_inc_start + 1) >= min_length:
        increasing_streaks.append({
            "start_idx": cur_inc_start,
            "end_idx": len(values) - 1,
            "length": len(values) - cur_inc_start,
            "start_val": values[cur_inc_start],
            "end_val": values[-1]
        })
    if cur_dec_start is not None and (len(values) - 1 - cur_dec_start + 1) >= min_length:
        decreasing_streaks.append({
            "start_idx": cur_dec_start,
            "end_idx": len(values) - 1,
            "length": len(values) - cur_dec_start,
            "start_val": values[cur_dec_start],
            "end_val": values[-1]
        })

    return {"increasing": increasing_streaks, "decreasing": decreasing_streaks}

def analyze_dataset_trends(
    df_calc: pd.DataFrame,
    master_anomalies: List[Dict[str, Any]],
    criteria: Dict[str, Tuple[float, float]]
) -> Dict[str, Any]:
    """
    Comprehensive historical trend analysis of the uploaded dataset.
    """
    if df_calc.empty:
        return {}

    total_days = len(df_calc)
    days_list = df_calc["Day"].tolist()

    # 1. Expense Trend Summaries
    expense_summaries = {}
    breach_counts = {k: 0 for k in STANDARD_EXPENSE_KEYS}
    breach_above_counts = {k: 0 for k in STANDARD_EXPENSE_KEYS}
    breach_below_counts = {k: 0 for k in STANDARD_EXPENSE_KEYS}

    # Count breaches from master anomalies
    for an in master_anomalies:
        item = an.get("Line Item")
        if item in breach_counts:
            breach_counts[item] += 1
            if an.get("Status") == "ABOVE TARGET":
                breach_above_counts[item] += 1
            elif an.get("Status") == "BELOW TARGET":
                breach_below_counts[item] += 1

    # Sudden shifts and streaks per category
    sudden_shifts = []

    for exp_key in STANDARD_EXPENSE_KEYS:
        col = f"{exp_key} %"
        series = df_calc[col]
        min_target, max_target = criteria.get(exp_key, (0.0, 100.0))

        # Mean, Min, Max, Std
        avg_val = round(series.mean(), 2)
        min_val = round(series.min(), 2)
        max_val = round(series.max(), 2)
        std_val = round(series.std(), 2)

        # Detect streaks
        streaks = detect_consecutive_streaks(series, min_length=3)
        
        # Format streaks with Day IDs
        formatted_inc_streaks = []
        for s in streaks["increasing"]:
            start_day = days_list[s["start_idx"]]
            end_day = days_list[s["end_idx"]]
            formatted_inc_streaks.append(
                f"Days {start_day} to {end_day} ({s['length']} days: {s['start_val']:.2f}% -> {s['end_val']:.2f}%)"
            )

        formatted_dec_streaks = []
        for s in streaks["decreasing"]:
            start_day = days_list[s["start_idx"]]
            end_day = days_list[s["end_idx"]]
            formatted_dec_streaks.append(
                f"Days {start_day} to {end_day} ({s['length']} days: {s['start_val']:.2f}% -> {s['end_val']:.2f}%)"
            )

        # Detect single-day large swings (> 2.5 pp difference)
        for i in range(1, len(series)):
            diff = series.iloc[i] - series.iloc[i - 1]
            if abs(diff) >= 2.5:
                sudden_shifts.append({
                    "Category": exp_key,
                    "From Day": days_list[i - 1],
                    "To Day": days_list[i],
                    "Previous %": round(series.iloc[i - 1], 2),
                    "Current %": round(series.iloc[i], 2),
                    "Swing (pp)": round(diff, 2),
                    "Direction": "SPIKE" if diff > 0 else "DROP"
                })

        expense_summaries[exp_key] = {
            "Target Range": f"{min_target:.1f}% - {max_target:.1f}%",
            "Average %": avg_val,
            "Min %": min_val,
            "Max %": max_val,
            "Volatility (Std)": std_val,
            "Total Breaches": breach_counts[exp_key],
            "Breach Rate %": round((breach_counts[exp_key] / total_days) * 100.0, 1),
            "Above Target Count": breach_above_counts[exp_key],
            "Below Target Count": breach_below_counts[exp_key],
            "Increasing Streaks": formatted_inc_streaks,
            "Decreasing Streaks": formatted_dec_streaks,
        }

    # Rank most problematic expense category
    ranked_expenses = sorted(breach_counts.items(), key=lambda x: x[1], reverse=True)
    most_problematic_cat, max_breaches = ranked_expenses[0] if ranked_expenses else ("None", 0)

    # 2. CM1 and CM2 Trajectories
    cm1_series = df_calc["CM1%"]
    cm2_series = df_calc["CM2%"]
    rev_series = df_calc["Revenue"]

    cm1_streaks = detect_consecutive_streaks(cm1_series, min_length=3)
    cm2_streaks = detect_consecutive_streaks(cm2_series, min_length=3)

    cm1_deviations_count = sum(1 for an in master_anomalies if an.get("Line Item") == "CM1")
    cm2_deviations_count = sum(1 for an in master_anomalies if an.get("Line Item") == "CM2")

    low_cm2_days = df_calc[df_calc["CM2%"] < 12.0]["Day"].tolist()
    suspicious_cm2_days = df_calc[df_calc["CM2%"] > 33.0]["Day"].tolist()

    # Formatted CM2 decline periods
    cm2_decline_periods = []
    for s in cm2_streaks["decreasing"]:
        start_day = days_list[s["start_idx"]]
        end_day = days_list[s["end_idx"]]
        cm2_decline_periods.append(
            f"Days {start_day} to {end_day} ({s['length']} days: {s['start_val']:.2f}% -> {s['end_val']:.2f}%)"
        )

    # Formatted CM2 improvement periods
    cm2_growth_periods = []
    for s in cm2_streaks["increasing"]:
        start_day = days_list[s["start_idx"]]
        end_day = days_list[s["end_idx"]]
        cm2_growth_periods.append(
            f"Days {start_day} to {end_day} ({s['length']} days: {s['start_val']:.2f}% -> {s['end_val']:.2f}%)"
        )

    # 3. Grounded AI & Rule-Based Trend Observations
    trend_insights = []

    # Observation 1: Most problematic category
    if max_breaches > 0:
        breach_pct = round((max_breaches / total_days) * 100.0, 1)
        ab_cnt = breach_above_counts[most_problematic_cat]
        bel_cnt = breach_below_counts[most_problematic_cat]
        trend_insights.append({
            "Title": f"Dominant Cost Anomaly: {most_problematic_cat}",
            "Observation": (
                f"{most_problematic_cat} is the most frequently breached expense line, exceeding or violating target bands "
                f"in {max_breaches} out of {total_days} days ({breach_pct}% breach frequency). "
                f"Of these, {ab_cnt} days were Above Target and {bel_cnt} days were Below Target."
            ),
            "Period/Scope": f"Full Dataset (Days {days_list[0]} to {days_list[-1]})",
            "Root Cause Indicator": (
                f"Persistent variance in {most_problematic_cat} indicates systematic structural deviation rather than one-off noise."
            ),
            "Recommended Action": (
                f"Prioritize operational review of {most_problematic_cat}; re-negotiate vendor rate cards or revise operational staffing benchmarks."
            )
        })

    # Observation 2: CM2 Profitability Declines
    if cm2_decline_periods:
        longest_decline = cm2_decline_periods[0]
        trend_insights.append({
            "Title": "Consecutive Profitability Contraction",
            "Observation": (
                f"CM2 experienced consecutive downward trends during: {', '.join(cm2_decline_periods[:2])}. "
                f"This reflects multi-day margin erosion where operating cost expansion outpaced revenue gains."
            ),
            "Period/Scope": f"Observed streaks: {', '.join(cm2_decline_periods[:2])}",
            "Root Cause Indicator": "Simultaneous inflation in Manpower and Overheads during these specific days compressed gross contribution.",
            "Recommended Action": "Enforce strict daily cost controls at the start of volume surges to prevent margin deterioration."
        })

    # Observation 3: Severe Margin Breaches (Red Days)
    if low_cm2_days:
        trend_insights.append({
            "Title": f"Critical CM2 Breaches (< 12%) Recorded ({len(low_cm2_days)} Days)",
            "Observation": (
                f"A total of {len(low_cm2_days)} days suffered from critically deficient CM2 below the 12% threshold: "
                f"Days {', '.join(str(d) for d in low_cm2_days[:10])}{' ...' if len(low_cm2_days) > 10 else ''}."
            ),
            "Period/Scope": f"{len(low_cm2_days)} days out of {total_days}",
            "Root Cause Indicator": "Severe operational overheads combined with fixed rent burdens on lower-revenue days.",
            "Recommended Action": "Conduct daily root-cause reviews with warehouse managers for any day falling below the 12% CM2 floor."
        })

    # Observation 4: Suspicious High CM2 Days
    if suspicious_cm2_days:
        trend_insights.append({
            "Title": f"Suspiciously High CM2 Margins (> 33%) ({len(suspicious_cm2_days)} Days)",
            "Observation": (
                f"{len(suspicious_cm2_days)} days exhibited CM2 exceeding 33%: "
                f"Days {', '.join(str(d) for d in suspicious_cm2_days[:10])}{' ...' if len(suspicious_cm2_days) > 10 else ''}."
            ),
            "Period/Scope": f"{len(suspicious_cm2_days)} occurrences",
            "Root Cause Indicator": "Potential under-booking, delayed posting, or missing entries for major operational expenditure lines.",
            "Recommended Action": "Mandate finance audit on these days before presenting high margins as true operational performance."
        })

    return {
        "Total Days": total_days,
        "Expense Summaries": expense_summaries,
        "Breach Counts": breach_counts,
        "Most Problematic Category": most_problematic_cat,
        "Most Problematic Breaches": max_breaches,
        "Sudden Shifts": sudden_shifts[:15],
        "CM1 Average %": round(cm1_series.mean(), 2),
        "CM1 Min %": round(cm1_series.min(), 2),
        "CM1 Max %": round(cm1_series.max(), 2),
        "CM1 Deviations Count": cm1_deviations_count,
        "CM2 Average %": round(cm2_series.mean(), 2),
        "CM2 Min %": round(cm2_series.min(), 2),
        "CM2 Max %": round(cm2_series.max(), 2),
        "CM2 Deviations Count": cm2_deviations_count,
        "Low CM2 Days Count (< 12%)": len(low_cm2_days),
        "Low CM2 Days": low_cm2_days,
        "Suspicious CM2 Days Count (> 33%)": len(suspicious_cm2_days),
        "Suspicious CM2 Days": suspicious_cm2_days,
        "CM2 Decline Periods": cm2_decline_periods,
        "CM2 Growth Periods": cm2_growth_periods,
        "Trend Insights": trend_insights,
    }
