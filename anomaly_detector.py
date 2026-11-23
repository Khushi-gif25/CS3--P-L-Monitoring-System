"""
PROZO P&L MONITOR - ANOMALY DETECTION ENGINE
Identifies line item variances, CM1/CM2 margin breaches,
computes financial impact in currency and percentage points,
and assigns severity levels.
"""

from typing import Dict, Any, List, Optional
import pandas as pd


# Mapping canonical expense names to data columns
EXPENSE_TO_COL = {
    "Manpower": ("Manpower Cost", "Manpower Cost %"),
    "Packaging": ("Packaging Cost", "Packaging Cost %"),
    "Power and Fuel": ("Power and Fuel", "Power and Fuel %"),
    "FC Rent": ("FC Rent", "FC Rent %"),
    "Equipment Rentals": ("Equipment Rentals", "Equipment Rentals %"),
    "Overheads": ("Overheads", "Overheads %"),
}


def detect_anomalies_for_dataset(
    df_metrics: pd.DataFrame,
    targets: Dict[str, Dict[str, float]],
    cm_threshold_pp: float = 3.0
) -> pd.DataFrame:
    """
    Evaluates every record against current active targets:
    1. Line item expenses: NORMAL, ABOVE TARGET, BELOW TARGET.
    2. CM1 & CM2 breaches > 3 percentage points beyond target corridor.
    3. Calculates financial variance in Rupee amount and Percentage Points.
    4. Categorizes severity (Critical, High, Medium, Low/Info).
    5. Flags suspicious high CM2.
    """
    anomalies: List[Dict[str, Any]] = []

    for idx, row in df_metrics.iterrows():
        record_date = str(row["Date"])
        fc_name = str(row["FC"])
        revenue = float(row["Revenue"])
        cm1_pct = float(row["CM1 %"])
        cm2_pct = float(row["CM2 %"])
        fc_status = str(row["FC Status"])

        # 1. Evaluate Line Items
        for cat_name, (amt_col, pct_col) in EXPENSE_TO_COL.items():
            if cat_name not in targets:
                continue

            target_min = float(targets[cat_name]["min"])
            target_max = float(targets[cat_name]["max"])
            actual_pct = float(row[pct_col])
            actual_amt = float(row[amt_col])
            target_range_str = f"{target_min:.1f}% - {target_max:.1f}%"

            status = "NORMAL"
            dev_pp = 0.0
            financial_impact = 0.0  # Rupee variance from target threshold
            severity = "Normal"
            explanation = ""

            if actual_pct > target_max:
                status = "ABOVE TARGET"
                dev_pp = round(actual_pct - target_max, 2)
                # Financial impact: how many rupees above maximum allowable target
                allowed_max_amt = (target_max / 100.0) * revenue
                financial_impact = round(actual_amt - allowed_max_amt, 2)

                # Severity based on deviation magnitude
                if dev_pp >= 5.0 or (financial_impact >= 25000 and dev_pp >= 3.0):
                    severity = "Critical"
                elif dev_pp >= 2.0:
                    severity = "High"
                else:
                    severity = "Medium"

                explanation = (
                    f"{cat_name} exceeded ceiling by +{dev_pp:.1f} pp "
                    f"({actual_pct:.1f}% vs max {target_max:.1f}%). "
                    f"Excess cost impact: ₹{financial_impact:,.2f}."
                )

            elif actual_pct < target_min:
                status = "BELOW TARGET"
                dev_pp = round(actual_pct - target_min, 2)  # negative
                allowed_min_amt = (target_min / 100.0) * revenue
                financial_impact = round(actual_amt - allowed_min_amt, 2)  # negative impact

                # Deviation below target: can indicate savings OR underreported/missing bills
                if abs(dev_pp) >= 5.0:
                    severity = "High"
                    explanation = (
                        f"{cat_name} is substantially below target floor by {dev_pp:.1f} pp "
                        f"({actual_pct:.1f}% vs min {target_min:.1f}%). "
                        f"Caution: Check for delayed vendor invoices or underreported costs."
                    )
                else:
                    severity = "Low"
                    explanation = (
                        f"{cat_name} came in below minimum target by {dev_pp:.1f} pp "
                        f"({actual_pct:.1f}% vs min {target_min:.1f}%). Positive cost control or accrual lag."
                    )

            if status != "NORMAL":
                anomalies.append({
                    "Date": record_date,
                    "FC": fc_name,
                    "Line Item": cat_name,
                    "Actual %": actual_pct,
                    "Actual Amount": actual_amt,
                    "Target Range": target_range_str,
                    "Target Min %": target_min,
                    "Target Max %": target_max,
                    "Status": status,
                    "Deviation (pp)": dev_pp,
                    "Financial Impact (₹)": financial_impact,
                    "Anomaly Category": f"{cat_name} Cost Variance",
                    "Severity": severity,
                    "Explanation": explanation,
                    "FC Status": fc_status,
                    "Metric Type": "Line Item"
                })

        # 2. Evaluate CM1 Profitability Target (Target: 65% - 70%, flag if > 3pp outside)
        if "CM1" in targets:
            cm1_min = float(targets["CM1"]["min"])
            cm1_max = float(targets["CM1"]["max"])
            cm1_range_str = f"{cm1_min:.1f}% - {cm1_max:.1f}%"

            cm1_breached = False
            dev_pp = 0.0
            cm1_status = "NORMAL"
            severity = "Normal"
            explanation = ""

            if cm1_pct < (cm1_min - cm_threshold_pp):
                cm1_breached = True
                dev_pp = round(cm1_pct - cm1_min, 2)
                cm1_status = "BELOW TARGET"
                severity = "Critical" if abs(dev_pp) >= 5.0 else "High"
                deficit_amt = round(((cm1_min - cm1_pct) / 100.0) * revenue, 2)
                explanation = (
                    f"CM1 breached lower margin threshold by {abs(dev_pp):.1f} pp "
                    f"({cm1_pct:.1f}% vs min {cm1_min:.1f}%). Margin deficit: ₹{deficit_amt:,.2f}. "
                    f"Driven by high direct costs (Manpower or Packaging)."
                )
            elif cm1_pct > (cm1_max + cm_threshold_pp):
                cm1_breached = True
                dev_pp = round(cm1_pct - cm1_max, 2)
                cm1_status = "ABOVE TARGET"
                severity = "Medium"
                explanation = (
                    f"CM1 exceeded upper target corridor by +{dev_pp:.1f} pp "
                    f"({cm1_pct:.1f}% vs max {cm1_max:.1f}%)."
                )

            if cm1_breached:
                anomalies.append({
                    "Date": record_date,
                    "FC": fc_name,
                    "Line Item": "Contribution Margin 1 (CM1)",
                    "Actual %": cm1_pct,
                    "Actual Amount": float(row["CM1"]),
                    "Target Range": cm1_range_str,
                    "Target Min %": cm1_min,
                    "Target Max %": cm1_max,
                    "Status": cm1_status,
                    "Deviation (pp)": dev_pp,
                    "Financial Impact (₹)": round((dev_pp / 100.0) * revenue, 2),
                    "Anomaly Category": "CM1 Margin Variance",
                    "Severity": severity,
                    "Explanation": explanation,
                    "FC Status": fc_status,
                    "Metric Type": "Profitability (CM1)"
                })

        # 3. Evaluate CM2 Profitability Target (Target: 15% - 30%, flag if > 3pp outside)
        if "CM2" in targets:
            cm2_min = float(targets["CM2"]["min"])
            cm2_max = float(targets["CM2"]["max"])
            cm2_range_str = f"{cm2_min:.1f}% - {cm2_max:.1f}%"

            cm2_breached = False
            dev_pp = 0.0
            cm2_status = "NORMAL"
            severity = "Normal"
            explanation = ""

            if cm2_pct < (cm2_min - cm_threshold_pp):
                cm2_breached = True
                dev_pp = round(cm2_pct - cm2_min, 2)
                cm2_status = "BELOW TARGET"
                severity = "Critical" if cm2_pct < 12.0 else "High"
                deficit_amt = round(((cm2_min - cm2_pct) / 100.0) * revenue, 2)
                explanation = (
                    f"CM2 fell below operational target corridor by {abs(dev_pp):.1f} pp "
                    f"({cm2_pct:.1f}% vs min {cm2_min:.1f}%). Profitability shortfall: ₹{deficit_amt:,.2f}."
                )
            elif cm2_pct > (cm2_max + cm_threshold_pp):
                cm2_breached = True
                dev_pp = round(cm2_pct - cm2_max, 2)
                cm2_status = "ABOVE TARGET"
                # If CM2 > 33%, it triggers SUSPICIOUS
                if cm2_pct > 33.0:
                    severity = "Critical"
                    explanation = (
                        f"SUSPICIOUS: CM2 ({cm2_pct:.1f}%) is abnormally high (+{dev_pp:.1f} pp over target). "
                        f"Likely missing, omitted, or unbilled indirect expenses (e.g. rent or power)."
                    )
                else:
                    severity = "Medium"
                    explanation = (
                        f"CM2 exceeded upper bound by +{dev_pp:.1f} pp "
                        f"({cm2_pct:.1f}% vs max {cm2_max:.1f}%)."
                    )

            if cm2_breached:
                anomalies.append({
                    "Date": record_date,
                    "FC": fc_name,
                    "Line Item": "Contribution Margin 2 (CM2)",
                    "Actual %": cm2_pct,
                    "Actual Amount": float(row["CM2"]),
                    "Target Range": cm2_range_str,
                    "Target Min %": cm2_min,
                    "Target Max %": cm2_max,
                    "Status": cm2_status,
                    "Deviation (pp)": dev_pp,
                    "Financial Impact (₹)": round((dev_pp / 100.0) * revenue, 2),
                    "Anomaly Category": "CM2 Margin Variance",
                    "Severity": severity,
                    "Explanation": explanation,
                    "FC Status": fc_status,
                    "Metric Type": "Profitability (CM2)"
                })

    df_anomalies = pd.DataFrame(anomalies)
    if not df_anomalies.empty:
        # Sort anomalies by absolute financial impact descending
        df_anomalies["Abs Impact"] = df_anomalies["Financial Impact (₹)"].abs()
        df_anomalies = df_anomalies.sort_values(by=["Abs Impact", "Severity"], ascending=[False, True]).drop(columns=["Abs Impact"])

    return df_anomalies


class Tuple_AnomalyOutput:
    pass
