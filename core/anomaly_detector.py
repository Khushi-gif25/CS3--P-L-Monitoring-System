"""
Business Rules & Anomaly Detection Engine
Prozo Daily FC P&L Monitoring System

Implements the exact business rules from Case Study 3:
- Expense-level target validation against dynamic criteria
- CM1 and CM2 validation (> 3 percentage points deviation from boundary)
- FC Colour Classification (RED, YELLOW, GREEN, SUSPICIOUS, REVIEW REQUIRED)
- Complete isolation between CM2 colour classification and individual expense breaches
- Detailed anomaly traceability
"""

from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
from .data_extractor import CM_TARGETS, STANDARD_EXPENSE_KEYS

def classify_fc_status(cm2_pct: float) -> Tuple[str, str]:
    """
    Classify Day/FC colour based on CM2%:
    - RED: CM2 below 12% (< 12.0%)
    - YELLOW: CM2 from 12% to 14.9% OR above 30% up to 32.9%
    - GREEN: CM2 from 15% to 30%, inclusive (15.0% <= CM2 <= 30.0%)
    - SUSPICIOUS: CM2 above 33% (> 33.0%)
    - REVIEW REQUIRED: Boundary gaps (14.9% < CM2 < 15.0% or 32.9% < CM2 <= 33.0%)
    
    Returns:
        (status, description)
    """
    val = round(cm2_pct, 4)

    if val < 12.0:
        return "RED", "Critical profitability breach: CM2 is below 12.0%."
    elif 12.0 <= val <= 14.9:
        return "YELLOW", "Sub-optimal margin: CM2 is between 12.0% and 14.9%."
    elif 15.0 <= val <= 30.0:
        return "GREEN", "Healthy operating profitability: CM2 is within the 15.0% - 30.0% target band."
    elif 30.0 < val <= 32.9:
        return "YELLOW", "Elevated margin: CM2 is between 30.1% and 32.9% (approaching review threshold)."
    elif val > 33.0:
        return "SUSPICIOUS", "Suspiciously high CM2 (> 33.0%): Potentially missing or underreported operational expenses."
    else:
        # Boundary gaps: 14.9 < val < 15.0 OR 32.9 < val <= 33.0
        return "REVIEW REQUIRED", f"Boundary gap condition: CM2 at {val:.2f}% requires managerial review."

def detect_expense_anomalies(
    day_record: Dict[str, Any],
    criteria: Dict[str, Tuple[float, float]]
) -> List[Dict[str, Any]]:
    """
    Validate individual expense percentages against company targets.
    Records anomalies ONLY when actual percentage is strictly outside the target range.
    """
    anomalies = []
    day_id = day_record["Day"]
    fc_id = day_record.get("FC Identifier") or "N/A"

    for exp_key in STANDARD_EXPENSE_KEYS:
        pct_key = f"{exp_key} %"
        if pct_key not in day_record:
            continue

        actual_pct = float(day_record[pct_key])
        if exp_key not in criteria:
            continue

        min_pct, max_pct = criteria[exp_key]

        if actual_pct < min_pct:
            dev = round(min_pct - actual_pct, 2)
            anomalies.append({
                "Day": day_id,
                "FC Identifier": fc_id,
                "Line Item": exp_key,
                "Type": "Expense Anomaly",
                "Actual %": actual_pct,
                "Target Range": f"{min_pct:.1f}% - {max_pct:.1f}%",
                "Deviation (pp)": dev,
                "Status": "BELOW TARGET",
                "Severity": "Warning",
                "Message": f"{exp_key} was {actual_pct:.2f}%, which is {dev:.2f} pp below target minimum of {min_pct:.1f}%."
            })
        elif actual_pct > max_pct:
            dev = round(actual_pct - max_pct, 2)
            anomalies.append({
                "Day": day_id,
                "FC Identifier": fc_id,
                "Line Item": exp_key,
                "Type": "Expense Anomaly",
                "Actual %": actual_pct,
                "Target Range": f"{min_pct:.1f}% - {max_pct:.1f}%",
                "Deviation (pp)": dev,
                "Status": "ABOVE TARGET",
                "Severity": "Critical" if dev > 2.0 else "Warning",
                "Message": f"{exp_key} was {actual_pct:.2f}%, which is {dev:.2f} pp above target maximum of {max_pct:.1f}%."
            })

    return anomalies

def detect_cm_deviations(
    day_record: Dict[str, Any],
    cm_targets: Optional[Dict[str, Tuple[float, float]]] = None
) -> List[Dict[str, Any]]:
    """
    Validate CM1 and CM2 against dynamic or standard targets.
    Rule: Flag a deviation when the actual percentage is MORE than 3 percentage points
    outside its target range.
    A deviation of exactly 3 percentage points is NOT flagged.
    """
    deviations = []
    day_id = day_record["Day"]
    fc_id = day_record.get("FC Identifier") or "N/A"
    targets = cm_targets if cm_targets else CM_TARGETS

    # CM1 Validation
    cm1_pct = float(day_record["CM1%"])
    cm1_min, cm1_max = targets.get("CM1", CM_TARGETS["CM1"])
    if cm1_pct < cm1_min:
        dev = round(cm1_min - cm1_pct, 2)
        if dev > 3.0:
            deviations.append({
                "Day": day_id,
                "FC Identifier": fc_id,
                "Line Item": "CM1",
                "Type": "CM1 Deviation",
                "Actual %": cm1_pct,
                "Target Range": f"{cm1_min:.1f}% - {cm1_max:.1f}%",
                "Deviation (pp)": dev,
                "Status": "BELOW TARGET",
                "Severity": "Critical",
                "Message": f"CM1 was {cm1_pct:.2f}%, breaching the lower threshold by {dev:.2f} percentage points (> 3.0 pp allowed margin)."
            })
    elif cm1_pct > cm1_max:
        dev = round(cm1_pct - cm1_max, 2)
        if dev > 3.0:
            deviations.append({
                "Day": day_id,
                "FC Identifier": fc_id,
                "Line Item": "CM1",
                "Type": "CM1 Deviation",
                "Actual %": cm1_pct,
                "Target Range": f"{cm1_min:.1f}% - {cm1_max:.1f}%",
                "Deviation (pp)": dev,
                "Status": "ABOVE TARGET",
                "Severity": "Warning",
                "Message": f"CM1 was {cm1_pct:.2f}%, exceeding the upper threshold by {dev:.2f} percentage points (> 3.0 pp allowed margin)."
            })

    # CM2 Validation
    cm2_pct = float(day_record["CM2%"])
    cm2_min, cm2_max = targets.get("CM2", CM_TARGETS["CM2"])
    if cm2_pct < cm2_min:
        dev = round(cm2_min - cm2_pct, 2)
        if dev > 3.0:
            deviations.append({
                "Day": day_id,
                "FC Identifier": fc_id,
                "Line Item": "CM2",
                "Type": "CM2 Deviation",
                "Actual %": cm2_pct,
                "Target Range": f"{cm2_min:.1f}% - {cm2_max:.1f}%",
                "Deviation (pp)": dev,
                "Status": "BELOW TARGET",
                "Severity": "Critical",
                "Message": f"CM2 was {cm2_pct:.2f}%, breaching the lower threshold by {dev:.2f} percentage points (> 3.0 pp allowed margin)."
            })
    elif cm2_pct > cm2_max:
        dev = round(cm2_pct - cm2_max, 2)
        if dev > 3.0:
            deviations.append({
                "Day": day_id,
                "FC Identifier": fc_id,
                "Line Item": "CM2",
                "Type": "CM2 Deviation",
                "Actual %": cm2_pct,
                "Target Range": f"{cm2_min:.1f}% - {cm2_max:.1f}%",
                "Deviation (pp)": dev,
                "Status": "ABOVE TARGET",
                "Severity": "Warning" if cm2_pct <= 33.0 else "Suspicious",
                "Message": f"CM2 was {cm2_pct:.2f}%, exceeding the upper threshold by {dev:.2f} percentage points (> 3.0 pp allowed margin)."
            })

    return deviations

def analyze_day_record(
    day_record: Dict[str, Any],
    criteria: Dict[str, Tuple[float, float]],
    cm_targets: Optional[Dict[str, Tuple[float, float]]] = None
) -> Dict[str, Any]:
    """
    Combined anomaly detection and status classification for a single day.
    """
    cm2_pct = float(day_record["CM2%"])
    status, status_desc = classify_fc_status(cm2_pct)

    expense_anomalies = detect_expense_anomalies(day_record, criteria)
    cm_deviations = detect_cm_deviations(day_record, cm_targets=cm_targets)

    all_anomalies = expense_anomalies + cm_deviations
    is_suspicious = (status == "SUSPICIOUS")

    return {
        "Day": day_record["Day"],
        "FC Identifier": day_record.get("FC Identifier") or "N/A",
        "Overall Status": status,
        "Status Description": status_desc,
        "Is Suspicious": is_suspicious,
        "Expense Anomalies": expense_anomalies,
        "CM Deviations": cm_deviations,
        "All Anomalies": all_anomalies,
        "Total Anomaly Count": len(all_anomalies),
    }

def analyze_full_dataset(
    df_calc: pd.DataFrame,
    criteria: Dict[str, Tuple[float, float]],
    cm_targets: Optional[Dict[str, Tuple[float, float]]] = None
) -> Tuple[pd.DataFrame, List[Dict[str, Any]], Dict[str, int]]:
    """
    Process full calculated dataset:
    - Adds 'Overall Status' to DataFrame
    - Generates master anomaly log
    - Computes status counts
    """
    records = df_calc.to_dict(orient="records")
    master_anomalies = []
    statuses = []
    
    status_counts = {
        "RED": 0,
        "YELLOW": 0,
        "GREEN": 0,
        "SUSPICIOUS": 0,
        "REVIEW REQUIRED": 0,
        "NORMAL": 0,
    }

    for r in records:
        day_analysis = analyze_day_record(r, criteria, cm_targets=cm_targets)
        st = day_analysis["Overall Status"]
        statuses.append(st)
        if st in status_counts:
            status_counts[st] += 1

        master_anomalies.extend(day_analysis["All Anomalies"])

    df_result = df_calc.copy()
    df_result["Overall Status"] = statuses

    # Add count of clean normal days (Green with 0 expense anomalies)
    green_clean_count = 0
    for r in records:
        day_analysis = analyze_day_record(r, criteria, cm_targets=cm_targets)
        if day_analysis["Overall Status"] == "GREEN" and day_analysis["Total Anomaly Count"] == 0:
            green_clean_count += 1
    status_counts["CLEAN_GREEN"] = green_clean_count

    return df_result, master_anomalies, status_counts
