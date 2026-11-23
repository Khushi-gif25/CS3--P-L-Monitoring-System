"""
Financial Calculation Engine
Prozo Daily FC P&L Monitoring System

Deterministic financial calculations:
- Individual expense percentages (% of Revenue)
- CM1 (Absolute ₹) and CM1%
- CM2 (Absolute ₹) and CM2%
- Revenue-weighted portfolio aggregates (never simple averages)
"""

from typing import Any, Dict, List, Tuple
import pandas as pd
from .data_extractor import STANDARD_EXPENSE_KEYS

def calculate_row_financials(row: Dict[str, Any]) -> Dict[str, Any]:
    """
    Calculate deterministic financial metrics for a single day record.
    """
    rev = float(row["Revenue"])
    if rev <= 0:
        raise ValueError(f"Revenue must be strictly positive, got {rev}")

    calc = dict(row)

    # 1. Individual Expense Percentages
    total_expenses = 0.0
    for exp_key in STANDARD_EXPENSE_KEYS:
        amt = float(row[exp_key])
        pct = (amt / rev) * 100.0
        calc[f"{exp_key} %"] = round(pct, 2)
        total_expenses += amt

    # 2. CM1: Revenue - Manpower - Packaging
    manpower = float(row["Manpower"])
    packaging = float(row["Packaging"])
    cm1 = rev - manpower - packaging
    cm1_pct = (cm1 / rev) * 100.0
    calc["CM1"] = round(cm1, 2)
    calc["CM1%"] = round(cm1_pct, 2)

    # 3. CM2: CM1 - Power and Fuel - FC Rent - Equipment Rentals - Overheads
    power_fuel = float(row["Power and Fuel"])
    fc_rent = float(row["FC Rent"])
    equip = float(row["Equipment Rentals"])
    overheads = float(row["Overheads"])
    cm2 = cm1 - power_fuel - fc_rent - equip - overheads
    cm2_pct = (cm2 / rev) * 100.0
    calc["CM2"] = round(cm2, 2)
    calc["CM2%"] = round(cm2_pct, 2)

    calc["Total Expenses"] = round(total_expenses, 2)
    calc["Total Expense %"] = round((total_expenses / rev) * 100.0, 2)

    return calc

def calculate_dataframe_financials(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply financial calculations deterministically across the entire DataFrame.
    """
    records = df.to_dict(orient="records")
    calculated_records = [calculate_row_financials(r) for r in records]
    return pd.DataFrame(calculated_records)

def calculate_portfolio_aggregates(df_calc: pd.DataFrame) -> Dict[str, Any]:
    """
    Compute revenue-weighted portfolio aggregates.
    Overall CM1% = Total CM1 / Total Revenue * 100
    Overall CM2% = Total CM2 / Total Revenue * 100
    Strictly avoids simple arithmetic averaging of daily percentages.
    """
    if df_calc.empty:
        return {}

    total_rev = float(df_calc["Revenue"].sum())
    total_cm1 = float(df_calc["CM1"].sum())
    total_cm2 = float(df_calc["CM2"].sum())
    
    if total_rev <= 0:
        overall_cm1_pct = 0.0
        overall_cm2_pct = 0.0
    else:
        overall_cm1_pct = (total_cm1 / total_rev) * 100.0
        overall_cm2_pct = (total_cm2 / total_rev) * 100.0

    aggregates = {
        "Total Days": len(df_calc),
        "Total Revenue": round(total_rev, 2),
        "Total CM1": round(total_cm1, 2),
        "Total CM2": round(total_cm2, 2),
        "Overall CM1%": round(overall_cm1_pct, 2),
        "Overall CM2%": round(overall_cm2_pct, 2),
        "Expense Totals": {},
        "Expense Overall Percentages": {},
    }

    for exp_key in STANDARD_EXPENSE_KEYS:
        tot_exp = float(df_calc[exp_key].sum())
        exp_pct = (tot_exp / total_rev) * 100.0 if total_rev > 0 else 0.0
        aggregates["Expense Totals"][exp_key] = round(tot_exp, 2)
        aggregates["Expense Overall Percentages"][exp_key] = round(exp_pct, 2)

    return aggregates
