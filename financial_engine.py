"""
PROZO P&L MONITOR - FINANCIAL CALCULATION ENGINE
Deterministic, programmatic calculation of Expense %, CM1, CM2,
and strict FC color classification according to Prozo business rules.
"""

from typing import Dict, Any, List, Tuple
import pandas as pd
import numpy as np


EXPENSE_FIELD_MAP = {
    "Manpower Cost": "Manpower",
    "Packaging Cost": "Packaging",
    "Power and Fuel": "Power and Fuel",
    "FC Rent": "FC Rent",
    "Equipment Rentals": "Equipment Rentals",
    "Overheads": "Overheads",
}


def classify_fc_status(cm2_pct: float) -> Tuple[str, str, str, str]:
    """
    Classify FC color status strictly based on CM2%:
    - RED: CM2 below 12.0%
    - YELLOW: CM2 between 12.0% and 14.9% OR CM2 above 30.0% up to 32.9%
    - GREEN: CM2 between 15.0% and 30.0%
    - SUSPICIOUS: CM2 above 33.0%
    - REVIEW REQUIRED: Boundary gaps (14.9% < CM2 < 15.0% or 32.9% < CM2 <= 33.0%)
    
    Returns (status_label, hex_color, badge_style, warning_message)
    """
    cm2_val = round(cm2_pct, 2)

    if cm2_val < 12.0:
        return (
            "RED",
            "#DC2626",  # Red
            "background-color: #FEE2E2; color: #991B1B; font-weight: bold; border-radius: 4px; padding: 2px 8px;",
            "Critical Profitability Breach: CM2 is below the 12% operational survival threshold."
        )
    elif 12.0 <= cm2_val <= 14.9:
        return (
            "YELLOW",
            "#D97706",  # Amber / Yellow
            "background-color: #FEF3C7; color: #92400E; font-weight: bold; border-radius: 4px; padding: 2px 8px;",
            "Sub-optimal Profitability: CM2 is between 12% and 14.9%, underperforming target."
        )
    elif 15.0 <= cm2_val <= 30.0:
        return (
            "GREEN",
            "#16A34A",  # Emerald Green
            "background-color: #DCFCE7; color: #166534; font-weight: bold; border-radius: 4px; padding: 2px 8px;",
            "Healthy Performance: CM2 is within the standard target corridor (15% - 30%)."
        )
    elif 30.0 < cm2_val <= 32.9:
        return (
            "YELLOW",
            "#D97706",  # Amber
            "background-color: #FEF3C7; color: #92400E; font-weight: bold; border-radius: 4px; padding: 2px 8px;",
            "Above Normal Target: CM2 is between 30% and 32.9%, verify cost accruals."
        )
    elif cm2_val > 33.0:
        return (
            "SUSPICIOUS",
            "#9333EA",  # Purple / Warning
            "background-color: #F3E8FF; color: #6B21A8; font-weight: bold; border-radius: 4px; padding: 2px 8px;",
            "Suspiciously High CM2 (>33%): High probability of missing or underreported expense accruals (e.g. unbilled rent or power)."
        )
    else:
        # Boundary gaps: 14.9% < CM2 < 15.0% OR 32.9% < CM2 <= 33.0%
        return (
            "REVIEW REQUIRED",
            "#6B7280",  # Gray
            "background-color: #F3F4F6; color: #374151; font-weight: bold; border-radius: 4px; padding: 2px 8px;",
            "Boundary Gap: CM2 falls in transitional boundary zone. Manual management review required."
        )


class Tuple_Status:
    pass


def calculate_financial_metrics(df_clean: pd.DataFrame) -> pd.DataFrame:
    """
    Perform deterministic financial calculations on standardized records:
    1. Expense % = (Expense Amount / Revenue) * 100 for all 6 categories
    2. CM1 = Revenue - Manpower - Packaging
    3. CM1% = (CM1 / Revenue) * 100
    4. CM2 = CM1 - Power & Fuel - FC Rent - Equipment Rentals - Overheads
    5. CM2% = (CM2 / Revenue) * 100
    6. FC Colour Status and Description
    """
    df = df_clean.copy()

    # Verify positive revenue
    if (df["Revenue"] <= 0).any():
        raise ValueError("Revenue contains non-positive values. Validate records before calculation.")

    # Calculate Expense Percentages
    for exp_col in EXPENSE_FIELD_MAP.keys():
        pct_col = f"{exp_col} %"
        df[pct_col] = ((df[exp_col] / df["Revenue"]) * 100.0).round(2)

    # CM1 Calculation
    df["CM1"] = (df["Revenue"] - df["Manpower Cost"] - df["Packaging Cost"]).round(2)
    df["CM1 %"] = ((df["CM1"] / df["Revenue"]) * 100.0).round(2)

    # CM2 Calculation
    df["CM2"] = (
        df["CM1"]
        - df["Power and Fuel"]
        - df["FC Rent"]
        - df["Equipment Rentals"]
        - df["Overheads"]
    ).round(2)
    df["CM2 %"] = ((df["CM2"] / df["Revenue"]) * 100.0).round(2)

    # Total Operating Expense
    df["Total Opex"] = (
        df["Manpower Cost"]
        + df["Packaging Cost"]
        + df["Power and Fuel"]
        + df["FC Rent"]
        + df["Equipment Rentals"]
        + df["Overheads"]
    ).round(2)
    df["Total Opex %"] = ((df["Total Opex"] / df["Revenue"]) * 100.0).round(2)

    # FC Color Classification
    statuses = []
    colors = []
    badges = []
    notes = []

    for cm2_val in df["CM2 %"]:
        status, color, badge, note = classify_fc_status(cm2_val)
        statuses.append(status)
        colors.append(color)
        badges.append(badge)
        notes.append(note)

    df["FC Status"] = statuses
    df["Status Color"] = colors
    df["Status Badge"] = badges
    df["Status Warning"] = notes

    return df
