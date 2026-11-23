"""
PROZO P&L MONITOR - SYNTHETIC DATA GENERATOR & TEST CASES
Implements the 4 required synthetic test scenarios and multi-day multi-FC benchmark datasets.
Also provides Excel file exporters for sample files.
"""

import io
from typing import Dict, Any, List, Tuple
import pandas as pd
import openpyxl


def get_standard_ideal_criteria() -> Dict[str, Dict[str, float]]:
    """Standard predefined Prozo expense targets."""
    return {
        "Manpower": {"min": 25.0, "max": 27.0},
        "Packaging": {"min": 5.0, "max": 7.0},
        "Power and Fuel": {"min": 10.0, "max": 12.0},
        "FC Rent": {"min": 15.0, "max": 18.0},
        "Equipment Rentals": {"min": 5.0, "max": 8.0},
        "Overheads": {"min": 10.0, "max": 12.0},
        "CM1": {"min": 65.0, "max": 70.0},
        "CM2": {"min": 15.0, "max": 30.0},
    }


def generate_ideal_criteria_excel() -> bytes:
    """Generates a formatted Ideal FC Criteria Excel file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ideal FC Criteria"

    # Add corporate header
    ws.append(["PROZO LOGISTICS - TARGET COST & PROFITABILITY CRITERIA"])
    ws.append(["Expense Category", "Minimum Target (%)", "Maximum Target (%)", "Target Type"])

    targets = get_standard_ideal_criteria()
    for cat, val in targets.items():
        cat_type = "Contribution Margin" if "CM" in cat else "Operational Expense"
        ws.append([cat, f"{val['min']}%", f"{val['max']}%", cat_type])

    # Widths
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 24

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def get_test_scenario_records(scenario_id: str) -> List[Dict[str, Any]]:
    """
    Returns mathematically validated records for the 4 core test scenarios.
    """
    revenue = 500000.0

    if scenario_id == "clean":
        # Test 1: Clean Day
        # All expenses within target ranges, CM1 ~ 68.0%, CM2 = 24.0% (GREEN, 0 anomalies)
        return [{
            "Date": "2026-05-20",
            "FC": "FC-Bhiwandi-01",
            "Revenue": revenue,
            "Manpower Cost": 130000.0,      # 26.0% (25-27%)
            "Packaging Cost": 30000.0,       # 6.0%  (5-7%)
            "Power and Fuel": 55000.0,       # 11.0% (10-12%)
            "FC Rent": 80000.0,              # 16.0% (15-18%)
            "Equipment Rentals": 30000.0,    # 6.0%  (5-8%)
            "Overheads": 55000.0             # 11.0% (10-12%)
        }]

    elif scenario_id == "cm2_breach":
        # Test 2: CM2 Breach
        # CM2 below 12%, RED status, at least one expense anomaly
        # Manpower: 31.0% (breach), Power: 14.5% (breach), Overheads: 14.0% (breach)
        # CM1 = 63.0%, CM2 = 10.5% (< 12%) -> RED
        return [{
            "Date": "2026-05-21",
            "FC": "FC-Farukhnagar-02",
            "Revenue": revenue,
            "Manpower Cost": 155000.0,      # 31.0% (>27%)
            "Packaging Cost": 30000.0,       # 6.0%  (5-7%)
            "Power and Fuel": 72500.0,       # 14.5% (>12%)
            "FC Rent": 85000.0,              # 17.0% (15-18%)
            "Equipment Rentals": 35000.0,    # 7.0%  (5-8%)
            "Overheads": 70000.0             # 14.0% (>12%)
        }]

    elif scenario_id == "line_item_breach":
        # Test 3: Line-item-only Breach
        # Exactly one expense outside range (Packaging: 8.5% > 7%), CM1 = 66.0%, CM2 = 23.5% (GREEN)
        return [{
            "Date": "2026-05-22",
            "FC": "FC-Hoskote-03",
            "Revenue": revenue,
            "Manpower Cost": 127500.0,      # 25.5% (25-27%)
            "Packaging Cost": 42500.0,       # 8.5%  (>7% - Single Breach)
            "Power and Fuel": 52500.0,       # 10.5% (10-12%)
            "FC Rent": 77500.0,              # 15.5% (15-18%)
            "Equipment Rentals": 30000.0,    # 6.0%  (5-8%)
            "Overheads": 52500.0             # 10.5% (10-12%)
        }]

    elif scenario_id == "suspicious_high":
        # Test 4: Suspiciously High CM2
        # CM2 above 33%, SUSPICIOUS status, underreported expenses warning
        # Revenue = 500,000, Total Opex = 41.5%, CM1 = 69.5%, CM2 = 58.5% (>33%)
        return [{
            "Date": "2026-05-23",
            "FC": "FC-Chakan-04",
            "Revenue": revenue,
            "Manpower Cost": 125000.0,      # 25.0% (25-27%)
            "Packaging Cost": 27500.0,       # 5.5%  (5-7%)
            "Power and Fuel": 15000.0,       # 3.0%  (<10% - Missing/Underreported)
            "FC Rent": 10000.0,              # 2.0%  (<15% - Missing/Underreported)
            "Equipment Rentals": 10000.0,    # 2.0%  (<5% - Missing/Underreported)
            "Overheads": 20000.0             # 4.0%  (<10% - Missing/Underreported)
        }]

    elif scenario_id == "all_scenarios_combined":
        # Combines all 4 test scenarios plus additional operational records
        recs = []
        for s in ["clean", "cm2_breach", "line_item_breach", "suspicious_high"]:
            recs.extend(get_test_scenario_records(s))
        
        # Add 3 additional historical days for rich trend analysis
        recs.extend([
            {
                "Date": "2026-05-24",
                "FC": "FC-Bhiwandi-01",
                "Revenue": 520000.0,
                "Manpower Cost": 135200.0,  # 26.0%
                "Packaging Cost": 31200.0,   # 6.0%
                "Power and Fuel": 57200.0,   # 11.0%
                "FC Rent": 83200.0,          # 16.0%
                "Equipment Rentals": 31200.0,# 6.0%
                "Overheads": 57200.0         # 11.0%
            },
            {
                "Date": "2026-05-25",
                "FC": "FC-Farukhnagar-02",
                "Revenue": 480000.0,
                "Manpower Cost": 139200.0,  # 29.0% (Yellow/Minor Breach)
                "Packaging Cost": 31200.0,   # 6.5%
                "Power and Fuel": 57600.0,   # 12.0%
                "FC Rent": 81600.0,          # 17.0%
                "Equipment Rentals": 33600.0,# 7.0%
                "Overheads": 62400.0         # 13.0%
            },
            {
                "Date": "2026-05-26",
                "FC": "FC-Hoskote-03",
                "Revenue": 610000.0,
                "Manpower Cost": 158600.0,  # 26.0%
                "Packaging Cost": 39650.0,   # 6.5%
                "Power and Fuel": 67100.0,   # 11.0%
                "FC Rent": 97600.0,          # 16.0%
                "Equipment Rentals": 42700.0,# 7.0%
                "Overheads": 67100.0         # 11.0%
            }
        ])
        return recs

    return []


def generate_daily_pnl_excel(scenario_id: str = "all_scenarios_combined") -> bytes:
    """Generates a formatted Daily P&L Excel file for testing."""
    records = get_test_scenario_records(scenario_id)
    df = pd.DataFrame(records)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily FC P&L"

    # Add subtitle or blank row to verify header row detection works seamlessly!
    ws.append(["PROZO FULFILLMENT NETWORK - DAILY FC OPERATIONAL FINANCIAL REPORT"])
    ws.append(["Generated for Internal Audit & P&L Monitoring"])
    ws.append([])  # Blank row so header is at row 4!

    # Headers
    headers = list(df.columns)
    ws.append(headers)

    # Rows
    for r in df.itertuples(index=False):
        ws.append(list(r))

    for col in ws.columns:
        ws.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = 18

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def generate_single_workbook_excel(scenario_id: str = "all_scenarios_combined") -> bytes:
    """
    Generates a single Excel workbook (.xlsx) containing TWO separate worksheets:
    1. 'Ideal FC Criteria'
    2. 'Day Wise PnL'
    Matches the exact layout of Case Study 3.xlsx.
    """
    wb = openpyxl.Workbook()
    
    # Sheet 1: Ideal FC Criteria
    ws1 = wb.active
    ws1.title = "Ideal FC Criteria"
    ws1.append(["#", "Expense Line Item", "Range (as % of Revenue)"])
    
    criteria_rows = [
        [1, "Manpower", "25 - 27%"],
        [2, "Packaging", "5 - 7%"],
        [3, "Power and Fuel", "10 - 12%"],
        [4, "FC Rent", "15 - 18%"],
        [5, "Equipment Rentals", "5 - 8%"],
        [6, "Overheads", "10 - 12%"],
    ]
    for r in criteria_rows:
        ws1.append(r)
        
    ws1.column_dimensions["A"].width = 8
    ws1.column_dimensions["B"].width = 24
    ws1.column_dimensions["C"].width = 24

    # Sheet 2: Day Wise PnL
    ws2 = wb.create_sheet(title="Day Wise PnL")
    records = get_test_scenario_records(scenario_id)
    
    headers = ["Day", "Revenue", "Manpower", "Packaging", "Power and Fuel", "FC Rent", "Equipment Rentals", "Overheads"]
    ws2.append(headers)
    
    for idx, r in enumerate(records, 1):
        day_val = r.get("Date", idx)
        if isinstance(day_val, str) and day_val.startswith("2026"):
            day_val = idx
        ws2.append([
            day_val,
            r["Revenue"],
            r.get("Manpower Cost", r.get("Manpower", 0)),
            r.get("Packaging Cost", r.get("Packaging", 0)),
            r.get("Power and Fuel", 0),
            r.get("FC Rent", 0),
            r.get("Equipment Rentals", 0),
            r.get("Overheads", 0)
        ])

    for col in ws2.columns:
        ws2.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = 18

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
