"""
Test Scenarios Suite Module
Prozo Daily FC P&L Monitoring System

Executes the four mandatory assignment test scenarios in complete isolation
from the production uploaded dataset:

Scenario 1: Clean Day (All metrics within target -> CM1 68%, CM2 24%, GREEN)
Scenario 2: CM2 Breach (Widespread inflation -> CM1 60%, CM2 4%, RED)
Scenario 3: Line-Item-Only Breach (Packaging inflated -> CM1 65%, CM2 21%, GREEN, 1 anomaly)
Scenario 4: Suspiciously High CM2 (Deflated expenses -> CM1 77%, CM2 52%, SUSPICIOUS)

Verifies all mathematical formulas, bounds, deviations, and classifications deterministically.
"""

from typing import Any, Dict, List, Tuple
from .data_extractor import DEFAULT_CRITERIA
from .financial_engine import calculate_row_financials
from .anomaly_detector import analyze_day_record
from .ai_insights import generate_insights_for_day

SCENARIO_INPUTS = {
    1: {
        "title": "Scenario 1: Clean Day",
        "description": "All individual expenses operate strictly within targets; healthy profitability.",
        "record": {
            "Day": "TEST-SCENARIO-1",
            "FC Identifier": "FC-TEST-ALPHA",
            "Revenue": 100000.0,
            "Manpower": 26000.0,
            "Packaging": 6000.0,
            "Power and Fuel": 11000.0,
            "FC Rent": 16000.0,
            "Equipment Rentals": 6000.0,
            "Overheads": 11000.0,
        },
        "expected": {
            "CM1": 68000.0,
            "CM1%": 68.0,
            "CM2": 24000.0,
            "CM2%": 24.0,
            "Overall Status": "GREEN",
            "Anomaly Count": 0,
            "Expense Anomaly Count": 0,
            "CM Deviations Count": 0,
            "Is Suspicious": False,
        }
    },
    2: {
        "title": "Scenario 2: CM2 Breach",
        "description": "Operating expenses severely exceed targets; CM2 collapses to critical red zone.",
        "record": {
            "Day": "TEST-SCENARIO-2",
            "FC Identifier": "FC-TEST-BETA",
            "Revenue": 100000.0,
            "Manpower": 30000.0,
            "Packaging": 10000.0,
            "Power and Fuel": 13000.0,
            "FC Rent": 20000.0,
            "Equipment Rentals": 9000.0,
            "Overheads": 14000.0,
        },
        "expected": {
            "CM1": 60000.0,
            "CM1%": 60.0,
            "CM2": 4000.0,
            "CM2%": 4.0,
            "Overall Status": "RED",
            "Anomaly Count": 8,  # 6 expense + 1 CM1 + 1 CM2
            "Expense Anomaly Count": 6,
            "CM Deviations Count": 2,
            "Is Suspicious": False,
        }
    },
    3: {
        "title": "Scenario 3: Line-Item-Only Breach",
        "description": "Packaging alone exceeds target band; overall CM1 and CM2 stay healthy and within bounds.",
        "record": {
            "Day": "TEST-SCENARIO-3",
            "FC Identifier": "FC-TEST-GAMMA",
            "Revenue": 100000.0,
            "Manpower": 26000.0,
            "Packaging": 9000.0,
            "Power and Fuel": 11000.0,
            "FC Rent": 16000.0,
            "Equipment Rentals": 6000.0,
            "Overheads": 11000.0,
        },
        "expected": {
            "CM1": 65000.0,
            "CM1%": 65.0,
            "CM2": 21000.0,
            "CM2%": 21.0,
            "Overall Status": "GREEN",
            "Anomaly Count": 1,  # Only packaging
            "Expense Anomaly Count": 1,
            "CM Deviations Count": 0,
            "Is Suspicious": False,
        }
    },
    4: {
        "title": "Scenario 4: Suspiciously High CM2",
        "description": "Expenses are artificially low; CM2 exceeds 33%, triggering fraud/missing cost alert.",
        "record": {
            "Day": "TEST-SCENARIO-4",
            "FC Identifier": "FC-TEST-DELTA",
            "Revenue": 100000.0,
            "Manpower": 20000.0,
            "Packaging": 3000.0,
            "Power and Fuel": 7000.0,
            "FC Rent": 10000.0,
            "Equipment Rentals": 3000.0,
            "Overheads": 5000.0,
        },
        "expected": {
            "CM1": 77000.0,
            "CM1%": 77.0,
            "CM2": 52000.0,
            "CM2%": 52.0,
            "Overall Status": "SUSPICIOUS",
            "Anomaly Count": 8,  # 6 expense below target + 1 CM1 + 1 CM2
            "Expense Anomaly Count": 6,
            "CM Deviations Count": 2,
            "Is Suspicious": True,
        }
    }
}

def execute_scenario(scenario_id: int) -> Dict[str, Any]:
    """
    Run calculation, anomaly detection, and AI insight generation on a scenario.
    Validates expected results and returns detailed audit bundle.
    """
    if scenario_id not in SCENARIO_INPUTS:
        raise ValueError(f"Unknown scenario ID: {scenario_id}")

    spec = SCENARIO_INPUTS[scenario_id]
    record = spec["record"]
    expected = spec["expected"]

    # 1. Financial calculation
    calc = calculate_row_financials(record)

    # 2. Anomaly detection using default criteria
    analysis = analyze_day_record(calc, DEFAULT_CRITERIA)

    # 3. AI Insights
    insights = generate_insights_for_day(calc, analysis)

    # 4. Assertions & Validation Checks
    checks = []
    
    # Check CM1 & CM2 values
    cm1_pass = abs(calc["CM1"] - expected["CM1"]) < 0.01
    cm1_pct_pass = abs(calc["CM1%"] - expected["CM1%"]) < 0.01
    cm2_pass = abs(calc["CM2"] - expected["CM2"]) < 0.01
    cm2_pct_pass = abs(calc["CM2%"] - expected["CM2%"]) < 0.01
    status_pass = (analysis["Overall Status"] == expected["Overall Status"])
    suspicious_pass = (analysis["Is Suspicious"] == expected["Is Suspicious"])
    anom_count_pass = (analysis["Total Anomaly Count"] == expected["Anomaly Count"])

    checks.append({"metric": "CM1 (₹)", "actual": calc["CM1"], "expected": expected["CM1"], "passed": cm1_pass})
    checks.append({"metric": "CM1 %", "actual": f"{calc['CM1%']:.1f}%", "expected": f"{expected['CM1%']:.1f}%", "passed": cm1_pct_pass})
    checks.append({"metric": "CM2 (₹)", "actual": calc["CM2"], "expected": expected["CM2"], "passed": cm2_pass})
    checks.append({"metric": "CM2 %", "actual": f"{calc['CM2%']:.1f}%", "expected": f"{expected['CM2%']:.1f}%", "passed": cm2_pct_pass})
    checks.append({"metric": "Overall Status", "actual": analysis["Overall Status"], "expected": expected["Overall Status"], "passed": status_pass})
    checks.append({"metric": "Is Suspicious", "actual": str(analysis["Is Suspicious"]), "expected": str(expected["Is Suspicious"]), "passed": suspicious_pass})
    checks.append({"metric": "Total Anomalies", "actual": analysis["Total Anomaly Count"], "expected": expected["Anomaly Count"], "passed": anom_count_pass})

    all_passed = all(c["passed"] for c in checks)

    return {
        "id": scenario_id,
        "title": spec["title"],
        "description": spec["description"],
        "input_record": record,
        "calculated_record": calc,
        "analysis": analysis,
        "insights": insights,
        "checks": checks,
        "all_passed": all_passed,
    }

def run_all_scenarios() -> Dict[int, Dict[str, Any]]:
    """Execute all 4 scenarios and return combined results."""
    return {s_id: execute_scenario(s_id) for s_id in SCENARIO_INPUTS}
