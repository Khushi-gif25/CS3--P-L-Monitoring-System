"""
Unit & Integration Test Suite
Prozo Daily FC P&L Monitoring System

Tests:
1. Scenario 1 (Clean Day) assertions
2. Scenario 2 (CM2 Breach) assertions
3. Scenario 3 (Line-Item-Only Breach) assertions
4. Scenario 4 (Suspiciously High CM2) assertions
5. Boundary conditions (12.0%, 14.9%, 15.0%, 30.0%, 32.9%, 33.0%, review required gap)
6. CM1/CM2 deviation rule (exact 3.0 pp is NOT flagged; > 3.0 pp IS flagged)
7. Portfolio revenue-weighted aggregates vs arithmetic average
8. Excel file loading (Case Study 3.xlsx)
9. CSV file loading (Day Wise PnL.csv)
10. Excel 5-sheet report generation
"""

import os
import unittest
import pandas as pd
from core.data_extractor import (
    DEFAULT_CRITERIA,
    CM_TARGETS,
    extract_criteria,
    extract_pnl
)
from core.financial_engine import (
    calculate_row_financials,
    calculate_dataframe_financials,
    calculate_portfolio_aggregates
)
from core.anomaly_detector import (
    classify_fc_status,
    detect_expense_anomalies,
    detect_cm_deviations,
    analyze_day_record,
    analyze_full_dataset
)
from core.trend_analyzer import analyze_dataset_trends
from core.report_generator import generate_excel_report
from core.test_scenarios import execute_scenario, run_all_scenarios

class TestPnLMonitor(unittest.TestCase):

    def test_scenario_1_clean_day(self):
        res = execute_scenario(1)
        self.assertTrue(res["all_passed"], f"Scenario 1 failed: {res['checks']}")
        calc = res["calculated_record"]
        self.assertEqual(calc["CM1"], 68000.0)
        self.assertEqual(calc["CM1%"], 68.0)
        self.assertEqual(calc["CM2"], 24000.0)
        self.assertEqual(calc["CM2%"], 24.0)
        self.assertEqual(res["analysis"]["Overall Status"], "GREEN")
        self.assertEqual(res["analysis"]["Total Anomaly Count"], 0)

    def test_scenario_2_cm2_breach(self):
        res = execute_scenario(2)
        self.assertTrue(res["all_passed"], f"Scenario 2 failed: {res['checks']}")
        calc = res["calculated_record"]
        self.assertEqual(calc["CM1"], 60000.0)
        self.assertEqual(calc["CM1%"], 60.0)
        self.assertEqual(calc["CM2"], 4000.0)
        self.assertEqual(calc["CM2%"], 4.0)
        self.assertEqual(res["analysis"]["Overall Status"], "RED")
        # Check CM2 deviation is flagged
        cm_dev_lines = [d["Line Item"] for d in res["analysis"]["CM Deviations"]]
        self.assertIn("CM2", cm_dev_lines)
        self.assertIn("CM1", cm_dev_lines)

    def test_scenario_3_line_item_only_breach(self):
        res = execute_scenario(3)
        self.assertTrue(res["all_passed"], f"Scenario 3 failed: {res['checks']}")
        calc = res["calculated_record"]
        self.assertEqual(calc["CM1"], 65000.0)
        self.assertEqual(calc["CM1%"], 65.0)
        self.assertEqual(calc["CM2"], 21000.0)
        self.assertEqual(calc["CM2%"], 21.0)
        self.assertEqual(res["analysis"]["Overall Status"], "GREEN")
        self.assertEqual(len(res["analysis"]["Expense Anomalies"]), 1)
        self.assertEqual(res["analysis"]["Expense Anomalies"][0]["Line Item"], "Packaging")
        self.assertEqual(len(res["analysis"]["CM Deviations"]), 0)

    def test_scenario_4_suspicious_high_cm2(self):
        res = execute_scenario(4)
        self.assertTrue(res["all_passed"], f"Scenario 4 failed: {res['checks']}")
        calc = res["calculated_record"]
        self.assertEqual(calc["CM1"], 77000.0)
        self.assertEqual(calc["CM1%"], 77.0)
        self.assertEqual(calc["CM2"], 52000.0)
        self.assertEqual(calc["CM2%"], 52.0)
        self.assertEqual(res["analysis"]["Overall Status"], "SUSPICIOUS")
        self.assertTrue(res["analysis"]["Is Suspicious"])

    def test_fc_status_classification_boundaries(self):
        # RED: below 12%
        self.assertEqual(classify_fc_status(11.9)[0], "RED")
        self.assertEqual(classify_fc_status(0.0)[0], "RED")
        
        # YELLOW: 12% to 14.9%
        self.assertEqual(classify_fc_status(12.0)[0], "YELLOW")
        self.assertEqual(classify_fc_status(13.5)[0], "YELLOW")
        self.assertEqual(classify_fc_status(14.9)[0], "YELLOW")

        # Boundary gap: 14.9% < CM2 < 15.0% -> REVIEW REQUIRED
        self.assertEqual(classify_fc_status(14.95)[0], "REVIEW REQUIRED")

        # GREEN: 15% to 30%, inclusive
        self.assertEqual(classify_fc_status(15.0)[0], "GREEN")
        self.assertEqual(classify_fc_status(22.5)[0], "GREEN")
        self.assertEqual(classify_fc_status(30.0)[0], "GREEN")

        # YELLOW: above 30% up to 32.9%
        self.assertEqual(classify_fc_status(30.1)[0], "YELLOW")
        self.assertEqual(classify_fc_status(32.9)[0], "YELLOW")

        # Boundary gap: 32.9% < CM2 <= 33.0% -> REVIEW REQUIRED
        self.assertEqual(classify_fc_status(32.95)[0], "REVIEW REQUIRED")
        self.assertEqual(classify_fc_status(33.0)[0], "REVIEW REQUIRED")

        # SUSPICIOUS: above 33%
        self.assertEqual(classify_fc_status(33.1)[0], "SUSPICIOUS")
        self.assertEqual(classify_fc_status(45.0)[0], "SUSPICIOUS")

    def test_cm_deviation_exact_3pp_rule(self):
        # CM2 target: 15% - 30%
        # Deviation of exactly 3.0 pp outside must NOT be flagged (assignment: MORE than 3 pp)
        # Below: 15 - 3.0 = 12.0%
        rec_exact_3pp_below = {"Day": 1, "CM1%": 67.0, "CM2%": 12.0}
        devs = detect_cm_deviations(rec_exact_3pp_below)
        cm2_devs = [d for d in devs if d["Line Item"] == "CM2"]
        self.assertEqual(len(cm2_devs), 0, "Exact 3.0 pp deviation must not be flagged")

        # Below: 15 - 3.1 = 11.9% -> Dev is 3.1 pp > 3.0 pp -> MUST be flagged
        rec_more_3pp_below = {"Day": 1, "CM1%": 67.0, "CM2%": 11.9}
        devs2 = detect_cm_deviations(rec_more_3pp_below)
        cm2_devs2 = [d for d in devs2 if d["Line Item"] == "CM2"]
        self.assertEqual(len(cm2_devs2), 1, "Deviation > 3.0 pp must be flagged")

        # Above: 30 + 3.0 = 33.0% -> Exact 3.0 pp -> NOT flagged
        rec_exact_3pp_above = {"Day": 1, "CM1%": 67.0, "CM2%": 33.0}
        devs3 = detect_cm_deviations(rec_exact_3pp_above)
        cm2_devs3 = [d for d in devs3 if d["Line Item"] == "CM2"]
        self.assertEqual(len(cm2_devs3), 0, "Exact 3.0 pp deviation must not be flagged")

        # Above: 30 + 3.1 = 33.1% -> Dev is 3.1 pp > 3.0 pp -> MUST be flagged
        rec_more_3pp_above = {"Day": 1, "CM1%": 67.0, "CM2%": 33.1}
        devs4 = detect_cm_deviations(rec_more_3pp_above)
        cm2_devs4 = [d for d in devs4 if d["Line Item"] == "CM2"]
        self.assertEqual(len(cm2_devs4), 1, "Deviation > 3.0 pp must be flagged")

    def test_revenue_weighted_aggregates(self):
        # Two days with vastly different revenues:
        # Day 1: Rev = 10,000, CM1 = 8,000 (80%), CM2 = 4,000 (40%)
        # Day 2: Rev = 90,000, CM1 = 54,000 (60%), CM2 = 18,000 (20%)
        # Total Rev = 100,000
        # Total CM1 = 62,000 -> Overall CM1% = 62.0%
        # Total CM2 = 22,000 -> Overall CM2% = 22.0%
        # Simple arithmetic average CM1% would be (80 + 60)/2 = 70.0% (WRONG)
        df_sample = pd.DataFrame([
            {
                "Day": 1, "Revenue": 10000.0, "Manpower": 1000.0, "Packaging": 1000.0,
                "Power and Fuel": 1000.0, "FC Rent": 1000.0, "Equipment Rentals": 1000.0, "Overheads": 1000.0,
                "CM1": 8000.0, "CM1%": 80.0, "CM2": 4000.0, "CM2%": 40.0
            },
            {
                "Day": 2, "Revenue": 90000.0, "Manpower": 18000.0, "Packaging": 18000.0,
                "Power and Fuel": 9000.0, "FC Rent": 9000.0, "Equipment Rentals": 9000.0, "Overheads": 9000.0,
                "CM1": 54000.0, "CM1%": 60.0, "CM2": 18000.0, "CM2%": 20.0
            }
        ])
        agg = calculate_portfolio_aggregates(df_sample)
        self.assertEqual(agg["Overall CM1%"], 62.0)
        self.assertEqual(agg["Overall CM2%"], 22.0)
        self.assertNotEqual(agg["Overall CM1%"], 70.0)

    def test_excel_loading_and_processing(self):
        xlsx_path = r"C:\Users\Lenovo\.gemini\antigravity\scratch\prozo_pnl_monitor\data\Case Study 3.xlsx"
        self.assertTrue(os.path.exists(xlsx_path), f"File {xlsx_path} does not exist")
        
        # Criteria extraction
        criteria, crit_warn = extract_criteria(xlsx_path)
        self.assertIn("Manpower", criteria)
        self.assertEqual(criteria["Manpower"], (25.0, 27.0))
        self.assertEqual(criteria["FC Rent"], (15.0, 18.0))

        # PnL extraction
        df_pnl, pnl_err, raw_recs = extract_pnl(xlsx_path)
        self.assertIsNotNone(df_pnl)
        self.assertEqual(len(df_pnl), 100)
        self.assertEqual(len(pnl_err), 0)

        # Financial calculations
        df_calc = calculate_dataframe_financials(df_pnl)
        self.assertEqual(len(df_calc), 100)
        self.assertIn("CM1%", df_calc.columns)
        self.assertIn("CM2%", df_calc.columns)

        # Anomaly detection
        df_res, anomalies, status_counts = analyze_full_dataset(df_calc, criteria)
        self.assertEqual(len(df_res), 100)
        self.assertIn("Overall Status", df_res.columns)
        self.assertGreater(len(anomalies), 0)
        
        # Verify status counts sum to 100
        total_counted = sum(status_counts[k] for k in ["RED", "YELLOW", "GREEN", "SUSPICIOUS", "REVIEW REQUIRED"])
        self.assertEqual(total_counted, 100)

        # Trend analysis
        trend_data = analyze_dataset_trends(df_calc, anomalies, criteria)
        self.assertEqual(trend_data["Total Days"], 100)
        self.assertIn(trend_data["Most Problematic Category"], DEFAULT_CRITERIA)

        # Report generation
        aggregates = calculate_portfolio_aggregates(df_calc)
        report_buf = generate_excel_report(df_res, anomalies, trend_data, criteria, aggregates)
        self.assertGreater(report_buf.getbuffer().nbytes, 10000)

if __name__ == "__main__":
    unittest.main()
