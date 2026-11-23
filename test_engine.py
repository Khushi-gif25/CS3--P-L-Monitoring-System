"""
PROZO P&L MONITOR - COMPREHENSIVE TEST SUITE
Validates Single-Workbook architecture (Case Study 3.xlsx), 
synthetic scenarios, calculations, anomaly detection, AI insights, and all 6 deliverables.
"""

import io
import os
import openpyxl
import pandas as pd
from data_processor import inspect_and_read_single_workbook
from financial_engine import calculate_financial_metrics, classify_fc_status
from anomaly_detector import detect_anomalies_for_dataset
from ai_insights import generate_all_ai_insights, generate_overall_trend_ai_observations
from report_generator import (
    generate_complete_pnl_excel, generate_anomaly_excel,
    generate_insights_pdf, generate_daily_digest_pdf, generate_trend_pdf,
    generate_zip_bundle
)
from synthetic_data import (
    get_standard_ideal_criteria, get_test_scenario_records,
    generate_single_workbook_excel
)


def run_all_tests():
    print("=" * 65)
    print("RUNNING PROZO P&L MONITOR SINGLE-WORKBOOK TEST SUITE")
    print("=" * 65)

    targets = get_standard_ideal_criteria()

    # -------------------------------------------------------------
    # 1. Test Ingestion of Reference Case Study 3.xlsx
    # -------------------------------------------------------------
    case_path = os.path.join(os.path.dirname(__file__), "data", "Case Study 3.xlsx")
    print(f"\n[Test 1] Testing ingestion of actual Case Study 3.xlsx ({case_path})...")
    assert os.path.exists(case_path), "Case Study 3.xlsx not found in data directory"
    
    df_case, case_targets, case_info = inspect_and_read_single_workbook(case_path)
    assert case_info["is_valid"], f"Case Study 3 validation failed: {case_info['errors']}"
    assert case_info["criteria_sheet_name"] == "Ideal FC Criteria"
    assert case_info["pnl_sheet_name"] == "Day Wise PnL"
    assert case_info["row_count"] == 100, f"Expected 100 rows, got {case_info['row_count']}"
    assert "Manpower" in case_targets and case_targets["Manpower"]["min"] == 25.0 and case_targets["Manpower"]["max"] == 27.0
    assert "CM1" in case_targets and "CM2" in case_targets
    print(f"✓ Case Study 3.xlsx PASSED: 100 records extracted, 2 sheets validated independently!")

    # -------------------------------------------------------------
    # 2. Test Synthetic 2-Sheet Workbook Generator
    # -------------------------------------------------------------
    print("\n[Test 2] Testing synthetic single-workbook generator...")
    wb_bytes = generate_single_workbook_excel("all_scenarios_combined")
    df_syn, syn_targets, syn_info = inspect_and_read_single_workbook(io.BytesIO(wb_bytes))
    assert syn_info["is_valid"], f"Synthetic workbook validation failed: {syn_info['errors']}"
    assert len(df_syn) >= 7
    print(f"✓ Synthetic single workbook PASSED: {len(df_syn)} records extracted successfully!")

    # -------------------------------------------------------------
    # 3. Test Missing Sheet Error Handling
    # -------------------------------------------------------------
    print("\n[Test 3] Testing missing worksheet error handling...")
    wb_broken = openpyxl.Workbook()
    wb_broken.active.title = "Random Sheet"
    broken_bytes = io.BytesIO()
    wb_broken.save(broken_bytes)
    
    _, _, broken_info = inspect_and_read_single_workbook(io.BytesIO(broken_bytes.getvalue()))
    assert not broken_info["is_valid"], "Expected validation failure for missing sheets"
    assert len(broken_info["errors"]) >= 2
    print("✓ Missing worksheet validation PASSED: Reported clear error messages.")

    # -------------------------------------------------------------
    # 4. Test Scenario 1: Clean Day
    # -------------------------------------------------------------
    print("\n[Test 4] Testing Scenario 1: Clean Day...")
    recs_clean = get_test_scenario_records("clean")
    df_clean = pd.DataFrame(recs_clean)
    df_m1 = calculate_financial_metrics(df_clean)
    anom_1 = detect_anomalies_for_dataset(df_m1, targets)
    
    assert df_m1.iloc[0]["FC Status"] == "GREEN"
    assert df_m1.iloc[0]["CM1 %"] == 68.0
    assert df_m1.iloc[0]["CM2 %"] == 24.0
    assert len(anom_1) == 0
    print("✓ Scenario 1 PASSED: GREEN status, CM1=68%, CM2=24%, 0 anomalies.")

    # -------------------------------------------------------------
    # 5. Test Scenario 2: CM2 Breach (RED status)
    # -------------------------------------------------------------
    print("\n[Test 5] Testing Scenario 2: CM2 Breach (RED)...")
    recs_breach = get_test_scenario_records("cm2_breach")
    df_breach = pd.DataFrame(recs_breach)
    df_m2 = calculate_financial_metrics(df_breach)
    anom_2 = detect_anomalies_for_dataset(df_m2, targets)

    assert df_m2.iloc[0]["FC Status"] == "RED"
    assert df_m2.iloc[0]["CM2 %"] < 12.0
    assert len(anom_2) > 0
    print(f"✓ Scenario 2 PASSED: RED status (CM2={df_m2.iloc[0]['CM2 %']}%), {len(anom_2)} anomalies detected.")

    # -------------------------------------------------------------
    # 6. Test Scenario 3: Line-item-only Breach
    # -------------------------------------------------------------
    print("\n[Test 6] Testing Scenario 3: Line-item-only Breach...")
    recs_line = get_test_scenario_records("line_item_breach")
    df_line = pd.DataFrame(recs_line)
    df_m3 = calculate_financial_metrics(df_line)
    anom_3 = detect_anomalies_for_dataset(df_m3, targets)

    assert df_m3.iloc[0]["FC Status"] == "GREEN"
    line_anoms = anom_3[anom_3["Metric Type"] == "Line Item"]
    assert len(line_anoms) == 1
    assert line_anoms.iloc[0]["Line Item"] == "Packaging"
    print(f"✓ Scenario 3 PASSED: Overall GREEN status with exactly 1 line item breach ({line_anoms.iloc[0]['Line Item']}).")

    # -------------------------------------------------------------
    # 7. Test Scenario 4: Suspiciously High CM2
    # -------------------------------------------------------------
    print("\n[Test 7] Testing Scenario 4: Suspiciously High CM2...")
    recs_susp = get_test_scenario_records("suspicious_high")
    df_susp = pd.DataFrame(recs_susp)
    df_m4 = calculate_financial_metrics(df_susp)
    anom_4 = detect_anomalies_for_dataset(df_m4, targets)

    assert df_m4.iloc[0]["FC Status"] == "SUSPICIOUS"
    assert df_m4.iloc[0]["CM2 %"] > 33.0
    print(f"✓ Scenario 4 PASSED: SUSPICIOUS status triggered (CM2={df_m4.iloc[0]['CM2 %']}%).")

    # -------------------------------------------------------------
    # 8. Test Full Pipeline & All 6 Deliverables on Case Study 3.xlsx
    # -------------------------------------------------------------
    print("\n[Test 8] Running Full Processing Pipeline on Case Study 3.xlsx...")
    df_metrics = calculate_financial_metrics(df_case)
    df_anomalies = detect_anomalies_for_dataset(df_metrics, case_targets)
    insights = generate_all_ai_insights(df_metrics, df_anomalies, case_targets)
    trends = generate_overall_trend_ai_observations(df_metrics, df_anomalies, case_targets)
    
    assert len(df_metrics) == 100
    assert len(df_anomalies) > 0
    assert len(insights) > 0
    print(f"✓ Pipeline execution PASSED: Processed 100 days ({len(df_anomalies)} anomalies, {len(insights)} AI insights).")

    print("\n[Test 9] Generating File 1: Complete_FC_PnL_Report.xlsx (8 worksheets)...")
    file1_excel = generate_complete_pnl_excel(df_metrics, df_anomalies, insights, trends, case_targets)
    wb = openpyxl.load_workbook(io.BytesIO(file1_excel))
    expected_sheets = [
        "Daily P&L Summary", "Expense Breakdown", "CM1 and CM2 Analysis",
        "Anomaly Log", "AI Insights", "Daily FC Digest", "Trend Analysis", "Ideal FC Criteria"
    ]
    for s in expected_sheets:
        assert s in wb.sheetnames, f"Worksheet {s} not found!"
    print(f"✓ File 1 PASSED: 8 worksheets verified ({len(file1_excel):,} bytes).")

    print("\n[Test 10] Generating File 2: Anomaly_Report.xlsx...")
    file2_excel = generate_anomaly_excel(df_anomalies)
    assert len(file2_excel) > 5000
    print(f"✓ File 2 PASSED ({len(file2_excel):,} bytes).")

    print("\n[Test 11] Generating PDF Reports (Files 3, 4, 5)...")
    file3_pdf = generate_insights_pdf(insights, trends, df_anomalies)
    file4_pdf = generate_daily_digest_pdf(df_metrics, df_anomalies, insights)
    file5_pdf = generate_trend_pdf(df_metrics, df_anomalies, trends, case_targets)
    assert len(file3_pdf) > 5000
    assert len(file4_pdf) > 5000
    assert len(file5_pdf) > 2000
    print(f"✓ Files 3, 4, 5 PDF Generation PASSED (Insights: {len(file3_pdf):,} B, Digest: {len(file4_pdf):,} B, Trend: {len(file5_pdf):,} B).")

    print("\n[Test 12] Generating File 6: All_Reports.zip...")
    zip_bundle = generate_zip_bundle(file1_excel, file2_excel, file3_pdf, file4_pdf, file5_pdf)
    assert len(zip_bundle) > 20000
    print(f"✓ File 6 All_Reports.zip PASSED ({len(zip_bundle):,} bytes).")

    print("\n" + "=" * 65)
    print("ALL 12 SINGLE-WORKBOOK TESTS PASSED WITH 100% SUCCESS!")
    print("=" * 65)


if __name__ == "__main__":
    run_all_tests()
