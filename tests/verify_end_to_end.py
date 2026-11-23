"""
End-to-End Verification Script
Prozo Daily FC P&L Monitoring System

Executes complete pipeline:
1. Dynamic Criteria Extraction
2. Day-wise P&L Extraction (100 days)
3. Deterministic Calculations
4. Anomaly Detection
5. AI Insight Generation
6. Trend Analysis
7. 5-Sheet Excel Workbook Export
8. 4 Required Test Scenarios
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import openpyxl
import pandas as pd
from core.data_extractor import extract_criteria, extract_pnl
from core.financial_engine import calculate_dataframe_financials, calculate_portfolio_aggregates
from core.anomaly_detector import analyze_full_dataset
from core.trend_analyzer import analyze_dataset_trends
from core.report_generator import generate_excel_report
from core.test_scenarios import run_all_scenarios

def run_verification():
    print("=================================================================")
    print("PROZO DAILY FC P&L MONITOR - END-TO-END VERIFICATION")
    print("=================================================================")

    xlsx_path = r"C:\Users\Lenovo\.gemini\antigravity\scratch\prozo_pnl_monitor\data\Case Study 3.xlsx"
    assert os.path.exists(xlsx_path), f"File {xlsx_path} missing!"

    # 1. Extraction
    criteria, c_warn = extract_criteria(xlsx_path)
    print(f"\n[1] Extracted FC Criteria ({len(criteria)} categories):")
    for cat, (lo, hi) in criteria.items():
        print(f"    - {cat}: {lo}% - {hi}%")

    df_pnl, p_err, raw_records = extract_pnl(xlsx_path)
    assert df_pnl is not None and len(df_pnl) == 100, f"Expected 100 rows, got {len(df_pnl) if df_pnl is not None else 0}"
    print(f"\n[2] Extracted Day-Wise PnL: {len(df_pnl)} valid days extracted.")
    print(f"    Validation errors: {len(p_err)}")

    # 2. Financial Calculations
    df_calc = calculate_dataframe_financials(df_pnl)
    aggregates = calculate_portfolio_aggregates(df_calc)
    print(f"\n[3] Portfolio Aggregates (Revenue-Weighted):")
    print(f"    - Total Revenue : ₹{aggregates['Total Revenue']:,.2f}")
    print(f"    - Total CM1     : ₹{aggregates['Total CM1']:,.2f} ({aggregates['Overall CM1%']:.2f}%)")
    print(f"    - Total CM2     : ₹{aggregates['Total CM2']:,.2f} ({aggregates['Overall CM2%']:.2f}%)")

    # 3. Anomaly Detection
    df_full, master_anomalies, status_counts = analyze_full_dataset(df_calc, criteria)
    print(f"\n[4] Status Classification & Anomaly Detection:")
    for st_name, count in status_counts.items():
        print(f"    - {st_name:15}: {count} days")
    print(f"    - Total Anomalies Detected: {len(master_anomalies)}")

    # 4. Trend Analysis
    trend_data = analyze_dataset_trends(df_full, master_anomalies, criteria)
    print(f"\n[5] Trend Intelligence:")
    print(f"    - Most Problematic Category: {trend_data['Most Problematic Category']} ({trend_data['Most Problematic Breaches']} breaches)")
    print(f"    - CM1 Average : {trend_data['CM1 Average %']:.2f}%")
    print(f"    - CM2 Average : {trend_data['CM2 Average %']:.2f}%")
    low_cm2_cnt = trend_data.get('Low CM2 Days Count (< 12%)', 0)
    susp_cm2_cnt = trend_data.get('Suspicious CM2 Days Count (> 33%)', 0)
    print(f"    - Low CM2 Days (< 12%)        : {low_cm2_cnt}")
    print(f"    - Suspicious CM2 Days (> 33%) : {susp_cm2_cnt}")
    print(f"    - Grounded Trend Insights     : {len(trend_data['Trend Insights'])}")

    # 5. Export Report
    buf = generate_excel_report(df_full, master_anomalies, trend_data, criteria, aggregates)
    wb = openpyxl.load_workbook(buf)
    print(f"\n[6] Generated Excel Report Worksheets:")
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"    - Sheet '{sheet_name}': {ws.max_row} rows, {ws.max_column} cols")
    assert len(wb.sheetnames) == 5, f"Expected 5 sheets, got {len(wb.sheetnames)}"

    # 6. Test Scenarios
    print(f"\n[7] Verifying 4 Required Test Scenarios:")
    all_scenarios = run_all_scenarios()
    for s_id, s_res in all_scenarios.items():
        pass_str = "PASSED" if s_res["all_passed"] else "FAILED"
        print(f"    - Scenario {s_id} ({s_res['title']}): {pass_str}")
        assert s_res["all_passed"], f"Scenario {s_id} failed: {s_res['checks']}"

    print("\n=================================================================")
    print("ALL VERIFICATIONS COMPLETED SUCCESSFULLY!")
    print("=================================================================")

if __name__ == "__main__":
    run_verification()
