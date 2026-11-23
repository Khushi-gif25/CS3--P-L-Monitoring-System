"""
PROZO – AI-POWERED DAILY FC P&L MONITOR & ANOMALY DETECTION SYSTEM
Main Streamlit Application - Single Excel Workbook Architecture
"""

import io
import os
import time
from typing import Dict, Any, List
import streamlit as st
import pandas as pd
import numpy as np

# Internal analytical engines
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

# -----------------------------------------------------------------------------
# Page Configuration & Corporate Styling
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Prozo – Daily FC P&L Monitor & Anomaly Detection",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

CUSTOM_CSS = """
<style>
    /* Global Brand Header */
    .prozo-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #1E40AF 50%, #3B82F6 100%);
        padding: 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .prozo-title {
        font-size: 26px;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
        color: #FFFFFF;
    }
    .prozo-subtitle {
        font-size: 14px;
        color: #E2E8F0;
        margin-top: 6px;
        margin-bottom: 0;
    }
    /* Metric Cards */
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-val {
        font-size: 22px;
        font-weight: 700;
        color: #1E293B;
    }
    .metric-label {
        font-size: 12px;
        color: #64748B;
        text-transform: uppercase;
        font-weight: 600;
        margin-bottom: 4px;
    }
    /* Status Badges */
    .badge-green {
        background-color: #DCFCE7;
        color: #166534;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 6px;
        display: inline-block;
        border: 1px solid #86EFAC;
    }
    .badge-red {
        background-color: #FEE2E2;
        color: #991B1B;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 6px;
        display: inline-block;
        border: 1px solid #FCA5A5;
    }
    .badge-yellow {
        background-color: #FEF3C7;
        color: #92400E;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 6px;
        display: inline-block;
        border: 1px solid #FCD34D;
    }
    .badge-suspicious {
        background-color: #F3E8FF;
        color: #6B21A8;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 6px;
        display: inline-block;
        border: 1px solid #D8B4FE;
    }
    .badge-review {
        background-color: #F1F5F9;
        color: #334155;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 6px;
        display: inline-block;
        border: 1px solid #CBD5E1;
    }
    /* Executive Insight Box */
    .insight-box {
        border-left: 4px solid #1E3A8A;
        background-color: #F8FAFC;
        padding: 16px;
        border-radius: 0 8px 8px 0;
        margin-bottom: 16px;
        border-top: 1px solid #E2E8F0;
        border-right: 1px solid #E2E8F0;
        border-bottom: 1px solid #E2E8F0;
    }
    /* Sheet Validation Pill */
    .validation-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 12px;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# Initialize Session State
# -----------------------------------------------------------------------------
def init_session_state():
    defaults = {
        "workbook_file_name": None,
        "workbook_validation_info": None,
        "pnl_df_raw": None,
        "criteria_targets_original": None,
        "criteria_targets_active": None,
        "processing_complete": False,
        "df_metrics": None,
        "df_anomalies": None,
        "ai_insights": None,
        "trend_data": None,
        "report_files": {},
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

init_session_state()


# -----------------------------------------------------------------------------
# Pipeline Execution Engine
# -----------------------------------------------------------------------------
def run_financial_pipeline():
    """Executes the 9-stage complete financial processing pipeline."""
    if st.session_state["pnl_df_raw"] is None or st.session_state["criteria_targets_active"] is None:
        st.error("Cannot run processing: Both 'Ideal FC Criteria' and 'Day Wise PnL' worksheets must be validated first.")
        return False

    progress_bar = st.progress(0.0)
    status_text = st.empty()

    steps = [
        (0.11, "1. Reading Excel workbook and extracting worksheets..."),
        (0.22, "2. Validating daily P&L records and schema..."),
        (0.33, "3. Extracting target cost corridors and CM1/CM2 margins..."),
        (0.44, "4. Calculating financial metrics programmatically (CM1, CM2, Expense %)..."),
        (0.55, "5. Detecting financial anomalies against target corridors..."),
        (0.66, "6. Generating AI executive insights and root-cause rankings..."),
        (0.77, "7. Analysing multi-day profitability trends..."),
        (0.88, "8. Preparing daily colour-coded FC digests..."),
        (1.00, "9. Generating downloadable Excel, PDF reports and ZIP bundle..."),
    ]

    try:
        # Step 1 - 3: Ingestion confirmed
        for pct, desc in steps[:3]:
            status_text.markdown(f"**Processing:** {desc}")
            progress_bar.progress(pct)
            time.sleep(0.06)

        # Step 4: Financial Metrics Calculation
        status_text.markdown(f"**Processing:** {steps[3][1]}")
        progress_bar.progress(steps[3][0])
        df_metrics = calculate_financial_metrics(st.session_state["pnl_df_raw"])
        st.session_state["df_metrics"] = df_metrics
        time.sleep(0.06)

        # Step 5: Anomaly Detection
        status_text.markdown(f"**Processing:** {steps[4][1]}")
        progress_bar.progress(steps[4][0])
        active_targets = st.session_state["criteria_targets_active"]
        df_anomalies = detect_anomalies_for_dataset(df_metrics, active_targets)
        st.session_state["df_anomalies"] = df_anomalies
        time.sleep(0.06)

        # Step 6: AI Insights Generation
        status_text.markdown(f"**Processing:** {steps[5][1]}")
        progress_bar.progress(steps[5][0])
        ai_insights = generate_all_ai_insights(df_metrics, df_anomalies, active_targets)
        st.session_state["ai_insights"] = ai_insights
        time.sleep(0.06)

        # Step 7: Trend Analysis
        status_text.markdown(f"**Processing:** {steps[6][1]}")
        progress_bar.progress(steps[6][0])
        trend_data = generate_overall_trend_ai_observations(df_metrics, df_anomalies, active_targets)
        st.session_state["trend_data"] = trend_data
        time.sleep(0.06)

        # Step 8: Digest
        status_text.markdown(f"**Processing:** {steps[7][1]}")
        progress_bar.progress(steps[7][0])
        time.sleep(0.06)

        # Step 9: Report Generation
        status_text.markdown(f"**Processing:** {steps[8][1]}")
        progress_bar.progress(steps[8][0])
        
        file1_excel_complete = generate_complete_pnl_excel(df_metrics, df_anomalies, ai_insights, trend_data, active_targets)
        file2_excel_anom = generate_anomaly_excel(df_anomalies)
        file3_pdf_insights = generate_insights_pdf(ai_insights, trend_data, df_anomalies)
        file4_pdf_digest = generate_daily_digest_pdf(df_metrics, df_anomalies, ai_insights)
        file5_pdf_trend = generate_trend_pdf(df_metrics, df_anomalies, trend_data, active_targets)
        zip_bundle = generate_zip_bundle(file1_excel_complete, file2_excel_anom, file3_pdf_insights, file4_pdf_digest, file5_pdf_trend)

        st.session_state["report_files"] = {
            "Complete_FC_PnL_Report.xlsx": (file1_excel_complete, len(file1_excel_complete), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "Anomaly_Report.xlsx": (file2_excel_anom, len(file2_excel_anom), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "AI_Insights_Report.pdf": (file3_pdf_insights, len(file3_pdf_insights), "application/pdf"),
            "Daily_FC_Digest.pdf": (file4_pdf_digest, len(file4_pdf_digest), "application/pdf"),
            "Trend_Analysis.pdf": (file5_pdf_trend, len(file5_pdf_trend), "application/pdf"),
            "All_Reports.zip": (zip_bundle, len(zip_bundle), "application/zip"),
        }

        st.session_state["processing_complete"] = True
        status_text.success("🎉 Financial Processing Pipeline & Deliverable Reports Generated Successfully!")
        progress_bar.progress(1.0)
        return True

    except Exception as e:
        status_text.error(f"❌ Processing Pipeline Error: {str(e)}")
        st.session_state["processing_complete"] = False
        return False


def load_case_study_direct():
    """Loads Case Study 3.xlsx directly from the local repository."""
    case_path = os.path.join(os.path.dirname(__file__), "data", "Case Study 3.xlsx")
    if os.path.exists(case_path):
        df_pnl, targets, info = inspect_and_read_single_workbook(case_path)
        if info["is_valid"]:
            st.session_state["workbook_file_name"] = "Case Study 3.xlsx"
            st.session_state["workbook_validation_info"] = info
            st.session_state["pnl_df_raw"] = df_pnl
            st.session_state["criteria_targets_original"] = {k: v.copy() for k, v in targets.items()}
            st.session_state["criteria_targets_active"] = {k: v.copy() for k, v in targets.items()}
            st.session_state["processing_complete"] = False
            st.session_state["report_files"] = {}
            return True
    return False


def load_synthetic_scenario(scenario_key: str):
    """Loads a pre-packaged synthetic test scenario as a single 2-sheet workbook."""
    wb_bytes = generate_single_workbook_excel(scenario_key)
    df_pnl, targets, info = inspect_and_read_single_workbook(io.BytesIO(wb_bytes))
    if info["is_valid"]:
        st.session_state["workbook_file_name"] = f"Synthetic_{scenario_key}.xlsx"
        st.session_state["workbook_validation_info"] = info
        st.session_state["pnl_df_raw"] = df_pnl
        st.session_state["criteria_targets_original"] = {k: v.copy() for k, v in targets.items()}
        st.session_state["criteria_targets_active"] = {k: v.copy() for k, v in targets.items()}
        st.session_state["processing_complete"] = False
        st.session_state["report_files"] = {}


# -----------------------------------------------------------------------------
# Sidebar Navigation & Brand Identity
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🏢 **PROZO P&L MONITOR**")
    st.markdown("**Daily FC Financial Governance**")
    st.caption("Single Workbook • Two Worksheets • AI Directives")
    st.divider()

    nav_selection = st.radio(
        "Navigation Menu",
        [
            "1. Overview & Architecture",
            "2. Upload FC Financial Data",
            "3. Daily FC P&L Dashboard",
            "4. Flagged Anomaly Log",
            "5. AI Insights & Action Plans",
            "6. Daily FC Colour Digest",
            "7. Trend Analysis",
            "8. Active FC Targets",
            "9. Download Reports",
            "10. System Architecture & Roadmap",
        ],
        index=1 if not st.session_state["processing_complete"] else 2
    )

    st.divider()
    st.markdown("#### 📁 Reference Case Study")
    if st.button("📊 Load Case Study 3.xlsx", type="primary", use_container_width=True, help="Load the actual Case Study 3.xlsx (100 days of data across both sheets)"):
        if load_case_study_direct():
            st.rerun()

    st.markdown("#### ⚡ Quick Test Scenarios")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("🟢 Clean Day", use_container_width=True, help="Test 1: All metrics in range, GREEN status"):
            load_synthetic_scenario("clean")
            st.rerun()
    with col_t2:
        if st.button("🔴 CM2 Breach", use_container_width=True, help="Test 2: CM2 < 12%, RED status"):
            load_synthetic_scenario("cm2_breach")
            st.rerun()

    col_t3, col_t4 = st.columns(2)
    with col_t3:
        if st.button("📦 Line Breach", use_container_width=True, help="Test 3: Packaging breach, GREEN CM2"):
            load_synthetic_scenario("line_item_breach")
            st.rerun()
    with col_t4:
        if st.button("🟣 Suspicious", use_container_width=True, help="Test 4: CM2 > 33%, underreporting warning"):
            load_synthetic_scenario("suspicious_high")
            st.rerun()

    st.divider()
    st.caption("Prozo Supply Chain Network • P&L Monitor v3.0")


# -----------------------------------------------------------------------------
# Header Banner
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="prozo-header">
        <h1 class="prozo-title">PROZO – DAILY FC P&L MONITOR & ANOMALY DETECTION SYSTEM</h1>
        <p class="prozo-subtitle">Autonomous financial monitoring, line-item variance detection, and plain-language executive directives for Prozo Fulfillment Centres</p>
    </div>
    """,
    unsafe_allow_html=True
)


# =============================================================================
# TAB 1: OVERVIEW & ARCHITECTURE
# =============================================================================
if nav_selection == "1. Overview & Architecture":
    st.markdown("### 🌐 System Architecture & User Journey")
    st.markdown(
        """
        The **Prozo FC P&L Monitor** operates on a streamlined single-file ingestion model:
        **One Excel Workbook (`.xlsx`) containing two separate worksheets**:
        1. **`Ideal FC Criteria`**: Target expense percentage ranges (e.g. Manpower 25-27%, CM1 65-70%, CM2 15-30%).
        2. **`Day Wise PnL`**: Daily operating revenue and cost breakdown.
        """
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-label">Step 1</div>
                <div class="metric-val">1 Workbook</div>
                <small>Two Worksheets Ingested</small>
            </div>
            """, unsafe_allow_html=True
        )
    with c2:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-label">Step 2</div>
                <div class="metric-val">1-Click Pipeline</div>
                <small>9-Stage Calculation Engine</small>
            </div>
            """, unsafe_allow_html=True
        )
    with c3:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-label">Step 3</div>
                <div class="metric-val">AI Governance</div>
                <small>Impact-Ranked Directives</small>
            </div>
            """, unsafe_allow_html=True
        )
    with c4:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-label">Step 4</div>
                <div class="metric-val">Download Pack</div>
                <small>5 Pro Reports + All_Reports.zip</small>
            </div>
            """, unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("#### 📐 Mathematical Governance Corridor")
    
    col_crit1, col_crit2 = st.columns(2)
    with col_crit1:
        st.markdown("**Predefined Expense Target Corridors (% of Revenue):**")
        df_crit_table = pd.DataFrame([
            {"Expense Category": "Manpower Cost", "Target Min": "25.0%", "Target Max": "27.0%", "Operational Driver": "Direct Labor & Picking Shifts"},
            {"Expense Category": "Packaging Cost", "Target Min": "5.0%", "Target Max": "7.0%", "Operational Driver": "Cartons, Dunnage, Tape"},
            {"Expense Category": "Power & Fuel", "Target Min": "10.0%", "Target Max": "12.0%", "Operational Driver": "HVAC, Conveyors, DG Run-Hours"},
            {"Expense Category": "FC Rent", "Target Min": "15.0%", "Target Max": "18.0%", "Operational Driver": "Facility Base Lease"},
            {"Expense Category": "Equipment Rentals", "Target Min": "5.0%", "Target Max": "8.0%", "Operational Driver": "Forklifts, MHE, Reach Trucks"},
            {"Expense Category": "Overheads", "Target Min": "10.0%", "Target Max": "12.0%", "Operational Driver": "Repairs, Admin, Security"},
        ])
        st.dataframe(df_crit_table, hide_index=True, use_container_width=True)

    with col_crit2:
        st.markdown("**Contribution Margin & FC Colour Classifications:**")
        st.markdown(
            """
            - **CM1 (Contribution Margin 1):** `Revenue - Manpower Cost - Packaging Cost`  
              *Target:* **65.0% – 70.0%** (Flagged if deviation > 3.0 percentage points)
            - **CM2 (Contribution Margin 2):** `CM1 - Power & Fuel - FC Rent - Equipment Rentals - Overheads`  
              *Target:* **15.0% – 30.0%**
            - **FC Colour Status Classification Rules:**
              - 🔴 **RED:** CM2 < 12.0% (Critical Profitability Breach)
              - 🟡 **YELLOW:** 12.0% ≤ CM2 ≤ 14.9% OR 30.0% < CM2 ≤ 32.9%
              - 🟢 **GREEN:** 15.0% ≤ CM2 ≤ 30.0% (Operational Target Corridor)
              - 🟣 **SUSPICIOUS:** CM2 > 33.0% (Likely missing or underreported expense accruals)
              - ⚪ **REVIEW REQUIRED:** 14.9% < CM2 < 15.0% OR 32.9% < CM2 ≤ 33.0% (Boundary gaps)
            """
        )

    st.info("💡 **Ready to begin?** Navigate to **'2. Upload FC Financial Data'** or click **'Load Case Study 3.xlsx'** in the left sidebar.")


# =============================================================================
# TAB 2: UPLOAD FC FINANCIAL DATA
# =============================================================================
elif nav_selection == "2. Upload FC Financial Data":
    st.markdown("### 📁 Input Section – Single Excel Workbook Upload")
    st.markdown("Upload **one Excel workbook (.xlsx)** containing both **`Ideal FC Criteria`** and **`Day Wise PnL`** worksheets.")

    # Reference Template Download
    case_path = os.path.join(os.path.dirname(__file__), "data", "Case Study 3.xlsx")
    with st.expander("📥 Download Reference Excel Template (Case Study 3.xlsx)", expanded=False):
        if os.path.exists(case_path):
            with open(case_path, "rb") as f:
                case_bytes = f.read()
            st.download_button(
                label="📄 Download Case Study 3.xlsx (Reference Structure)",
                data=case_bytes,
                file_name="Case Study 3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.download_button(
                label="📄 Download Sample 2-Sheet Workbook (.xlsx)",
                data=generate_single_workbook_excel("all_scenarios_combined"),
                file_name="Sample_FC_Financial_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # SINGLE EXCEL UPLOADER
    uploaded_file = st.file_uploader("Upload FC Financial Data (.xlsx)", type=["xlsx", "xls"], key="uploader_single_wb")

    if uploaded_file is not None:
        df_pnl, targets, info = inspect_and_read_single_workbook(uploaded_file)
        st.session_state["workbook_file_name"] = uploaded_file.name
        st.session_state["workbook_validation_info"] = info

        if info["is_valid"]:
            st.session_state["pnl_df_raw"] = df_pnl
            st.session_state["criteria_targets_original"] = {k: v.copy() for k, v in targets.items()}
            st.session_state["criteria_targets_active"] = {k: v.copy() for k, v in targets.items()}
        else:
            st.session_state["pnl_df_raw"] = None
            st.session_state["criteria_targets_active"] = None

    # DISPLAY VALIDATION STATUS
    if st.session_state["workbook_file_name"] is not None:
        info = st.session_state["workbook_validation_info"]
        fname = st.session_state["workbook_file_name"]

        if info and info["is_valid"]:
            st.success(f"✓ Workbook **{fname}** validated successfully! Found worksheets: `{', '.join(info['sheet_names'])}`")
            
            c_v1, c_v2 = st.columns(2)
            with c_v1:
                st.markdown(
                    f"""
                    <div class="validation-card">
                        <div style="font-weight:700; color:#166534; margin-bottom:4px;">✓ Sheet 1: {info['criteria_sheet_name']} (Validated)</div>
                        <div style="font-size:13px; color:#334155;">Extracted <b>{len(info['extracted_targets'])}</b> target cost and margin corridors (Manpower 25-27%, Packaging 5-7%, CM1 65-70%, CM2 15-30%).</div>
                    </div>
                    """, unsafe_allow_html=True
                )
            with c_v2:
                fc_note = "FC Identifier detected" if info['has_fc'] else "FC identifier absent (Day-wise aggregated analysis enabled)"
                st.markdown(
                    f"""
                    <div class="validation-card">
                        <div style="font-weight:700; color:#166534; margin-bottom:4px;">✓ Sheet 2: {info['pnl_sheet_name']} (Validated)</div>
                        <div style="font-size:13px; color:#334155;">Extracted <b>{info['row_count']}</b> operational daily records. {fc_note}.</div>
                    </div>
                    """, unsafe_allow_html=True
                )

            if info.get("warnings"):
                with st.expander("ℹ️ Data Ingestion Notes & Warnings", expanded=False):
                    for w in info["warnings"]:
                        st.caption(f"• {w}")
        else:
            st.error(f"❌ Validation Failed for **{fname}**:")
            if info and info.get("errors"):
                for err in info["errors"]:
                    st.markdown(f"- {err}")

    st.markdown("---")

    # GENERATE OUTPUT BUTTON
    can_generate = (
        st.session_state["pnl_df_raw"] is not None and
        st.session_state["criteria_targets_active"] is not None
    )

    if can_generate:
        st.markdown("### 🚀 Step 2: Generate Output")
        st.markdown("Workbook successfully validated. Click the button below to execute the complete analytical pipeline.")
        
        generate_clicked = st.button(
            "⚡ GENERATE OUTPUT",
            type="primary",
            use_container_width=True,
            help="Executes 9-stage financial calculation, anomaly detection, AI insights, and report generation"
        )

        if generate_clicked:
            success = run_financial_pipeline()
            if success:
                st.balloons()
                st.info("👉 Processing complete! Explore the **Daily FC P&L Dashboard**, **Flagged Anomaly Log**, **AI Insights**, and **Download Reports** tabs.")
    else:
        st.warning("⚠️ Please upload a valid workbook (.xlsx) with both worksheets or click 'Load Case Study 3.xlsx' in the sidebar.")

    # Preview of loaded data
    if st.session_state["pnl_df_raw"] is not None:
        with st.expander("👀 View Ingested Daily P&L Records Preview", expanded=False):
            st.dataframe(st.session_state["pnl_df_raw"].head(10), use_container_width=True)


# =============================================================================
# TAB 3: DAILY FC P&L DASHBOARD
# =============================================================================
elif nav_selection == "3. Daily FC P&L Dashboard":
    if not st.session_state["processing_complete"] or st.session_state["df_metrics"] is None:
        st.warning("⚠️ Processing not yet run. Please go to **'2. Upload FC Financial Data'** and click **GENERATE OUTPUT**.")
    else:
        df_m = st.session_state["df_metrics"]
        st.markdown("### 📊 Required Output 1 – Per-Day, Per-FC P&L Analysis")
        st.caption("Granular daily financial performance across fulfillment centres with deterministic CM1 and CM2 margins.")

        # KPI Metrics Cards
        k1, k2, k3, k4, k5 = st.columns(5)
        total_rev = df_m["Revenue"].sum()
        avg_cm1 = df_m["CM1 %"].mean()
        avg_cm2 = df_m["CM2 %"].mean()
        green_count = (df_m["FC Status"] == "GREEN").sum()
        red_count = (df_m["FC Status"] == "RED").sum()
        susp_count = (df_m["FC Status"] == "SUSPICIOUS").sum()

        with k1:
            st.markdown(f'<div class="metric-card"><div class="metric-label">Total Revenue</div><div class="metric-val">₹{total_rev:,.0f}</div><small>{len(df_m)} records</small></div>', unsafe_allow_html=True)
        with k2:
            st.markdown(f'<div class="metric-card"><div class="metric-label">Avg CM1 %</div><div class="metric-val">{avg_cm1:.1f}%</div><small>Target: 65% - 70%</small></div>', unsafe_allow_html=True)
        with k3:
            st.markdown(f'<div class="metric-card"><div class="metric-label">Avg CM2 %</div><div class="metric-val">{avg_cm2:.1f}%</div><small>Target: 15% - 30%</small></div>', unsafe_allow_html=True)
        with k4:
            st.markdown(f'<div class="metric-card"><div class="metric-label">Green Days</div><div class="metric-val" style="color:#16A34A">{green_count}</div><small>Target Achieved</small></div>', unsafe_allow_html=True)
        with k5:
            st.markdown(f'<div class="metric-card"><div class="metric-label">Critical / Suspicious</div><div class="metric-val" style="color:#DC2626">{red_count + susp_count}</div><small>{red_count} Red / {susp_count} Susp</small></div>', unsafe_allow_html=True)

        st.markdown("---")

        # Interactive Filters
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1:
            all_fcs = list(df_m["FC"].unique())
            sel_fcs = st.multiselect("Filter by Fulfillment Centre (FC)", options=all_fcs, default=all_fcs)
        with c_f2:
            all_statuses = list(df_m["FC Status"].unique())
            sel_statuses = st.multiselect("Filter by FC Status", options=all_statuses, default=all_statuses)
        with c_f3:
            search_query = st.text_input("Search (Day, Date, FC, etc.)", "")

        df_filtered = df_m[
            (df_m["FC"].isin(sel_fcs)) &
            (df_m["FC Status"].isin(sel_statuses))
        ]
        if search_query:
            df_filtered = df_filtered[
                df_filtered.apply(lambda row: search_query.lower() in str(row.values).lower(), axis=1)
            ]

        # Format Table Display
        display_cols = [
            "Date", "FC", "Revenue",
            "Manpower Cost", "Manpower Cost %",
            "Packaging Cost", "Packaging Cost %",
            "Power and Fuel", "Power and Fuel %",
            "FC Rent", "FC Rent %",
            "Equipment Rentals", "Equipment Rentals %",
            "Overheads", "Overheads %",
            "CM1", "CM1 %",
            "CM2", "CM2 %",
            "FC Status"
        ]

        def highlight_status(val):
            if val == "GREEN":
                return "background-color: #DCFCE7; color: #166534; font-weight: bold;"
            elif val == "RED":
                return "background-color: #FEE2E2; color: #991B1B; font-weight: bold;"
            elif val == "YELLOW":
                return "background-color: #FEF3C7; color: #92400E; font-weight: bold;"
            elif val == "SUSPICIOUS":
                return "background-color: #F3E8FF; color: #6B21A8; font-weight: bold;"
            return "background-color: #F1F5F9; color: #334155;"

        styled_df = df_filtered[display_cols].style.format({
            "Revenue": "₹{:,.2f}",
            "Manpower Cost": "₹{:,.2f}",
            "Packaging Cost": "₹{:,.2f}",
            "Power and Fuel": "₹{:,.2f}",
            "FC Rent": "₹{:,.2f}",
            "Equipment Rentals": "₹{:,.2f}",
            "Overheads": "₹{:,.2f}",
            "CM1": "₹{:,.2f}",
            "CM2": "₹{:,.2f}",
            "Manpower Cost %": "{:.1f}%",
            "Packaging Cost %": "{:.1f}%",
            "Power and Fuel %": "{:.1f}%",
            "FC Rent %": "{:.1f}%",
            "Equipment Rentals %": "{:.1f}%",
            "Overheads %": "{:.1f}%",
            "CM1 %": "{:.1f}%",
            "CM2 %": "{:.1f}%",
        }).map(highlight_status, subset=["FC Status"])

        st.dataframe(styled_df, use_container_width=True, height=400)

        # Record Deep Dive Inspector
        with st.expander("🔍 Inspect Individual Record Breakdown", expanded=False):
            rec_options = [f"{r['Date']} | {r['FC']} | CM2: {r['CM2 %']:.1f}% ({r['FC Status']})" for _, r in df_filtered.iterrows()]
            if rec_options:
                sel_rec = st.selectbox("Select Record to Inspect", rec_options)
                sel_idx = rec_options.index(sel_rec)
                r_detail = df_filtered.iloc[sel_idx]
                
                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    st.markdown(f"**Period / Day:** `{r_detail['Date']}` &nbsp;&nbsp; **FC:** `{r_detail['FC']}`")
                    st.markdown(f"**Gross Revenue:** ₹{r_detail['Revenue']:,.2f}")
                    st.markdown(f"**CM1:** ₹{r_detail['CM1']:,.2f} (`{r_detail['CM1 %']:.1f}%`)")
                    st.markdown(f"**CM2:** ₹{r_detail['CM2']:,.2f} (`{r_detail['CM2 %']:.1f}%`)")
                    st.markdown(f"**Status:** `{r_detail['FC Status']}`")
                    st.info(r_detail["Status Warning"])
                
                with c_d2:
                    st.markdown("**Cost Allocation Distribution:**")
                    cost_breakdown_df = pd.DataFrame([
                        {"Expense Line": "Manpower Cost", "Amount (₹)": r_detail["Manpower Cost"], "% of Rev": f"{r_detail['Manpower Cost %']:.1f}%"},
                        {"Expense Line": "Packaging Cost", "Amount (₹)": r_detail["Packaging Cost"], "% of Rev": f"{r_detail['Packaging Cost %']:.1f}%"},
                        {"Expense Line": "Power and Fuel", "Amount (₹)": r_detail["Power and Fuel"], "% of Rev": f"{r_detail['Power and Fuel %']:.1f}%"},
                        {"Expense Line": "FC Rent", "Amount (₹)": r_detail["FC Rent"], "% of Rev": f"{r_detail['FC Rent %']:.1f}%"},
                        {"Expense Line": "Equipment Rentals", "Amount (₹)": r_detail["Equipment Rentals"], "% of Rev": f"{r_detail['Equipment Rentals %']:.1f}%"},
                        {"Expense Line": "Overheads", "Amount (₹)": r_detail["Overheads"], "% of Rev": f"{r_detail['Overheads %']:.1f}%"},
                    ])
                    st.dataframe(cost_breakdown_df, hide_index=True, use_container_width=True)


# =============================================================================
# TAB 4: FLAGGED ANOMALY LOG
# =============================================================================
elif nav_selection == "4. Flagged Anomaly Log":
    if not st.session_state["processing_complete"] or st.session_state["df_anomalies"] is None:
        st.warning("⚠️ Processing not yet run. Please run **GENERATE OUTPUT**.")
    else:
        df_anom = st.session_state["df_anomalies"]
        st.markdown("### 🚨 Required Output 2 – Flagged Financial Anomaly Log")
        st.caption("Comprehensive log of all cost ceiling overruns, margin breaches, and suspicious deviations.")

        if df_anom.empty:
            st.success("🎉 No financial anomalies detected across any FC records! All metrics conform to active criteria.")
        else:
            c_a1, c_a2, c_a3, c_a4 = st.columns(4)
            total_anom = len(df_anom)
            crit_anom = (df_anom["Severity"] == "Critical").sum()
            high_anom = (df_anom["Severity"] == "High").sum()
            total_leakage = df_anom[df_anom["Financial Impact (₹)"] > 0]["Financial Impact (₹)"].sum()

            with c_a1:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Total Flagged</div><div class="metric-val">{total_anom}</div><small>Anomalies</small></div>', unsafe_allow_html=True)
            with c_a2:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Critical Severity</div><div class="metric-val" style="color:#DC2626">{crit_anom}</div><small>Immediate Action</small></div>', unsafe_allow_html=True)
            with c_a3:
                st.markdown(f'<div class="metric-card"><div class="metric-label">High Severity</div><div class="metric-val" style="color:#D97706">{high_anom}</div><small>Review Required</small></div>', unsafe_allow_html=True)
            with c_a4:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Total Cost Overrun</div><div class="metric-val">₹{total_leakage:,.2f}</div><small>Excess vs Target</small></div>', unsafe_allow_html=True)

            st.markdown("---")

            col_fa1, col_fa2, col_fa3, col_fa4 = st.columns(4)
            with col_fa1:
                fcs = ["All"] + list(df_anom["FC"].unique())
                sel_fc = st.selectbox("Filter by FC", fcs)
            with col_fa2:
                categories = ["All"] + list(df_anom["Line Item"].unique())
                sel_cat = st.selectbox("Filter by Line Item / Margin", categories)
            with col_fa3:
                severities = ["All"] + list(df_anom["Severity"].unique())
                sel_sev = st.selectbox("Filter by Severity", severities)
            with col_fa4:
                statuses = ["All"] + list(df_anom["FC Status"].unique())
                sel_st = st.selectbox("Filter by FC Status", statuses)

            filtered_anom = df_anom.copy()
            if sel_fc != "All":
                filtered_anom = filtered_anom[filtered_anom["FC"] == sel_fc]
            if sel_cat != "All":
                filtered_anom = filtered_anom[filtered_anom["Line Item"] == sel_cat]
            if sel_sev != "All":
                filtered_anom = filtered_anom[filtered_anom["Severity"] == sel_sev]
            if sel_st != "All":
                filtered_anom = filtered_anom[filtered_anom["FC Status"] == sel_st]

            show_cols = [
                "Date", "FC", "Line Item", "Actual %", "Target Range",
                "Status", "Deviation (pp)", "Actual Amount", "Financial Impact (₹)", "Severity", "Explanation"
            ]

            def highlight_anom_status(val):
                if val == "ABOVE TARGET":
                    return "background-color: #FEE2E2; color: #991B1B; font-weight: bold;"
                elif val == "BELOW TARGET":
                    return "background-color: #FEF3C7; color: #92400E; font-weight: bold;"
                return ""

            def highlight_sev(val):
                if val == "Critical":
                    return "color: #DC2626; font-weight: bold;"
                elif val == "High":
                    return "color: #D97706; font-weight: bold;"
                return "color: #2563EB;"

            styled_anom = filtered_anom[show_cols].style.format({
                "Actual %": "{:.1f}%",
                "Deviation (pp)": "{:+.1f} pp",
                "Actual Amount": "₹{:,.2f}",
                "Financial Impact (₹)": "₹{:,.2f}",
            }).map(highlight_anom_status, subset=["Status"]).map(highlight_sev, subset=["Severity"])

            st.dataframe(styled_anom, use_container_width=True, height=450)


# =============================================================================
# TAB 5: AI INSIGHTS & ACTION PLANS
# =============================================================================
elif nav_selection == "5. AI Insights & Action Plans":
    if not st.session_state["processing_complete"] or st.session_state["ai_insights"] is None:
        st.warning("⚠️ Processing not yet run. Please run **GENERATE OUTPUT**.")
    else:
        insights = st.session_state["ai_insights"]
        st.markdown("### 🧠 Required Output 3 – AI-Powered Executive Insights")
        st.caption(f"Deterministic analytical explanations, root-cause categorization, and concrete management directives ({len(insights)} records analyzed).")

        # Top 20 insights display with pagination / expanders if dataset is large (like Case Study 3 with 99 records)
        max_display = min(25, len(insights))
        st.caption(f"Displaying top {max_display} priority anomalies (all {len(insights)} are included in the downloadable PDF and Excel reports):")

        for idx, ins in enumerate(insights[:max_display], 1):
            status = ins["fc_status"]
            badge_class = "badge-green" if status == "GREEN" else (
                "badge-red" if status == "RED" else (
                    "badge-yellow" if status == "YELLOW" else "badge-suspicious"
                )
            )

            st.markdown(
                f"""
                <div class="insight-box">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                        <h4 style="margin:0; color:#1E3A8A;">Period: {ins['date']} &nbsp;|&nbsp; FC: {ins['fc']} &nbsp;|&nbsp; Revenue: ₹{ins['revenue']:,.2f}</h4>
                        <span class="{badge_class}">{status} (CM2: {ins['cm2_pct']:.1f}%)</span>
                    </div>
                    <p style="margin-bottom:6px;"><b>1. What Happened:</b> {ins['what_happened']}</p>
                    <p style="margin-bottom:6px;"><b>2. Size of Deviation:</b> {ins['deviation_size']}</p>
                    <p style="margin-bottom:6px;"><b>3. Breached Line Items:</b> {ins['breached_items']}</p>
                    <p style="margin-bottom:6px;"><b>4. Top Expense Contributors (Sorted by Financial Impact):</b><br/>
                        {'<br/>'.join(['&nbsp;&nbsp;&nbsp;&nbsp;• ' + c['summary'] for c in ins['ranked_contributors']]) if ins['ranked_contributors'] else '&nbsp;&nbsp;&nbsp;&nbsp;• None'}
                    </p>
                    <p style="margin-bottom:6px;"><b>5. Why the Issue May Have Occurred:</b><br/>
                        <b>Confirmed Drivers:</b> {' '.join(ins['confirmed_causes'])}<br/>
                        <b>Possible Operational Drivers:</b> {' '.join(ins['possible_causes']) if ins['possible_causes'] else 'None identified.'}
                    </p>
                    <p style="margin-bottom:0px; color:#1E3A8A;"><b>6. Recommended Management Action:</b><br/>
                        {'<br/>'.join(['&nbsp;&nbsp;&nbsp;&nbsp;👉 ' + act for act in ins['recommended_actions']])}
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )


# =============================================================================
# TAB 6: DAILY FC COLOUR DIGEST
# =============================================================================
elif nav_selection == "6. Daily FC Colour Digest":
    if not st.session_state["processing_complete"] or st.session_state["df_metrics"] is None:
        st.warning("⚠️ Processing not yet run. Please run **GENERATE OUTPUT**.")
    else:
        df_m = st.session_state["df_metrics"]
        df_anom = st.session_state["df_anomalies"]
        insights = st.session_state["ai_insights"]
        
        st.markdown("### 📋 Required Output 4 – Daily Colour-Coded FC Digest")
        st.caption("Central management dashboard card view across all fulfillment centres and operational days.")

        st.info("ℹ️ **Integration Status Notice:** External messaging connectors (Automated Email Dispatch, WhatsApp Business API, Slack Webhooks) are currently disabled. Digests are securely published to this dashboard and available for local PDF/Excel download.")

        # Performance Matrix
        st.markdown("#### 🗺️ Performance Matrix")
        pivot_status = df_m.pivot_table(index="FC", columns="Date", values="CM2 %", aggfunc="first")
        pivot_status_df = pd.DataFrame(pivot_status)

        st.dataframe(
            pivot_status_df.style.format("{:.1f}%").map(
                lambda val: "background-color: #DCFCE7; color: #166534; font-weight: bold;" if (pd.notna(val) and 15.0 <= val <= 30.0) else (
                    "background-color: #FEE2E2; color: #991B1B; font-weight: bold;" if (pd.notna(val) and val < 12.0) else (
                        "background-color: #F3E8FF; color: #6B21A8; font-weight: bold;" if (pd.notna(val) and val > 33.0) else (
                            "background-color: #FEF3C7; color: #92400E; font-weight: bold;" if pd.notna(val) else ""
                        )
                    )
                )
            ),
            use_container_width=True
        )

        st.markdown("#### 📑 Daily Performance Cards (Preview)")
        max_cards = min(20, len(df_m))
        for idx, row in df_m.head(max_cards).iterrows():
            d_val = str(row["Date"])
            fc_val = str(row["FC"])
            status = str(row["FC Status"])
            
            anom_sub = df_anom[(df_anom["Date"] == d_val) & (df_anom["FC"] == fc_val)] if not df_anom.empty else pd.DataFrame()
            anom_cnt = len(anom_sub)
            
            matched_ins = next((i for i in insights if i["date"] == d_val and i["fc"] == fc_val), None)
            summary_txt = matched_ins["what_happened"] if matched_ins else f"Operations within target limits. CM2: {row['CM2 %']:.1f}%."
            rec_action = matched_ins["recommended_actions"][0] if (matched_ins and matched_ins["recommended_actions"]) else "Continue standard procedures."

            badge_style = "badge-green" if status == "GREEN" else (
                "badge-red" if status == "RED" else (
                    "badge-yellow" if status == "YELLOW" else "badge-suspicious"
                )
            )

            with st.container():
                st.markdown(
                    f"""
                    <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:8px; padding:16px; margin-bottom:12px;">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span style="font-size:16px; font-weight:700; color:#1E3A8A;">{fc_val} &nbsp;•&nbsp; {d_val}</span>
                            <span class="{badge_style}">{status}</span>
                        </div>
                        <div style="display:flex; gap:24px; margin-top:8px; font-size:13px; color:#475569;">
                            <span><b>Revenue:</b> ₹{row['Revenue']:,.2f}</span>
                            <span><b>CM1:</b> {row['CM1 %']:.1f}%</span>
                            <span><b>CM2:</b> {row['CM2 %']:.1f}%</span>
                            <span><b>Anomalies:</b> {anom_cnt}</span>
                        </div>
                        <div style="margin-top:8px; font-size:13px; color:#334155;">
                            <b>Summary:</b> {summary_txt}
                        </div>
                        <div style="margin-top:4px; font-size:13px; color:#1E3A8A;">
                            <b>Key Action:</b> {rec_action}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )


# =============================================================================
# TAB 7: TREND ANALYSIS
# =============================================================================
elif nav_selection == "7. Trend Analysis":
    if not st.session_state["processing_complete"] or st.session_state["df_metrics"] is None:
        st.warning("⚠️ Processing not yet run. Please run **GENERATE OUTPUT**.")
    else:
        df_m = st.session_state["df_metrics"]
        df_anom = st.session_state["df_anomalies"]
        trend_data = st.session_state["trend_data"]

        st.markdown("### 📈 Required Output 5 – Multi-Day Trend Analysis")
        st.caption("Historical cost evolution, margin trajectory, and recurring variance detection across actual uploaded records.")

        st.markdown("#### 🤖 Strategic AI Trend Observations")
        for obs in trend_data.get("observations", []):
            st.markdown(f"- {obs}")

        st.markdown("---")

        st.markdown("#### 💰 Profitability & Margin Trajectory (CM1% vs CM2%)")
        cm_trend_df = df_m.copy()
        cm_trend_df["Record Identifier"] = cm_trend_df["Date"]
        chart_data = cm_trend_df.set_index("Record Identifier")[["CM1 %", "CM2 %"]]
        st.line_chart(chart_data)

        st.markdown("---")

        st.markdown("#### 📊 Operational Expense Trends (% of Revenue)")
        exp_chart_cols = [
            "Manpower Cost %", "Packaging Cost %", "Power and Fuel %",
            "FC Rent %", "Equipment Rentals %", "Overheads %"
        ]
        exp_trend_df = cm_trend_df.set_index("Record Identifier")[[c for c in exp_chart_cols if c in cm_trend_df.columns]]
        st.line_chart(exp_trend_df)

        st.markdown("---")

        st.markdown("#### 🔄 Anomaly Frequency & Breached Line Items")
        if not df_anom.empty:
            anom_freq = df_anom["Line Item"].value_counts().reset_index()
            anom_freq.columns = ["Line Item / Margin", "Breach Count"]
            
            c_tr1, c_tr2 = st.columns([1, 1])
            with c_tr1:
                st.dataframe(anom_freq, hide_index=True, use_container_width=True)
            with c_tr2:
                st.bar_chart(anom_freq.set_index("Line Item / Margin"))
        else:
            st.success("No anomalies recorded across any historical days.")


# =============================================================================
# TAB 8: ACTIVE FC TARGETS MANAGEMENT
# =============================================================================
elif nav_selection == "8. Active FC Targets":
    st.markdown("### 🎯 Target Range Management")
    st.caption("Dynamically adjust cost ceilings and profitability thresholds. Saving changes automatically updates calculations and reports.")

    targets = st.session_state["criteria_targets_active"] or get_standard_ideal_criteria()
    orig_targets = st.session_state["criteria_targets_original"] or get_standard_ideal_criteria()

    st.markdown("#### ✏️ Edit Active Criteria (% of Revenue)")

    with st.form("targets_edit_form"):
        col_ed1, col_ed2 = st.columns(2)
        new_targets = {}

        categories = list(targets.keys())
        half = len(categories) // 2

        with col_ed1:
            for cat in categories[:half]:
                st.markdown(f"**{cat}**")
                c_min, c_max = st.columns(2)
                with c_min:
                    min_val = st.number_input(f"Min % ({cat})", min_value=0.0, max_value=100.0, value=float(targets[cat]["min"]), step=0.5, key=f"min_{cat}")
                with c_max:
                    max_val = st.number_input(f"Max % ({cat})", min_value=0.0, max_value=100.0, value=float(targets[cat]["max"]), step=0.5, key=f"max_{cat}")
                new_targets[cat] = {"min": min_val, "max": max_val}

        with col_ed2:
            for cat in categories[half:]:
                st.markdown(f"**{cat}**")
                c_min, c_max = st.columns(2)
                with c_min:
                    min_val = st.number_input(f"Min % ({cat})", min_value=0.0, max_value=100.0, value=float(targets[cat]["min"]), step=0.5, key=f"min_{cat}")
                with c_max:
                    max_val = st.number_input(f"Max % ({cat})", min_value=0.0, max_value=100.0, value=float(targets[cat]["max"]), step=0.5, key=f"max_{cat}")
                new_targets[cat] = {"min": min_val, "max": max_val}

        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            save_clicked = st.form_submit_button("💾 Save Target Changes", type="primary", use_container_width=True)
        with c_btn2:
            reset_clicked = st.form_submit_button("🔄 Reset to Original Excel Targets", use_container_width=True)

    if save_clicked:
        invalid_pairs = [cat for cat, val in new_targets.items() if val["min"] > val["max"]]
        if invalid_pairs:
            st.error(f"Validation Error: Minimum cannot exceed Maximum for: {', '.join(invalid_pairs)}")
        else:
            st.session_state["criteria_targets_active"] = new_targets
            st.success("✓ Active targets updated successfully! Recalculating pipeline...")
            if st.session_state["pnl_df_raw"] is not None:
                run_financial_pipeline()
            st.rerun()

    if reset_clicked:
        st.session_state["criteria_targets_active"] = {k: v.copy() for k, v in orig_targets.items()}
        st.success("✓ Targets reset to original Excel values! Recalculating pipeline...")
        if st.session_state["pnl_df_raw"] is not None:
            run_financial_pipeline()
        st.rerun()


# =============================================================================
# TAB 9: DOWNLOAD REPORTS
# =============================================================================
elif nav_selection == "9. Download Reports":
    st.markdown("### 📥 GENERATED REPORTS – DOWNLOAD YOUR FILES")
    st.caption("Production-grade, publication-ready financial deliverables. Real binary files generated dynamically from uploaded data.")

    if not st.session_state["processing_complete"] or not st.session_state["report_files"]:
        st.warning("⚠️ No reports generated yet. Please upload the Excel workbook and click **GENERATE OUTPUT**.")
    else:
        reports = st.session_state["report_files"]

        # All_Reports.zip
        if "All_Reports.zip" in reports:
            zip_bytes, zip_size, zip_mime = reports["All_Reports.zip"]
            st.markdown(
                f"""
                <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:10px; padding:20px; margin-bottom:24px;">
                    <h4 style="margin:0; color:#1E3A8A;">📦 File 6: All_Reports.zip (Consolidated Package)</h4>
                    <p style="margin:6px 0 12px 0; font-size:13px; color:#3B82F6;">Download all 5 Excel and PDF reports in a single compressed archive ({zip_size / 1024:.1f} KB).</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            st.download_button(
                label="📦 Download All_Reports.zip",
                data=zip_bytes,
                file_name="All_Reports.zip",
                mime=zip_mime,
                type="primary",
                use_container_width=True
            )

        st.markdown("---")
        st.markdown("#### 📄 Individual Report Downloads")

        file_metadata = [
            (
                "Complete_FC_PnL_Report.xlsx",
                "📊 File 1: Complete_FC_PnL_Report.xlsx",
                "Full 8-worksheet Excel workbook containing Daily Summary, Expense Breakdown, CM1/CM2 Analysis, Anomaly Log, AI Insights, Daily Digest, Trends, and Ideal Criteria with professional styles."
            ),
            (
                "Anomaly_Report.xlsx",
                "🚨 File 2: Anomaly_Report.xlsx",
                "Dedicated workbook listing all flagged cost variances, deviation percentage points, target corridors, and operational severity ratings."
            ),
            (
                "AI_Insights_Report.pdf",
                "🧠 File 3: AI_Insights_Report.pdf",
                "Executive PDF report with 6-part analytical explanations, rupee-impact rankings, and recommended management actions."
            ),
            (
                "Daily_FC_Digest.pdf",
                "📋 File 4: Daily_FC_Digest.pdf",
                "Colour-coded daily summary across fulfillment centers with operational status and key action summaries."
            ),
            (
                "Trend_Analysis.pdf",
                "📈 File 5: Trend_Analysis.pdf",
                "Multi-day profitability trajectory, expense category averages, and recurrence frequency analysis."
            ),
        ]

        for fname, label, desc in file_metadata:
            if fname in reports:
                fbytes, fsize, fmime = reports[fname]
                c_desc, c_btn = st.columns([3, 1])
                with c_desc:
                    st.markdown(f"**{label}** &nbsp;`({fsize / 1024:.1f} KB)`")
                    st.caption(desc)
                with c_btn:
                    st.download_button(
                        label=f"⬇️ Download {fname.split('.')[-1].upper()}",
                        data=fbytes,
                        file_name=fname,
                        mime=fmime,
                        use_container_width=True
                    )
                st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)


# =============================================================================
# TAB 10: SYSTEM ARCHITECTURE & ROADMAP
# =============================================================================
elif nav_selection == "10. System Architecture & Roadmap":
    st.markdown("### 🚀 System Architecture & Future Enhancements")
    st.markdown(
        """
        The **Prozo FC P&L Monitor** operates on a zero-database, in-memory architecture designed for instant validation, 
        mathematical determinism, and seamless single-workbook ingestion.
        """
    )

    st.markdown("#### 🔮 Future Evolution Roadmap (How the System Gets Smarter Over Time)")
    
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.markdown(
            """
            ##### 1. Persistent Cross-Upload Historical Memory
            - **Time-Series Warehouse:** Integrate a managed columnar datastore (ClickHouse / TimescaleDB) to accumulate daily uploads into longitudinal FC histories.
            - **Baseline Normalization:** Automatically calculate rolling 30-day and 90-day moving averages for each FC to detect seasonal shifts (e.g. festive volume surges vs monsoon slowdowns).
            - **Peer FC Benchmarking:** Compare Bhiwandi vs Farukhnagar vs Hoskote normalized cost-per-pick.

            ##### 2. Human-in-the-Loop Manager Feedback
            - **Anomaly Override & Annotation:** Allow FC Managers to provide contextual annotations directly on flagged anomalies (e.g. "Overtime approved due to Client X flash sale").
            - **False-Positive Tuning:** When a manager validates a legitimate variance, the active target engine recalibrates tolerance bands automatically.
            """
        )
    with col_r2:
        st.markdown(
            """
            ##### 3. Predictive Forecasting & Proactive Alerts
            - **Intra-Day Drift Warning:** Predict end-of-day CM2 at 2:00 PM based on real-time WMS pick hours and order volumes before shifts close.
            - **Automated Dispatch Integration:** Connect webhooks to Prozo Slack channels, WhatsApp Business alerts for GMs, and automated executive email digests.
            - **Vendor Contract Auditing:** Cross-check equipment rental invoices directly against telemetry hours logged by warehouse MHE trackers.

            ##### 4. Continuous Model Distillation
            - Use fine-tuned domain LLMs to ingest regional tariff updates, labor laws, and freight regulations for deeper root-cause synthesis.
            """
        )

    st.info("📌 **Current Implementation Note:** In compliance with current technical parameters, all analyses are performed in-memory directly on uploaded Excel records without persistent database dependencies.")
