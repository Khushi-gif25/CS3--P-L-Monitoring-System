"""
Report Generation Module
Prozo Daily FC P&L Monitoring System

Generates a publication-grade, 5-sheet formatted Excel workbook (.xlsx):
- Sheet 1: Daily P&L Summary
- Sheet 2: Anomaly Log
- Sheet 3: AI Insights
- Sheet 4: Trend Analysis
- Sheet 5: FC Criteria

Formatting & Styling:
- Prozo Indigo/Purple brand headers with bold white text
- Status badge fills (Red, Yellow, Green, Suspicious, Review Required)
- Currency formatting for amounts (₹#,##0)
- Percentage formatting for rates (0.00%)
- Auto-fitted column widths with clean padding
- Aligned cells with subtle grid borders
"""

import io
from typing import Any, Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .data_extractor import CM_TARGETS, STANDARD_EXPENSE_KEYS
from .ai_insights import generate_insights_for_day, generate_rule_based_anomaly_insight

# Brand Theme Colors
COLOR_HEADER_BG = "4F46E5"        # Indigo/Purple
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_ZEBRA_BG = "F8FAFC"          # Light Slate
COLOR_BORDER = "E2E8F0"

# Status Colors
STATUS_COLORS = {
    "GREEN": {"bg": "DCFCE7", "text": "166534"},
    "YELLOW": {"bg": "FEF3C7", "text": "92400E"},
    "RED": {"bg": "FEE2E2", "text": "991B1B"},
    "SUSPICIOUS": {"bg": "EDE9FE", "text": "5B21B6"},
    "REVIEW REQUIRED": {"bg": "FFEDD5", "text": "9A3412"},
    "ABOVE TARGET": {"bg": "FEE2E2", "text": "991B1B"},
    "BELOW TARGET": {"bg": "FEF3C7", "text": "92400E"},
    "NORMAL": {"bg": "DCFCE7", "text": "166534"},
}

def create_thin_border() -> Border:
    thin = Side(border_style="thin", color=COLOR_BORDER)
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def auto_fit_columns(ws: openpyxl.worksheet.worksheet.Worksheet, max_col_width: int = 45):
    """Dynamically set column widths based on cell content."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            # If cell has newline, take length of longest line
            lines = val_str.split("\n")
            line_len = max(len(l) for l in lines) if lines else 0
            if line_len > max_len:
                max_len = line_len
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), max_col_width)

def generate_excel_report(
    df_calc: openpyxl.Workbook,
    master_anomalies: List[Dict[str, Any]],
    trend_data: Dict[str, Any],
    criteria: Dict[str, Tuple[float, float]],
    aggregates: Dict[str, Any],
    cm_targets: Optional[Dict[str, Tuple[float, float]]] = None
) -> io.BytesIO:
    """
    Build 5-sheet styled Excel workbook in-memory and return BytesIO buffer.
    """
    wb = openpyxl.Workbook()
    thin_border = create_thin_border()
    header_font = Font(name="Segoe UI", size=11, bold=True, color=COLOR_HEADER_TEXT)
    header_fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    data_font = Font(name="Segoe UI", size=10)
    bold_data_font = Font(name="Segoe UI", size=10, bold=True)

    # =========================================================================
    # SHEET 1: Daily P&L Summary
    # =========================================================================
    ws_pnl = wb.active
    ws_pnl.title = "Daily P&L Summary"

    headers_pnl = [
        "Day", "FC Identifier", "Revenue (₹)",
        "Manpower (₹)", "Packaging (₹)", "Power and Fuel (₹)", "FC Rent (₹)", "Equipment Rentals (₹)", "Overheads (₹)",
        "Manpower %", "Packaging %", "Power and Fuel %", "FC Rent %", "Equipment Rentals %", "Overheads %",
        "CM1 (₹)", "CM1 %",
        "CM2 (₹)", "CM2 %",
        "Overall Status"
    ]
    ws_pnl.append(headers_pnl)

    for row_idx, r in enumerate(df_calc.to_dict(orient="records"), start=2):
        row_vals = [
            r["Day"],
            r.get("FC Identifier") or "N/A",
            r["Revenue"],
            r["Manpower"],
            r["Packaging"],
            r["Power and Fuel"],
            r["FC Rent"],
            r["Equipment Rentals"],
            r["Overheads"],
            r["Manpower %"] / 100.0,
            r["Packaging %"] / 100.0,
            r["Power and Fuel %"] / 100.0,
            r["FC Rent %"] / 100.0,
            r["Equipment Rentals %"] / 100.0,
            r["Overheads %"] / 100.0,
            r["CM1"],
            r["CM1%"] / 100.0,
            r["CM2"],
            r["CM2%"] / 100.0,
            r["Overall Status"]
        ]
        ws_pnl.append(row_vals)

        # Style data cells
        for col_idx in range(1, len(row_vals) + 1):
            cell = ws_pnl.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            
            # Formatting
            if col_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="center")
            elif 3 <= col_idx <= 9 or col_idx in [16, 18]:
                cell.number_format = '"₹"#,##0'
                cell.alignment = Alignment(horizontal="right")
            elif 10 <= col_idx <= 15 or col_idx in [17, 19]:
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 20: # Status
                st = str(cell.value)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if st in STATUS_COLORS:
                    c_info = STATUS_COLORS[st]
                    cell.fill = PatternFill(start_color=c_info["bg"], end_color=c_info["bg"], fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color=c_info["text"])

    # Style header row
    for cell in ws_pnl[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    auto_fit_columns(ws_pnl)

    # =========================================================================
    # SHEET 2: Anomaly Log
    # =========================================================================
    ws_anom = wb.create_sheet(title="Anomaly Log")
    headers_anom = [
        "Day", "FC Identifier", "Line Item / Metric", "Anomaly Type",
        "Actual % of Revenue", "Target Range", "Deviation (pp)", "Status", "Explanation"
    ]
    ws_anom.append(headers_anom)

    for row_idx, a in enumerate(master_anomalies, start=2):
        row_vals = [
            a["Day"],
            a.get("FC Identifier") or "N/A",
            a["Line Item"],
            a.get("Type", "Anomaly"),
            a["Actual %"] / 100.0,
            a["Target Range"],
            f"{a['Deviation (pp)']:.2f} pp",
            a["Status"],
            a.get("Message", "")
        ]
        ws_anom.append(row_vals)

        for col_idx in range(1, len(row_vals) + 1):
            cell = ws_anom.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [1, 2, 4, 6, 7]:
                cell.alignment = Alignment(horizontal="center")
            elif col_idx == 5:
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 8: # Status
                st = str(cell.value)
                cell.alignment = Alignment(horizontal="center")
                if st in STATUS_COLORS:
                    c_info = STATUS_COLORS[st]
                    cell.fill = PatternFill(start_color=c_info["bg"], end_color=c_info["bg"], fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color=c_info["text"])
            elif col_idx == 9:
                cell.alignment = Alignment(horizontal="left", wrap_text=True)

    for cell in ws_anom[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    auto_fit_columns(ws_anom, max_col_width=50)

    # =========================================================================
    # SHEET 3: AI Insights
    # =========================================================================
    ws_insights = wb.create_sheet(title="AI Insights")
    headers_insights = [
        "Day", "Affected Metric", "What Happened", "Actual Value",
        "Target Range", "Deviation", "Possible Contributing Factors", "Recommended Action"
    ]
    ws_insights.append(headers_insights)

    # Generate insights from anomalies and suspicious days
    insight_row_idx = 2
    for a in master_anomalies:
        ins = generate_rule_based_anomaly_insight(a, {})
        row_vals = [
            ins["Day"],
            ins["Affected Metric"],
            ins["What Happened"],
            ins["Actual Value"],
            ins["Target Range"],
            ins["Deviation"],
            ins["Possible Contributing Factors"],
            ins["Recommended Action"]
        ]
        ws_insights.append(row_vals)
        for col_idx in range(1, len(row_vals) + 1):
            cell = ws_insights.cell(row=insight_row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [1, 2, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="top")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        insight_row_idx += 1

    for cell in ws_insights[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    auto_fit_columns(ws_insights, max_col_width=55)

    # =========================================================================
    # SHEET 4: Trend Analysis
    # =========================================================================
    ws_trends = wb.create_sheet(title="Trend Analysis")
    
    # Section A: Expense Trend Summary Table
    ws_trends.append(["Expense Category Trend Summary"])
    ws_trends.cell(row=1, column=1).font = Font(name="Segoe UI", size=12, bold=True, color="1E1B4B")
    
    headers_t1 = [
        "Expense Category", "Target Range", "Average %", "Min %", "Max %",
        "Volatility (Std)", "Total Breaches", "Breach Rate %", "Above Target Count", "Below Target Count"
    ]
    ws_trends.append(headers_t1)
    for cell in ws_trends[2]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    exp_sums = trend_data.get("Expense Summaries", {})
    t_row = 3
    for cat, s in exp_sums.items():
        ws_trends.append([
            cat,
            s["Target Range"],
            s["Average %"] / 100.0,
            s["Min %"] / 100.0,
            s["Max %"] / 100.0,
            s["Volatility (Std)"],
            s["Total Breaches"],
            s["Breach Rate %"] / 100.0,
            s["Above Target Count"],
            s["Below Target Count"]
        ])
        for c in range(1, 11):
            cell = ws_trends.cell(row=t_row, column=c)
            cell.font = data_font
            cell.border = thin_border
            if c in [3, 4, 5, 8]:
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal="right")
            elif c in [1, 2]:
                cell.alignment = Alignment(horizontal="left")
            else:
                cell.alignment = Alignment(horizontal="center")
        t_row += 1

    # Section B: CM1 & CM2 Trajectory Observations
    t_row += 2
    ws_trends.cell(row=t_row, column=1, value="Profitability Trajectory Metrics").font = Font(name="Segoe UI", size=12, bold=True, color="1E1B4B")
    t_row += 1
    
    cm_metrics = [
        ("CM1 Average %", f"{trend_data.get('CM1 Average %', 0):.2f}%"),
        ("CM1 Deviations Count (> 3 pp breach)", str(trend_data.get("CM1 Deviations Count", 0))),
        ("CM2 Average %", f"{trend_data.get('CM2 Average %', 0):.2f}%"),
        ("CM2 Deviations Count (> 3 pp breach)", str(trend_data.get("CM2 Deviations Count", 0))),
        ("Low CM2 Days (< 12% - RED)", str(trend_data.get("Low CM2 Days Count (< 12%)", 0))),
        ("Suspicious CM2 Days (> 33% - SUSPICIOUS)", str(trend_data.get("Suspicious CM2 Days Count (> 33%)", 0))),
        ("Most Problematic Expense Line", f"{trend_data.get('Most Problematic Category', 'N/A')} ({trend_data.get('Most Problematic Breaches', 0)} breaches)"),
    ]
    for lbl, val in cm_metrics:
        ws_trends.cell(row=t_row, column=1, value=lbl).font = bold_data_font
        ws_trends.cell(row=t_row, column=2, value=val).font = data_font
        ws_trends.cell(row=t_row, column=1).border = thin_border
        ws_trends.cell(row=t_row, column=2).border = thin_border
        t_row += 1

    # Section C: AI-Generated Trend Insights
    t_row += 2
    ws_trends.cell(row=t_row, column=1, value="AI-Powered Trend Insights & Recommendations").font = Font(name="Segoe UI", size=12, bold=True, color="1E1B4B")
    t_row += 1
    headers_t3 = ["Insight Title", "Period / Scope", "Observation", "Root Cause Indicator", "Recommended Action"]
    ws_trends.append(headers_t3)
    for cell in ws_trends[t_row]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    t_row += 1

    for ti in trend_data.get("Trend Insights", []):
        ws_trends.append([
            ti.get("Title", ""),
            ti.get("Period/Scope", ""),
            ti.get("Observation", ""),
            ti.get("Root Cause Indicator", ""),
            ti.get("Recommended Action", "")
        ])
        for c in range(1, 6):
            cell = ws_trends.cell(row=t_row, column=c)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        t_row += 1

    auto_fit_columns(ws_trends, max_col_width=50)

    # =========================================================================
    # SHEET 5: FC Criteria
    # =========================================================================
    ws_crit = wb.create_sheet(title="FC Criteria")
    headers_crit = ["Category / Metric", "Minimum Target (%)", "Maximum Target (%)", "Target Type", "Tolerance / Threshold"]
    ws_crit.append(headers_crit)

    crit_rows = []
    for cat in STANDARD_EXPENSE_KEYS:
        low, high = criteria.get(cat, (0.0, 0.0))
        crit_rows.append([cat, low / 100.0, high / 100.0, "Expense % of Revenue", "Exact Boundary"])
    
    # Add CM1 & CM2 criteria
    active_cm = cm_targets if cm_targets else CM_TARGETS
    cm1_bounds = active_cm.get("CM1", CM_TARGETS["CM1"])
    cm2_bounds = active_cm.get("CM2", CM_TARGETS["CM2"])
    crit_rows.append(["CM1", cm1_bounds[0] / 100.0, cm1_bounds[1] / 100.0, "Gross Contribution %", "Flagged if > 3.0 pp outside"])
    crit_rows.append(["CM2", cm2_bounds[0] / 100.0, cm2_bounds[1] / 100.0, "Net Contribution %", "Flagged if > 3.0 pp outside"])

    for r_idx, cr in enumerate(crit_rows, start=2):
        ws_crit.append(cr)
        for col_idx in range(1, len(cr) + 1):
            cell = ws_crit.cell(row=r_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [2, 3]:
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 1:
                cell.alignment = Alignment(horizontal="left")
                cell.font = bold_data_font
            else:
                cell.alignment = Alignment(horizontal="center")

    for cell in ws_crit[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    auto_fit_columns(ws_crit)

    # Save to buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
