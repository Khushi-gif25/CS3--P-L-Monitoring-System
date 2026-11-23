"""
PROZO P&L MONITOR - REPORT GENERATION ENGINE
Generates professional multi-tab Excel files, styled executive PDF reports,
and consolidated ZIP archives containing all deliverables.
"""

import io
import zipfile
from typing import Dict, Any, List, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY


# Prozo Corporate Color Palette
NAVY_HEADER = "1E3A8A"
NAVY_ACCENT = "3B82F6"
BG_LIGHT = "F8FAFC"
BORDER_GRAY = "CBD5E1"

STATUS_COLORS_HEX = {
    "GREEN": ("DCFCE7", "166534"),
    "RED": ("FEE2E2", "991B1B"),
    "YELLOW": ("FEF3C7", "92400E"),
    "SUSPICIOUS": ("F3E8FF", "6B21A8"),
    "REVIEW REQUIRED": ("F3F4F6", "374151")
}


def apply_excel_styling(ws, title: str):
    """Applies modern corporate styling, header formats, auto column width to an openpyxl worksheet."""
    # Freeze top row
    ws.freeze_panes = "A2"
    
    # Styles
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    regular_font = Font(name="Calibri", size=10)
    border_thin = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    # Style Header Row
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

    # Style Data Rows
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 20
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = regular_font
            cell.border = border_thin
            header_name = str(ws.cell(row=1, column=col_idx).value or "")

            # Alignments and number formats
            val = cell.value
            if isinstance(val, (int, float)):
                if "%" in header_name or "pp" in header_name or "Pct" in header_name:
                    cell.number_format = '0.0"%"'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif any(curr_kw in header_name for curr_kw in ["Revenue", "Cost", "Rent", "Overhead", "CM1", "CM2", "Impact", "Amount", "Total"]):
                    cell.number_format = '₹#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                # String status color badges
                str_val = str(val).strip().upper()
                if str_val in STATUS_COLORS_HEX and ("STATUS" in header_name.upper() or "FC" in header_name.upper()):
                    bg_col, txt_col = STATUS_COLORS_HEX[str_val]
                    cell.fill = PatternFill(start_color=bg_col, end_color=bg_col, fill_type="solid")
                    cell.font = Font(name="Calibri", size=10, bold=True, color=txt_col)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

    # Auto adjust column widths with padding
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            v_str = str(cell.value or "")
            if len(v_str) > max_len:
                max_len = len(v_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


def generate_complete_pnl_excel(
    df_metrics: pd.DataFrame,
    df_anomalies: pd.DataFrame,
    insights_list: List[Dict[str, Any]],
    trend_data: Dict[str, Any],
    targets: Dict[str, Dict[str, float]]
) -> bytes:
    """
    Generates File 1: Complete_FC_PnL_Report.xlsx with 8 distinct worksheets:
    1. Daily P&L Summary
    2. Expense Breakdown
    3. CM1 and CM2 Analysis
    4. Anomaly Log
    5. AI Insights
    6. Daily FC Digest
    7. Trend Analysis
    8. Ideal FC Criteria
    """
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # 1. Daily P&L Summary
    ws1 = wb.create_sheet(title="Daily P&L Summary")
    summary_cols = [
        "Date", "FC", "Revenue", "Total Opex", "CM1", "CM1 %", "CM2", "CM2 %", "FC Status", "Status Warning"
    ]
    sub_df1 = df_metrics[[c for c in summary_cols if c in df_metrics.columns]]
    ws1.append(list(sub_df1.columns))
    for r in sub_df1.itertuples(index=False):
        ws1.append(list(r))
    apply_excel_styling(ws1, "Daily P&L Summary")

    # 2. Expense Breakdown
    ws2 = wb.create_sheet(title="Expense Breakdown")
    exp_cols = [
        "Date", "FC", "Revenue",
        "Manpower Cost", "Manpower Cost %",
        "Packaging Cost", "Packaging Cost %",
        "Power and Fuel", "Power and Fuel %",
        "FC Rent", "FC Rent %",
        "Equipment Rentals", "Equipment Rentals %",
        "Overheads", "Overheads %",
        "Total Opex", "Total Opex %"
    ]
    sub_df2 = df_metrics[[c for c in exp_cols if c in df_metrics.columns]]
    ws2.append(list(sub_df2.columns))
    for r in sub_df2.itertuples(index=False):
        ws2.append(list(r))
    apply_excel_styling(ws2, "Expense Breakdown")

    # 3. CM1 and CM2 Analysis
    ws3 = wb.create_sheet(title="CM1 and CM2 Analysis")
    cm_cols = [
        "Date", "FC", "Revenue",
        "Manpower Cost", "Packaging Cost",
        "CM1", "CM1 %",
        "Power and Fuel", "FC Rent", "Equipment Rentals", "Overheads",
        "CM2", "CM2 %",
        "FC Status"
    ]
    sub_df3 = df_metrics[[c for c in cm_cols if c in df_metrics.columns]]
    ws3.append(list(sub_df3.columns))
    for r in sub_df3.itertuples(index=False):
        ws3.append(list(r))
    apply_excel_styling(ws3, "CM1 and CM2 Analysis")

    # 4. Anomaly Log
    ws4 = wb.create_sheet(title="Anomaly Log")
    if not df_anomalies.empty:
        ws4.append(list(df_anomalies.columns))
        for r in df_anomalies.itertuples(index=False):
            ws4.append(list(r))
    else:
        ws4.append(["Status", "Message"])
        ws4.append(["NORMAL", "No financial anomalies flagged across operational records."])
    apply_excel_styling(ws4, "Anomaly Log")

    # 5. AI Insights
    ws5 = wb.create_sheet(title="AI Insights")
    ws5.append(["Date", "FC", "Revenue", "CM2 %", "FC Status", "What Happened", "Deviation Size", "Breached Line Items", "Top Expense Contributors", "Confirmed Drivers", "Recommended Actions"])
    for ins in insights_list:
        ranked_str = "; ".join([c["summary"] for c in ins.get("ranked_contributors", [])])
        confirmed_str = "; ".join(ins.get("confirmed_causes", []))
        rec_str = "; ".join(ins.get("recommended_actions", []))
        ws5.append([
            ins["date"], ins["fc"], ins["revenue"], ins["cm2_pct"], ins["fc_status"],
            ins["what_happened"], ins["deviation_size"], ins["breached_items"],
            ranked_str, confirmed_str, rec_str
        ])
    apply_excel_styling(ws5, "AI Insights")

    # 6. Daily FC Digest
    ws6 = wb.create_sheet(title="Daily FC Digest")
    ws6.append(["Date", "FC", "Revenue", "CM1 %", "CM2 %", "Status", "Anomaly Count", "Breached Categories", "Executive Summary", "Key Action"])
    for idx, row in df_metrics.iterrows():
        d_val = str(row["Date"])
        fc_val = str(row["FC"])
        anom_subset = df_anomalies[(df_anomalies["Date"] == d_val) & (df_anomalies["FC"] == fc_val)] if not df_anomalies.empty else pd.DataFrame()
        anom_count = len(anom_subset)
        breached_cats = ", ".join(anom_subset["Line Item"].unique()) if not anom_subset.empty else "None"
        
        # Match insight
        matched_ins = next((i for i in insights_list if i["date"] == d_val and i["fc"] == fc_val), None)
        exec_sum = matched_ins["what_happened"] if matched_ins else f"Operations within targets. CM2: {row['CM2 %']:.1f}%."
        act_rec = "; ".join(matched_ins["recommended_actions"]) if matched_ins else "Continue standard procedures."

        ws6.append([
            d_val, fc_val, row["Revenue"], row["CM1 %"], row["CM2 %"], row["FC Status"],
            anom_count, breached_cats, exec_sum, act_rec
        ])
    apply_excel_styling(ws6, "Daily FC Digest")

    # 7. Trend Analysis
    ws7 = wb.create_sheet(title="Trend Analysis")
    ws7.append(["Metric / Category", "Portfolio Average %", "Target Min %", "Target Max %", "Status Assessment"])
    for cat in ["Manpower", "Packaging", "Power and Fuel", "FC Rent", "Equipment Rentals", "Overheads"]:
        col_name = f"{cat} Cost %" if f"{cat} Cost %" in df_metrics.columns else f"{cat} %"
        if col_name in df_metrics.columns:
            avg_pct = round(df_metrics[col_name].mean(), 2)
            t_min = targets.get(cat, {}).get("min", 0.0)
            t_max = targets.get(cat, {}).get("max", 0.0)
            status_txt = "NORMAL"
            if avg_pct > t_max:
                status_txt = "ABOVE TARGET"
            elif avg_pct < t_min:
                status_txt = "BELOW TARGET"
            ws7.append([cat, avg_pct, t_min, t_max, status_txt])

    # Add CM1 and CM2 trends
    ws7.append(["Contribution Margin 1 (CM1)", round(df_metrics["CM1 %"].mean(), 2), targets.get("CM1", {}).get("min", 65.0), targets.get("CM1", {}).get("max", 70.0), "Portfolio CM1 Average"])
    ws7.append(["Contribution Margin 2 (CM2)", round(df_metrics["CM2 %"].mean(), 2), targets.get("CM2", {}).get("min", 15.0), targets.get("CM2", {}).get("max", 30.0), "Portfolio CM2 Average"])
    
    # Append Observations
    ws7.append([])
    ws7.append(["Trend Observations & Strategic Findings", "", "", "", ""])
    for obs in trend_data.get("observations", []):
        ws7.append([obs, "", "", "", ""])

    apply_excel_styling(ws7, "Trend Analysis")

    # 8. Ideal FC Criteria
    ws8 = wb.create_sheet(title="Ideal FC Criteria")
    ws8.append(["Expense Category / Metric", "Minimum Target (%)", "Maximum Target (%)", "Target Type"])
    for cat, val in targets.items():
        cat_type = "Profitability Margin" if "CM" in cat else "Operating Expense Ceiling"
        ws8.append([cat, val["min"], val["max"], cat_type])
    apply_excel_styling(ws8, "Ideal FC Criteria")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def generate_anomaly_excel(df_anomalies: pd.DataFrame) -> bytes:
    """Generates File 2: Anomaly_Report.xlsx."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Anomaly Report"

    if not df_anomalies.empty:
        ws.append(list(df_anomalies.columns))
        for r in df_anomalies.itertuples(index=False):
            ws.append(list(r))
    else:
        ws.append(["Status", "Message"])
        ws.append(["NORMAL", "No financial anomalies identified in current dataset."])

    apply_excel_styling(ws, "Anomaly Report")
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def get_pdf_styles():
    """Builds clean, modern executive styles for PDF generation."""
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=colors.HexColor('#1E3A8A'),
        spaceAfter=6
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#64748B'),
        spaceAfter=14
    )
    h1_style = ParagraphStyle(
        'Heading1_Custom',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=17,
        textColor=colors.HexColor('#1E3A8A'),
        spaceBefore=12,
        spaceAfter=6
    )
    body_style = ParagraphStyle(
        'Body_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor('#1E293B')
    )
    bullet_style = ParagraphStyle(
        'Bullet_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        leftIndent=12,
        spaceAfter=4,
        textColor=colors.HexColor('#334155')
    )
    badge_green = ParagraphStyle('BG_Green', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.HexColor('#166534'))
    badge_red = ParagraphStyle('BG_Red', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.HexColor('#991B1B'))
    badge_yellow = ParagraphStyle('BG_Yellow', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.HexColor('#92400E'))
    badge_suspicious = ParagraphStyle('BG_Suspicious', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.HexColor('#6B21A8'))

    return {
        "title": title_style,
        "subtitle": subtitle_style,
        "h1": h1_style,
        "body": body_style,
        "bullet": bullet_style,
        "green": badge_green,
        "red": badge_red,
        "yellow": badge_yellow,
        "suspicious": badge_suspicious,
    }


def generate_insights_pdf(
    insights_list: List[Dict[str, Any]],
    trend_data: Dict[str, Any],
    df_anomalies: pd.DataFrame
) -> bytes:
    """Generates File 3: AI_Insights_Report.pdf."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    styles = get_pdf_styles()
    story = []

    # Title & Header
    story.append(Paragraph("PROZO FULFILLMENT NETWORK", styles["subtitle"]))
    story.append(Paragraph("AI Executive Insights & Anomaly Analysis Report", styles["title"]))
    story.append(Paragraph("Automated Financial Governance • Root Cause Identification • Action Directives", styles["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=14))

    # Executive Overview
    story.append(Paragraph("1. Executive Portfolio Summary", styles["h1"]))
    obs_list = trend_data.get("observations", [])
    for obs in obs_list:
        story.append(Paragraph(f"• {obs}", styles["bullet"]))
    story.append(Spacer(1, 10))

    # Detailed FC-Level Insights
    story.append(Paragraph("2. Deep-Dive Anomaly Investigations & Recommended Actions", styles["h1"]))

    for ins in insights_list:
        # Build insight card
        status_color = colors.HexColor("#16A34A") if ins["fc_status"] == "GREEN" else (
            colors.HexColor("#DC2626") if ins["fc_status"] == "RED" else (
                colors.HexColor("#D97706") if ins["fc_status"] == "YELLOW" else colors.HexColor("#9333EA")
            )
        )
        
        card_content = [
            [
                Paragraph(f"<b>FC:</b> {ins['fc']} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Date:</b> {ins['date']} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Revenue:</b> ₹{ins['revenue']:,.2f}", styles["body"]),
                Paragraph(f"<b>STATUS: {ins['fc_status']}</b> (CM2: {ins['cm2_pct']:.1f}%)", styles["body"])
            ],
            [
                Paragraph(f"<b>What Happened:</b> {ins['what_happened']}", styles["body"]),
                ""
            ],
            [
                Paragraph(f"<b>Line Item Breaches:</b> {ins['breached_items']}", styles["body"]),
                ""
            ],
            [
                Paragraph(f"<b>Top Contributors by Rupee Impact:</b> " + ("; ".join([c['summary'] for c in ins.get('ranked_contributors', [])]) or "None"), styles["body"]),
                ""
            ],
            [
                Paragraph(f"<b>Root Cause Analysis:</b> " + ("; ".join(ins.get('possible_causes', [])) or "Routine variance"), styles["body"]),
                ""
            ],
            [
                Paragraph(f"<b>Recommended Management Action:</b> " + ("; ".join(ins.get('recommended_actions', [])) or "Monitor normal ops"), styles["body"]),
                ""
            ]
        ]
        
        card_table = Table(card_content, colWidths=[400, 140])
        card_table.setStyle(TableStyle([
            ('SPAN', (0, 1), (1, 1)),
            ('SPAN', (0, 2), (1, 2)),
            ('SPAN', (0, 3), (1, 3)),
            ('SPAN', (0, 4), (1, 4)),
            ('SPAN', (0, 5), (1, 5)),
            ('BACKGROUND', (0, 0), (1, 0), colors.HexColor("#F1F5F9")),
            ('LINELEFT', (0, 0), (0, 5), 3, status_color),
            ('BOX', (0, 0), (1, 5), 0.5, colors.HexColor("#CBD5E1")),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
        ]))
        
        story.append(card_table)
        story.append(Spacer(1, 10))

    doc.build(story)
    return buffer.getvalue()


def generate_daily_digest_pdf(
    df_metrics: pd.DataFrame,
    df_anomalies: pd.DataFrame,
    insights_list: List[Dict[str, Any]]
) -> bytes:
    """Generates File 4: Daily_FC_Digest.pdf."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=30,
        rightMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    styles = get_pdf_styles()
    story = []

    # Title & Header
    story.append(Paragraph("PROZO FULFILLMENT NETWORK", styles["subtitle"]))
    story.append(Paragraph("Daily Fulfillment Centre (FC) Financial Digest", styles["title"]))
    story.append(Paragraph("Daily Operational Performance • Profitability Status • Key Breaches & Directives", styles["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=12))

    # Table of Daily Digest
    table_data = [
        [
            Paragraph("<b>Date</b>", styles["body"]),
            Paragraph("<b>FC Name</b>", styles["body"]),
            Paragraph("<b>Revenue (₹)</b>", styles["body"]),
            Paragraph("<b>CM1 %</b>", styles["body"]),
            Paragraph("<b>CM2 %</b>", styles["body"]),
            Paragraph("<b>Status</b>", styles["body"]),
            Paragraph("<b>Anomalies</b>", styles["body"]),
            Paragraph("<b>AI Summary & Action</b>", styles["body"]),
        ]
    ]

    for idx, row in df_metrics.iterrows():
        d_val = str(row["Date"])
        fc_val = str(row["FC"])
        status = str(row["FC Status"])
        
        anom_subset = df_anomalies[(df_anomalies["Date"] == d_val) & (df_anomalies["FC"] == fc_val)] if not df_anomalies.empty else pd.DataFrame()
        anom_count = len(anom_subset)

        matched_ins = next((i for i in insights_list if i["date"] == d_val and i["fc"] == fc_val), None)
        summary_txt = matched_ins["what_happened"] if matched_ins else f"Healthy operations. CM2 at {row['CM2 %']:.1f}%."
        action_txt = matched_ins["recommended_actions"][0] if (matched_ins and matched_ins["recommended_actions"]) else "Standard protocol."

        combined_text = f"{summary_txt}<br/><font color='#1E3A8A'><b>Action:</b></font> {action_txt}"

        status_style = styles["green"] if status == "GREEN" else (
            styles["red"] if status == "RED" else (styles["yellow"] if status == "YELLOW" else styles["suspicious"])
        )

        table_data.append([
            Paragraph(d_val, styles["body"]),
            Paragraph(fc_val, styles["body"]),
            Paragraph(f"₹{row['Revenue']:,.0f}", styles["body"]),
            Paragraph(f"{row['CM1 %']:.1f}%", styles["body"]),
            Paragraph(f"{row['CM2 %']:.1f}%", styles["body"]),
            Paragraph(f"<b>{status}</b>", status_style),
            Paragraph(str(anom_count), styles["body"]),
            Paragraph(combined_text, styles["body"]),
        ])

    digest_table = Table(
        table_data,
        colWidths=[65, 80, 75, 50, 50, 65, 45, 300]
    )
    digest_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F8FAFC")]),
    ]))

    story.append(digest_table)
    doc.build(story)
    return buffer.getvalue()


def generate_trend_pdf(
    df_metrics: pd.DataFrame,
    df_anomalies: pd.DataFrame,
    trend_data: Dict[str, Any],
    targets: Dict[str, Dict[str, float]]
) -> bytes:
    """Generates File 5: Trend_Analysis.pdf."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    styles = get_pdf_styles()
    story = []

    # Title & Header
    story.append(Paragraph("PROZO FULFILLMENT NETWORK", styles["subtitle"]))
    story.append(Paragraph("Multi-Day Financial Trend & Margin Stability Analysis", styles["title"]))
    story.append(Paragraph("Historical P&L Trajectories • Variance Recurrence • Efficiency Benchmarks", styles["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=14))

    # Trend Observations
    story.append(Paragraph("1. Strategic AI Trend Synthesis", styles["h1"]))
    for obs in trend_data.get("observations", []):
        story.append(Paragraph(f"• {obs}", styles["bullet"]))
    story.append(Spacer(1, 10))

    # Category Averages vs Targets Table
    story.append(Paragraph("2. Portfolio Cost Structure vs Ideal Targets", styles["h1"]))
    cat_rows = [
        [
            Paragraph("<b>Expense Category</b>", styles["body"]),
            Paragraph("<b>Target Corridor</b>", styles["body"]),
            Paragraph("<b>Portfolio Average %</b>", styles["body"]),
            Paragraph("<b>Variance Status</b>", styles["body"])
        ]
    ]

    for cat in ["Manpower", "Packaging", "Power and Fuel", "FC Rent", "Equipment Rentals", "Overheads"]:
        col_name = f"{cat} Cost %" if f"{cat} Cost %" in df_metrics.columns else f"{cat} %"
        if col_name in df_metrics.columns:
            avg_pct = round(df_metrics[col_name].mean(), 2)
            t_min = targets.get(cat, {}).get("min", 0.0)
            t_max = targets.get(cat, {}).get("max", 0.0)
            status_txt = "NORMAL"
            style_tag = styles["green"]
            if avg_pct > t_max:
                status_txt = f"ABOVE TARGET (+{round(avg_pct - t_max, 1)} pp)"
                style_tag = styles["red"]
            elif avg_pct < t_min:
                status_txt = f"BELOW TARGET ({round(avg_pct - t_min, 1)} pp)"
                style_tag = styles["yellow"]

            cat_rows.append([
                Paragraph(cat, styles["body"]),
                Paragraph(f"{t_min:.1f}% - {t_max:.1f}%", styles["body"]),
                Paragraph(f"{avg_pct:.1f}%", styles["body"]),
                Paragraph(f"<b>{status_txt}</b>", style_tag)
            ])

    # Append CM1 & CM2
    cat_rows.append([
        Paragraph("<b>Contribution Margin 1 (CM1)</b>", styles["body"]),
        Paragraph(f"{targets.get('CM1', {}).get('min', 65.0):.1f}% - {targets.get('CM1', {}).get('max', 70.0):.1f}%", styles["body"]),
        Paragraph(f"{round(df_metrics['CM1 %'].mean(), 1):.1f}%", styles["body"]),
        Paragraph("Direct Margin", styles["body"])
    ])
    cat_rows.append([
        Paragraph("<b>Contribution Margin 2 (CM2)</b>", styles["body"]),
        Paragraph(f"{targets.get('CM2', {}).get('min', 15.0):.1f}% - {targets.get('CM2', {}).get('max', 30.0):.1f}%", styles["body"]),
        Paragraph(f"{round(df_metrics['CM2 %'].mean(), 1):.1f}%", styles["body"]),
        Paragraph("Operating Margin", styles["body"])
    ])

    cat_table = Table(cat_rows, colWidths=[150, 110, 110, 170])
    cat_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#F1F5F9")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F8FAFC")]),
    ]))
    story.append(cat_table)
    story.append(Spacer(1, 14))

    # Anomaly Frequency Table
    story.append(Paragraph("3. Anomaly Frequency & Recurrence Analysis", styles["h1"]))
    if not df_anomalies.empty:
        anom_counts = df_anomalies["Line Item"].value_counts().reset_index()
        anom_counts.columns = ["Line Item", "Breach Count"]
        
        freq_rows = [
            [
                Paragraph("<b>Line Item / Metric</b>", styles["body"]),
                Paragraph("<b>Total Breaches Count</b>", styles["body"]),
                Paragraph("<b>Recurrence Impact Level</b>", styles["body"])
            ]
        ]
        for _, r in anom_counts.iterrows():
            cnt = int(r["Breach Count"])
            risk_label = "Systemic / Recurring" if cnt >= 3 else ("Moderate" if cnt == 2 else "Isolated Event")
            risk_style = styles["red"] if cnt >= 3 else styles["yellow"]
            freq_rows.append([
                Paragraph(str(r["Line Item"]), styles["body"]),
                Paragraph(str(cnt), styles["body"]),
                Paragraph(f"<b>{risk_label}</b>", risk_style)
            ])
        freq_table = Table(freq_rows, colWidths=[220, 130, 190])
        freq_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#F1F5F9")),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(freq_table)
    else:
        story.append(Paragraph("No anomalies detected in the dataset.", styles["body"]))

    doc.build(story)
    return buffer.getvalue()


def generate_zip_bundle(
    file1_bytes: bytes,
    file2_bytes: bytes,
    file3_bytes: bytes,
    file4_bytes: bytes,
    file5_bytes: bytes
) -> bytes:
    """Consolidates all 5 generated report files into a single ZIP archive."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("Complete_FC_PnL_Report.xlsx", file1_bytes)
        zip_file.writestr("Anomaly_Report.xlsx", file2_bytes)
        zip_file.writestr("AI_Insights_Report.pdf", file3_bytes)
        zip_file.writestr("Daily_FC_Digest.pdf", file4_bytes)
        zip_file.writestr("Trend_Analysis.pdf", file5_bytes)
    return zip_buffer.getvalue()
