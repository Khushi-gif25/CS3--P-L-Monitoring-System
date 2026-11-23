"""
Extraction Edge Cases & Robustness Test Suite
Prozo Daily FC P&L Monitoring System

Validates:
1. Dynamic header detection with leading blank rows and metadata banners
2. Column A blank offset (leading empty column)
3. Case-insensitive and whitespace-insensitive sheet matching (e.g. 'Day-Wise P&L', 'FC Criteria')
4. Semantic column aliases ('Power & Fuel', 'Rent', 'Labor', 'Total Revenue', 'Equipment Rental')
5. Newlines and formatting artifacts in column names
6. Single file-stream reuse without EOF exhaustion
7. Error diagnostics when required columns are genuinely missing
"""

import io
import os
import unittest
import openpyxl
import pandas as pd

from core.data_extractor import (
    DEFAULT_CRITERIA,
    extract_criteria,
    extract_pnl,
    extract_workbook_data,
    detect_header_row,
    match_column_name
)

class TestExtractionEdgeCases(unittest.TestCase):

    def setUp(self):
        self.base_dir = r"C:\Users\Lenovo\.gemini\antigravity\scratch\prozo_pnl_monitor"
        self.standard_xlsx = os.path.join(self.base_dir, "data", "Case Study 3.xlsx")

    def test_standard_workbook(self):
        """Test extraction on baseline Case Study 3.xlsx."""
        df_pnl, criteria, warnings, errors, diag = extract_workbook_data(self.standard_xlsx)
        self.assertIsNotNone(df_pnl)
        self.assertEqual(len(df_pnl), 100)
        self.assertEqual(len(errors), 0)
        self.assertEqual(criteria["Manpower"], (25.0, 27.0))
        self.assertEqual(diag["pnl"]["sheet_selected"], "Day Wise PnL")
        self.assertEqual(diag["pnl"]["header_row_index"], 0)

    def test_leading_blank_rows_and_metadata(self):
        """Test workbook with 3 blank rows and a title banner before the actual header."""
        wb = openpyxl.Workbook()
        
        # Criteria sheet with leading blank rows
        ws_c = wb.active
        ws_c.title = "Ideal FC Criteria"
        ws_c.append([])  # Blank row 1
        ws_c.append(["CONFIDENTIAL - PROZO BENCHMARKS"])  # Title row 2
        ws_c.append([])  # Blank row 3
        ws_c.append(["#", "Expense Line Item", "Range (as % of Revenue)"]) # Header row 4
        ws_c.append([1, "Manpower", "25 - 27%"])
        ws_c.append([2, "Packaging", "5 - 7%"])
        ws_c.append([3, "Power and Fuel", "10 - 12%"])
        ws_c.append([4, "FC Rent", "15 - 18%"])
        ws_c.append([5, "Equipment Rentals", "5 - 8%"])
        ws_c.append([6, "Overheads", "10 - 12%"])

        # PnL sheet with leading metadata banner
        ws_p = wb.create_sheet(title="Day Wise PnL")
        ws_p.append(["PROZO FULFILLMENT CENTRE - DAILY P&L MONITOR"])  # Row 1: Title
        ws_p.append(["Generated: 2026-11-23", "Status: Approved"])    # Row 2: Metadata
        ws_p.append([])                                               # Row 3: Blank
        ws_p.append([])                                               # Row 4: Blank
        # Row 5: Header
        ws_p.append(["Day", "Revenue", "Manpower", "Packaging", "Power and Fuel", "FC Rent", "Equipment Rentals", "Overheads"])
        # Data
        ws_p.append([1, 100000, 26000, 6000, 11000, 16000, 6000, 11000])
        ws_p.append([2, 105000, 27000, 6500, 11500, 16500, 6500, 11500])

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        df_pnl, criteria, warnings, errors, diag = extract_workbook_data(buf)
        self.assertIsNotNone(df_pnl)
        self.assertEqual(len(df_pnl), 2)
        self.assertEqual(len(errors), 0)
        self.assertEqual(diag["pnl"]["header_row_index"], 4)  # 0-indexed row 4 is row 5
        self.assertEqual(diag["criteria"]["header_row_index"], 3) # 0-indexed row 3 is row 4

    def test_blank_column_a_offset(self):
        """Test workbook where column A is empty and data starts in Column B."""
        wb = openpyxl.Workbook()
        ws_p = wb.active
        ws_p.title = "Day Wise PnL"
        # None in Col A
        ws_p.append([None, "Day", "Revenue", "Manpower", "Packaging", "Power and Fuel", "FC Rent", "Equipment Rentals", "Overheads"])
        ws_p.append([None, 1, 100000, 26000, 6000, 11000, 16000, 6000, 11000])

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        df_pnl, criteria, warnings, errors, diag = extract_workbook_data(buf)
        self.assertIsNotNone(df_pnl)
        self.assertEqual(len(df_pnl), 1)
        self.assertEqual(df_pnl.iloc[0]["Revenue"], 100000.0)

    def test_sheet_name_variations_and_insensitivity(self):
        """Test alternative sheet names like 'FC Criteria' and 'Day-Wise P&L'."""
        wb = openpyxl.Workbook()
        ws_c = wb.active
        ws_c.title = "  fc_criteria (v1) "
        ws_c.append(["Expense Category", "Target Range"])
        ws_c.append(["Manpower", "25 - 27%"])
        ws_c.append(["Packaging", "5 - 7%"])
        ws_c.append(["Power and Fuel", "10 - 12%"])
        ws_c.append(["FC Rent", "15 - 18%"])
        ws_c.append(["Equipment Rentals", "5 - 8%"])
        ws_c.append(["Overheads", "10 - 12%"])

        ws_p = wb.create_sheet(title="Day-Wise P&L")
        ws_p.append(["Day", "Revenue", "Manpower", "Packaging", "Power and Fuel", "FC Rent", "Equipment Rentals", "Overheads"])
        ws_p.append([1, 100000, 26000, 6000, 11000, 16000, 6000, 11000])

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        df_pnl, criteria, warnings, errors, diag = extract_workbook_data(buf)
        self.assertIsNotNone(df_pnl)
        self.assertEqual(len(df_pnl), 1)
        self.assertEqual(diag["pnl"]["sheet_selected"], "Day-Wise P&L")
        self.assertEqual(diag["criteria"]["sheet_selected"], "  fc_criteria (v1) ")

    def test_semantic_column_aliases_and_newlines(self):
        """Test column aliases: 'Power & Fuel', 'Rent', 'Labor', 'Total Revenue', 'Equipment Rental', 'Overhead Cost'."""
        wb = openpyxl.Workbook()
        ws_p = wb.active
        ws_p.title = "Day Wise PnL"
        # Column headers with aliases and newlines
        ws_p.append([
            "Day Number\n(#)",
            "Total Revenue\n(INR)",
            "Labor Cost",
            "Packing Material",
            "Power & Fuel",
            "Warehouse Rent",
            "Equipment Rental",
            "Overhead Cost"
        ])
        ws_p.append([1, 100000, 26000, 6000, 11000, 16000, 6000, 11000])

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        df_pnl, criteria, warnings, errors, diag = extract_workbook_data(buf)
        self.assertIsNotNone(df_pnl)
        self.assertEqual(len(df_pnl), 1)
        self.assertEqual(df_pnl.iloc[0]["Manpower"], 26000.0)
        self.assertEqual(df_pnl.iloc[0]["Packaging"], 6000.0)
        self.assertEqual(df_pnl.iloc[0]["Power and Fuel"], 11000.0)
        self.assertEqual(df_pnl.iloc[0]["FC Rent"], 16000.0)
        self.assertEqual(df_pnl.iloc[0]["Equipment Rentals"], 6000.0)
        self.assertEqual(df_pnl.iloc[0]["Overheads"], 11000.0)

    def test_diagnostic_reporting_on_missing_columns(self):
        """Test that missing columns produce rich diagnostic information instead of generic errors."""
        wb = openpyxl.Workbook()
        ws_p = wb.active
        ws_p.title = "Day Wise PnL"
        # Missing 'FC Rent' and 'Overheads'
        ws_p.append(["Day", "Revenue", "Manpower", "Packaging", "Power and Fuel", "Equipment Rentals"])
        ws_p.append([1, 100000, 26000, 6000, 11000, 6000])

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        df_pnl, criteria, warnings, errors, diag = extract_workbook_data(buf)
        self.assertIsNone(df_pnl)
        self.assertGreater(len(errors), 0)
        self.assertIn("FC Rent", diag["pnl"]["missing_columns"])
        self.assertIn("Overheads", diag["pnl"]["missing_columns"])
        # Check diagnostic details exist in error message
        self.assertIn("Workbook Sheets Detected", errors[0])
        self.assertIn("Columns Found on Header Row", errors[0])

if __name__ == "__main__":
    unittest.main()
