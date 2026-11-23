"""
PROZO P&L MONITOR - DATA PROCESSOR MODULE
Handles intelligent Excel reading, header detection (tolerant of non-zero rows),
fuzzy column mapping, target criteria extraction, and data validation.
"""

import io
import re
from typing import Dict, List, Tuple, Any, Optional
import pandas as pd
import openpyxl


# Standard expense categories required by Prozo
EXPENSE_CATEGORIES = [
    "Manpower Cost",
    "Packaging Cost",
    "Power and Fuel",
    "FC Rent",
    "Equipment Rentals",
    "Overheads",
]

# Canonical column names
COL_DATE = "Date"
COL_FC = "FC"
COL_REVENUE = "Revenue"

# Synonyms for column matching
COLUMN_SYNONYMS = {
    COL_DATE: ["date", "day", "record_date", "txn_date", "dated", "posting_date", "period"],
    COL_FC: ["fc", "fc identifier", "fc name", "fulfillment centre", "fulfillment center", "warehouse", "facility", "hub", "location", "site"],
    COL_REVENUE: ["revenue", "sales", "gross revenue", "net revenue", "turnover", "total revenue", "billing"],
    "Manpower Cost": ["manpower cost", "manpower", "labor cost", "labour cost", "staff cost", "workforce cost", "payroll", "man power"],
    "Packaging Cost": ["packaging cost", "packaging", "packing cost", "packing material", "packaging material", "pkg cost"],
    "Power and Fuel": ["power and fuel", "power & fuel", "electricity & fuel", "utilities", "power cost", "fuel cost", "energy cost"],
    "FC Rent": ["fc rent", "rent", "facility rent", "warehouse rent", "lease", "building rent", "premises rent"],
    "Equipment Rentals": ["equipment rentals", "equipment rental", "mhe rental", "machinery rental", "forklift rental", "equipment lease"],
    "Overheads": ["overheads", "overhead", "overhead cost", "admin expenses", "general overheads", "other expenses", "misc overheads"]
}

# Synonyms for Ideal FC Criteria category matching
CRITERIA_CATEGORY_SYNONYMS = {
    "Manpower": ["manpower", "labor", "labour", "staff", "workforce", "manpower cost"],
    "Packaging": ["packaging", "packing", "packing material", "packaging cost"],
    "Power and Fuel": ["power and fuel", "power & fuel", "electricity", "fuel", "utilities", "energy"],
    "FC Rent": ["fc rent", "rent", "facility rent", "lease"],
    "Equipment Rentals": ["equipment rentals", "equipment rental", "mhe", "machinery"],
    "Overheads": ["overheads", "overhead", "admin", "other expenses"],
    "CM1": ["cm1", "contribution margin 1", "cm-1", "contribution margin 1 %", "cm 1"],
    "CM2": ["cm2", "contribution margin 2", "cm-2", "contribution margin 2 %", "cm 2"],
}


def clean_text(val: Any) -> str:
    """Normalize text for matching."""
    if val is None or pd.isna(val):
        return ""
    text = str(val).strip().lower()
    text = re.sub(r'[\r\n\t_]+', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def parse_percentage_value(val: Any) -> Optional[float]:
    """Parse percentage strings (e.g. '25%', '0.25', '25', '25.5 %') into float percentage (e.g. 25.0)."""
    if val is None or pd.isna(val):
        return None
    val_str = str(val).strip().replace('%', '').strip()
    try:
        num = float(val_str)
        # If user entered 0.25 instead of 25 for percentage
        if 0.0 < num <= 1.0:
            num = num * 100.0
        return round(num, 2)
    except ValueError:
        return None


def detect_header_row(df_raw: pd.DataFrame, target_synonyms: List[str]) -> int:
    """
    Search first 15 rows to find the row that has the highest concentration of target header names.
    Returns row index.
    """
    best_row = 0
    max_matches = 0

    for r_idx in range(min(15, len(df_raw))):
        row_vals = [clean_text(v) for v in df_raw.iloc[r_idx].values]
        matches = 0
        for val in row_vals:
            if not val:
                continue
            for syn in target_synonyms:
                if syn in val or val in syn:
                    matches += 1
                    break
        if matches > max_matches:
            max_matches = matches
            best_row = r_idx

    return best_row


def match_column(col_candidates: List[str], target_name: str) -> Optional[str]:
    """Match a column header against canonical target synonyms."""
    synonyms = COLUMN_SYNONYMS.get(target_name, [clean_text(target_name)])
    
    # 1. Exact cleaned match
    for col in col_candidates:
        c_clean = clean_text(col)
        # Prevent FC Rent from matching FC
        if target_name == COL_FC and "rent" in c_clean:
            continue
        if c_clean in synonyms:
            return col

    # 2. Substring match
    for col in col_candidates:
        c_clean = clean_text(col)
        # Prevent FC Rent from matching FC
        if target_name == COL_FC and "rent" in c_clean:
            continue
        for syn in synonyms:
            # For short tokens like 'fc', require word boundary or exact match
            if syn == "fc":
                if c_clean == "fc" or c_clean.startswith("fc ") or c_clean.endswith(" fc"):
                    return col
            elif syn in c_clean or c_clean in syn:
                return col

    return None


def inspect_and_read_single_workbook(file_bytes_or_buffer: Any) -> Tuple[Optional[pd.DataFrame], Optional[Dict[str, Dict[str, float]]], Dict[str, Any]]:
    """
    Reads a single Excel workbook containing TWO separate worksheets:
    1. 'Ideal FC Criteria'
    2. 'Day Wise PnL'
    
    Identifies sheets dynamically, extracts targets and records,
    validates both independently, and returns standardized structures.
    """
    validation_info = {
        "is_valid": False,
        "errors": [],
        "warnings": [],
        "sheet_names": [],
        "criteria_sheet_name": None,
        "pnl_sheet_name": None,
        "row_count": 0,
        "has_fc": False,
        "extracted_targets": {},
        "detected_headers": {},
    }

    try:
        excel_file = pd.ExcelFile(file_bytes_or_buffer)
        all_sheets = excel_file.sheet_names
        validation_info["sheet_names"] = all_sheets

        # 1. Identify "Ideal FC Criteria" worksheet
        criteria_sheet = None
        for s in all_sheets:
            s_clean = clean_text(s)
            if "ideal" in s_clean or "criteria" in s_clean or "target" in s_clean:
                criteria_sheet = s
                break

        # 2. Identify "Day Wise PnL" worksheet
        pnl_sheet = None
        for s in all_sheets:
            s_clean = clean_text(s)
            if "pnl" in s_clean or "day wise" in s_clean or "daily" in s_clean or "financial" in s_clean:
                pnl_sheet = s
                break

        # Fallback if names are reversed or exactly 2 sheets
        if not criteria_sheet and len(all_sheets) >= 2:
            criteria_sheet = all_sheets[0]
        if not pnl_sheet and len(all_sheets) >= 2:
            pnl_sheet = all_sheets[1]

        if not criteria_sheet:
            validation_info["errors"].append("Missing required worksheet: 'Ideal FC Criteria'. Found sheets: " + ", ".join(all_sheets))
        if not pnl_sheet:
            validation_info["errors"].append("Missing required worksheet: 'Day Wise PnL'. Found sheets: " + ", ".join(all_sheets))

        if validation_info["errors"]:
            return None, None, validation_info

        validation_info["criteria_sheet_name"] = criteria_sheet
        validation_info["pnl_sheet_name"] = pnl_sheet

        # Extract Targets from Criteria Sheet
        df_crit_raw = excel_file.parse(criteria_sheet, header=None)
        extracted_targets, crit_info = extract_targets_from_raw_df(df_crit_raw)
        validation_info["warnings"].extend(crit_info["warnings"])
        if not crit_info["is_valid"]:
            validation_info["errors"].extend(crit_info["errors"])
            return None, None, validation_info
        validation_info["extracted_targets"] = extracted_targets

        # Extract Records from PnL Sheet
        df_pnl_raw = excel_file.parse(pnl_sheet, header=None)
        std_pnl_df, pnl_info = parse_pnl_from_raw_df(df_pnl_raw, excel_file, pnl_sheet)
        validation_info["warnings"].extend(pnl_info["warnings"])
        if not pnl_info["is_valid"]:
            validation_info["errors"].extend(pnl_info["errors"])
            return None, None, validation_info

        validation_info["has_fc"] = pnl_info["has_fc"]
        validation_info["row_count"] = pnl_info["row_count"]
        validation_info["detected_headers"] = pnl_info["detected_headers"]
        validation_info["is_valid"] = True

        return std_pnl_df, extracted_targets, validation_info

    except Exception as e:
        validation_info["errors"].append(f"Failed to read Excel workbook: {str(e)}")
        return None, None, validation_info


def extract_targets_from_raw_df(df_raw: pd.DataFrame) -> Tuple[Optional[Dict[str, Dict[str, float]]], Dict[str, Any]]:
    """Extracts target percentage ranges from a raw criteria dataframe."""
    validation_info = {"is_valid": False, "errors": [], "warnings": {}}
    warnings_list = []
    
    default_targets = {
        "Manpower": {"min": 25.0, "max": 27.0},
        "Packaging": {"min": 5.0, "max": 7.0},
        "Power and Fuel": {"min": 10.0, "max": 12.0},
        "FC Rent": {"min": 15.0, "max": 18.0},
        "Equipment Rentals": {"min": 5.0, "max": 8.0},
        "Overheads": {"min": 10.0, "max": 12.0},
        "CM1": {"min": 65.0, "max": 70.0},
        "CM2": {"min": 15.0, "max": 30.0},
    }

    if df_raw.empty:
        validation_info["errors"] = ["The Criteria worksheet is empty."]
        validation_info["warnings"] = warnings_list
        return None, validation_info

    extracted_targets = {}

    for r_idx in range(len(df_raw)):
        row_items = [clean_text(v) for v in df_raw.iloc[r_idx].values]
        
        for cat_canonical, synonyms in CRITERIA_CATEGORY_SYNONYMS.items():
            if cat_canonical in extracted_targets:
                continue
            
            matched_col_idx = None
            for c_idx, cell_str in enumerate(row_items):
                for syn in synonyms:
                    if syn == cell_str or f" {syn} " in f" {cell_str} ":
                        matched_col_idx = c_idx
                        break
                if matched_col_idx is not None:
                    break

            if matched_col_idx is not None:
                # Look at adjacent cells in this row for min/max
                found_range = False
                # First check for explicit range string e.g. "25 - 27%" or "25% - 27%"
                for c_idx in range(len(df_raw.columns)):
                    if c_idx == matched_col_idx:
                        continue
                    cell_raw = df_raw.iloc[r_idx, c_idx]
                    cell_str = str(cell_raw).strip()
                    range_match = re.findall(r'(\d+(?:\.\d+)?)\s*%?\s*[-–to]+\s*(\d+(?:\.\d+)?)\s*%?', cell_str)
                    if range_match:
                        try:
                            v_min = float(range_match[0][0])
                            v_max = float(range_match[0][1])
                            if 0.0 < v_min <= 1.0 and 0.0 < v_max <= 1.0:
                                v_min *= 100.0
                                v_max *= 100.0
                            extracted_targets[cat_canonical] = {
                                "min": min(v_min, v_max),
                                "max": max(v_min, v_max)
                            }
                            found_range = True
                            break
                        except ValueError:
                            pass

                # If no range string, check for separate min and max numeric columns
                if not found_range:
                    numbers_in_row = []
                    for c_idx in range(len(df_raw.columns)):
                        if c_idx == matched_col_idx:
                            continue
                        # If header of this column looks like # or serial, skip
                        header_val = clean_text(df_raw.iloc[0, c_idx])
                        if header_val in ["#", "s.no", "sno", "sr no", "id"]:
                            continue
                        cell_raw = df_raw.iloc[r_idx, c_idx]
                        p_val = parse_percentage_value(cell_raw)
                        if p_val is not None:
                            numbers_in_row.append(p_val)

                    if len(numbers_in_row) >= 2:
                        extracted_targets[cat_canonical] = {
                            "min": min(numbers_in_row[0], numbers_in_row[1]),
                            "max": max(numbers_in_row[0], numbers_in_row[1])
                        }
                    elif len(numbers_in_row) == 1:
                        extracted_targets[cat_canonical] = {
                            "min": numbers_in_row[0],
                            "max": numbers_in_row[0]
                        }

    # Add default CM1 / CM2 targets if not present
    for k, v in default_targets.items():
        if k not in extracted_targets:
            warnings_list.append(f"Target for '{k}' not found in sheet. Using standard Prozo target: {v['min']}% - {v['max']}%.")
            extracted_targets[k] = v.copy()

    for cat, t in extracted_targets.items():
        if t["min"] > t["max"]:
            t["min"], t["max"] = t["max"], t["min"]

    validation_info["is_valid"] = True
    validation_info["warnings"] = warnings_list
    return extracted_targets, validation_info


def parse_pnl_from_raw_df(df_raw: pd.DataFrame, excel_file: pd.ExcelFile, sheet_name: str) -> Tuple[Optional[pd.DataFrame], Dict[str, Any]]:
    """Parses and standardizes P&L records from raw dataframe."""
    validation_info = {
        "is_valid": False,
        "errors": [],
        "warnings": [],
        "detected_headers": {},
        "has_fc": False,
        "row_count": 0
    }

    if df_raw.empty:
        validation_info["errors"].append("The P&L worksheet is empty.")
        return None, validation_info

    keywords = ["revenue", "manpower", "packaging", "power", "date", "day", "fc", "rent"]
    header_row = detect_header_row(df_raw, keywords)

    df = excel_file.parse(sheet_name, skiprows=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    raw_cols = [c for c in df.columns if not c.startswith("Unnamed")]

    matched_cols = {}
    for canonical in [COL_DATE, COL_FC, COL_REVENUE] + EXPENSE_CATEGORIES:
        found = match_column(raw_cols, canonical)
        if found:
            matched_cols[canonical] = found

    validation_info["detected_headers"] = matched_cols

    missing_critical = []
    if COL_DATE not in matched_cols:
        missing_critical.append("Date / Day")
    if COL_REVENUE not in matched_cols:
        missing_critical.append("Revenue")
    
    for exp in EXPENSE_CATEGORIES:
        if exp not in matched_cols:
            missing_critical.append(exp)

    if missing_critical:
        validation_info["errors"].append(
            f"Missing required columns in Day Wise PnL worksheet: {', '.join(missing_critical)}. "
            f"Found columns: {', '.join(raw_cols)}"
        )
        return None, validation_info

    has_fc = COL_FC in matched_cols
    validation_info["has_fc"] = has_fc
    if not has_fc:
        validation_info["warnings"].append(
            "FC identifier column not found. Performing Day-wise aggregated analysis across all operational records."
        )

    std_df = pd.DataFrame()
    raw_dates = df[matched_cols[COL_DATE]]
    
    # Clean Day / Date label
    def format_day_date(d):
        if pd.isna(d):
            return "Unknown"
        if isinstance(d, (int, float)):
            return f"Day {int(d)}"
        if hasattr(d, 'date'):
            return str(d.date())
        d_str = str(d).strip()
        if d_str.isdigit():
            return f"Day {d_str}"
        return d_str

    std_df["Date"] = raw_dates.apply(format_day_date)

    if has_fc:
        std_df["FC"] = df[matched_cols[COL_FC]].apply(lambda x: "All-FC" if pd.isna(x) or str(x).strip() == "" else str(x).strip())
    else:
        std_df["FC"] = "Prozo FC (Day-wise)"

    std_df["Revenue"] = pd.to_numeric(df[matched_cols[COL_REVENUE]], errors="coerce")

    for exp in EXPENSE_CATEGORIES:
        std_df[exp] = pd.to_numeric(df[matched_cols[exp]], errors="coerce")

    std_df = std_df.dropna(subset=["Date", "Revenue"])
    invalid_rev = std_df[std_df["Revenue"] <= 0]
    if not invalid_rev.empty:
        validation_info["warnings"].append(f"Ignored {len(invalid_rev)} rows with zero or negative revenue.")
        std_df = std_df[std_df["Revenue"] > 0]

    for exp in EXPENSE_CATEGORIES:
        null_count = std_df[exp].isna().sum()
        if null_count > 0:
            validation_info["warnings"].append(f"Filled {null_count} missing value(s) in '{exp}' with ₹0.")
            std_df[exp] = std_df[exp].fillna(0.0)

    if std_df.empty:
        validation_info["errors"].append("No valid records found in PnL worksheet after validating Date and Revenue.")
        return None, validation_info

    validation_info["row_count"] = len(std_df)
    validation_info["is_valid"] = True
    return std_df, validation_info

    """
    Reads the Daily P&L Excel file, detects headers even if not at row 0,
    maps columns, handles missing FC identification, and validates rows.
    """
    validation_info = {
        "is_valid": False,
        "errors": [],
        "warnings": [],
        "detected_headers": {},
        "has_fc": False,
        "row_count": 0,
        "date_range": (None, None),
    }

    try:
        # Load raw sheet without header to scan
        excel_file = pd.ExcelFile(file_bytes_or_buffer)
        sheet_name = excel_file.sheet_names[0]
        df_raw = excel_file.parse(sheet_name, header=None)

        if df_raw.empty:
            validation_info["errors"].append("The uploaded Daily P&L Excel file is empty.")
            return None, validation_info

        # Candidate keywords to find header row
        keywords = ["revenue", "manpower", "packaging", "power", "date", "fc", "rent"]
        header_row = detect_header_row(df_raw, keywords)

        # Parse with detected header
        df = excel_file.parse(sheet_name, skiprows=header_row)
        
        # Clean column names (strip whitespace and unname columns)
        df.columns = [str(c).strip() for c in df.columns]
        raw_cols = [c for c in df.columns if not c.startswith("Unnamed")]

        # Match columns
        matched_cols = {}
        for canonical in [COL_DATE, COL_FC, COL_REVENUE] + EXPENSE_CATEGORIES:
            found = match_column(raw_cols, canonical)
            if found:
                matched_cols[canonical] = found

        validation_info["detected_headers"] = matched_cols

        # Required fields check
        missing_critical = []
        if COL_DATE not in matched_cols:
            missing_critical.append("Date / Day")
        if COL_REVENUE not in matched_cols:
            missing_critical.append("Revenue")
        
        for exp in EXPENSE_CATEGORIES:
            if exp not in matched_cols:
                missing_critical.append(exp)

        if missing_critical:
            validation_info["errors"].append(
                f"Missing required columns in Daily P&L file: {', '.join(missing_critical)}. "
                f"Found columns: {', '.join(raw_cols)}"
            )
            return None, validation_info

        # Check FC identifier
        has_fc = COL_FC in matched_cols
        validation_info["has_fc"] = has_fc
        if not has_fc:
            validation_info["warnings"].append(
                "FC identifier column not found. System will perform Day-wise aggregated analysis."
            )

        # Construct standardized DataFrame
        std_df = pd.DataFrame()
        
        # Date processing
        raw_dates = df[matched_cols[COL_DATE]]
        # Try converting to datetime or clean string
        std_df["Date"] = raw_dates.apply(lambda d: str(d.date()) if hasattr(d, 'date') else str(d).strip())

        # FC
        if has_fc:
            std_df["FC"] = df[matched_cols[COL_FC]].apply(lambda x: "All-FC" if pd.isna(x) or str(x).strip() == "" else str(x).strip())
        else:
            std_df["FC"] = "FC-Consolidated (Day-wise)"

        # Revenue
        std_df["Revenue"] = pd.to_numeric(df[matched_cols[COL_REVENUE]], errors="coerce")

        # Expense categories
        for exp in EXPENSE_CATEGORIES:
            std_df[exp] = pd.to_numeric(df[matched_cols[exp]], errors="coerce")

        # Drop rows where Date is completely empty or Revenue is NaN
        initial_len = len(std_df)
        std_df = std_df.dropna(subset=["Date", "Revenue"])
        # Remove zero or negative revenue rows or flag them
        invalid_rev = std_df[std_df["Revenue"] <= 0]
        if not invalid_rev.empty:
            validation_info["warnings"].append(f"Ignored {len(invalid_rev)} rows with zero or negative revenue.")
            std_df = std_df[std_df["Revenue"] > 0]

        # Fill remaining missing expense values with 0.0 and warn
        for exp in EXPENSE_CATEGORIES:
            null_count = std_df[exp].isna().sum()
            if null_count > 0:
                validation_info["warnings"].append(f"Filled {null_count} missing value(s) in '{exp}' with ₹0.")
                std_df[exp] = std_df[exp].fillna(0.0)

        if std_df.empty:
            validation_info["errors"].append("No valid records found after validating Date and Revenue.")
            return None, validation_info

        validation_info["row_count"] = len(std_df)
        validation_info["date_range"] = (std_df["Date"].min(), std_df["Date"].max())
        validation_info["is_valid"] = True
        return std_df, validation_info

    except Exception as e:
        validation_info["errors"].append(f"Failed to parse Daily P&L file: {str(e)}")
        return None, validation_info


def inspect_and_read_criteria_excel(file_bytes_or_buffer: Any) -> Tuple[Optional[Dict[str, Dict[str, float]]], Dict[str, Any]]:
    """
    Reads the Ideal FC Criteria Excel file, extracts min and max target percentages
    for expense categories, CM1, and CM2.
    """
    validation_info = {
        "is_valid": False,
        "errors": [],
        "warnings": [],
        "extracted_targets": {},
    }

    # Default baseline criteria in case of fallback
    default_targets = {
        "Manpower": {"min": 25.0, "max": 27.0},
        "Packaging": {"min": 5.0, "max": 7.0},
        "Power and Fuel": {"min": 10.0, "max": 12.0},
        "FC Rent": {"min": 15.0, "max": 18.0},
        "Equipment Rentals": {"min": 5.0, "max": 8.0},
        "Overheads": {"min": 10.0, "max": 12.0},
        "CM1": {"min": 65.0, "max": 70.0},
        "CM2": {"min": 15.0, "max": 30.0},
    }

    try:
        excel_file = pd.ExcelFile(file_bytes_or_buffer)
        sheet_name = excel_file.sheet_names[0]
        df_raw = excel_file.parse(sheet_name, header=None)

        if df_raw.empty:
            validation_info["errors"].append("The Ideal FC Criteria Excel file is empty.")
            return None, validation_info

        extracted_targets = {}

        # Scan all rows and columns to find category names and their min/max values
        # Formats might be:
        # 1. Columns: [Category, Min, Max]
        # 2. Columns: [Category, Target Range] e.g. "25% - 27%"
        # 3. Transposed table

        # Convert entire raw table into strings for pattern search
        for r_idx in range(len(df_raw)):
            row_items = [clean_text(v) for v in df_raw.iloc[r_idx].values]
            
            for cat_canonical, synonyms in CRITERIA_CATEGORY_SYNONYMS.items():
                if cat_canonical in extracted_targets:
                    continue
                
                # Check if this row contains the category
                matched_col_idx = None
                for c_idx, cell_str in enumerate(row_items):
                    for syn in synonyms:
                        # Match exact or clean word
                        if syn == cell_str or f" {syn} " in f" {cell_str} ":
                            matched_col_idx = c_idx
                            break
                    if matched_col_idx is not None:
                        break

                if matched_col_idx is not None:
                    # Look at adjacent cells in this row for min/max
                    numbers_in_row = []
                    for c_idx in range(len(df_raw.columns)):
                        if c_idx == matched_col_idx:
                            continue
                        cell_raw = df_raw.iloc[r_idx, c_idx]
                        cell_str = str(cell_raw)
                        # Check if cell has range format e.g. "25% - 27%" or "25-27"
                        range_match = re.findall(r'(\d+(?:\.\d+)?)\s*%\s*[-–to]+\s*(\d+(?:\.\d+)?)\s*%', cell_str)
                        if not range_match:
                            range_match = re.findall(r'(\d+(?:\.\d+)?)\s*[-–to]+\s*(\d+(?:\.\d+)?)', cell_str)
                        if range_match:
                            try:
                                v_min = float(range_match[0][0])
                                v_max = float(range_match[0][1])
                                if 0.0 < v_min <= 1.0 and 0.0 < v_max <= 1.0:
                                    v_min *= 100.0
                                    v_max *= 100.0
                                numbers_in_row.extend([v_min, v_max])
                                break
                            except ValueError:
                                pass
                        else:
                            p_val = parse_percentage_value(cell_raw)
                            if p_val is not None:
                                numbers_in_row.append(p_val)

                    if len(numbers_in_row) >= 2:
                        extracted_targets[cat_canonical] = {
                            "min": min(numbers_in_row[0], numbers_in_row[1]),
                            "max": max(numbers_in_row[0], numbers_in_row[1])
                        }
                    elif len(numbers_in_row) == 1:
                        # Single target or equal
                        extracted_targets[cat_canonical] = {
                            "min": numbers_in_row[0],
                            "max": numbers_in_row[0]
                        }

        # If CM1 or CM2 were not found in the sheet, check default baseline or warn
        for k, v in default_targets.items():
            if k not in extracted_targets:
                validation_info["warnings"].append(
                    f"Target for '{k}' not found in Criteria file. Using standard Prozo target: {v['min']}% - {v['max']}%."
                )
                extracted_targets[k] = v.copy()

        # Validate that min <= max
        for cat, t in extracted_targets.items():
            if t["min"] > t["max"]:
                t["min"], t["max"] = t["max"], t["min"]

        validation_info["extracted_targets"] = extracted_targets
        validation_info["is_valid"] = True
        return extracted_targets, validation_info

    except Exception as e:
        validation_info["errors"].append(f"Failed to parse Ideal FC Criteria file: {str(e)}")
        return None, validation_info
