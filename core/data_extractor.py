"""
Data Extraction & Validation Module
Prozo Daily FC P&L Monitoring System

Features:
- Single-workbook ingestion: seamlessly extracts both 'Ideal FC Criteria' and 'Day Wise PnL' sheets
- Robust Stream Handling: accepts file path, bytes, or Streamlit UploadedFile with automatic pointer rewinding
- Intelligent Sheet Selection: Case-insensitive, whitespace-insensitive, symbol-insensitive matching
  with content-based fallback (inspecting cell contents if sheet names differ)
- Dynamic Header Detection: Automatically detects header row index even with leading blank rows,
  metadata rows, or title blocks (scans up to 25 rows)
- Semantic Column Normalization & Aliases:
  - Day: Day, Days, Day Number, Day #, DayNo, Day_No, D
  - Revenue: Revenue, Total Revenue, Gross Revenue, Turnover, Daily Revenue, Net Revenue, Rev
  - Manpower: Manpower, Manpower Cost, Manpower Expenses, Labor, Labour, Salary, Man Power
  - Packaging: Packaging, Packaging Cost, Packaging Expenses, Packing, Packing Material, Pack
  - Power and Fuel: Power and Fuel, Power & Fuel, Power & Fuel Cost, Electricity & Fuel, Utilities, Utility
  - FC Rent: FC Rent, Rent, Warehouse Rent, Facility Rent, FC Rental, Building Rent, Rental
  - Equipment Rentals: Equipment Rentals, Equipment Rental, Equipment Hire, MHE Rental, MHE Rentals, Machinery Rental
  - Overheads: Overheads, Overhead Cost, Overhead Expenses, Other Overheads, Admin & Overheads, Admin, Overhead
  - FC Identifier (optional): FC, FC ID, FC_ID, Fulfillment Centre, FC Name, Center, Hub, Warehouse
- Detailed Diagnostic Reporting: Sheet names found, sheet selected, found columns, missing columns, and validation errors
- Preserves raw input records without fabricating FC identifiers
"""

import io
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import numpy as np
import pandas as pd

# Default Company-Defined FC Criteria (used if sheet is missing or for CSV upload)
DEFAULT_CRITERIA: Dict[str, Tuple[float, float]] = {
    "Manpower": (25.0, 27.0),
    "Packaging": (5.0, 7.0),
    "Power and Fuel": (10.0, 12.0),
    "FC Rent": (15.0, 18.0),
    "Equipment Rentals": (5.0, 8.0),
    "Overheads": (10.0, 12.0),
}

# Standard CM Targets
CM_TARGETS = {
    "CM1": (65.0, 70.0),
    "CM2": (15.0, 30.0),
}

STANDARD_EXPENSE_KEYS = [
    "Manpower",
    "Packaging",
    "Power and Fuel",
    "FC Rent",
    "Equipment Rentals",
    "Overheads",
]

def ensure_file_bytes(file_or_path: Union[str, io.BytesIO, bytes, Any]) -> Tuple[Union[str, io.BytesIO], Optional[bytes]]:
    """
    Ensure the file stream can be read repeatedly without EOF exhaustion.
    Returns (readable_source, raw_bytes).
    """
    if isinstance(file_or_path, str):
        return file_or_path, None
    elif isinstance(file_or_path, bytes):
        return io.BytesIO(file_or_path), file_or_path
    elif hasattr(file_or_path, "read"):
        raw_bytes = file_or_path.read()
        if hasattr(file_or_path, "seek"):
            file_or_path.seek(0)
        return io.BytesIO(raw_bytes), raw_bytes
    return file_or_path, None

def normalize_text(text: Any) -> str:
    """Normalize text: lowercase, replace punctuation/newlines, strip whitespace."""
    if text is None or pd.isna(text):
        return ""
    s = str(text).strip().lower()
    s = s.replace("\n", " ").replace("\r", " ")
    s = s.replace("&", " and ")
    s = re.sub(r'[\s_]+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s.strip()

def parse_percentage_range(val: Any) -> Optional[Tuple[float, float]]:
    """
    Parse percentage range string (e.g. '25 - 27%', '25%-27%', '25 to 27', '0.25 - 0.27').
    Returns (min_pct, max_pct) in float percentages.
    """
    if val is None or pd.isna(val):
        return None
    val_str = str(val).strip().replace('%', '')
    m = re.search(r'([\d\.]+)\s*(?:-|to)\s*([\d\.]+)', val_str)
    if m:
        try:
            low = float(m.group(1))
            high = float(m.group(2))
            if 0 < low < 1.0 and 0 < high <= 1.0:
                low *= 100.0
                high *= 100.0
            return (min(low, high), max(low, high))
        except ValueError:
            return None
    return None

def match_column_name(col_name: str) -> Optional[str]:
    """
    Map raw column name to standardized line item name using semantic aliases.
    Strictly avoids ambiguous or unintended mappings.
    """
    norm = normalize_text(col_name)
    if not norm:
        return None

    # 1. Day Identifier
    if norm in ["day", "days", "d", "dayno", "day no", "daynum", "day num", "day number", "day id", "dayid"]:
        return "Day"

    # 2. Revenue
    if any(k in norm for k in ["revenue", "turnover", "gross sales", "total sales"]):
        return "Revenue"
    if norm in ["rev", "total rev", "gross rev"]:
        return "Revenue"

    # 3. Manpower
    if any(k in norm for k in ["manpower", "man power", "labor", "labour", "salary", "salaries", "workforce", "staffing"]):
        return "Manpower"

    # 4. Packaging
    if any(k in norm for k in ["packaging", "packing", "carton", "polybag"]):
        return "Packaging"
    if norm in ["pack", "packaging cost", "packing cost", "packaging material", "packing material"]:
        return "Packaging"

    # 5. Power and Fuel
    if "power" in norm and "fuel" in norm:
        return "Power and Fuel"
    if any(k in norm for k in ["electricity", "dg fuel", "diesel", "utility", "utilities"]):
        return "Power and Fuel"
    if norm in ["power", "fuel", "power fuel", "power and fuel cost", "power fuel cost"]:
        return "Power and Fuel"

    # 6. Equipment Rentals (check before FC Rent to avoid false 'rent' match)
    if any(k in norm for k in ["equipment", "mhe", "forklift", "reach truck", "bopt", "machinery"]):
        return "Equipment Rentals"
    if "rent" in norm and ("equip" in norm or "machine" in norm):
        return "Equipment Rentals"

    # 7. FC Rent (facility lease, excludes equipment)
    if any(k in norm for k in ["fc rent", "warehouse rent", "facility rent", "building rent", "fc rental", "shed rent"]):
        return "FC Rent"
    if norm in ["rent", "rental", "lease", "occupancy"]:
        return "FC Rent"

    # 8. Overheads
    if any(k in norm for k in ["overhead", "overheads", "admin", "administrative", "miscellaneous", "sundry"]):
        return "Overheads"

    # 9. Optional FC Identifier
    if norm in ["fc", "fc id", "fc_id", "fc name", "fulfillment centre", "fulfillment center", "facility", "center", "hub", "location", "wh", "warehouse"]:
        return "FC Identifier"

    return None

def detect_header_row(df_raw: pd.DataFrame, expected_type: str = "pnl", max_scan_rows: int = 25) -> Tuple[int, Dict[int, str], List[str]]:
    """
    Scans the raw DataFrame (read with header=None) up to max_scan_rows
    to detect the true header row dynamically.
    Returns:
        (header_row_index, col_index_to_mapped_name, raw_header_strings)
    """
    num_rows = min(len(df_raw), max_scan_rows)
    best_row_idx = 0
    best_score = 0
    best_mapping: Dict[int, str] = {}
    best_raw_headers: List[str] = []

    for r_idx in range(num_rows):
        row_vals = df_raw.iloc[r_idx].tolist()
        current_mapping = {}
        raw_headers = []
        
        for c_idx, val in enumerate(row_vals):
            val_str = str(val).strip() if pd.notna(val) else ""
            raw_headers.append(val_str)
            if not val_str:
                continue

            if expected_type == "pnl":
                matched = match_column_name(val_str)
                if matched and matched not in current_mapping.values():
                    current_mapping[c_idx] = matched
            elif expected_type == "criteria":
                norm = normalize_text(val_str)
                if any(k in norm for k in ["expense line item", "expense category", "line item", "category", "expense"]):
                    current_mapping[c_idx] = "category"
                elif any(k in norm for k in ["range", "target range", "allowed range", "% of revenue"]):
                    current_mapping[c_idx] = "range"
                elif "min" in norm:
                    current_mapping[c_idx] = "min"
                elif "max" in norm:
                    current_mapping[c_idx] = "max"

        score = len(current_mapping)
        if score > best_score:
            best_score = score
            best_row_idx = r_idx
            best_mapping = current_mapping
            best_raw_headers = raw_headers

        # Early exit if we found a strong match
        if expected_type == "pnl" and score >= 6:
            break
        if expected_type == "criteria" and score >= 2:
            break

    return best_row_idx, best_mapping, best_raw_headers

def find_sheet_by_name_or_content(
    xl: pd.ExcelFile,
    target_type: str
) -> Tuple[Optional[str], Optional[str]]:
    """
    Intelligently select the correct sheet for Criteria or PnL.
    1. Tests sheet names with case, whitespace, and punctuation insensitivity.
    2. Excludes criteria sheet candidates when searching for PnL.
    3. Content inspection fallback: reads first few rows to verify headers.
    Returns:
        (selected_sheet_name, match_reason)
    """
    sheet_names = xl.sheet_names
    if not sheet_names:
        return None, "Workbook contains no sheets."

    # Normalize all sheet names
    norm_sheets = {s: normalize_text(s) for s in sheet_names}

    if target_type == "criteria":
        criteria_keywords = ["idealfccriteria", "fccriteria", "criteria", "targetrange", "target", "ranges", "ideal"]
        for s, norm in norm_sheets.items():
            if any(k in norm for k in criteria_keywords):
                return s, f"Matched criteria keywords in sheet name '{s}'"
        
        # Content fallback
        for s in sheet_names:
            try:
                sample_df = xl.parse(s, nrows=10, header=None)
                _, mapping, _ = detect_header_row(sample_df, expected_type="criteria", max_scan_rows=10)
                if len(mapping) >= 2:
                    return s, f"Detected criteria table content in sheet '{s}'"
            except Exception:
                pass
        return None, "No criteria sheet found by name or content."

    elif target_type == "pnl":
        # First find criteria sheet to explicitly EXCLUDE it from PnL candidates
        crit_sheet, _ = find_sheet_by_name_or_content(xl, target_type="criteria")

        pnl_keywords = ["daywisepnl", "daywise", "dailyfc", "dailypnl", "pnl", "day wise", "daily", "days"]
        for s, norm in norm_sheets.items():
            if s == crit_sheet:
                continue
            if any(k in norm for k in ["criteria", "target", "ideal"]):
                continue
            if any(k in norm for k in pnl_keywords):
                return s, f"Matched P&L keywords in sheet name '{s}'"

        # Content inspection fallback
        for s in sheet_names:
            if s == crit_sheet:
                continue
            try:
                sample_df = xl.parse(s, nrows=20, header=None)
                _, mapping, _ = detect_header_row(sample_df, expected_type="pnl", max_scan_rows=20)
                if len(mapping) >= 4:
                    return s, f"Detected Day-Wise P&L table content in sheet '{s}' ({len(mapping)} columns matched)"
            except Exception:
                pass

        # If only 2 sheets and one is criteria, the other must be PnL
        if len(sheet_names) == 2 and crit_sheet:
            other = [s for s in sheet_names if s != crit_sheet][0]
            return other, f"Defaulted to remaining non-criteria sheet '{other}'"

        # If 1 sheet only, use it
        if len(sheet_names) == 1:
            return sheet_names[0], f"Single sheet in workbook '{sheet_names[0]}'"

        return None, "Could not unambiguously identify Day-Wise P&L sheet."

def extract_criteria(
    file_or_path: Union[str, io.BytesIO, bytes, Any],
    sheet_name: Optional[str] = None,
    return_diag: bool = False
) -> Union[Tuple[Dict[str, Tuple[float, float]], List[str]], Tuple[Dict[str, Tuple[float, float]], List[str], Dict[str, Any]]]:
    """
    Dynamically extract criteria from the Excel sheet.
    Returns:
        (criteria_dict, warnings) or (criteria_dict, warnings, diagnostic_info) if return_diag=True
    """
    criteria: Dict[str, Tuple[float, float]] = {}
    warnings: List[str] = []
    diag: Dict[str, Any] = {"status": "SUCCESS", "sheet_selected": None, "detected_sheets": []}

    try:
        source, _ = ensure_file_bytes(file_or_path)
        xl = pd.ExcelFile(source)
        diag["detected_sheets"] = xl.sheet_names

        # Select sheet
        target_sheet = sheet_name
        if not target_sheet or target_sheet not in xl.sheet_names:
            matched_sheet, reason = find_sheet_by_name_or_content(xl, "criteria")
            target_sheet = matched_sheet
            diag["match_reason"] = reason

        if not target_sheet:
            warnings.append(f"Criteria sheet not found among {xl.sheet_names}. Using default company criteria.")
            diag["status"] = "FALLBACK_DEFAULT"
            return (DEFAULT_CRITERIA.copy(), warnings, diag) if return_diag else (DEFAULT_CRITERIA.copy(), warnings)

        diag["sheet_selected"] = target_sheet

        # Read sheet with header=None to detect dynamic header row
        df_raw = xl.parse(target_sheet, header=None)
        h_idx, col_map, raw_headers = detect_header_row(df_raw, expected_type="criteria", max_scan_rows=20)
        diag["header_row_index"] = h_idx
        diag["raw_headers"] = raw_headers

        # Slice data rows
        df = df_raw.iloc[h_idx + 1:].copy()
        inv_map = {c_idx: name for c_idx, name in col_map.items()}

        cat_idx = [c for c, name in inv_map.items() if name == "category"]
        range_idx = [c for c, name in inv_map.items() if name == "range"]
        min_idx = [c for c, name in inv_map.items() if name == "min"]
        max_idx = [c for c, name in inv_map.items() if name == "max"]

        if not cat_idx:
            # Fallback: scan columns for text values matching standard categories
            for c_idx in range(df.shape[1]):
                col_text = df.iloc[:, c_idx].astype(str).tolist()
                if any("manpower" in normalize_text(v) for v in col_text):
                    cat_idx = [c_idx]
                    break

        if cat_idx:
            c_col = cat_idx[0]
            for _, row in df.iterrows():
                val_cat = row[c_col]
                if pd.isna(val_cat) or not str(val_cat).strip():
                    continue
                raw_cat = str(val_cat).strip()
                norm_cat = normalize_text(raw_cat)

                matched_key = None
                for std_k in STANDARD_EXPENSE_KEYS:
                    std_norm = normalize_text(std_k)
                    if std_norm in norm_cat or norm_cat in std_norm:
                        matched_key = std_k
                        break
                    if "power" in norm_cat and "fuel" in norm_cat and std_k == "Power and Fuel":
                        matched_key = std_k
                        break
                    if "rent" in norm_cat and "fc" in norm_cat and std_k == "FC Rent":
                        matched_key = std_k
                        break
                    if "equipment" in norm_cat and std_k == "Equipment Rentals":
                        matched_key = std_k
                        break

                if not matched_key:
                    continue

                range_tuple = None
                if range_idx:
                    range_tuple = parse_percentage_range(row[range_idx[0]])
                elif min_idx and max_idx:
                    try:
                        low_val = str(row[min_idx[0]]).replace('%', '')
                        high_val = str(row[max_idx[0]]).replace('%', '')
                        low = float(low_val)
                        high = float(high_val)
                        if 0 < low < 1.0 and 0 < high <= 1.0:
                            low *= 100.0
                            high *= 100.0
                        range_tuple = (min(low, high), max(low, high))
                    except (ValueError, TypeError):
                        pass

                # Search row values if range column index wasn't explicit
                if not range_tuple:
                    for cell_val in row:
                        parsed = parse_percentage_range(cell_val)
                        if parsed:
                            range_tuple = parsed
                            break

                if range_tuple:
                    criteria[matched_key] = range_tuple

        # Harmonize with defaults for any missing lines
        for std_k, def_range in DEFAULT_CRITERIA.items():
            if std_k not in criteria:
                criteria[std_k] = def_range
                warnings.append(f"Target for '{std_k}' missing in sheet. Applied default: {def_range[0]}% - {def_range[1]}%.")

        diag["criteria_extracted"] = {k: f"{v[0]}% - {v[1]}%" for k, v in criteria.items()}
        return (criteria, warnings, diag) if return_diag else (criteria, warnings)

    except Exception as e:
        warnings.append(f"Error parsing criteria: {str(e)}. Fallback to default criteria.")
        diag["status"] = "ERROR_FALLBACK"
        diag["error"] = str(e)
        return (DEFAULT_CRITERIA.copy(), warnings, diag) if return_diag else (DEFAULT_CRITERIA.copy(), warnings)

def extract_pnl(
    file_or_path: Union[str, io.BytesIO, bytes, Any],
    sheet_name: Optional[str] = None,
    is_csv: bool = False,
    return_diag: bool = False
) -> Any:
    """
    Extract and validate Day-Wise PnL records with dynamic sheet and header detection.
    Returns:
        (cleaned_df, validation_errors, raw_records) by default, or
        (cleaned_df, validation_errors, diagnostic_info) if return_diag=True.
    """
    validation_errors: List[str] = []
    diag: Dict[str, Any] = {
        "status": "SUCCESS",
        "sheet_selected": None,
        "detected_sheets": [],
        "found_columns": [],
        "mapped_columns": {},
        "missing_columns": [],
        "header_row_index": None
    }

    try:
        source, _ = ensure_file_bytes(file_or_path)

        if is_csv:
            df_raw = pd.read_csv(source, header=None)
            diag["detected_sheets"] = ["CSV File"]
            diag["sheet_selected"] = "CSV Upload"
        else:
            xl = pd.ExcelFile(source)
            diag["detected_sheets"] = xl.sheet_names

            # Select PnL Sheet
            target_sheet = sheet_name
            if not target_sheet or target_sheet not in xl.sheet_names:
                matched_sheet, reason = find_sheet_by_name_or_content(xl, "pnl")
                target_sheet = matched_sheet
                diag["match_reason"] = reason

            if not target_sheet:
                err_msg = (
                    f"Could not identify the Day-Wise P&L sheet. Available sheets in workbook: {xl.sheet_names}. "
                    f"Please ensure a sheet contains daily records or includes 'Day' and 'Revenue' columns."
                )
                diag["status"] = "SHEET_NOT_FOUND"
                return (None, [err_msg], diag) if return_diag else (None, [err_msg], [])

            diag["sheet_selected"] = target_sheet
            df_raw = xl.parse(target_sheet, header=None)

    except Exception as e:
        diag["status"] = "FILE_READ_ERROR"
        return (None, [f"Failed to read file: {str(e)}"], diag) if return_diag else (None, [f"Failed to read file: {str(e)}"], [])

    # Dynamic Header Detection
    h_idx, col_mapping, raw_headers = detect_header_row(df_raw, expected_type="pnl", max_scan_rows=25)
    diag["header_row_index"] = h_idx
    diag["found_columns"] = [h for h in raw_headers if h]
    diag["mapped_columns"] = {col_mapping[c_idx]: raw_headers[c_idx] for c_idx in col_mapping}

    # Verify required columns
    required_cols = ["Day", "Revenue"] + STANDARD_EXPENSE_KEYS
    missing_cols = [c for c in required_cols if c not in col_mapping.values()]
    diag["missing_columns"] = missing_cols

    if missing_cols:
        err_msg = (
            f"Missing required columns in dataset: {', '.join(missing_cols)}.\n"
            f"• Workbook Sheets Detected: {diag['detected_sheets']}\n"
            f"• Sheet Selected for P&L: '{diag['sheet_selected']}' (Header detected on Row {h_idx + 1})\n"
            f"• Columns Found on Header Row: {diag['found_columns']}\n"
            f"• Mapped Columns: {diag['mapped_columns']}"
        )
        diag["status"] = "MISSING_COLUMNS"
        return (None, [err_msg], diag) if return_diag else (None, [err_msg], [])

    # Slice data rows starting after header row
    df_data = df_raw.iloc[h_idx + 1:].copy()
    
    # Keep only mapped columns and rename to standard keys
    has_fc = "FC Identifier" in col_mapping.values()
    inv_col_map = {c_idx: std_name for c_idx, std_name in col_mapping.items()}
    
    cleaned_rows = []
    for r_idx, row in df_data.iterrows():
        excel_row_num = r_idx + 1

        # Check if entire row is empty
        if row.isna().all():
            continue

        day_raw = row[list(inv_col_map.keys())[list(inv_col_map.values()).index("Day")]]
        
        # Check if Day is empty or non-numeric header/footer (e.g. "Total" or metadata)
        if pd.isna(day_raw) or str(day_raw).strip() == "":
            continue
        day_str = str(day_raw).strip()
        if any(term in day_str.lower() for term in ["total", "summary", "average", "sum"]):
            continue

        try:
            day_clean = int(float(day_str))
        except (ValueError, TypeError):
            day_clean = day_str

        # Revenue
        rev_idx = list(inv_col_map.keys())[list(inv_col_map.values()).index("Revenue")]
        rev_raw = row[rev_idx]
        if pd.isna(rev_raw) or str(rev_raw).strip() == "":
            validation_errors.append(f"Day {day_clean} (Row {excel_row_num}): Revenue is missing.")
            continue
        try:
            revenue_num = float(str(rev_raw).replace(',', '').replace('₹', '').strip())
        except (ValueError, TypeError):
            validation_errors.append(f"Day {day_clean} (Row {excel_row_num}): Invalid revenue '{rev_raw}'.")
            continue

        if revenue_num <= 0:
            validation_errors.append(f"Day {day_clean} (Row {excel_row_num}): Revenue must be strictly positive (got {revenue_num}).")
            continue

        # Expenses
        expense_values = {}
        row_valid = True
        for exp_key in STANDARD_EXPENSE_KEYS:
            c_idx = list(inv_col_map.keys())[list(inv_col_map.values()).index(exp_key)]
            val_raw = row[c_idx]

            if pd.isna(val_raw) or str(val_raw).strip() == "":
                validation_errors.append(f"Day {day_clean} (Row {excel_row_num}): Missing value for '{exp_key}'. Not converted silently to 0.")
                row_valid = False
                break
            try:
                exp_num = float(str(val_raw).replace(',', '').replace('₹', '').strip())
                if exp_num < 0:
                    validation_errors.append(f"Day {day_clean} (Row {excel_row_num}): Negative expense '{exp_key}' = {exp_num}.")
                    row_valid = False
                    break
                expense_values[exp_key] = exp_num
            except (ValueError, TypeError):
                validation_errors.append(f"Day {day_clean} (Row {excel_row_num}): Invalid numeric value for '{exp_key}' = '{val_raw}'.")
                row_valid = False
                break

        if not row_valid:
            continue

        # Optional FC ID
        fc_id = None
        if has_fc:
            fc_c_idx = list(inv_col_map.keys())[list(inv_col_map.values()).index("FC Identifier")]
            fc_val = row[fc_c_idx]
            if pd.notna(fc_val) and str(fc_val).strip():
                fc_id = str(fc_val).strip()

        cleaned_rows.append({
            "Day": day_clean,
            "FC Identifier": fc_id,
            "Revenue": revenue_num,
            **expense_values
        })

    if not cleaned_rows:
        diag["status"] = "NO_VALID_DATA"
        errs = validation_errors + ["No valid financial day records could be parsed."]
        return (None, errs, diag) if return_diag else (None, errs, [])

    cleaned_df = pd.DataFrame(cleaned_rows)
    diag["total_records_extracted"] = len(cleaned_df)
    return (cleaned_df, validation_errors, diag) if return_diag else (cleaned_df, validation_errors, cleaned_rows)

def extract_workbook_data(
    file_or_path: Union[str, io.BytesIO, bytes, Any],
    is_csv: bool = False
) -> Tuple[Optional[pd.DataFrame], Dict[str, Tuple[float, float]], List[str], List[str], Dict[str, Any]]:
    """
    Unified extraction pipeline for a single uploaded workbook or CSV:
    1. Rewinds stream safely
    2. Dynamically extracts Criteria
    3. Dynamically extracts Day-Wise P&L
    4. Returns (cleaned_pnl_df, criteria_dict, warnings, errors, diagnostics)
    """
    all_warnings = []
    all_errors = []
    combined_diag = {}

    source, _ = ensure_file_bytes(file_or_path)

    if is_csv:
        criteria = DEFAULT_CRITERIA.copy()
        all_warnings.append("CSV source provided. Standard company criteria applied.")
        df_pnl, p_err, p_diag = extract_pnl(source, is_csv=True, return_diag=True)
        all_errors.extend(p_err)
        combined_diag["pnl"] = p_diag
        combined_diag["criteria"] = {"status": "CSV_DEFAULT"}
        return df_pnl, criteria, all_warnings, all_errors, combined_diag

    # Excel: Extract Criteria
    criteria, c_warn, c_diag = extract_criteria(source, return_diag=True)
    all_warnings.extend(c_warn)
    combined_diag["criteria"] = c_diag

    # Excel: Extract P&L (re-using source safely)
    df_pnl, p_err, p_diag = extract_pnl(source, is_csv=False, return_diag=True)
    all_errors.extend(p_err)
    combined_diag["pnl"] = p_diag

    return df_pnl, criteria, all_warnings, all_errors, combined_diag
