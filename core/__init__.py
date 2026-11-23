"""
Prozo Daily FC P&L Monitoring System - Core Modules
"""

from .data_extractor import (
    DEFAULT_CRITERIA,
    CM_TARGETS,
    STANDARD_EXPENSE_KEYS,
    extract_criteria,
    extract_pnl,
    extract_workbook_data,
    parse_percentage_range
)
from .financial_engine import (
    calculate_row_financials,
    calculate_dataframe_financials,
    calculate_portfolio_aggregates
)
from .anomaly_detector import (
    classify_fc_status,
    detect_expense_anomalies,
    detect_cm_deviations,
    analyze_day_record,
    analyze_full_dataset
)
from .ai_insights import (
    generate_rule_based_anomaly_insight,
    generate_suspicious_cm2_insight,
    generate_insights_for_day
)
from .trend_analyzer import (
    analyze_dataset_trends
)
from .report_generator import (
    generate_excel_report
)
from .test_scenarios import (
    run_all_scenarios,
    execute_scenario
)
