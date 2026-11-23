"""
AI Insight & Recommendation Engine
Prozo Daily FC P&L Monitoring System

Generates natural language explanations and actionable recommendations for:
- Individual expense breaches
- CM1/CM2 deviations
- Suspiciously high CM2 (> 33%)
- Comprehensive daily summaries

Features:
- Structured output: What happened, Affected metric, Actual %, Expected %, Deviation, Contributing factors, Recommended action.
- LLM integration: Uses Google Gemini API if GEMINI_API_KEY / GOOGLE_API_KEY is available in env.
- Deterministic Rule-Based Fallback Engine: High quality financial logic guarantees full functionality offline or without API keys.
- Mathematical integrity: Passes pre-calculated values; LLM cannot alter numbers or rules.
"""

import os
import json
from typing import Any, Dict, List, Optional
import requests

# Knowledge base of domain-specific operational contributing factors & actions for Prozo FC operations
CONTRIBUTING_FACTORS_KB = {
    "Manpower": {
        "ABOVE TARGET": [
            "Excessive overtime hours or surge staffing for peak inbound/outbound volumes.",
            "Higher ratio of contractual/temp labor hired at premium rates without volume justification.",
            "Unplanned warehouse downtime or conveyor stalls leading to idle paid hours.",
            "Under-utilization of pick-and-pack labor due to uneven order flow."
        ],
        "BELOW TARGET": [
            "Severe understaffing leading to potential order dispatch delays and SLA risks.",
            "Delayed contractor billing or labor accrual entries not yet posted for the day.",
            "Temporary workforce shortages during festival or local holidays."
        ],
        "RECOMMENDATION_HIGH": "Audit daily roster vs dispatched units; enforce productivity benchmarks (Units Per Hour); verify overtime approvals.",
        "RECOMMENDATION_LOW": "Audit labor timesheets to confirm all worker shifts were booked; monitor dispatch SLA compliance to ensure packing quality was not compromised."
    },
    "Packaging": {
        "ABOVE TARGET": [
            "Excessive box sizing or inappropriate carton selection leading to higher packaging unit costs.",
            "High bubble wrap / filler consumption or packaging material wastage during packing.",
            "Batch procurement of specialty corrugated boxes billed on a single day.",
            "Damaged packaging material written off directly to daily operational expense."
        ],
        "BELOW TARGET": [
            "Bulk packing supplies consumed from pre-existing stock without corresponding ERP goods issue.",
            "Shipper carton stock-out resulting in poly-mailer substitutions.",
            "Delayed supplier packaging invoice posting."
        ],
        "RECOMMENDATION_HIGH": "Inspect standard packaging operating procedures (SOPs); verify carton-to-SKU volumetric mapping; conduct physical scrap audit.",
        "RECOMMENDATION_LOW": "Cross-check physical packaging inventory consumption against daily dispatches; verify ERP consumable ledger postings."
    },
    "Power and Fuel": {
        "ABOVE TARGET": [
            "High diesel generator (DG) running hours during grid power cuts.",
            "Continuous HVAC / temperature control operation during high ambient weather conditions.",
            "Diesel fuel tank refilling billed as a lump sum on this single operational day.",
            "Material handling equipment (MHE) battery charging cycles overlapping peak utility tariff hours."
        ],
        "BELOW TARGET": [
            "Favorable grid uptime minimizing backup generator reliance.",
            "Utility billing split or meter reading timing discrepancy.",
            "Solar captive generation offsetting grid consumption."
        ],
        "RECOMMENDATION_HIGH": "Inspect DG logbooks and fuel consumption per run hour; review tariff time-of-day slots for battery recharging.",
        "RECOMMENDATION_LOW": "Confirm that fuel delivery slips and power bills for the day are fully captured in the ledger."
    },
    "FC Rent": {
        "ABOVE TARGET": [
            "Daily revenue dipped significantly below standard run-rate while rent remains fixed.",
            "Inclusion of short-term overflow warehouse or mezzanine rental surcharge.",
            "Quarterly CAM (Common Area Maintenance) or property tax adjustment accounted on this day."
        ],
        "BELOW TARGET": [
            "Exceptionally high revenue surge diluting fixed rent percentage.",
            "Rental escalation rebate or vendor credit adjustment passed in ledger."
        ],
        "RECOMMENDATION_HIGH": "Analyze capacity utilization; since rent is largely fixed, focus on driving throughput and utilization to absorb fixed overheads.",
        "RECOMMENDATION_LOW": "Ensure rent amortization entry was calculated over the correct day count."
    },
    "Equipment Rentals": {
        "ABOVE TARGET": [
            "Additional short-term forklift, reach truck, or BOPT hires for handling pallet surges.",
            "Equipment repair, battery replacement, or scheduled maintenance billed under rental line item.",
            "Low revenue volume exaggerating leased equipment fixed cost proportion."
        ],
        "BELOW TARGET": [
            "Equipment downtime or off-hiring of redundant material handling machinery.",
            "Vendor rental invoice unposted."
        ],
        "RECOMMENDATION_HIGH": "Review MHE utilization logs; return standby equipment if pallet movements do not warrant additional units.",
        "RECOMMENDATION_LOW": "Verify all rented material handling assets are active and invoiced."
    },
    "Overheads": {
        "ABOVE TARGET": [
            "Unscheduled facility maintenance, pest control, or security guard surge billing.",
            "Office consumables, cafeteria supplies, or IT hardware replacement purchases.",
            "One-time municipal license, compliance, or regulatory renewal fee."
        ],
        "BELOW TARGET": [
            "Deferred maintenance or timing delay in miscellaneous vendor bill submissions."
        ],
        "RECOMMENDATION_HIGH": "Review sundry voucher register; mandate pre-approval thresholds for administrative and non-recurring expenses.",
        "RECOMMENDATION_LOW": "Confirm administrative expenses were not booked to a suspense account."
    }
}

def generate_rule_based_anomaly_insight(anomaly: Dict[str, Any], day_summary: Dict[str, Any]) -> Dict[str, str]:
    """
    Deterministic rule-based generator fulfilling the 7 mandatory requirements:
    1. What happened?
    2. Which expense category or metric is affected?
    3. Actual percentage or value.
    4. Expected target range.
    5. Deviation from target.
    6. Possible contributing factors.
    7. Recommended action for management.
    """
    line_item = anomaly["Line Item"]
    status = anomaly["Status"]
    actual_pct = anomaly["Actual %"]
    target_range = anomaly["Target Range"]
    dev = anomaly["Deviation (pp)"]
    day_id = anomaly["Day"]

    if anomaly["Type"] == "Expense Anomaly":
        direction = "exceeded the maximum threshold" if status == "ABOVE TARGET" else "fell below the minimum threshold"
        what_happened = (
            f"On Day {day_id}, {line_item} expenses stood at {actual_pct:.2f}% of daily revenue, "
            f"which {direction} of {target_range} by {dev:.2f} percentage points."
        )
        
        kb_entry = CONTRIBUTING_FACTORS_KB.get(line_item, {})
        factors_list = kb_entry.get(status, ["Operational volume variance or timing difference in expense billing."])
        contributing_factors = "; ".join(factors_list[:2])

        if status == "ABOVE TARGET":
            action = kb_entry.get("RECOMMENDATION_HIGH", f"Investigate root cause of {line_item} overspend and tighten operational controls.")
        else:
            action = kb_entry.get("RECOMMENDATION_LOW", f"Verify {line_item} booking completeness to ensure expenses are not under-recorded.")

    elif anomaly["Type"] == "CM1 Deviation":
        if status == "BELOW TARGET":
            what_happened = (
                f"On Day {day_id}, CM1 was recorded at {actual_pct:.2f}%, breaching the lower threshold of {target_range} "
                f"by {dev:.2f} percentage points (well beyond the 3.0 pp allowance)."
            )
            contributing_factors = (
                "Direct fulfillment costs (Manpower and/or Packaging) were disproportionately high relative to revenue throughput, "
                "compressing gross operational contribution."
            )
            action = "Immediately audit floor productivity (Units Per Hour) and packaging material consumption rates."
        else:
            what_happened = (
                f"On Day {day_id}, CM1 was exceptionally strong at {actual_pct:.2f}%, exceeding the upper target of {target_range} "
                f"by {dev:.2f} percentage points."
            )
            contributing_factors = (
                "High revenue surge with lean direct staffing, or possible lag in logging contractual labor and packaging invoices."
            )
            action = "Verify completeness of labor and packaging ledger entries before booking as true margin gain."

    elif anomaly["Type"] == "CM2 Deviation":
        if status == "BELOW TARGET":
            what_happened = (
                f"On Day {day_id}, CM2 collapsed to {actual_pct:.2f}%, critically missing the expected {target_range} target "
                f"by {dev:.2f} percentage points."
            )
            contributing_factors = (
                "Cumulative operating cost leakage across fixed and semi-variable lines (FC Rent, Utilities, Equipment, Overheads) "
                "severely eroded net contribution."
            )
            action = "Conduct an immediate P&L review; identify top variance expense lines and enforce cost rationalization measures."
        else:
            what_happened = (
                f"On Day {day_id}, CM2 registered at {actual_pct:.2f}%, surpassing the upper target {target_range} by {dev:.2f} percentage points."
            )
            contributing_factors = "High operational leverage during volume peak, or unrecorded operating overheads."
            action = "Verify all facility rental, utility, and maintenance provisions are accurately reflected."

    else:
        what_happened = f"On Day {day_id}, an operational variance was flagged on {line_item}."
        contributing_factors = "Fluctuations in daily operational volume and expense booking cycles."
        action = "Review line item ledger entries and operational logs."

    return {
        "Day": str(day_id),
        "Affected Metric": line_item,
        "What Happened": what_happened,
        "Actual Value": f"{actual_pct:.2f}%",
        "Target Range": target_range,
        "Deviation": f"{dev:.2f} pp",
        "Possible Contributing Factors": contributing_factors,
        "Recommended Action": action,
    }

def generate_suspicious_cm2_insight(day_record: Dict[str, Any]) -> Dict[str, str]:
    """
    Dedicated insight for Suspiciously High CM2 (> 33%).
    Direct assignment requirement:
    'Do not interpret suspiciously high CM2 as automatically excellent performance.
     Explain that some expenses may be missing or underreported.
     Recommend verifying financial records.'
    """
    day_id = day_record["Day"]
    cm2_pct = float(day_record["CM2%"])
    dev = round(cm2_pct - 30.0, 2)

    what_happened = (
        f"On Day {day_id}, CM2 was recorded at an unusually elevated level of {cm2_pct:.2f}%, "
        f"exceeding the standard healthy target of 15.0% - 30.0% by {dev:.2f} percentage points."
    )
    factors = (
        "This abnormal margin is likely artificial: one or more major operational cost categories "
        "(such as Manpower, FC Rent, Power & Fuel, or Overheads) may be missing, omitted, or underreported in the uploaded ledger."
    )
    action = (
        "DO NOT treat this as genuine operational excellence. Audit the day's expense ledger entries with the finance controller; "
        "verify that utility, rental accrual, and outsourced labor bills have been completely accounted for before sign-off."
    )

    return {
        "Day": str(day_id),
        "Affected Metric": "CM2 (Suspiciously High)",
        "What Happened": what_happened,
        "Actual Value": f"{cm2_pct:.2f}%",
        "Target Range": "15.0% - 30.0%",
        "Deviation": f"{dev:.2f} pp",
        "Possible Contributing Factors": factors,
        "Recommended Action": action,
    }

def call_gemini_api_if_available(prompt_text: str) -> Optional[str]:
    """
    Attempts to call Google Gemini API if GEMINI_API_KEY or GOOGLE_API_KEY is present in env.
    Returns None on failure or if key is absent (graceful fallback).
    """
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        return None

    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
        payload = {
            "contents": [{"parts": [{"text": prompt_text}]}],
            "generationConfig": {"temperature": 0.2, "maxOutputTokens": 600}
        }
        resp = requests.post(url, json=payload, timeout=8)
        if resp.status_code == 200:
            data = resp.json()
            candidates = data.get("candidates", [])
            if candidates:
                text = candidates[0].get("content", {}).get("parts", [{}])[0].get("text")
                return text
    except Exception:
        pass
    return None

def generate_insights_for_day(day_record: Dict[str, Any], day_analysis: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Generate comprehensive insights for all issues identified on a given day.
    """
    insights = []
    
    # Check if suspicious CM2
    if day_analysis.get("Is Suspicious"):
        suspicious_insight = generate_suspicious_cm2_insight(day_record)
        insights.append(suspicious_insight)

    # Individual anomalies
    for anomaly in day_analysis.get("All Anomalies", []):
        ins = generate_rule_based_anomaly_insight(anomaly, day_record)
        insights.append(ins)

    return insights
