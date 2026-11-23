# PROZO – AI-POWERED DAILY FC P&L MONITOR & ANOMALY DETECTION SYSTEM

An autonomous, end-to-end financial intelligence platform designed for **Prozo Fulfillment Centres (FCs)**. The application automates daily P&L ingestion, cost-ratio calculation, anomaly detection against ideal operational corridors, executive AI insight generation, multi-day trend analysis, and one-click production report downloads.

---

## 🚀 Key Highlights & User Journey

The system requires **zero manual calculations or Excel modifications**:

```
ONE EXCEL WORKBOOK UPLOAD ➡️ VALIDATE BOTH SHEETS ➡️ GENERATE OUTPUT ➡️ VIEW DASHBOARD ➡️ DOWNLOAD ALL REPORTS
```

1. **Step 1: Upload One Single Excel Workbook (.xlsx)**
   - **Single Uploader:** `Upload FC Financial Data (.xlsx)`
   - **Worksheet 1: `Ideal FC Criteria`:** Contains expense categories, min/max percentage corridors, and CM1/CM2 margins.
   - **Worksheet 2: `Day Wise PnL`:** Contains Day/Date, Revenue, Manpower, Packaging, Power and Fuel, FC Rent, Equipment Rentals, and Overheads.
   - *Intelligent Ingestion:* Tested directly on `Case Study 3.xlsx` reference structure; tolerant to non-zero header rows, column name variations, and missing FC identifiers (switches gracefully to Day-wise analysis).
2. **Step 2: Prominent "GENERATE OUTPUT" Execution**
   - Transparent 9-stage pipeline progress tracking.
3. **Step 3: Deterministic Financial Engine**
   - Calculates Expense %, CM1, CM1%, CM2, CM2%, and FC status colors (RED, YELLOW, GREEN, SUSPICIOUS, REVIEW REQUIRED).
   - Zero LLM math hallucination.
4. **Step 4: Anomaly Detection Engine**
   - Identifies line-item breaches, CM1/CM2 corridor breaches (> 3 percentage points), calculates Rupee impact and severity.
5. **Step 5: Executive AI Insights**
   - 6-part plain-language analytical breakdown:
     1. What happened?
     2. Size of deviation
     3. Breached line items
     4. Top expense contributors sorted by actual Rupee impact
     5. Why the issue may have occurred (Confirmed vs. Possible causes)
     6. Recommended management action
6. **Step 6: Production-Grade Downloadable Deliverables**
   - **File 1:** `Complete_FC_PnL_Report.xlsx` (8 distinct, professionally styled worksheets)
   - **File 2:** `Anomaly_Report.xlsx`
   - **File 3:** `AI_Insights_Report.pdf`
   - **File 4:** `Daily_FC_Digest.pdf`
   - **File 5:** `Trend_Analysis.pdf`
   - **File 6:** `All_Reports.zip` (consolidated archive containing all reports)

---

## 📐 Mathematical Formulas & Business Rules

### Expense Percentages
$$\text{Expense } \% = \left(\frac{\text{Expense Amount}}{\text{Revenue}}\right) \times 100$$

### Contribution Margins
$$\text{CM1} = \text{Revenue} - \text{Manpower Cost} - \text{Packaging Cost}$$
$$\text{CM1 } \% = \left(\frac{\text{CM1}}{\text{Revenue}}\right) \times 100$$

$$\text{CM2} = \text{CM1} - \text{Power and Fuel} - \text{FC Rent} - \text{Equipment Rentals} - \text{Overheads}$$
$$\text{CM2 } \% = \left(\frac{\text{CM2}}{\text{Revenue}}\right) \times 100$$

### Ideal Expense Targets (% of Revenue)
- **Manpower:** 25.0% – 27.0%
- **Packaging:** 5.0% – 7.0%
- **Power and Fuel:** 10.0% – 12.0%
- **FC Rent:** 15.0% – 18.0%
- **Equipment Rentals:** 5.0% – 8.0%
- **Overheads:** 10.0% – 12.0%
- **CM1 Corridor:** 65.0% – 70.0%
- **CM2 Corridor:** 15.0% – 30.0%

### FC Colour Classification Logic
- 🔴 **RED:** $\text{CM2} < 12.0\%$ (Critical Profitability Breach)
- 🟡 **YELLOW:** $12.0\% \le \text{CM2} \le 14.9\%$ OR $30.0\% < \text{CM2} \le 32.9\%$ (Sub-optimal or High Caution)
- 🟢 **GREEN:** $15.0\% \le \text{CM2} \le 30.0\%$ (Healthy Target Performance)
- 🟣 **SUSPICIOUS:** $\text{CM2} > 33.0\%$ (High probability of unposted / underreported accruals)
- ⚪ **REVIEW REQUIRED:** $14.9\% < \text{CM2} < 15.0\%$ OR $32.9\% < \text{CM2} \le 33.0\%$ (Boundary gaps)

> **Key Rule:** Individual expense status and overall CM2 status are kept separate. A day can have an overall **GREEN** status while still flagging an individual expense anomaly.

---

## 🧪 Certified Test Scenarios

The system includes pre-configured, mathematically certified synthetic test cases accessible directly from the application sidebar:

1. **Test 1: Clean Day**
   - All expenses within target corridor. CM1 = 68.0%, CM2 = 24.0%. **GREEN** status, 0 anomalies.
2. **Test 2: CM2 Breach**
   - CM2 = 10.5% (< 12.0%). **RED** status. Multiple cost overruns flagged in Manpower and Power & Fuel.
3. **Test 3: Line-item-only Breach**
   - Packaging is 8.5% (> 7.0% ceiling). CM1 = 66.0%, CM2 = 23.5%. Overall status remains **GREEN** while isolating the packaging variance.
4. **Test 4: Suspiciously High CM2**
   - CM2 = 58.5% (> 33.0%). **SUSPICIOUS** status. Triggers warnings for missing rent and utility expenses.
5. **Multi-Day 3-FC Benchmark**
   - 7 days of operational data across Bhiwandi, Farukhnagar, and Hoskote FCs.

---

## 🔮 Future Enhancements & System Evolution

As documented in Requirement 13, here is how the system is architected to scale into an autonomous enterprise intelligence engine:

1. **Persistent Historical Memory & Columnar Lakehouse:**
   - Integrate a persistent time-series datastore (e.g. ClickHouse or TimescaleDB) to accumulate daily uploads into longitudinal historical records.
   - Automatically compute rolling 30, 60, and 90-day seasonal baselines to isolate systemic cost creeps from one-off volume surges.
2. **Manager Feedback & Learning Loop:**
   - Provide an in-app review workflow allowing FC Managers to annotate anomalies (e.g. *"Overtime approved for Amazon Prime Day volume surge"*).
   - Use manager confirmations to auto-tune tolerance thresholds and suppress routine operational variances.
3. **Automated Notification Microservices:**
   - Webhook integrations for Slack channels (`#fc-operations-audit`), WhatsApp Business API for regional GMs, and automated executive email digests with attached PDFs.
4. **Intra-Day Predictive Forecasting:**
   - Connect live WMS telemetry to forecast end-of-day CM2 by 2:00 PM, allowing shift supervisors to adjust labor rosters before cost overruns occur.

---

## 💻 Installation & Running Locally

### Prerequisites
- Python 3.10+
- Dependencies: `streamlit`, `pandas`, `openpyxl`, `reportlab`, `matplotlib`

### Launch the Application
```bash
# Navigate to project directory
cd C:\Users\Lenovo\.gemini\antigravity\scratch\prozo_pnl_monitor

# Run Streamlit Web Application
streamlit run app.py
```

Open `http://localhost:8501` in your browser.
