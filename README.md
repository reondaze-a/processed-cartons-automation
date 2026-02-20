# 🏥 Hospital Processed Cartons Automation

## Overview

The **Hospital Processed Cartons Automation** is a production-grade Power Automate + Excel + Office Scripts system designed to eliminate manual logging, enforce data integrity, and create deterministic shift-based reconciliation for warehouse hospital processing operations.

This system automates the capture, calculation, and persistence of processed carton metrics across shifts and brands, converting operational dashboard reports into structured historical records.

---

## 🎯 Purpose

- Eliminate manual entry errors in processed carton logging  
- Separate Start-of-Day (SoD) and End-of-Day (EoD) logic  
- Automatically calculate second-shift processed values  
- Persist brand-level processed metrics directly from dashboard reports  
- Enforce deterministic date-keyed updates  
- Improve auditability and reporting reliability  

---

## ⚙️ System Architecture

### Recurrence-Based Shift Logic

The flow runs on a recurrence trigger at:

- **3:30 PM** → Start of Day (SoD) branch  
- **1:30 AM** → End of Day (EoD) branch  

Timezone-aware evaluation ensures correct branching regardless of DST changes.

Recurrence
↓
Evaluate run hour (EST)
Case 15 → SoD Snapshot Logic
Case 1 → EoD Reconciliation Logic


---

## 🕒 3:30 PM – SoD Snapshot Logic

At 3:30 PM, the flow:

1. Pulls current processed values from the Hospital Dashboard via Office Script.
2. Appends a new row into the `Hospital Cartons Processed` table.
3. Logs:
   - `workDate`
   - `Start of Day (SoD)` processed value

This establishes the baseline for calculating actual second-shift processing.

---

## 🌙 1:30 AM – EoD Reconciliation Logic

At 1:30 AM, the flow:

1. Pulls updated processed totals from the Hospital Dashboard.
2. Locates the existing row using `workDate` as a deterministic key.
3. Updates:
   - `Total` (EoD processed value)
   - `Start of Day – EoD Delta` (true second-shift processed count)

This ensures accurate shift reconciliation without relying on manual subtraction or user input.

---

## 🏷 Brand-Level Automation

During the EoD branch:

- Each brand’s processed value (CBI, Oved, SLV) is appended to the `Brands_Processed` table.
- Data is sourced strictly from the official dashboard report.
- Manual brand-level entry has been fully eliminated.

This reduces operational risk and converts reporting into structured database-backed logging.

---

## 🧠 Key Design Decisions

### Deterministic Date Keying
Excel date serial values are normalized to ensure stable row matching.

### Single-Loop Append Pattern
Array handling is structured to avoid nested `Apply to each` duplication issues.

### Script-Based Data Extraction
Office Scripts extract structured arrays from dashboard sheets, ensuring consistent schema contracts between Excel and Power Automate.

### Data Integrity Enforcement
- No manual input dependency
- No copy-paste override risks
- Controlled update architecture

---

## 📊 Data Model

### `Hospital Cartons Processed`

| Column | Description |
|--------|------------|
| workDate | Operational date key |
| SoD | Snapshot at 3:30 PM |
| Total | Final EoD total |
| SoD-EoD Delta | True second-shift processed |

---

### `Brands_Processed`

| Column | Description |
|--------|------------|
| workDate | Operational date |
| Brand | CBI / Oved / SLV |
| Processed | Brand-specific processed count |
| Timestamp | Logging time |

---

## 🚀 Impact

- Fully automated shift reconciliation
- Zero manual brand-level entry
- Accurate second-shift processed calculation
- Stable recurrence-driven logging
- Reduced operational fragility
- Improved audit traceability

---

## 🛠 Tech Stack

- **Microsoft Power Automate**
- **Excel Online (Business)**
- **Office Scripts (TypeScript)**
- Recurrence triggers
- Array filtering & deterministic branching
- Expression-based numeric reconciliation

---

## 🔮 Future Improvements

- Shift variance anomaly detection
- SLA-based threshold alerts
- Migration to SharePoint List or SQL backend
- Historical analytics dashboard

---

## Screenshots

![SoD Branch](assets/Daily_Processed_Cartons_1.png)

## Author

**Abraham Efraim**  
Automation Engineer – Warehouse Operations  
Power Platform | Office Scripts | Process Engineering

---
