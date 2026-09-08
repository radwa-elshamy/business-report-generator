[README.md](https://github.com/user-attachments/files/31952676/README.md)
# 📊 Business Report Generator

> Turn raw data into boardroom-ready reports in minutes.

A desktop app that transforms any Excel or CSV file into a professional business report with charts, KPIs, and a polished layout — automatically.

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Output](https://img.shields.io/badge/Output-Excel%20%7C%20Charts-green)
![Data](https://img.shields.io/badge/Input-Excel%20%7C%20CSV-orange)

---

## 📸 What It Looks Like

```
┌────────────────────────────────────────────────────────┐
│  📊 Business Report Generator                          │
├──────────────────┬─────────────────────────────────────┤
│  📁 Data Source  │  📊 Data Preview                    │
│  ✅ sales.csv    │  Month  │ Product  │ Revenue │ Units│
│  (24 rows)       │  ───────┼──────────┼─────────┼──────│
│                  │  Jan    │ Widget A │ 15,000  │ 120  │
│  Company: ACME   │  Jan    │ Widget B │ 12,000  │ 95   │
│  Title: Q1 Sales │  Feb    │ Widget A │ 17,000  │ 135  │
│  Period: Q1 2025 │  ...                               │
│  Type: Sales Rep.│                                     │
│                  │  📈 Charts Preview                  │
│  Label: product  │  [Bar chart: Revenue by Product]    │
│  Value: revenue  │                                     │
│  Chart: Bar      │  Widget A ████████████ $111,500    │
│                  │  Widget B ████████     $82,000     │
│  [⚡ Generate]   │  Gadget X ██████       $64,000     │
│  [💾 Save Excel] │  Gadget Y ████         $47,500     │
└──────────────────┴─────────────────────────────────────┘
```

---

## ✨ Features

- **Any data** — Excel (.xlsx) or CSV files
- **Auto-detect** numeric vs category columns
- **6 report types** — Sales, HR, Finance, Marketing, Inventory, Customer
- **5 chart types** — Bar, Line, Pie, Horizontal Bar, Area
- **KPI metrics** — Total, Average, Min, Max, Growth %
- **Professional Excel output** — summary sheet + raw data sheet
- **Formatted charts** embedded in the Excel report
- **Company branding** — name, date range, report title
- **Data preview** — see your data before generating

---

## 🚀 Quick Start

```bash
pip install customtkinter openpyxl matplotlib
python report_generator.py
```

### How to use:
1. Click **Load Data** → select your Excel or CSV file
2. Fill in Company Name, Report Title, Date Range
3. Choose your Label column (X axis) and Value column (Y axis)
4. Select chart type
5. Click **⚡ Generate Report**
6. Click **💾 Save as Excel**

---

## 📋 Sample Data Included

The `sample_data/` folder includes:

### sales_data.csv
```
month,product,category,revenue,units_sold,region
Jan,Widget A,Electronics,15000,120,Cairo
Jan,Widget B,Electronics,12000,95,Alexandria
Feb,Widget A,Electronics,17000,135,Cairo
...
```

### employees.csv
```
name,department,salary,performance_score,years
Ahmed,Engineering,18000,4.5,3
Sara,Marketing,14000,4.2,2
...
```

---

## 📊 What the Generated Report Contains

### Sheet 1: Executive Summary
- Company name and report title
- Date range and generation timestamp
- **KPI metrics table:**

| Metric | Revenue | Units Sold |
|--------|---------|------------|
| Count | 24 | 24 |
| Total | $305,000 | 4,221 |
| Average | $12,708 | 175 |
| Min | $6,500 | 88 |
| Max | $23,000 | 310 |
| Growth % | +53% | +42% |

- Embedded charts (bar + pie)

### Sheet 2: Raw Data
- Full data with styled headers
- Auto-width columns
- Filters and freeze panes

---

## 📁 File Structure

```
business-report-generator/
├── report_generator.py     # Main app
├── sample_data/
│   ├── sales_data.csv      # Sample sales data (24 rows)
│   └── employees.csv       # Sample HR data (10 rows)
├── report_charts/          # Auto-created, stores chart images
├── requirements.txt
└── README.md
```

---

## 🛠 Requirements

```
customtkinter>=5.2.0
openpyxl>=3.1.0
matplotlib>=3.7.0
Pillow>=10.0.0  # for chart preview in app
```

---

## 💡 Use Cases

| Industry | Data | Report |
|----------|------|--------|
| Retail | Sales CSV from POS | Monthly sales report |
| HR | Employee spreadsheet | Performance report |
| Finance | Transaction data | Financial summary |
| Marketing | Campaign results | Analytics report |
| E-commerce | Orders export | Revenue report |

---

## 📞 Custom Reports
Need a report built from your data? Order on Fiverr: [your-link]
