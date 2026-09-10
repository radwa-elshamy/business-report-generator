[README.md](https://github.com/user-attachments/files/32056173/README.md)
# 📊 Sales Data Analysis Dashboard — Excel + Python

A complete sales data analysis pipeline: clean raw data with Python, then generate a professional interactive Excel dashboard with KPIs, charts, and an executive summary.

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Excel](https://img.shields.io/badge/Output-Excel%20Dashboard-green)
![Arabic](https://img.shields.io/badge/Language-Arabic%20%2B%20English-orange)

---

## 📸 Dashboard Preview

```
┌────────────────────────────────────────────────────────────────┐
│         تقرير تحليل المبيعات السنوي 2024                      │
│         إجمالي 490 طلب  |  تم الإنشاء: 2025-01-15             │
├──────────┬──────────┬──────────┬──────────┬──────────┬────────┤
│ 💰       │ 📈       │ 🎯       │ 📋       │ 🛒       │ 📦     │
│ إجمالي   │ إجمالي  │ هامش     │ عدد      │ متوسط    │ إجمالي │
│ الإيراد  │ الأرباح │ الربح    │ الطلبات  │ الطلب    │الوحدات │
│ 2.1M ج   │ 750K ج  │ 35.2%    │ 490      │ 4,285 ج  │ 2,847  │
├──────────┴──────────┴──────────┴──────────┴──────────┴────────┤
│  [Line Chart: Monthly Revenue & Profit Trend]                  │
│   Jan ▁▂▃▄▅▆▇█ Dec                                           │
├────────────────────────────────────────────────────────────────┤
│  🥇 آيفون 15    │  🥈 كاميرا Canon  │  🥉 لابتوب Dell         │
│  520K جنيه      │  485K جنيه        │  420K جنيه              │
└────────────────────────────────────────────────────────────────┘
```

---

## ✨ What's Included

### 📋 Sheet 1: Dashboard (Main)
- **6 KPI Cards** — Total Revenue, Total Profit, Profit Margin %, Orders, Avg Order Value, Total Units
- **Monthly Line Chart** — Revenue & Profit trend over 12 months
- **Top Products Bar Chart** — Revenue by product with medals 🥇🥈🥉
- Styled with professional colors, Arabic RTL layout

### 📊 Sheet 2: Clean Data
- 490 records after cleaning (removed 18 dirty rows)
- Auto-filter on all columns
- Frozen header row
- Alternating row colors
- Number formatting (currency, percentages)

### 📈 Sheet 3: Analysis
- **By Region** — Revenue & profit per governorate + Pie Chart
- **By Sales Channel** — Online / Branches / Distributors / Phone + Bar Chart
- **By Quarter** — Q1/Q2/Q3/Q4 comparison
- **By Category** — Product category breakdown

### 📝 Sheet 4: Executive Summary
- KPIs summary table
- Top 3 products, regions, channels
- Auto-generated recommendations
- Ready to share with management

---

## 🚀 Quick Start

```bash
pip install pandas openpyxl
python generate_dashboard.py
```

This reads `sales_data.csv` and generates `sales_dashboard.xlsx` automatically.

---

## 📁 File Structure

```
sales-dashboard-excel/
├── generate_dashboard.py    # Main script — generates full Excel dashboard
├── sales_data_sample.csv    # 500-row sample dataset (Arabic)
├── requirements.txt
└── README.md
```

---

## 📋 Sample Data Columns

| Column | Description |
|--------|-------------|
| رقم_الطلب | Order ID |
| التاريخ | Order date |
| المنتج | Product name |
| الفئة | Product category |
| المنطقة | Region/Governorate |
| قناة_البيع | Sales channel |
| الكمية | Quantity |
| الإيراد | Revenue |
| الربح | Profit |
| هامش_الربح_% | Profit margin % |

---

## 🛠 Requirements

```
pandas>=2.0.0
openpyxl>=3.1.0
```

---

## 💡 What Gets Cleaned

| Issue | How Fixed |
|-------|-----------|
| Duplicate orders | Removed by order ID |
| Empty quantities | Rows removed |
| Blank regions | Rows removed |
| Number formatting | Rounded to 2 decimals |
| Profit margin | Calculated and added |

---

## 🛒 Get the Full Version
💰 Order a custom dashboard built from YOUR data: [Fiverr link]
📧 Contact: radwa.elshamy16@gmail.com
