"""
Automated Business Report Generator
Generate professional PDF/Word reports from Excel/CSV data

Requirements:
    pip install customtkinter openpyxl matplotlib python-docx
"""

import customtkinter as ctk
import threading
import csv
import json
import os
import re
from datetime import datetime
from tkinter import messagebox, filedialog, colorchooser

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

REPORT_TYPES = [
    "Sales & Revenue Report",
    "Employee Performance Report",
    "Financial Summary Report",
    "Marketing Analytics Report",
    "Inventory Report",
    "Customer Analysis Report",
    "Custom Report",
]

CHART_TYPES = ["Bar Chart", "Line Chart", "Pie Chart", "Horizontal Bar", "Area Chart"]


def read_data(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        with open(filepath, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            return list(reader), reader.fieldnames or []
    elif ext in (".xlsx", ".xls") and HAS_OPENPYXL:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return [], []
        headers = [str(h) if h else f"Col{i}" for i, h in enumerate(rows[0])]
        data = [{headers[i]: str(row[i]) if row[i] is not None else "" for i in range(len(headers))} for row in rows[1:]]
        return data, headers
    return [], []


def safe_float(val):
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return None


def compute_stats(data, col):
    vals = [safe_float(row.get(col)) for row in data]
    vals = [v for v in vals if v is not None]
    if not vals:
        return {}
    n = len(vals)
    total = sum(vals)
    mean = total / n
    return {
        "Count": n, "Total": round(total, 2),
        "Average": round(mean, 2), "Min": round(min(vals), 2),
        "Max": round(max(vals), 2),
        "Growth %": round(((vals[-1] - vals[0]) / vals[0] * 100), 2) if len(vals) > 1 and vals[0] != 0 else 0,
    }


def generate_chart(data, label_col, value_col, chart_type, title, filepath, color="#2E86AB"):
    if not HAS_MATPLOTLIB or not data:
        return False
    try:
        labels = [str(row.get(label_col, ""))[:20] for row in data[:15]]
        values = [safe_float(row.get(value_col)) or 0 for row in data[:15]]

        fig, ax = plt.subplots(figsize=(10, 5), facecolor="#1a1a2e")
        ax.set_facecolor("#16213e")
        ax.tick_params(colors="white")
        ax.title.set_color("white")
        for spine in ax.spines.values():
            spine.set_edgecolor("#333")

        colors_list = plt.cm.Blues([0.4 + 0.6 * i / len(values) for i in range(len(values))])

        if chart_type == "Bar Chart":
            bars = ax.bar(range(len(labels)), values, color=colors_list)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=45, ha="right", color="white", fontsize=8)
            for bar, val in zip(bars, values):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                        f"{val:,.0f}", ha="center", va="bottom", fontsize=8, color="white")
        elif chart_type == "Line Chart":
            ax.plot(range(len(labels)), values, marker="o", linewidth=2, color=color, markersize=6)
            ax.fill_between(range(len(labels)), values, alpha=0.2, color=color)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=45, ha="right", color="white", fontsize=8)
        elif chart_type == "Pie Chart":
            wedges, texts, autotexts = ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
            for text in texts:
                text.set_color("white")
        elif chart_type == "Horizontal Bar":
            ax.barh(range(len(labels)), values, color=colors_list)
            ax.set_yticks(range(len(labels)))
            ax.set_yticklabels(labels, color="white", fontsize=9)

        ax.set_title(title, color="white", fontsize=14, fontweight="bold", pad=15)
        ax.set_xlabel(label_col, color="#aaa")
        ax.set_ylabel(value_col, color="#aaa")
        ax.yaxis.label.set_color("#aaa")

        plt.tight_layout()
        plt.savefig(filepath, dpi=150, bbox_inches="tight", facecolor="#1a1a2e")
        plt.close()
        return True
    except Exception as e:
        return False


def generate_excel_report(data, headers, stats_map, charts_dir, config, output_path):
    if not HAS_OPENPYXL:
        return

    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "Executive Summary"

    company = config.get("company", "Company Name")
    report_type = config.get("report_type", "Business Report")
    date_range = config.get("date_range", datetime.now().strftime("%B %Y"))

    # Title section
    ws_summary.merge_cells("A1:H1")
    title_cell = ws_summary["A1"]
    title_cell.value = f"{company} — {report_type}"
    title_cell.font = Font(size=18, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="1A5276")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[1].height = 45

    ws_summary.merge_cells("A2:H2")
    subtitle = ws_summary["A2"]
    subtitle.value = f"Period: {date_range}  |  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    subtitle.font = Font(size=11, color="7F8C8D")
    subtitle.alignment = Alignment(horizontal="center")
    ws_summary.row_dimensions[2].height = 20

    # KPI boxes
    row = 4
    ws_summary.cell(row=row, column=1, value="KEY METRICS").font = Font(size=13, bold=True, color="1A5276")
    ws_summary.row_dimensions[row].height = 25
    row += 1

    kpi_fill = PatternFill("solid", fgColor="EBF5FB")
    border = Border(left=Side("thin", color="AED6F1"), right=Side("thin", color="AED6F1"),
                    top=Side("thin", color="AED6F1"), bottom=Side("thin", color="AED6F1"))

    col = 1
    for metric_col, stats in stats_map.items():
        ws_summary.cell(row=row, column=col, value=metric_col).font = Font(bold=True, size=11)
        ws_summary.cell(row=row, column=col).fill = PatternFill("solid", fgColor="1A5276")
        ws_summary.cell(row=row, column=col).font = Font(bold=True, color="FFFFFF")
        row2 = row + 1
        for stat_name, stat_val in stats.items():
            ws_summary.cell(row=row2, column=col, value=stat_name).font = Font(bold=True, size=10)
            ws_summary.cell(row=row2, column=col + 1, value=stat_val)
            ws_summary.cell(row=row2, column=col).border = border
            ws_summary.cell(row=row2, column=col + 1).border = border
            row2 += 1
        col += 3

    # Charts
    if HAS_MATPLOTLIB:
        img_row = row + 3
        for fname in os.listdir(charts_dir):
            if fname.endswith(".png"):
                try:
                    from openpyxl.drawing.image import Image as XLImg
                    img = XLImg(os.path.join(charts_dir, fname))
                    img.width = 600
                    img.height = 350
                    ws_summary.add_image(img, f"A{img_row}")
                    img_row += 22
                except Exception:
                    pass

    # Data sheet
    ws_data = wb.create_sheet("Raw Data")
    h_font = Font(bold=True, color="FFFFFF")
    h_fill = PatternFill("solid", fgColor="2C3E50")
    for ci, h in enumerate(headers, 1):
        cell = ws_data.cell(row=1, column=ci, value=h)
        cell.font = h_font
        cell.fill = h_fill
        cell.alignment = Alignment(horizontal="center")

    alt = PatternFill("solid", fgColor="F2F3F4")
    for ri, row_d in enumerate(data, 2):
        for ci, h in enumerate(headers, 1):
            cell = ws_data.cell(row=ri, column=ci, value=row_d.get(h, ""))
            if ri % 2 == 0:
                cell.fill = alt

    for ci, h in enumerate(headers, 1):
        max_len = max(len(str(h)), *(len(str(r.get(h, ""))) for r in data[:50]))
        ws_data.column_dimensions[get_column_letter(ci)].width = min(max_len + 3, 35)

    ws_data.auto_filter.ref = ws_data.dimensions
    ws_data.freeze_panes = "A2"

    wb.save(output_path)


class ReportGeneratorApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("📊 Business Report Generator")
        self.geometry("1100x750")
        self.resizable(True, True)

        self.data = []
        self.headers = []
        self.filepath = ""
        self.charts = []

        self._build_ui()

    def _build_ui(self):
        header = ctk.CTkFrame(self, fg_color="#0D1B2A", corner_radius=0, height=65)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="📊 Business Report Generator",
                     font=ctk.CTkFont(size=20, weight="bold"), text_color="#F0B429").pack(side="left", padx=20, pady=16)
        ctk.CTkLabel(header, text="Turn raw data into professional reports",
                     font=ctk.CTkFont(size=12), text_color="#7F8C8D").pack(side="right", padx=20)

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=12)

        left = ctk.CTkFrame(main, width=320, corner_radius=12)
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)

        ls = ctk.CTkScrollableFrame(left, corner_radius=0, fg_color="transparent")
        ls.pack(fill="both", expand=True)

        ctk.CTkLabel(ls, text="📁 Data Source", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(12, 4), padx=16, anchor="w")

        self.file_label = ctk.CTkLabel(ls, text="No file loaded", text_color="#7F8C8D",
                                        font=ctk.CTkFont(size=11), wraplength=260)
        self.file_label.pack(padx=16, anchor="w")
        ctk.CTkButton(ls, text="📂 Load Data (Excel/CSV)", command=self._load_file, height=38).pack(padx=16, fill="x")

        ctk.CTkLabel(ls, text="─" * 36, text_color="#333").pack(padx=16, pady=8)
        ctk.CTkLabel(ls, text="🏢 Report Details", font=ctk.CTkFont(size=14, weight="bold")).pack(padx=16, anchor="w")

        for label, attr, placeholder in [
            ("Company Name", "company_entry", "My Company"),
            ("Report Title", "title_entry", "Q1 Sales Report"),
            ("Date Range", "date_entry", "January - March 2025"),
        ]:
            ctk.CTkLabel(ls, text=label, font=ctk.CTkFont(size=12)).pack(pady=(8, 2), padx=16, anchor="w")
            entry = ctk.CTkEntry(ls, placeholder_text=placeholder, height=34)
            entry.pack(padx=16, fill="x")
            setattr(self, attr, entry)

        ctk.CTkLabel(ls, text="Report Type", font=ctk.CTkFont(size=12)).pack(pady=(8, 2), padx=16, anchor="w")
        self.report_type_var = ctk.StringVar(value=REPORT_TYPES[0])
        ctk.CTkOptionMenu(ls, values=REPORT_TYPES, variable=self.report_type_var).pack(padx=16, fill="x")

        ctk.CTkLabel(ls, text="─" * 36, text_color="#333").pack(padx=16, pady=8)
        ctk.CTkLabel(ls, text="📊 Charts", font=ctk.CTkFont(size=14, weight="bold")).pack(padx=16, anchor="w")

        ctk.CTkLabel(ls, text="Label Column (X axis)", font=ctk.CTkFont(size=12)).pack(pady=(8, 2), padx=16, anchor="w")
        self.label_col_var = ctk.StringVar(value="")
        self.label_col_menu = ctk.CTkOptionMenu(ls, values=["Load data first"], variable=self.label_col_var)
        self.label_col_menu.pack(padx=16, fill="x")

        ctk.CTkLabel(ls, text="Value Column (Y axis)", font=ctk.CTkFont(size=12)).pack(pady=(8, 2), padx=16, anchor="w")
        self.value_col_var = ctk.StringVar(value="")
        self.value_col_menu = ctk.CTkOptionMenu(ls, values=["Load data first"], variable=self.value_col_var)
        self.value_col_menu.pack(padx=16, fill="x")

        ctk.CTkLabel(ls, text="Chart Type", font=ctk.CTkFont(size=12)).pack(pady=(8, 2), padx=16, anchor="w")
        self.chart_type_var = ctk.StringVar(value="Bar Chart")
        ctk.CTkOptionMenu(ls, values=CHART_TYPES, variable=self.chart_type_var).pack(padx=16, fill="x")

        ctk.CTkLabel(ls, text="─" * 36, text_color="#333").pack(padx=16, pady=8)

        ctk.CTkButton(ls, text="⚡ Generate Report", command=self._generate_report,
                      height=44, font=ctk.CTkFont(size=14, weight="bold"),
                      fg_color="#F0B429", text_color="black", hover_color="#D4A017").pack(padx=16, fill="x")

        ctk.CTkButton(ls, text="💾 Save as Excel", command=lambda: self._save_report("excel"),
                      height=36, fg_color="#145A32").pack(padx=16, pady=(8, 4), fill="x")
        ctk.CTkButton(ls, text="📄 Save as Word", command=lambda: self._save_report("word"),
                      height=36, fg_color="#1A5276").pack(padx=16, pady=(0, 16), fill="x")

        right = ctk.CTkFrame(main, corner_radius=12)
        right.pack(side="right", fill="both", expand=True)

        tabs = ctk.CTkTabview(right)
        tabs.pack(fill="both", expand=True, padx=12, pady=12)
        tabs.add("📊 Data Preview")
        tabs.add("📈 Charts Preview")
        tabs.add("📋 Report Summary")
        tabs.add("📋 Log")

        self.data_preview = ctk.CTkScrollableFrame(tabs.tab("📊 Data Preview"), corner_radius=6)
        self.data_preview.pack(fill="both", expand=True)

        self.charts_preview = ctk.CTkScrollableFrame(tabs.tab("📈 Charts Preview"), corner_radius=6)
        self.charts_preview.pack(fill="both", expand=True)

        self.summary_text = ctk.CTkTextbox(tabs.tab("📋 Report Summary"), font=ctk.CTkFont(family="Courier", size=12))
        self.summary_text.pack(fill="both", expand=True)

        self.log_text = ctk.CTkTextbox(tabs.tab("📋 Log"), font=ctk.CTkFont(family="Courier", size=11))
        self.log_text.pack(fill="both", expand=True)

        self.status_label = ctk.CTkLabel(right, text="Load data to begin",
                                          font=ctk.CTkFont(size=11), text_color="#7F8C8D")
        self.status_label.pack(pady=(0, 8))

        self.tabs_ref = tabs

    def _load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Data files", "*.csv *.xlsx *.xls"), ("All", "*.*")])
        if not filepath:
            return
        self._log(f"Loading: {filepath}")
        data, headers = read_data(filepath)
        if not data:
            messagebox.showerror("Error", "Could not read file or file is empty.")
            return

        self.data = data
        self.headers = headers
        self.filepath = filepath
        self.file_label.configure(text=f"✅ {os.path.basename(filepath)} ({len(data)} rows)", text_color="#2ECC71")

        self.label_col_menu.configure(values=headers)
        self.value_col_menu.configure(values=headers)
        if headers:
            self.label_col_var.set(headers[0])
        if len(headers) > 1:
            self.value_col_var.set(headers[1])

        self._render_data_preview()
        self._log(f"✅ Loaded {len(data)} rows, {len(headers)} columns")
        self.status_label.configure(text=f"✅ {len(data)} rows loaded")

    def _render_data_preview(self):
        for w in self.data_preview.winfo_children():
            w.destroy()

        if not self.data:
            return

        # Header row
        h_row = ctk.CTkFrame(self.data_preview, fg_color="#1A5276", corner_radius=4)
        h_row.pack(fill="x", pady=1)
        for h in self.headers[:6]:
            ctk.CTkLabel(h_row, text=h[:18], font=ctk.CTkFont(size=11, weight="bold"),
                         text_color="white", width=130, anchor="w").pack(side="left", padx=6, pady=5)

        for i, row in enumerate(self.data[:30]):
            r = ctk.CTkFrame(self.data_preview, fg_color="#1C2833" if i % 2 == 0 else "#212F3D", corner_radius=3)
            r.pack(fill="x", pady=1)
            for h in self.headers[:6]:
                ctk.CTkLabel(r, text=str(row.get(h, ""))[:18],
                             font=ctk.CTkFont(size=11), width=130, anchor="w").pack(side="left", padx=6, pady=3)

        if len(self.data) > 30:
            ctk.CTkLabel(self.data_preview, text=f"... and {len(self.data)-30} more rows",
                         text_color="#555").pack(pady=8)

    def _generate_report(self):
        if not self.data:
            messagebox.showerror("Missing", "Load data first.")
            return

        self._log("Generating report...")
        self.status_label.configure(text="Generating...")

        def worker():
            charts_dir = "report_charts"
            os.makedirs(charts_dir, exist_ok=True)

            label_col = self.label_col_var.get()
            value_col = self.value_col_var.get()
            chart_type = self.chart_type_var.get()
            company = self.company_entry.get() or "Company"
            title = self.title_entry.get() or "Business Report"

            # Stats
            numeric_cols = [h for h in self.headers if any(safe_float(r.get(h)) is not None for r in self.data[:10])]
            stats_map = {col: compute_stats(self.data, col) for col in numeric_cols[:3]}

            # Generate chart
            chart_path = os.path.join(charts_dir, "main_chart.png")
            generate_chart(self.data, label_col, value_col, chart_type, f"{value_col} by {label_col}", chart_path)

            # Build summary
            summary = f"REPORT SUMMARY\n{'='*50}\n"
            summary += f"Company: {company}\n"
            summary += f"Report: {title}\n"
            summary += f"Period: {self.date_entry.get() or 'N/A'}\n"
            summary += f"Records: {len(self.data)}\n\n"

            for col, stats in stats_map.items():
                summary += f"📊 {col}:\n"
                for k, v in stats.items():
                    summary += f"   {k}: {v:,}\n" if isinstance(v, (int, float)) else f"   {k}: {v}\n"
                summary += "\n"

            self.after(0, lambda: self._show_results(summary, charts_dir, stats_map))

        threading.Thread(target=worker, daemon=True).start()

    def _show_results(self, summary, charts_dir, stats_map):
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("end", summary)

        for w in self.charts_preview.winfo_children():
            w.destroy()

        for fname in os.listdir(charts_dir):
            if fname.endswith(".png"):
                try:
                    img_path = os.path.join(charts_dir, fname)
                    img = ctk.CTkImage(light_image=__import__('PIL').Image.open(img_path),
                                       dark_image=__import__('PIL').Image.open(img_path), size=(700, 400))
                    ctk.CTkLabel(self.charts_preview, image=img, text="").pack(pady=8)
                except Exception:
                    ctk.CTkLabel(self.charts_preview, text=f"Chart saved: {fname}",
                                 text_color="#2ECC71").pack(pady=8)

        self.stats_map = stats_map
        self.tabs_ref.set("📋 Report Summary")
        self.status_label.configure(text="✅ Report generated!")
        self._log("✅ Report generated!")

    def _save_report(self, fmt):
        if not self.data:
            messagebox.showerror("Missing", "Generate a report first.")
            return

        config = {
            "company": self.company_entry.get() or "Company",
            "report_type": self.report_type_var.get(),
            "date_range": self.date_entry.get() or datetime.now().strftime("%B %Y"),
        }

        if fmt == "excel":
            filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel", "*.xlsx")],
                                                     initialfile=f"report_{datetime.now().strftime('%Y%m%d')}")
            if filepath:
                generate_excel_report(self.data, self.headers, getattr(self, "stats_map", {}),
                                       "report_charts", config, filepath)
                messagebox.showinfo("Saved!", f"Report saved to {filepath}")
        else:
            messagebox.showinfo("Coming soon", "Word export coming in next version. Use Excel for now.")

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n")
        self.log_text.see("end")


if __name__ == "__main__":
    app = ReportGeneratorApp()
    app.mainloop()
