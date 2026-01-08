import pandas as pd
import random, os, calendar
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, PageBreak
)
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT

# =====================================================
# READ CONFIG
# =====================================================
with open("ui_config.txt") as f:
    BATCH_ID = f.readline().strip()
    DEPT_NAME = f.readline().strip()
    COMP_NAME = f.readline().strip()
    START_DATE = f.readline().strip()
    END_DATE = f.readline().strip()

start_date = datetime.strptime(START_DATE, "%d-%m-%Y")
end_date   = datetime.strptime(END_DATE, "%d-%m-%Y")

students_df = pd.read_excel("uploads/students.xlsx")
students = students_df.rename(columns=str.lower).to_dict("records")

# =====================================================
# HELPERS
# =====================================================
def hhmm(minutes):
    return f"{minutes//60:02d}:{minutes%60:02d}"

def weekday(dt):
    return ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][dt.weekday()]

# =====================================================
# MONTH ATTENDANCE GENERATOR
# =====================================================
def generate_month(year, month):
    month_start = datetime(year, month, 1)
    month_end = datetime(year, month, calendar.monthrange(year, month)[1])

    eff_start = max(start_date, month_start)
    eff_end   = min(end_date, month_end)

    days = calendar.monthrange(year, month)[1]

    rows = {"IN": [], "OUT": [], "WORK": [], "BREAK": [], "OT": [], "Status": []}
    weekdays = []

    present = absent = wo = hl = lv = 0
    total_work = total_ot = 0

    valid_days = []
    for d in range(1, days + 1):
        dt = datetime(year, month, d)
        weekdays.append(weekday(dt))
        if eff_start <= dt <= eff_end and dt.weekday() != 6:
            valid_days.append(d)
        if dt.weekday() == 6:
            wo += 1

    absent_days = random.sample(valid_days, min(random.randint(2, 4), len(valid_days)))

    for d in range(1, days + 1):
        dt = datetime(year, month, d)

        if dt < eff_start or dt > eff_end or dt.weekday() == 6:
            for k in rows:
                rows[k].append("---")
            continue

        if d in absent_days:
            rows["IN"].append("---")
            rows["OUT"].append("---")
            rows["WORK"].append("00:00")
            rows["BREAK"].append("00:00")
            rows["OT"].append("00:00")
            rows["Status"].append("A")
            absent += 1
            continue

        in_m = random.randint(5, 40)
        out_m = random.randint(0, 45)
        work = (14 * 60 + out_m) - (9 * 60 + in_m)

        rows["IN"].append(f"09:{in_m:02d}")
        rows["OUT"].append(f"14:{out_m:02d}")
        rows["WORK"].append(hhmm(work))
        rows["BREAK"].append("00:00")
        rows["OT"].append(hhmm(work))
        rows["Status"].append("P")

        present += 1
        total_work += work
        total_ot += work

    df = pd.DataFrame(rows).T
    df.columns = [str(i) for i in range(1, days + 1)]

    summary = {
        "Present": present,
        "Absent": absent,
        "WO": wo,
        "HL": hl,
        "LV": lv,
        "TotalWorkOT": hhmm(total_work),
        "TotalOT": hhmm(total_ot)
    }

    return df, weekdays, summary

# =====================================================
# MONTH LOOP
# =====================================================
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

cur = start_date.replace(day=1)
months = []

while cur <= end_date:
    months.append(cur)
    cur = (cur + timedelta(days=32)).replace(day=1)

# =====================================================
# EXCEL + PDF PER MONTH
# =====================================================
styles = getSampleStyleSheet()
header_style = ParagraphStyle(
    "header",
    fontSize=10,
    leading=14,
    alignment=TA_LEFT,
    spaceAfter=8
)

for m in months:
    
    year, month = m.year, m.month
    month_label = m.strftime("%B-%Y")
    file_prefix = month_label.replace(" ", "_")

    excel_file = f"{OUTPUT_DIR}/{file_prefix}.xlsx"
    pdf_file   = f"{OUTPUT_DIR}/{file_prefix}.pdf"

    # ---------------- EXCEL ----------------
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        row_ptr = 0

        for stu in students:
            df, weekdays, summary = generate_month(year, month)

            header = [
                ["Empcode", "Batch ID:", BATCH_ID, stu["empcode"], "Name", stu["name"],"Month", month_label],
                ["Present", summary["Present"], "Absent", summary["Absent"]],"month", month_label,
                [""] + df.columns.tolist(),
                [""] + weekdays
            ]

            pd.DataFrame(header).to_excel(
                writer, sheet_name=month_label,
                startrow=row_ptr, index=False, header=False
            )

            df.to_excel(
                writer, sheet_name=month_label,
                startrow=row_ptr + 4
            )

            row_ptr += len(df) + 8

    # ---------------- PDF ----------------
    pdf = SimpleDocTemplate(
        pdf_file,
        pagesize=landscape(A3),
        leftMargin=10,
        rightMargin=10,
        topMargin=10,
        bottomMargin=10
    )

    elements = [Paragraph(f"<b>{month_label}</b>", styles["Heading2"])]

    for stu in students:
        month_label = m.strftime("%B-%Y")   # e.g. July-2025
        df, weekdays, summary = generate_month(year, month)
        

        elements.append(Paragraph(

            f"""
            <b>Batch ID:</b> {BATCH_ID} &nbsp;&nbsp;
            <b>Course:</b> {DEPT_NAME} &nbsp;&nbsp;
            <b>Company:</b> {COMP_NAME} &nbsp;&nbsp;
            <b>Empcode:</b> {stu['empcode']} &nbsp;&nbsp;
            <b>Name:</b> {stu['name']}<br/>
            <b>Month:</b> <b>{month_label}</b><br/>
            <b>Present:</b> <font color="green">{summary['Present']}</font> &nbsp;&nbsp;
            <b>Absent:</b> <font color="red">{summary['Absent']}</font> &nbsp;&nbsp;
            <b>Total OT:</b> {summary['TotalOT']}
            """,
            header_style
        ))

        table_data = [[""] + df.columns.tolist(), [""] + weekdays]
        for idx, row in df.iterrows():
            table_data.append([idx] + row.tolist())

        table = Table(table_data, repeatRows=2)

        style_cmds = [
            ("GRID", (0,0), (-1,-1), 0.4, colors.black),
            ("ALIGN", (1,0), (-1,-1), "CENTER"),
            ("BACKGROUND", (0,0), (-1,1), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,1), "Helvetica-Bold"),
        ]

        status_row = len(table_data) - 1
        for c in range(1, len(table_data[0])):
            if table_data[status_row][c] == "P":
                style_cmds.append(("TEXTCOLOR", (c,status_row), (c,status_row), colors.green))
            elif table_data[status_row][c] == "A":
                style_cmds.append(("TEXTCOLOR", (c,status_row), (c,status_row), colors.red))

        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
        elements.append(PageBreak())

    pdf.build(elements)

print("✅ Month-wise consolidated attendance generated successfully")
