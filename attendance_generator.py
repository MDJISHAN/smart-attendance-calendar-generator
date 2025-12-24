import pandas as pd
import random
import calendar
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document

# ================= USER INPUT =================
# ===== READ FROM UI =====
# ===== READ FROM UI CONFIG =====
with open("ui_config.txt") as f:
    BATCH_ID = f.readline().strip()
    EMP_CODE = f.readline().strip()
    EMP_NAME = f.readline().strip()
    DEPT_NAME = f.readline().strip()
    COMP_NAME = f.readline().strip()
    START_DATE = f.readline().strip()
    END_DATE = f.readline().strip()


# =============================================

start_date = datetime.strptime(START_DATE, "%Y-%m-%d")
end_date   = datetime.strptime(END_DATE, "%Y-%m-%d")

def hhmm(minutes):
    return f"{minutes//60:02d}:{minutes%60:02d}"

def google_weekday(dt):
    return ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][dt.weekday()]

def generate_month(year, month):
    days = calendar.monthrange(year, month)[1]

    rows = {"IN": [], "OUT": [], "WORK": [], "BREAK": [], "OT": [], "Status": []}
    weekday_row = []

    present = absent = wo = hl = lv = 0
    total_work = total_ot = 0

    valid_days = []
    for d in range(1, days + 1):
        dt = datetime(year, month, d)
        weekday_row.append(google_weekday(dt))
        if start_date <= dt <= end_date and dt.weekday() != 6:
            valid_days.append(d)
        if dt.weekday() == 6:
            wo += 1

    absent_days = random.sample(valid_days, min(random.randint(2, 4), len(valid_days)))

    for d in range(1, days + 1):
        dt = datetime(year, month, d)

        if d not in valid_days:
            for k in rows:
                rows[k].append("---")
        elif d in absent_days:
            rows["IN"].append("---")
            rows["OUT"].append("---")
            rows["WORK"].append("00:00")
            rows["BREAK"].append("00:00")
            rows["OT"].append("00:00")
            rows["Status"].append("A")
            absent += 1
        else:
            in_m = random.randint(5, 40)
            out_m = random.randint(0, 45)
            work = (14 * 60 + out_m) - (8 * 60 + in_m)

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

    return df, weekday_row, summary

# ================= GENERATE MONTHS =================
month_blocks = []
cur = start_date.replace(day=1)

while cur <= end_date:
    df, weekdays, summary = generate_month(cur.year, cur.month)
    month_blocks.append((cur.strftime("%B-%Y"), df, weekdays, summary))
    cur = (cur + timedelta(days=32)).replace(day=1)

# ================= EXCEL =================
excel_file = "attendance_calendar.xlsx"
with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    for month, df, weekdays, summary in month_blocks:
        header = [
            ["Dept. Name", DEPT_NAME, "", "CompName", COMP_NAME, "", "Report Month", month],
            ["Empcode", EMP_CODE, "Name", EMP_NAME, "Present", summary["Present"], "Absent", summary["Absent"]],
            ["WO", summary["WO"], "HL", summary["HL"], "LV", summary["LV"], "Tot. Work+OT", summary["TotalWorkOT"]],
            ["", "", "", "", "", "", "Total OT", summary["TotalOT"]],
            [""] + df.columns.tolist(),
            [""] + weekdays
        ]
        pd.DataFrame(header).to_excel(writer, sheet_name=month, index=False, header=False)
        df.to_excel(writer, sheet_name=month, startrow=6)

# Format Excel
wb = load_workbook(excel_file)
for ws in wb.worksheets:
    for r in range(1, 5):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)

    for row in ws.iter_rows(min_row=1, max_row=6):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 12
    for c in range(2, 35):
        ws.column_dimensions[get_column_letter(c)].width = 6

wb.save(excel_file)

# ================= PDF =================
pdf = SimpleDocTemplate(
    "attendance_calendar.pdf",
    pagesize=landscape(A3),
    leftMargin=10,
    rightMargin=10,
    topMargin=10,
    bottomMargin=10
)

styles = getSampleStyleSheet()
elements = []

for month, df, weekdays, summary in month_blocks:
    elements.append(Paragraph(f"<b>{month}</b>", styles["Heading2"]))
    elements.append(Paragraph(
        f"Dept. Name: {DEPT_NAME} &nbsp;&nbsp; CompName: {COMP_NAME} "
        f"&nbsp;&nbsp; Report Month: {month}<br/>"
        f"Empcode: {EMP_CODE} &nbsp;&nbsp; Name: {EMP_NAME} "
        f"&nbsp;&nbsp; Present: {summary['Present']} "
        f"&nbsp;&nbsp; Absent: {summary['Absent']} "
        f"&nbsp;&nbsp; WO: {summary['WO']} HL: {summary['HL']} LV: {summary['LV']} "
        f"&nbsp;&nbsp; Tot. Work+OT: {summary['TotalWorkOT']} "
        f"&nbsp;&nbsp; Total OT: {summary['TotalOT']}",
        styles["Normal"]
    ))
    elements.append(Spacer(1, 10))

    table_data = [
        [""] + df.columns.tolist(),
        [""] + weekdays
    ]
    for idx, row in df.iterrows():
        table_data.append([idx] + row.tolist())

    table = Table(table_data, colWidths=[45] + [35]*31, repeatRows=2)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (0, 0), (-1, 1), colors.lightgrey)
    ]))

    elements.append(table)
    elements.append(Spacer(1, 25))

pdf.build(elements)

# ================= WORD =================
doc = Document()
doc.add_heading("Attendance Report", level=1)

for month, df, weekdays, summary in month_blocks:
    doc.add_heading(month, level=2)
    doc.add_paragraph(
        f"Dept. Name: {DEPT_NAME}    CompName: {COMP_NAME}    Report Month: {month}\n"
        f"Empcode: {EMP_CODE}    Name: {EMP_NAME}\n"
        f"Present: {summary['Present']}  Absent: {summary['Absent']}  "
        f"WO: {summary['WO']}  HL: {summary['HL']}  LV: {summary['LV']}  "
        f"Tot. Work+OT: {summary['TotalWorkOT']}  Total OT: {summary['TotalOT']}"
    )

    table = doc.add_table(rows=len(df)+2, cols=len(df.columns)+1)
    table.rows[0].cells[0].text = ""
    for i, c in enumerate(df.columns):
        table.rows[0].cells[i+1].text = c
        table.rows[1].cells[i+1].text = weekdays[i]

    for r, (idx, row) in enumerate(df.iterrows(), start=2):
        table.rows[r].cells[0].text = idx
        for c, val in enumerate(row):
            table.rows[r].cells[c+1].text = val

doc.save("attendance_calendar.docx")

print("✅ FINAL attendance generated (Excel, PDF, Word) — perfectly aligned")
