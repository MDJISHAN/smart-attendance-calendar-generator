import pandas as pd
import random, os
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
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT


# ================= READ FROM UI CONFIG =================
with open("ui_config.txt") as f:
    BATCH_ID = f.readline().strip()
    EMP_CODE = f.readline().strip()
    EMP_NAME = f.readline().strip()
    DEPT_NAME = f.readline().strip()
    COMP_NAME = f.readline().strip()
    START_DATE = f.readline().strip()
    END_DATE = f.readline().strip()

start_date = datetime.strptime(START_DATE, "%d-%m-%Y")
end_date   = datetime.strptime(END_DATE, "%d-%m-%Y")

def hhmm(minutes):
    return f"{minutes//60:02d}:{minutes%60:02d}"

def google_weekday(dt):
    return ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][dt.weekday()]

# ================= MONTH GENERATION =================
def generate_month(year, month, global_start, global_end):
    month_start = datetime(year, month, 1)
    month_end = datetime(year, month, calendar.monthrange(year, month)[1])

    effective_start = max(global_start, month_start)
    effective_end = min(global_end, month_end)

    days = calendar.monthrange(year, month)[1]

    rows = {"IN": [], "OUT": [], "WORK": [], "BREAK": [], "OT": [], "Status": []}
    weekday_row = []

    present = absent = wo = hl = lv = 0
    total_work = total_ot = 0

    valid_days = []

    for d in range(1, days + 1):
        dt = datetime(year, month, d)
        weekday_row.append(google_weekday(dt))

        if effective_start <= dt <= effective_end and dt.weekday() != 6:
            valid_days.append(d)

        if dt.weekday() == 6:
            wo += 1

    absent_days = random.sample(valid_days, min(random.randint(3, 6), len(valid_days)))

    for d in range(1, days + 1):
        dt = datetime(year, month, d)

        # OUTSIDE COURSE RANGE
        if dt < effective_start or dt > effective_end:
            for k in rows:
                rows[k].append("---")
            continue

        # SUNDAY
        if dt.weekday() == 6:
            for k in rows:
                rows[k].append("---")
            continue

        # ABSENT
        if d in absent_days:
            rows["IN"].append("---")
            rows["OUT"].append("---")
            rows["WORK"].append("00:00")
            rows["BREAK"].append("00:00")
            rows["OT"].append("00:00")
            rows["Status"].append("A")
            absent += 1
            continue

        # PRESENT
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

    return df, weekday_row, summary

# ================= COLLECT MONTHS (FIXED) =================
month_blocks = []
cur = start_date.replace(day=1)

while cur <= end_date:
    df, weekdays, summary = generate_month(
        cur.year,
        cur.month,
        start_date,
        end_date
    )
    month_blocks.append((cur.strftime("%B-%Y"), df, weekdays, summary))
    cur = (cur + timedelta(days=32)).replace(day=1)

# ================= OUTPUT DIR =================
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

REPORT_MONTH = start_date.strftime("%B-%Y")
filename_prefix = f"{EMP_NAME}_{REPORT_MONTH}".replace(" ", "_")

# ================= EXCEL =================
excel_file = f"{OUTPUT_DIR}/{filename_prefix}.xlsx"
with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    for month, df, weekdays, summary in month_blocks:
        header = [
            ["Course Name", DEPT_NAME, "", "Company Name", COMP_NAME, "", "Report Month", month],
            ["Reg", EMP_CODE, "Name", EMP_NAME, "Present", summary["Present"], "Absent", summary["Absent"]],
            ["WO", summary["WO"], "HL", summary["HL"], "LV", summary["LV"], "Tot. Work+OT", summary["TotalWorkOT"]],
            ["", "", "", "", "", "", "Total OT", summary["TotalOT"]],
            [""] + df.columns.tolist(),
            [""] + weekdays
        ]
        pd.DataFrame(header).to_excel(writer, sheet_name=month, index=False, header=False)
        df.to_excel(writer, sheet_name=month, startrow=6)

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
pdf_file = f"{OUTPUT_DIR}/{filename_prefix}.pdf"
styles = getSampleStyleSheet()
elements = []

header_style = ParagraphStyle(
    name="HeaderStyle",
    fontSize=10,
    leading=14,
    alignment=TA_LEFT,
    spaceAfter=6
)

pdf = SimpleDocTemplate(
    pdf_file,
    pagesize=landscape(A3),
    leftMargin=10,
    rightMargin=10,
    topMargin=10,
    bottomMargin=10
)

for month, df, weekdays, summary in month_blocks:
    elements.append(Paragraph(f"<b>{month}</b>", styles["Heading2"]))
    elements.append(Paragraph(
    f"""
    <b>Course Name:</b> {DEPT_NAME}&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <b>Company Name:</b> {COMP_NAME}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <b>Report Month:</b> {month}<br/>

    <b>REG:</b> {EMP_CODE}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <b>Name:</b> {EMP_NAME}<br/>

    <b>Present:</b>
    <font color="green"><b>{summary['Present']}</b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

    <b>Absent:</b>
    <font color="red"><b>{summary['Absent']}</b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

    <b>WO:</b> {summary['WO']}&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp;
    <b>HL:</b> {summary['HL']}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <b>LV:</b> {summary['LV']}<br/>

    <b>Total Work + OT:</b> {summary['TotalWorkOT']}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <b>Total OT:</b> {summary['TotalOT']}
    """,
    header_style
))

    elements.append(Spacer(1, 12))

    table_data = [[""] + df.columns.tolist(), [""] + weekdays]
    for idx, row in df.iterrows():
        table_data.append([idx] + row.tolist())

    table = Table(table_data, colWidths=[45] + [35]*31, repeatRows=2)

    style_commands = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (0, 0), (-1, 1), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
    ]

    # 🔹 Status row index (last row)
    status_row_index = len(table_data) - 1

    # 🔹 Apply colors to P / A
    for col in range(1, len(table_data[0])):
        val = table_data[status_row_index][col]

        if val == "P":
            style_commands.append(
                ("TEXTCOLOR", (col, status_row_index), (col, status_row_index), colors.green)
            )
            style_commands.append(
                ("FONTNAME", (col, status_row_index), (col, status_row_index), "Helvetica-Bold")
            )

        elif val == "A":
            style_commands.append(
                ("TEXTCOLOR", (col, status_row_index), (col, status_row_index), colors.red)
            )
            style_commands.append(
                ("FONTNAME", (col, status_row_index), (col, status_row_index), "Helvetica-Bold")
            )

    table.setStyle(TableStyle(style_commands))


    elements.append(table)
    elements.append(Spacer(1, 25))

pdf.build(elements)



print("✅ FINAL attendance generated correctly (month boundaries respected)")
