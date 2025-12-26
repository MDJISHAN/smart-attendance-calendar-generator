from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
import sys
import subprocess
import pandas as pd
import shutil
from datetime import datetime
from dateutil.relativedelta import relativedelta
from werkzeug.utils import secure_filename

# --------------------------------------------------
# FLASK APP CONFIG
# --------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "attendance-secret-key")

BASE_DIR = os.getcwd()
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --------------------------------------------------
# ROUTES
# --------------------------------------------------
def read_students_from_excel(excel_path):
    df = pd.read_excel(excel_path)

    # Normalize column names
    df.columns = [c.strip().lower() for c in df.columns]

    if len(df.columns) < 2:
        raise ValueError("Excel must have Empcode in 1st column and Name in 2nd column")

    students = []
    for _, row in df.iterrows():
        empcode = str(row.iloc[0]).strip()
        name = str(row.iloc[1]).strip()

        if empcode and name:
            students.append({
                "empcode": empcode,
                "name": name
            })

    return students

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        # -----------------------------
        # FORM DATA
        # -----------------------------
        batch_id = request.form["batch_id"]
        dept_name = request.form["dept_name"]
        company_name = request.form["company_name"]
        start_date = request.form["start_date"]
        end_date = request.form["end_date"]

        # -----------------------------
        # FILE UPLOAD
        # -----------------------------
        file = request.files["student_file"]
        if not file:
            flash("❌ Please upload student Excel file")
            return redirect(url_for("index"))

        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_DIR, filename)
        file.save(file_path)

        # -----------------------------
        # READ STUDENTS (Empcode + Name)
        # -----------------------------
        students = read_students_from_excel(file_path)

        if not students:
            flash("❌ No valid students found in Excel")
            return redirect(url_for("index"))

        # -----------------------------
        # DATE PARSING (HTML DATE INPUT)
        # -----------------------------
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt   = datetime.strptime(end_date, "%Y-%m-%d")

        # -----------------------------
        # MONTH LIST
        # -----------------------------
        months = []
        current = start_dt.replace(day=1)
        while current <= end_dt:
            months.append(current)
            current += relativedelta(months=1)

        # -----------------------------
        # CLEAN OLD OUTPUT
        # -----------------------------
        shutil.rmtree(OUTPUT_DIR, ignore_errors=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # -----------------------------
        # GENERATE ATTENDANCE
        # -----------------------------
        for month in months:
            month_start = max(start_dt, month.replace(day=1))
            month_end = min(
                end_dt,
                (month + relativedelta(months=1)) - relativedelta(days=1)
            )

            for student in students:
                with open("ui_config.txt", "w") as f:
                    f.write(batch_id + "\n")
                    f.write(student["empcode"] + "\n")   # ✅ REAL EMPCODE
                    f.write(student["name"] + "\n")      # ✅ STUDENT NAME
                    f.write(dept_name + "\n")
                    f.write(company_name + "\n")
                    f.write(month_start.strftime("%d-%m-%Y") + "\n")
                    f.write(month_end.strftime("%d-%m-%Y") + "\n")

                subprocess.run(
                    [sys.executable, "attendance_generator.py"],
                    check=True
                )

        # -----------------------------
        # ZIP OUTPUT
        # -----------------------------
        zip_path = shutil.make_archive(
            "attendance_batch_output",
            "zip",
            OUTPUT_DIR
        )

        flash("✅ Attendance generated successfully for entire batch!")
        return redirect(url_for("index"))


    except Exception as e:
        flash(f"❌ Error: {str(e)}")
        return redirect(url_for("index"))

from flask import send_from_directory
@app.route("/download/zip")
def download_zip():
    zip_path = "attendance_batch_output.zip"
    if not os.path.exists(zip_path):
        flash("❌ No generated file found")
        return redirect(url_for("index"))
    return send_file(zip_path, as_attachment=True)



# --------------------------------------------------
# RUN APP (RENDER SAFE)
# --------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
