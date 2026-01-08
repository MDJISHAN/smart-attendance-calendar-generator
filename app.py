from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
import sys
import subprocess
import pandas as pd
import shutil
from datetime import datetime
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
# HELPER: READ STUDENT EXCEL
# --------------------------------------------------
def read_students_from_excel(excel_path):
    df = pd.read_excel(excel_path)
    df.columns = [c.strip().lower() for c in df.columns]

    if len(df.columns) < 2:
        raise ValueError("Excel must have Empcode in 1st column and Name in 2nd column")

    students = []
    for _, row in df.iterrows():
        empcode = str(row.iloc[0]).strip()
        name = str(row.iloc[1]).strip()
        if empcode and name:
            students.append({"empcode": empcode, "name": name})

    return students

# --------------------------------------------------
# ROUTES
# --------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        # ---------------- FORM DATA ----------------
        batch_id = request.form["batch_id"]
        dept_name = request.form["dept_name"]
        company_name = request.form["company_name"]
        start_date = request.form["start_date"]  # YYYY-MM-DD
        end_date = request.form["end_date"]      # YYYY-MM-DD

        # ---------------- FILE UPLOAD ----------------
        file = request.files.get("student_file")
        if not file or file.filename == "":
            flash("❌ Please upload student Excel file")
            return redirect(url_for("index"))

        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_DIR, filename)
        file.save(file_path)

        # ---------------- READ STUDENTS ----------------
        students = read_students_from_excel(file_path)
        if not students:
            flash("❌ No valid students found in Excel")
            return redirect(url_for("index"))

        # ---------------- DATE FORMAT FIX ----------------
        # HTML input -> generator format
        start_date_fmt = datetime.strptime(start_date, "%Y-%m-%d").strftime("%d-%m-%Y")
        end_date_fmt   = datetime.strptime(end_date, "%Y-%m-%d").strftime("%d-%m-%Y")

        # ---------------- CLEAN OLD OUTPUT ----------------
        shutil.rmtree(OUTPUT_DIR, ignore_errors=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # ---------------- WRITE CONFIG ----------------
        with open("ui_config.txt", "w") as f:
            f.write(batch_id + "\n")
            f.write(dept_name + "\n")
            f.write(company_name + "\n")
            f.write(start_date_fmt + "\n")
            f.write(end_date_fmt + "\n")

        # Save student list for generator
        students_file = os.path.join(UPLOAD_DIR, "students.xlsx")
        pd.DataFrame(students).to_excel(students_file, index=False)

        # ---------------- RUN GENERATOR ----------------
        subprocess.run(
            [sys.executable, "attendance_generator.py"],
            check=True
        )


        flash("Hey Mr. Sandip ✅ Attendance generated successfully for entire batch!")

        # ---------------- ZIP OUTPUT ----------------
        shutil.make_archive("attendance_batch_output", "zip", OUTPUT_DIR)

        flash("✅ Attendance generated successfully for entire batch!")

        return redirect(url_for("index"))

    except Exception as e:
        flash(f"❌ Error: {str(e)}")
        return redirect(url_for("index"))


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
