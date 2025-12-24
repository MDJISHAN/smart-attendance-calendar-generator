from flask import Flask, render_template, request, render_template, request, redirect, url_for, flash, send_file
import subprocess
import sys
import os

app = Flask(__name__)
app.secret_key = "smart_attendance_secret_2025"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    data = [
        request.form["batch_id"],
        request.form["emp_id"],
        request.form["emp_name"],
        request.form["dept_name"],
        request.form["company_name"],
        request.form["start_date"],
        request.form["end_date"]
    ]

    # Write config
    with open("ui_config.txt", "w") as f:
        for v in data:
            f.write(v + "\n")

    subprocess.run([sys.executable, "attendance_generator.py"], check=True)

    flash("✅ Attendance generated successfully! Download your files below.")
    return redirect(url_for("index"))

@app.route("/download/<filetype>")
def download(filetype):
    files = {
        "excel": "attendance_calendar.xlsx",
        "pdf": "attendance_calendar.pdf",
        "docx": "attendance_calendar.docx"
    }

    if filetype not in files:
        return "Invalid file type"

    return send_file(files[filetype], as_attachment=True)    

if __name__ == "__main__":
    app.run(debug=True)
