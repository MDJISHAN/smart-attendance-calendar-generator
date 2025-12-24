import tkinter as tk
from tkinter import messagebox
import subprocess
import sys
import os

# ---------------- UI APP ----------------

def generate_attendance():
    batch_id = entry_batch.get().strip()
    emp_id = entry_empid.get().strip()
    emp_name = entry_empname.get().strip()
    dept_name = entry_dept.get().strip()
    company_name = entry_company.get().strip()
    start_date = entry_start.get().strip()
    end_date = entry_end.get().strip()

    # Basic validation
    if not all([batch_id, emp_id, emp_name, dept_name, company_name, start_date, end_date]):
        messagebox.showerror("Error", "All fields are required")
        return

    # Date validation
    if "YYYY" in start_date or "YYYY" in end_date:
        messagebox.showerror("Error", "Please enter valid Start Date and End Date")
        return

    # Write config safely
    with open("ui_config.txt", "w") as f:
        f.write(batch_id + "\n")
        f.write(emp_id + "\n")
        f.write(emp_name + "\n")
        f.write(dept_name + "\n")
        f.write(company_name + "\n")
        f.write(start_date + "\n")
        f.write(end_date + "\n")

    try:
        subprocess.run(
            [sys.executable, "attendance_generator.py"],
            check=True
        )
        messagebox.showinfo(
            "Success",
            "Attendance generated successfully!\n\nFiles created:\nExcel | PDF | Word"
        )
    except Exception as e:
        messagebox.showerror("Error", str(e))

# ---------------- WINDOW ----------------
root = tk.Tk()
root.title("Smart Attendance Calendar Generator")
root.geometry("820x660")
root.configure(bg="#f4f6f8")
root.resizable(False, False)

# ---------------- HEADING ----------------
title = tk.Label(
    root,
    text="Smart Attendance Calendar Generator",
    font=("Segoe UI", 16, "bold"),
    bg="#f4f6f8",
    fg="#2c3e50"
)
title.pack(pady=20)

subtitle = tk.Label(
    root,
    text="Generate professional biometric-style attendance reports",
    font=("Segoe UI", 10),
    bg="#f4f6f8",
    fg="#555"
)
subtitle.pack(pady=5)

# ---------------- FORM ----------------
# ---------------- FORM ----------------
frame = tk.Frame(root, bg="#f4f6f8")
frame.pack(pady=20)

def make_row(label, row, default=""):
    tk.Label(
        frame,
        text=label,
        font=("Segoe UI", 10, "bold"),
        bg="#f4f6f8",
        anchor="w",
        width=15
    ).grid(row=row, column=0, pady=6, sticky="w")

    entry = tk.Entry(frame, width=30, font=("Segoe UI", 10))
    entry.grid(row=row, column=1, pady=6)
    if default:
        entry.insert(0, default)
    return entry
entry_company = make_row("Company Name", 4, "SAKSHAM")
entry_batch   = make_row("Batch ID", 0)
entry_empid   = make_row("Employee ID", 1)
entry_empname = make_row("Employee Name", 2)
entry_dept    = make_row("Dept. Name", 3, "Students")
entry_start   = make_row("Start Date", 5, "YYYY-MM-DD")
entry_end     = make_row("End Date", 6, "YYYY-MM-DD")


# ---------------- BUTTON ----------------
btn = tk.Button(
    root,
    text="Generate Attendance",
    font=("Segoe UI", 11, "bold"),
    bg="#27ae60",
    fg="white",
    padx=20,
    pady=10,
    command=generate_attendance
)
btn.pack(pady=25)

# ---------------- FOOTER ----------------
footer = tk.Label(
    root,
    text="© Smart Attendance Tool",
    font=("Segoe UI", 8),
    bg="#f4f6f8",
    fg="#888"
)
footer.pack(side="bottom", pady=10)

root.mainloop()
