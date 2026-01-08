import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import sys
import os
import shutil

# ---------------- UI APP ----------------

def generate_attendance():
    batch_id = entry_batch.get().strip()
    dept_name = entry_dept.get().strip()
    company_name = entry_company.get().strip()
    start_date = entry_start.get().strip()
    end_date = entry_end.get().strip()

    # ---------------- VALIDATION ----------------
    if not all([batch_id, dept_name, company_name, start_date, end_date]):
        messagebox.showerror("Error", "All fields are required")
        return

    if "YYYY" in start_date or "YYYY" in end_date:
        messagebox.showerror("Error", "Please enter valid dates (dd-mm-yyyy)")
        return

    if not student_file_path.get():
        messagebox.showerror("Error", "Please select student Excel file")
        return

    # ---------------- PREPARE FOLDERS ----------------
    os.makedirs("uploads", exist_ok=True)

    # Copy selected student file
    try:
        shutil.copy(student_file_path.get(), "uploads/students.xlsx")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to copy student file\n{e}")
        return

    # ---------------- WRITE CONFIG ----------------
    with open("ui_config.txt", "w") as f:
        f.write(batch_id + "\n")
        f.write(dept_name + "\n")
        f.write(company_name + "\n")
        f.write(start_date + "\n")
        f.write(end_date + "\n")

    # ---------------- RUN GENERATOR ----------------
    try:
        subprocess.run(
            [sys.executable, "attendance_generator.py"],
            check=True
        )
        messagebox.showinfo(
            "Success",
            "✅ Batch attendance generated successfully!\n\n"
            "✔ One PDF per month\n"
            "✔ One Excel per month\n"
            "✔ All students included"
        )
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", str(e))


def browse_student_file():
    file_path = filedialog.askopenfilename(
        title="Select Student Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        student_file_path.set(file_path)


# ---------------- WINDOW ----------------
root = tk.Tk()
root.title("Smart Attendance Calendar Generator (Batch Mode)")
root.geometry("820x620")
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
    text="Batch-wise biometric attendance (Month-wise consolidated)",
    font=("Segoe UI", 10),
    bg="#f4f6f8",
    fg="#555"
)
subtitle.pack(pady=5)

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
        width=18
    ).grid(row=row, column=0, pady=6, sticky="w")

    entry = tk.Entry(frame, width=32, font=("Segoe UI", 10))
    entry.grid(row=row, column=1, pady=6)
    if default:
        entry.insert(0, default)
    return entry

entry_batch   = make_row("Batch ID", 0)
entry_dept    = make_row("Dept. / Course Name", 1, "Students")
entry_company = make_row("Company Name", 2, "SAKSHAM")
entry_start   = make_row("Start Date (dd-mm-yyyy)", 3, "12-09-2025")
entry_end     = make_row("End Date (dd-mm-yyyy)", 4, "12-12-2025")

# ---------------- STUDENT FILE ----------------
student_file_path = tk.StringVar()

tk.Label(
    frame,
    text="Student Excel File",
    font=("Segoe UI", 10, "bold"),
    bg="#f4f6f8",
    anchor="w",
    width=18
).grid(row=5, column=0, pady=10, sticky="w")

tk.Entry(
    frame,
    textvariable=student_file_path,
    width=32,
    font=("Segoe UI", 10),
    state="readonly"
).grid(row=5, column=1, pady=10)

tk.Button(
    frame,
    text="Browse",
    command=browse_student_file,
    bg="#3498db",
    fg="white",
    font=("Segoe UI", 9, "bold"),
    padx=10
).grid(row=5, column=2, padx=8)

# ---------------- BUTTON ----------------
btn = tk.Button(
    root,
    text="Generate Batch Attendance",
    font=("Segoe UI", 11, "bold"),
    bg="#27ae60",
    fg="white",
    padx=25,
    pady=12,
    command=generate_attendance
)
btn.pack(pady=30)

# ---------------- FOOTER ----------------
footer = tk.Label(
    root,
    text="© Smart Attendance Tool | Batch-wise Monthly Reports",
    font=("Segoe UI", 8),
    bg="#f4f6f8",
    fg="#888"
)
footer.pack(side="bottom", pady=12)

root.mainloop()
