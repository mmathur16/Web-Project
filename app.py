from flask import Flask, render_template, request, send_file
import sqlite3
from datetime import datetime
from openpyxl import Workbook

app = Flask(__name__)

DB_NAME = "teacher_portal.db"

# -------------------------
# CREATE DATABASE
# -------------------------
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_name TEXT NOT NULL,
            semester_type TEXT NOT NULL,
            semester INTEGER NOT NULL,
            department TEXT NOT NULL,

            co1_direct TEXT,
            co1_indirect TEXT,
            co2_direct TEXT,
            co2_indirect TEXT,
            co3_direct TEXT,
            co3_indirect TEXT,
            co4_direct TEXT,
            co4_indirect TEXT,

            created_at TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()


# -------------------------
# HOME PAGE
# -------------------------
@app.route("/")
def home():
    return render_template("index.html")


# -------------------------
# SUBMIT FORM
# -------------------------
@app.route("/submit", methods=["POST"])
def submit():

    teacher_name = request.form.get("teacher_name")
    semester_type = request.form.get("semester_type")
    semester = request.form.get("semester")
    department = request.form.get("department")

    co1_direct = request.form.get("co1_direct")
    co1_indirect = request.form.get("co1_indirect")
    co2_direct = request.form.get("co2_direct")
    co2_indirect = request.form.get("co2_indirect")
    co3_direct = request.form.get("co3_direct")
    co3_indirect = request.form.get("co3_indirect")
    co4_direct = request.form.get("co4_direct")
    co4_indirect = request.form.get("co4_indirect")

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO submissions (
            teacher_name,
            semester_type,
            semester,
            department,

            co1_direct,
            co1_indirect,
            co2_direct,
            co2_indirect,
            co3_direct,
            co3_indirect,
            co4_direct,
            co4_indirect,

            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        teacher_name,
        semester_type,
        semester,
        department,

        co1_direct,
        co1_indirect,
        co2_direct,
        co2_indirect,
        co3_direct,
        co3_indirect,
        co4_direct,
        co4_indirect,

        created_at
    ))

    conn.commit()
    conn.close()

    return """
        <h3>Submitted Successfully!</h3>
        <a href='/'>Back</a> |
        <a href='/view'>View Submissions</a> |
        <a href='/export'>Download Excel</a>
    """


# -------------------------
# VIEW SUBMISSIONS
# -------------------------
@app.route("/view")
def view():

    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()

    cur.execute("""
        SELECT * FROM submissions
        ORDER BY id DESC
    """)

    rows = cur.fetchall()

    conn.close()

    return render_template("view.html", rows=rows)


# -------------------------
# EXPORT TO EXCEL
# -------------------------
@app.route("/export")
def export_excel():

    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()

    cur.execute("SELECT * FROM submissions ORDER BY id DESC")

    rows = cur.fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Teacher Responses"

    headers = [
        "ID",
        "Teacher Name",
        "Semester Type",
        "Semester",
        "Department",
        "CO1 Direct",
        "CO1 Indirect",
        "CO2 Direct",
        "CO2 Indirect",
        "CO3 Direct",
        "CO3 Indirect",
        "CO4 Direct",
        "CO4 Indirect",
        "Created At"
    ]

    ws.append(headers)

    for row in rows:
        ws.append(row)

    file_name = "teacher_responses.xlsx"

    wb.save(file_name)

    return send_file(file_name, as_attachment=True)


# -------------------------
# MAIN
# -------------------------
if __name__ == "__main__":
    init_db()
    app.run(debug=True)