"""
Payslip Generator
------------------
Upload an Excel sheet (one row per employee) and get back a single PDF
with 4 payslips per A4 page, ready to print and cut.

No calculations are performed. Every value shown on the payslip is taken
directly from the matching Excel column.
"""

import io
import re
from datetime import datetime

import openpyxl
from flask import Flask, request, send_file, render_template, flash, redirect, url_for
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

app = Flask(__name__)
app.secret_key = "change-this-secret-key"  # only used for flash messages

# ---------------------------------------------------------------------------
# 1. Column mapping: matches many possible Excel header spellings to a
#    fixed internal field name. Add new aliases here if your headers change.
# ---------------------------------------------------------------------------

ALIASES = {
    "emp id": "emp_id",
    "employee id": "emp_id",
    "empid": "emp_id",
    "emp name": "emp_name",
    "employee name": "emp_name",
    "name": "emp_name",
    "actual hrs": "actual_hours",
    "actual hours": "actual_hours",
    "bonus hours": "bonus_hours",
    "bonus hrs": "bonus_hours",
    "total": "total_hours",
    "total hours": "total_hours",
    "absent leave": "absent_leave",
    "absent": "absent_leave",
    "leave": "absent_leave",
    "rate per month": "rate_per_month",
    "rate": "rate_per_month",
    "salary": "salary",
    "bonus alowns": "bonus_allowance",
    "bonus allowances": "bonus_allowance",
    "bonus allowance": "bonus_allowance",
    "total salary": "total_salary",
    "pending salary advance": "pending_salary",
    "pending salary": "pending_salary",
    "adj round off": "round_off",
    "round off": "round_off",
    "adv ded": "adv_ded",
    "advance deductions": "adv_ded",
    "advance deduction": "adv_ded",
    "c3 payment": "net_salary",
    "net salary": "net_salary",
    "month": "month",
    "employee welfare fund": "welfare_fund",
    "welfare fund": "welfare_fund",
}

# Fields shown in the "Working Hours" box, in order
HOURS_FIELDS = [
    ("actual_hours", "Actual Hours"),
    ("bonus_hours", "Bonus Hours"),
    ("total_hours", "Total Hours"),
    ("absent_leave", "Absent/Leave"),
    ("rate_per_month", "Rate per month (AED)"),
]

# Fields shown in the "Salary Details & Deductions" box, in order
SALARY_FIELDS = [
    ("salary", "Salary"),
    ("bonus_allowance", "Bonus/Allowances"),
    ("total_salary", "Total Salary"),
    ("pending_salary", "Pending Salary/Advance"),
    ("round_off", "Round Off"),
    ("adv_ded", "Advance/Deductions"),
    ("welfare_fund", "Employee Welfare Fund"),
]

CURRENCY_FIELDS = {
    "salary", "bonus_allowance", "total_salary", "pending_salary",
    "round_off", "adv_ded", "welfare_fund", "net_salary", "rate_per_month",
}


def normalize(text):
    """Lowercase, strip punctuation/extra spaces, for fuzzy header matching."""
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def build_column_map(header_row):
    """Given the raw header cells from Excel, return {field_key: column_index}."""
    col_map = {}
    for idx, header in enumerate(header_row):
        if header is None:
            continue
        norm = normalize(header)
        if norm in ALIASES:
            col_map[ALIASES[norm]] = idx
    return col_map


def read_employee_rows(file_stream):
    """Read the uploaded Excel file and return a list of employee dicts."""
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("The uploaded sheet is empty.")

    header_row = rows[0]
    col_map = build_column_map(header_row)

    if "emp_name" not in col_map:
        raise ValueError(
            "Could not find an Employee Name column. "
            "Please check your Excel header row."
        )

    employees = []
    for row in rows[1:]:
        if row is None or all(cell is None for cell in row):
            continue  # skip fully blank rows
        record = {}
        for field_key, col_idx in col_map.items():
            if col_idx < len(row):
                record[field_key] = row[col_idx]
        if record.get("emp_name") is None:
            continue  # skip rows with no employee name
        employees.append(record)

    return employees


def fmt_value(field_key, value):
    """Format a raw cell value for display, with no calculation."""
    if value is None or (isinstance(value, float) and value != value):  # NaN check
        return "AED 0.00" if field_key in CURRENCY_FIELDS else "0"
    if field_key in CURRENCY_FIELDS:
        try:
            return f"AED {float(value):,.2f}"
        except (TypeError, ValueError):
            return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


# ---------------------------------------------------------------------------
# 2. PDF generation: draws 4 payslips per A4 page (2x2 grid)
# ---------------------------------------------------------------------------

PAGE_W, PAGE_H = A4
MARGIN = 8 * mm
QUAD_W = (PAGE_W - 2 * MARGIN) / 2
QUAD_H = (PAGE_H - 2 * MARGIN) / 2

QUAD_ORIGINS = [
    (MARGIN, PAGE_H / 2),                # top-left
    (MARGIN + QUAD_W, PAGE_H / 2),       # top-right
    (MARGIN, MARGIN),                    # bottom-left
    (MARGIN + QUAD_W, MARGIN),           # bottom-right
]


def draw_cut_lines(c):
    c.saveState()
    c.setDash(3, 3)
    c.setStrokeGray(0.6)
    c.line(PAGE_W / 2, MARGIN, PAGE_W / 2, PAGE_H - MARGIN)
    c.line(MARGIN, PAGE_H / 2, PAGE_W - MARGIN, PAGE_H / 2)
    c.restoreState()


def draw_payslip(c, ox, oy, employee, month_label):
    """Draw a single payslip inside the quadrant starting at (ox, oy)."""
    pad = 5 * mm
    x0, y0 = ox + pad, oy + pad
    w, h = QUAD_W - 2 * pad, QUAD_H - 2 * pad
    top = y0 + h

    c.saveState()
    c.setLineWidth(0.75)
    c.rect(x0, y0, w, h)

    # Title
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(x0 + w / 2, top - 12, "PAYSLIP")

    # Header: name / id / month
    c.setFont("Helvetica-Bold", 7.5)
    name = str(employee.get("emp_name") or "")
    c.drawString(x0 + 4, top - 24, f"Employee Name: {name}")
    c.setFont("Helvetica", 7.5)
    month_text = str(employee.get("month") or month_label or "")
    c.drawRightString(x0 + w - 4, top - 24, f"Month: {month_text}")
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(x0 + 4, top - 34, f"Employee ID: {employee.get('emp_id', '')}")

    col_gap = 4 * mm
    col_w = (w - col_gap) / 2
    left_x = x0
    right_x = x0 + col_w + col_gap
    section_top = top - 46

    # Working Hours box
    c.setFont("Helvetica-Bold", 7)
    c.drawString(left_x + 2, section_top, "Working Hours")
    c.setFont("Helvetica", 6.5)
    line_h = 9
    ty = section_top - line_h
    for field_key, label in HOURS_FIELDS:
        val = fmt_value(field_key, employee.get(field_key))
        c.drawString(left_x + 2, ty, f"{label}:")
        c.drawRightString(left_x + col_w - 2, ty, val)
        ty -= line_h
    c.rect(left_x, ty - 2, col_w, (section_top + line_h) - (ty - 2))

    # Salary Details box
    c.setFont("Helvetica-Bold", 6.5)
    c.drawString(right_x + 2, section_top, "Salary Details (AED)")
    c.setFont("Helvetica", 6.5)
    sy = section_top - line_h
    for field_key, label in SALARY_FIELDS:
        if field_key not in employee:
            continue  # skip rows for optional fields not present in this sheet
        val = fmt_value(field_key, employee.get(field_key))
        c.drawString(right_x + 2, sy, f"{label}:")
        c.drawRightString(right_x + col_w - 2, sy, val)
        sy -= line_h
    box_bottom = min(ty, sy) - 2
    c.rect(right_x, sy - 2, col_w, (section_top + line_h) - (sy - 2))

    # Net salary (bold line)
    net_y = box_bottom - 12
    c.setFont("Helvetica-Bold", 8.5)
    net_val = fmt_value("net_salary", employee.get("net_salary"))
    c.drawCentredString(x0 + w / 2, net_y, f"Net Salary (AED): {net_val}")

    # Footer disclaimer
    c.setFont("Helvetica-Oblique", 5.5)
    c.drawCentredString(
        x0 + w / 2, y0 + 4,
        "This is a computer-generated payslip and does not require a signature."
    )

    c.restoreState()


def generate_pdf(employees, month_label=""):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    for i in range(0, len(employees), 4):
        batch = employees[i:i + 4]
        draw_cut_lines(c)
        for slot, employee in enumerate(batch):
            ox, oy = QUAD_ORIGINS[slot]
            draw_payslip(c, ox, oy, employee, month_label)
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# 3. Routes
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    uploaded_file = request.files.get("excel_file")
    if not uploaded_file or uploaded_file.filename == "":
        flash("Please choose an Excel file to upload.")
        return redirect(url_for("index"))

    if not uploaded_file.filename.lower().endswith((".xlsx", ".xlsm")):
        flash("Please upload a .xlsx file.")
        return redirect(url_for("index"))

    month_label = request.form.get("month_label", "").strip()

    try:
        employees = read_employee_rows(uploaded_file.stream)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for("index"))

    if not employees:
        flash("No employee rows were found in the uploaded file.")
        return redirect(url_for("index"))

    pdf_buffer = generate_pdf(employees, month_label)
    filename = f"payslips_{month_label or datetime.now().strftime('%b-%y')}.pdf"

    return send_file(
        pdf_buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
