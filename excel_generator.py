import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def _money(n: float) -> float:
    try:
        return float(n)
    except Exception:
        return 0.0

def generate_excel(
    employee_name: str,
    employee_email: str,
    location: str,
    depart,
    return_date,
    purpose: str,
    per_diem_rate: float,
    per_diem_days: int,
    expenses: list[dict]
) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Expense Report"

    # Styles
    title_font = Font(size=16, bold=True)
    header_font = Font(bold=True, color="FFFFFF")
    bold_font = Font(bold=True)
    center = Alignment(vertical="center")
    header_fill = PatternFill("solid", fgColor="B10000")  # Performa red vibe
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title
    ws["A1"] = "PERFORMA"
    ws["A1"].font = title_font
    ws["A2"] = "Expense Report"
    ws["A2"].font = Font(size=12, bold=True)

    # Trip summary block
    ws["A4"] = "Employee"
    ws["B4"] = employee_name
    ws["A5"] = "Employee Email"
    ws["B5"] = employee_email
    ws["A6"] = "Trip Location"
    ws["B6"] = location
    ws["A7"] = "Departure Date"
    ws["B7"] = str(depart)
    ws["A8"] = "Return Date"
    ws["B8"] = str(return_date)
    ws["A9"] = "Business Purpose"
    ws["B9"] = purpose

    for r in range(4, 10):
        ws[f"A{r}"].font = bold_font
        ws[f"A{r}"].alignment = center
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")

    # Per diem
    ws["A11"] = "Per Diem Rate"
    ws["B11"] = per_diem_rate
    ws["A12"] = "Per Diem Days"
    ws["B12"] = per_diem_days
    ws["A13"] = "Per Diem Total"
    ws["B13"] = per_diem_rate * per_diem_days
    for r in range(11, 14):
        ws[f"A{r}"].font = bold_font
        ws[f"B{r}"].number_format = '"$"#,##0.00' if r in (11, 13) else '0'

    # Line items header
    start_row = 15
    headers = ["Category", "Date", "Description", "Amount", "Paid By", "Reimbursable", "Receipt Attached"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Line items
    row = start_row + 1
    employee_paid_total = 0.0
    company_paid_total = 0.0
    total_spend = 0.0

    for e in expenses:
        amt = _money(e.get("amount", 0))
        paid_by = e.get("paid_by", "Employee")
        reimb = amt if paid_by == "Employee" else 0.0
        has_receipt = "Yes" if e.get("receipt_name") else "No"

        ws.cell(row=row, column=1, value=e.get("category", "Other"))
        ws.cell(row=row, column=2, value=str(e.get("date", "")))
        ws.cell(row=row, column=3, value=e.get("description", ""))
        ws.cell(row=row, column=4, value=amt)
        ws.cell(row=row, column=5, value=paid_by)
        ws.cell(row=row, column=6, value=reimb)
        ws.cell(row=row, column=7, value=has_receipt)

        ws.cell(row=row, column=4).number_format = '"$"#,##0.00'
        ws.cell(row=row, column=6).number_format = '"$"#,##0.00'

        for c in range(1, 8):
            ws.cell(row=row, column=c).border = border
            ws.cell(row=row, column=c).alignment = Alignment(vertical="top", wrap_text=True)

        total_spend += amt
        if paid_by == "Employee":
            employee_paid_total += amt
        else:
            company_paid_total += amt

        row += 1

    per_diem_total = per_diem_rate * per_diem_days
    reimbursement_due = per_diem_total + employee_paid_total

    # Totals block
    totals_row = row + 1
    ws[f"E{totals_row}"] = "Total Spend"
    ws[f"F{totals_row}"] = total_spend
    ws[f"E{totals_row+1}"] = "Company Paid"
    ws[f"F{totals_row+1}"] = company_paid_total
    ws[f"E{totals_row+2}"] = "Employee Paid"
    ws[f"F{totals_row+2}"] = employee_paid_total
    ws[f"E{totals_row+3}"] = "Per Diem"
    ws[f"F{totals_row+3}"] = per_diem_total
    ws[f"E{totals_row+4}"] = "Reimbursement Due"
    ws[f"F{totals_row+4}"] = reimbursement_due

    for r in range(totals_row, totals_row + 5):
        ws[f"E{r}"].font = bold_font
        ws[f"F{r}"].number_format = '"$"#,##0.00'
        ws[f"E{r}"].alignment = Alignment(horizontal="right")
        ws[f"F{r}"].alignment = Alignment(horizontal="right")

    # Column widths
    widths = [18, 14, 40, 12, 12, 14, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A16"

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream
