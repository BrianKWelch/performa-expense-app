import base64
import os
import re
from datetime import date
from io import BytesIO
from typing import List, Dict, Any, Optional

import streamlit as st
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Mail,
    Email,
    To,
    Cc,
    Attachment,
    FileContent,
    FileName,
    FileType,
    Disposition,
)

from docx import Document  # for NN template population
from excel_generator import generate_excel  # Guam flow stays unchanged


# -----------------------------
# Config and constants
# -----------------------------
st.set_page_config(page_title="Performa Expense Report", layout="wide")

PER_DIEM_RATE = float(st.secrets.get("PER_DIEM_RATE", 100))
MAX_ATTACHMENT_MB = float(st.secrets.get("MAX_ATTACHMENT_MB", 18))

TEMPLATES_DIR = "templates"  # NN docx templates live here

CATEGORIES = [
    "Airfare",
    "Airport Parking",
    "Taxi or Uber to Airport",
    "Hotel",
    "Rental Car",
    "Gas for Rental Car",
    "Other",
]


# -----------------------------
# Helpers
# -----------------------------
def bytes_from_uploaded_file(uploaded_file) -> bytes:
    if uploaded_file is None:
        return b""
    return uploaded_file.getvalue()


def total_receipt_bytes(expenses: List[Dict[str, Any]]) -> int:
    total = 0
    for e in expenses:
        f = e.get("receipt_file")
        if f is not None:
            total += len(bytes_from_uploaded_file(f))
    return total


def calc_trip_days(departure: date, ret: date) -> int:
    if not departure or not ret:
        return 0
    if ret < departure:
        return 0
    return (ret - departure).days + 1


def calc_totals(expenses: List[Dict[str, Any]]) -> Dict[str, float]:
    total_spend = 0.0
    company_paid = 0.0
    employee_paid = 0.0

    for e in expenses:
        amt = float(e.get("amount") or 0)
        total_spend += amt
        if e.get("paid_by") == "Performa":
            company_paid += amt
        else:
            employee_paid += amt

    return {
        "total_spend": total_spend,
        "company_paid": company_paid,
        "employee_paid": employee_paid,
    }


def build_email_html(
    employee_name: str,
    employee_email: str,
    location: str,
    purpose: str,
    departure_date: date,
    return_date: date,
    per_diem_total: float,
    total_spend: float,
    company_paid: float,
    employee_paid: float,
    reimbursement_due: float,
    expenses: List[Dict[str, Any]],
) -> str:
    # Simple, clean HTML that reads well in Outlook and Gmail
    # No em dashes used
    def esc(x: Optional[str]) -> str:
        if x is None:
            return ""
        return (
            str(x)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    lines_html = ""
    if expenses:
        rows = []
        for i, e in enumerate(expenses, start=1):
            rows.append(
                f"""
                <tr>
                  <td style="padding:6px 8px;border-bottom:1px solid #eee;">{i}</td>
                  <td style="padding:6px 8px;border-bottom:1px solid #eee;">{esc(e.get("category",""))}</td>
                  <td style="padding:6px 8px;border-bottom:1px solid #eee;">{esc(e.get("expense_date",""))}</td>
                  <td style="padding:6px 8px;border-bottom:1px solid #eee;">{esc(e.get("description",""))}</td>
                  <td style="padding:6px 8px;border-bottom:1px solid #eee;">{esc(e.get("paid_by",""))}</td>
                  <td style="padding:6px 8px;border-bottom:1px solid #eee;text-align:right;">${float(e.get("amount") or 0):,.2f}</td>
                  <td style="padding:6px 8px;border-bottom:1px solid #eee;">{"Yes" if e.get("receipt_file") else "No"}</td>
                </tr>
                """
            )
        lines_html = f"""
        <p><strong>Line items:</strong></p>
        <table style="border-collapse:collapse;width:100%;font-family:Arial, sans-serif;font-size:13px;">
          <thead>
            <tr>
              <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;">#</th>
              <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;">Category</th>
              <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;">Date</th>
              <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;">Description</th>
              <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;">Paid By</th>
              <th style="text-align:right;padding:6px 8px;border-bottom:2px solid #ddd;">Amount</th>
              <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;">Receipt</th>
            </tr>
          </thead>
          <tbody>
            {''.join(rows)}
          </tbody>
        </table>
        """

    html = f"""
    <div style="font-family:Arial, sans-serif;font-size:14px;color:#111;">
      <p>Dear Performa Finance,</p>

      <p>Please find attached the submitted expense report for <strong>{esc(employee_name)}</strong> and accompanying receipts.</p>

      <p><strong>Details below:</strong></p>

      <table style="border-collapse:collapse;font-family:Arial, sans-serif;font-size:13px;">
        <tr><td style="padding:4px 10px 4px 0;"><strong>Employee Name:</strong></td><td style="padding:4px 0;">{esc(employee_name)}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Employee Email:</strong></td><td style="padding:4px 0;">{esc(employee_email)}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Trip Location:</strong></td><td style="padding:4px 0;">{esc(location)}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Business Purpose:</strong></td><td style="padding:4px 0;">{esc(purpose)}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Departure Date:</strong></td><td style="padding:4px 0;">{esc(departure_date)}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Return Date:</strong></td><td style="padding:4px 0;">{esc(return_date)}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Per Diem Total:</strong></td><td style="padding:4px 0;">${per_diem_total:,.2f}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Total Spend:</strong></td><td style="padding:4px 0;">${total_spend:,.2f}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Company Paid:</strong></td><td style="padding:4px 0;">${company_paid:,.2f}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Employee Paid:</strong></td><td style="padding:4px 0;">${employee_paid:,.2f}</td></tr>
        <tr><td style="padding:4px 10px 4px 0;"><strong>Reimbursement Due:</strong></td><td style="padding:4px 0;">${reimbursement_due:,.2f}</td></tr>
      </table>

      {lines_html}

      <p>Please let me know if any additional information is required.</p>

      <p>Best regards,<br>{esc(employee_name)}</p>
    </div>
    """
    return html


def send_email_with_attachments(
    subject: str,
    html_body: str,
    employee_email: str,
    attachments: List[Dict[str, Any]],
) -> int:
    """
    attachments = [{ "filename": str, "content_bytes": bytes, "mime_type": str }]
    """
    sg = SendGridAPIClient(st.secrets["SENDGRID_API_KEY"])

    msg = Mail(
        from_email=Email(st.secrets["SENDER_EMAIL"]),
        to_emails=To(st.secrets["FINANCE_EMAIL"]),
        subject=subject,
        html_content=html_body,
    )

    # CC approver and employee (employee is dynamic from the form)
    msg.add_cc(Cc(st.secrets["APPROVER_EMAIL"]))
    msg.add_cc(Cc(employee_email))

    # Add attachments
    for a in attachments:
        b = a["content_bytes"]
        encoded = base64.b64encode(b).decode("utf-8")
        msg.add_attachment(
            Attachment(
                FileContent(encoded),
                FileName(a["filename"]),
                FileType(a["mime_type"]),
                Disposition("attachment"),
            )
        )

    resp = sg.send(msg)
    return resp.status_code


# -----------------------------
# NN Template helpers
# -----------------------------
_PLACEHOLDER_PATTERN = re.compile(r"{{\s*([A-Z0-9_]+)\s*}}")


def list_docx_templates() -> List[str]:
    if not os.path.isdir(TEMPLATES_DIR):
        return []
    return sorted([f for f in os.listdir(TEMPLATES_DIR) if f.lower().endswith(".docx")])


def _replace_placeholders_in_paragraph(paragraph, values: Dict[str, str]) -> None:
    if not paragraph.runs:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    if "{{" not in full_text:
        return

    def repl(match):
        key = match.group(1)
        return str(values.get(key, match.group(0)))

    new_text = _PLACEHOLDER_PATTERN.sub(repl, full_text)
    for run in paragraph.runs:
        run.text = ""
    paragraph.runs[0].text = new_text


def replace_placeholders_in_doc(doc: Document, values: Dict[str, str]) -> None:
    for p in doc.paragraphs:
        _replace_placeholders_in_paragraph(p, values)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_placeholders_in_paragraph(p, values)


def fill_expense_table_if_present(doc: Document, expenses: List[Dict[str, Any]]) -> bool:
    """
    If the template contains a table where the header row has any of:
    Date, Category, Description, Paid By, Amount
    then it will wipe existing body rows and insert expenses.
    """
    expected = {"date", "category", "description", "paid by", "paid_by", "amount"}

    for table in doc.tables:
        if len(table.rows) < 1:
            continue

        header_cells = [c.text.strip().lower() for c in table.rows[0].cells]
        header_set = set(header_cells)

        if not (expected & header_set):
            continue

        # Remove all rows except header
        while len(table.rows) > 1:
            table._tbl.remove(table.rows[1]._tr)

        for e in expenses:
            row_cells = table.add_row().cells
            mapping = {
                "date": str(e.get("expense_date", "")),
                "category": str(e.get("category", "")),
                "description": str(e.get("description", "")),
                "paid by": str(e.get("paid_by", "")),
                "paid_by": str(e.get("paid_by", "")),
                "amount": f"{float(e.get('amount') or 0):,.2f}",
            }

            for i, h in enumerate(header_cells):
                h_clean = h.strip().lower()
                if h_clean in mapping and i < len(row_cells):
                    row_cells[i].text = mapping[h_clean]

        return True

    return False


def insert_expense_table_at_marker(doc: Document, expenses: List[Dict[str, Any]], marker: str = "{{EXPENSE_TABLE}}") -> bool:
    """
    If the template contains a paragraph with {{EXPENSE_TABLE}},
    replace it by inserting a table right after it.
    """
    for i, p in enumerate(doc.paragraphs):
        if marker in p.text:
            p.text = p.text.replace(marker, "").strip()

            table = doc.add_table(rows=1, cols=5)
            hdr = table.rows[0].cells
            hdr[0].text = "Date"
            hdr[1].text = "Category"
            hdr[2].text = "Description"
            hdr[3].text = "Paid By"
            hdr[4].text = "Amount"

            for e in expenses:
                cells = table.add_row().cells
                cells[0].text = str(e.get("expense_date", ""))
                cells[1].text = str(e.get("category", ""))
                cells[2].text = str(e.get("description", ""))
                cells[3].text = str(e.get("paid_by", ""))
                cells[4].text = f"{float(e.get('amount') or 0):,.2f}"

            doc._body._body.insert(i + 1, table._tbl)
            return True

    return False


def build_nn_docx_bytes(
    template_path: str,
    trip_info: Dict[str, Any],
    expenses: List[Dict[str, Any]],
) -> bytes:
    """
    Loads NN template docx, replaces placeholders, fills expense table, returns bytes.
    """
    values = {
        "EMPLOYEE_NAME": str(trip_info.get("employee_name", "")),
        "EMPLOYEE_EMAIL": str(trip_info.get("employee_email", "")),
        "LOCATION": str(trip_info.get("location", "")),
        "PURPOSE": str(trip_info.get("purpose", "")),
        "DEPARTURE_DATE": str(trip_info.get("departure_date", "")),
        "RETURN_DATE": str(trip_info.get("return_date", "")),
        "TRIP_DAYS": str(trip_info.get("trip_days", "")),
        "PER_DIEM_RATE": f"{float(trip_info.get('per_diem_rate') or 0):,.2f}",
        "PER_DIEM_TOTAL": f"{float(trip_info.get('per_diem_total') or 0):,.2f}",
        "TOTAL_SPEND": f"{float(trip_info.get('total_spend') or 0):,.2f}",
        "COMPANY_PAID": f"{float(trip_info.get('company_paid') or 0):,.2f}",
        "EMPLOYEE_PAID": f"{float(trip_info.get('employee_paid') or 0):,.2f}",
        "REIMBURSEMENT_DUE": f"{float(trip_info.get('reimbursement_due') or 0):,.2f}",
        "REPORT_DATE": str(date.today()),
    }

    doc = Document(template_path)
    replace_placeholders_in_doc(doc, values)

    # Fill a pre-existing table if present, otherwise insert at marker if present
    filled = fill_expense_table_if_present(doc, expenses)
    if not filled:
        insert_expense_table_at_marker(doc, expenses, marker="{{EXPENSE_TABLE}}")

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# -----------------------------
# App state
# -----------------------------
if "expenses" not in st.session_state:
    st.session_state.expenses = []


# -----------------------------
# UI
# -----------------------------
st.title("Performa Expense Report")
st.caption("Phase 1, generates Excel plus receipts for Guam, generates populated NN DOCX plus receipts for NN, emails Finance, Approver, and Employee archive.")

# Selector: Guam vs NN
st.subheader("Report Type")
report_type = st.radio(
    "Select report type",
    ["Guam", "NN (Navajo Nation)"],
    horizontal=True,
)

# If NN, show template selector (does not affect Guam behavior)
selected_nn_template = None
templates = []
if report_type == "NN (Navajo Nation)":
    templates = list_docx_templates()
    if not templates:
        st.warning("No .docx templates found in the templates folder. Add the NN template to /templates in your repo.")
    else:
        default_idx = 0
        for idx, t in enumerate(templates):
            if "navajo" in t.lower() or re.search(r"\bnn\b", t.lower()):
                default_idx = idx
                break
        selected_nn_template = st.selectbox("Select NN Template", templates, index=default_idx)

st.subheader("Trip Information")

col1, col2 = st.columns(2)
with col1:
    employee_name = st.text_input("Employee Name")
    employee_email = st.text_input("Employee Email")
with col2:
    location = st.text_input("Trip Location")
    purpose = st.text_area("Business Purpose", height=80)

col3, col4 = st.columns(2)
with col3:
    departure_date = st.date_input("Departure Date", value=date.today())
with col4:
    return_date = st.date_input("Return Date", value=date.today())

trip_days = calc_trip_days(departure_date, return_date)
per_diem_total = PER_DIEM_RATE * trip_days
st.info(f"Per diem is ${PER_DIEM_RATE:,.0f} per day, {trip_days} day(s), total ${per_diem_total:,.2f}")

st.subheader("Expenses")

with st.expander("Add an expense", expanded=True):
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        category = st.selectbox("Category", CATEGORIES)
    with c2:
        expense_date = st.date_input("Expense Date", value=date.today())
    with c3:
        paid_by = st.radio("Paid By", ["Employee", "Performa"], horizontal=True)

    description = st.text_input("Description (optional)")
    amount = st.number_input("Amount", min_value=0.0, value=0.0, step=1.0, format="%.2f")

    receipt_file = st.file_uploader(
        "Receipt (optional)",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=False,
        help="Accepted: PDF, JPG, JPEG, PNG",
    )

    if st.button("Add Expense"):
        st.session_state.expenses.append(
            {
                "category": category,
                "expense_date": expense_date,
                "paid_by": paid_by,
                "description": description,
                "amount": float(amount),
                "receipt_file": receipt_file,
            }
        )
        st.success("Expense added.")


st.subheader("Summary")

totals = calc_totals(st.session_state.expenses)
total_spend = totals["total_spend"]
company_paid = totals["company_paid"]
employee_paid = totals["employee_paid"]

reimbursement_due = per_diem_total + employee_paid

s1, s2, s3, s4 = st.columns(4)
s1.metric("Total Spend", f"${total_spend:,.2f}")
s2.metric("Company Paid", f"${company_paid:,.2f}")
s3.metric("Employee Paid", f"${employee_paid:,.2f}")
s4.metric("Reimbursement Due", f"${reimbursement_due:,.2f}")

st.subheader("Current Line Items")
if not st.session_state.expenses:
    st.write("No expenses added yet.")
else:
    for idx, e in enumerate(st.session_state.expenses, start=1):
        receipt_note = "Receipt attached" if e.get("receipt_file") else "No receipt"
        st.write(
            f"{idx}. {e['category']} on {e['expense_date']}, {e['description'] or '$0'}, "
            f"${float(e['amount']):,.2f}, Paid by {e['paid_by']}, {receipt_note}"
        )

    remove_idx = st.number_input(
        "Remove line item number",
        min_value=0,
        max_value=len(st.session_state.expenses),
        value=0,
        step=1,
        help="Enter the line number to remove, 0 means do nothing.",
    )
    if st.button("Remove Selected Line Item"):
        if remove_idx == 0:
            st.info("No line item selected.")
        else:
            st.session_state.expenses.pop(int(remove_idx) - 1)
            st.success("Removed.")


# Attachment sizing and submit
st.divider()

st.caption(
    f"Attachment limit enforced at {MAX_ATTACHMENT_MB:,.0f} MB total for receipts plus the report file."
)

submit = st.button("Submit Expense Report", type="primary")

if submit:
    # Basic validation
    missing = []
    if not employee_name.strip():
        missing.append("Employee Name")
    if not employee_email.strip():
        missing.append("Employee Email")
    if not location.strip():
        missing.append("Trip Location")
    if not purpose.strip():
        missing.append("Business Purpose")

    if return_date < departure_date:
        missing.append("Return Date must be on or after Departure Date")

    if report_type == "NN (Navajo Nation)":
        if not selected_nn_template:
            missing.append("NN Template selection")

    if missing:
        st.error("Please complete the following fields: " + ", ".join(missing))
        st.stop()

    # Trip info object (used by both Guam and NN paths)
    trip_info = {
        "employee_name": employee_name,
        "employee_email": employee_email,
        "location": location,
        "purpose": purpose,
        "departure_date": departure_date,
        "return_date": return_date,
        "trip_days": trip_days,
        "per_diem_rate": PER_DIEM_RATE,
        "per_diem_total": per_diem_total,
        "total_spend": total_spend,
        "company_paid": company_paid,
        "employee_paid": employee_paid,
        "reimbursement_due": reimbursement_due,
    }

    # Prepare attachments: Guam uses Excel, NN uses populated DOCX, receipts always included
    attachments: List[Dict[str, Any]] = []

    if report_type == "Guam":
        # -----------------------------
        # Guam flow, unchanged
        # -----------------------------
        # generate_excel should return bytes
        excel_bytes = generate_excel(trip_info, st.session_state.expenses)

        attachments.append(
            {
                "filename": f"Expense_Report_{employee_name.replace(' ', '_')}.xlsx",
                "content_bytes": excel_bytes,
                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
        )

    else:
        # -----------------------------
        # NN flow, populate template docx
        # -----------------------------
        template_path = os.path.join(TEMPLATES_DIR, selected_nn_template)
        try:
            nn_docx_bytes = build_nn_docx_bytes(
                template_path=template_path,
                trip_info=trip_info,
                expenses=st.session_state.expenses,
            )
        except Exception as ex:
            st.error(f"Failed to populate NN template: {ex}")
            st.stop()

        attachments.append(
            {
                "filename": f"NN_Expense_Report_{employee_name.replace(' ', '_')}.docx",
                "content_bytes": nn_docx_bytes,
                "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            }
        )

    # Receipts (unchanged logic)
    for i, e in enumerate(st.session_state.expenses, start=1):
        f = e.get("receipt_file")
        if f is None:
            continue
        b = bytes_from_uploaded_file(f)
        ext = (f.name.split(".")[-1] or "").lower()
        if ext in ["jpg", "jpeg"]:
            mime = "image/jpeg"
        elif ext == "png":
            mime = "image/png"
        else:
            mime = "application/pdf"

        safe_cat = str(e.get("category", "Receipt")).replace(" ", "_")
        filename = f"{i:02d}_{safe_cat}_{employee_name.replace(' ', '_')}.{ext if ext else 'pdf'}"

        attachments.append(
            {"filename": filename, "content_bytes": b, "mime_type": mime}
        )

    # Enforce max total size
    total_bytes = sum(len(a["content_bytes"]) for a in attachments)
    max_bytes = int(MAX_ATTACHMENT_MB * 1024 * 1024)
    if total_bytes > max_bytes:
        st.error(
            f"Attachments are too large: {total_bytes/1024/1024:,.2f} MB. "
            f"Limit is {MAX_ATTACHMENT_MB:,.0f} MB. Remove some receipts or compress them."
        )
        st.stop()

    # Email content (same HTML body used for both)
    subject = (
        f"Expense Report Submitted, {employee_name}, {location}, "
        f"{departure_date} to {return_date}"
    )
    if report_type == "NN (Navajo Nation)":
        subject = (
            f"NN Expense Report Submitted, {employee_name}, {location}, "
            f"{departure_date} to {return_date}"
        )

    html_body = build_email_html(
        employee_name=employee_name,
        employee_email=employee_email,
        location=location,
        purpose=purpose,
        departure_date=departure_date,
        return_date=return_date,
        per_diem_total=per_diem_total,
        total_spend=total_spend,
        company_paid=company_paid,
        employee_paid=employee_paid,
        reimbursement_due=reimbursement_due,
        expenses=st.session_state.expenses,
    )

    try:
        status_code = send_email_with_attachments(
            subject=subject,
            html_body=html_body,
            employee_email=employee_email,
            attachments=attachments,
        )

        if 200 <= int(status_code) < 300:
            st.success("Submitted successfully. Check your email for the package.")
        else:
            st.error(f"SendGrid returned status code: {status_code}")
    except Exception as ex:
        st.error(f"Email failed: {ex}")
