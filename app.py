import base64
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

from excel_generator import generate_excel


# -----------------------------
# Config and constants
# -----------------------------
st.set_page_config(page_title="Performa Expense Report", layout="wide")

PER_DIEM_RATE = float(st.secrets.get("PER_DIEM_RATE", 100))
MAX_ATTACHMENT_MB = float(st.secrets.get("MAX_ATTACHMENT_MB", 18))

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
# App state
# -----------------------------
if "expenses" not in st.session_state:
    st.session_state.expenses = []


# -----------------------------
# UI
# -----------------------------
st.title("Performa Expense Report")
st.caption("Phase 1, generates Excel plus receipts, emails Finance, Approver, and Employee archive.")

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
    f"Attachment limit enforced at {MAX_ATTACHMENT_MB:,.0f} MB total for receipts plus the Excel file."
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

    if missing:
        st.error("Please complete the following fields: " + ", ".join(missing))
        st.stop()

    # Build Excel
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

    # generate_excel should return bytes
    excel_bytes = generate_excel(trip_info, st.session_state.expenses)

    # Prepare attachments: Excel + receipts
    attachments: List[Dict[str, Any]] = []
    attachments.append(
        {
            "filename": f"Expense_Report_{employee_name.replace(' ', '_')}.xlsx",
            "content_bytes": excel_bytes,
            "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
    )

    # Receipts
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

    # Email content
    subject = (
        f"Expense Report Submitted, {employee_name}, {location}, "
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
