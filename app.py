import streamlit as st
from datetime import date
from excel_generator import generate_excel
from email_utils import send_email

st.set_page_config(page_title="Performa Expense Report", layout="wide")

PER_DIEM_RATE = float(st.secrets.get("PER_DIEM_RATE", 100))
MAX_ATTACHMENT_MB = float(st.secrets.get("MAX_ATTACHMENT_MB", 18))

st.title("Performa Expense Report")

st.caption("Phase 1, generates Excel plus receipts, emails Finance, Approver, and Employee archive.")

# Trip info
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
    depart = st.date_input("Departure Date", value=date.today())
with col4:
    return_date = st.date_input("Return Date", value=date.today())

if return_date < depart:
    st.error("Return date cannot be before departure date.")
    st.stop()

per_diem_days = (return_date - depart).days + 1
per_diem_total = PER_DIEM_RATE * per_diem_days

st.info(f"Per diem is ${PER_DIEM_RATE:.0f} per day, {per_diem_days} day(s), total ${per_diem_total:.2f}")

# Expenses
st.subheader("Expenses")

if "expenses" not in st.session_state:
    st.session_state.expenses = []

with st.expander("Add an expense", expanded=True):
    with st.form("add_expense_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 1, 2])
        with c1:
            category = st.selectbox("Category", ["Airfare", "Airport Parking", "Taxi/Uber", "Hotel", "Rental Car", "Gas", "Other"])
        with c2:
            expense_date = st.date_input("Expense Date", value=depart)
        with c3:
            paid_by = st.radio("Paid By", ["Employee", "Performa"], horizontal=True)

        description = st.text_input("Description")
        amount = st.number_input("Amount", min_value=0.0, step=0.01, format="%.2f")
        receipt = st.file_uploader("Receipt (optional)", type=["pdf", "jpg", "jpeg", "png"])

        add_btn = st.form_submit_button("Add Expense")

        if add_btn:
            st.session_state.expenses.append({
                "category": category,
                "date": expense_date,
                "description": description,
                "amount": float(amount),
                "paid_by": paid_by,
                "receipt_name": receipt.name if receipt else None,
                "receipt_file": receipt
            })

# Show expenses and summary
if st.session_state.expenses:
    reimbursable_expenses = sum(e["amount"] for e in st.session_state.expenses if e["paid_by"] == "Employee")
    company_paid = sum(e["amount"] for e in st.session_state.expenses if e["paid_by"] == "Performa")
    total_spend = sum(e["amount"] for e in st.session_state.expenses)
    reimbursement_due = reimbursable_expenses + per_diem_total

    st.markdown("### Summary")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Total Spend", f"${total_spend:,.2f}")
    s2.metric("Company Paid", f"${company_paid:,.2f}")
    s3.metric("Employee Paid", f"${reimbursable_expenses:,.2f}")
    s4.metric("Reimbursement Due", f"${reimbursement_due:,.2f}")

    st.markdown("### Current Line Items")
    for idx, e in enumerate(st.session_state.expenses, start=1):
        left, right = st.columns([4, 1])
        with left:
            receipt_label = e["receipt_name"] if e["receipt_name"] else "No receipt"
            st.write(f"{idx}. {e['category']} on {e['date']} , {e['description']} , ${e['amount']:.2f} , Paid by {e['paid_by']} , {receipt_label}")
        with right:
            if st.button("Remove", key=f"rm_{idx}"):
                st.session_state.expenses.pop(idx - 1)
                st.rerun()

    st.divider()

    submit_col1, submit_col2 = st.columns([2, 3])
    with submit_col1:
        submit = st.button("Submit Expense Report", type="primary")
    with submit_col2:
        st.caption(f"Attachment limit enforced at {MAX_ATTACHMENT_MB:.0f} MB total for receipts plus the Excel file.")

    if submit:
        # Basic validation
        if not employee_name.strip() or not employee_email.strip() or not location.strip():
            st.error("Please fill in Employee Name, Employee Email, and Trip Location before submitting.")
            st.stop()

        # Build receipt attachments, enforce size limit
        attachments: list[tuple[str, bytes]] = []
        total_size_mb = 0.0

        for e in st.session_state.expenses:
            rf = e.get("receipt_file")
            if rf is not None:
                file_bytes = rf.getvalue()
                size_mb = len(file_bytes) / (1024 * 1024)
                total_size_mb += size_mb
                attachments.append((rf.name, file_bytes))

        if total_size_mb > MAX_ATTACHMENT_MB:
            st.error(f"Receipts total {total_size_mb:.2f} MB, which exceeds the {MAX_ATTACHMENT_MB:.0f} MB limit. Compress receipts or remove some, then resubmit.")
            st.stop()

        # Generate Excel
        excel_stream = generate_excel(
            employee_name=employee_name,
            employee_email=employee_email,
            location=location,
            depart=depart,
            return_date=return_date,
            purpose=purpose,
            per_diem_rate=PER_DIEM_RATE,
            per_diem_days=per_diem_days,
            expenses=[{k: v for k, v in e.items() if k != "receipt_file"} for e in st.session_state.expenses]
        )

        report_name = f"ExpenseReport_{employee_name.replace(' ', '')}_{location.replace(' ', '')}_{depart.strftime('%Y%m%d')}-{return_date.strftime('%Y%m%d')}.xlsx"
        attachments.append((report_name, excel_stream.getvalue()))

        subject = f"Expense Report Submitted, {employee_name}, {location}, {depart} to {return_date}, Reimbursement ${reimbursement_due:,.2f}"
        body = (
            "Expense Report Submitted\n\n"
            f"Employee: {employee_name}\n"
            f"Employee Email: {employee_email}\n"
            f"Trip: {location}\n"
            f"Dates: {depart} to {return_date}\n"
            f"Per Diem: ${per_diem_total:,.2f}\n"
            f"Employee Paid Expenses: ${reimbursable_expenses:,.2f}\n"
            f"Company Paid Expenses: ${company_paid:,.2f}\n"
            f"Reimbursement Due: ${reimbursement_due:,.2f}\n\n"
            "Excel report attached. Receipts attached.\n"
        )

        try:
            send_email(subject=subject, body=body, attachments=attachments)
            st.success("Submitted successfully. Check your email for the package.")
            st.session_state.expenses = []
        except Exception as ex:
            st.error("Submit failed. Open Streamlit logs to see the error details.")
            st.exception(ex)

else:
    st.warning("Add at least one expense line item to enable submit.")
