subject = f"Expense Report Submitted – {employee_name} – {location} – {departure_date} to {return_date}"

body = f"""
<p>Dear Performa Finance,</p>

<p>Please find attached the submitted expense report for <strong>{employee_name}</strong> and accompanying receipts.</p>

<p><strong>Details below:</strong></p>

<ul>
<li><strong>Employee Name:</strong> {employee_name}</li>
<li><strong>Employee Email:</strong> {employee_email}</li>
<li><strong>Trip Location:</strong> {location}</li>
<li><strong>Business Purpose:</strong> {purpose}</li>
<li><strong>Departure Date:</strong> {departure_date}</li>
<li><strong>Return Date:</strong> {return_date}</li>
<li><strong>Total Spend:</strong> ${total_spend:.2f}</li>
<li><strong>Company Paid:</strong> ${company_paid:.2f}</li>
<li><strong>Employee Paid:</strong> ${employee_paid:.2f}</li>
<li><strong>Reimbursement Due:</strong> ${reimbursement_due:.2f}</li>
</ul>

<p>Please let me know if any additional information is required.</p>

<p>Best regards,<br>
{employee_name}</p>
"""
