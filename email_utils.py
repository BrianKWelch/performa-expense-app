from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Email, To, Cc, Attachment, FileContent, FileName, FileType, Disposition
import base64
import streamlit as st


def send_email(subject, body, attachment_bytes, attachment_filename, employee_email):
    try:
        message = Mail(
            from_email=Email(st.secrets["SENDER_EMAIL"]),
            to_emails=To(st.secrets["FINANCE_EMAIL"]),
            subject=subject,
            html_content=body
        )

        # Add CC recipients
        message.add_cc(Cc(st.secrets["APPROVER_EMAIL"]))
        message.add_cc(Cc(employee_email))

        # Attach file
        encoded_file = base64.b64encode(attachment_bytes).decode()

        attachment = Attachment(
            FileContent(encoded_file),
            FileName(attachment_filename),
            FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            Disposition("attachment")
        )

        message.attachment = attachment

        sg = SendGridAPIClient(st.secrets["SENDGRID_API_KEY"])
        response = sg.send(message)

        return response.status_code

    except Exception as e:
        return str(e)
