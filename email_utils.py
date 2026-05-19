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

        message.add_cc(Cc(st.secrets["APPROVER_EMAIL"]))
        message.add_cc(Cc(employee_email))

        encoded_file = base64.b64encode(attachment_bytes).decode()

        attachment = Attachment(
            FileContent(encoded_file),
            FileName(attachment_filename),
            FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            Disposition("attachment")
        )

        message.attachment = attachment

        from sendgrid_secrets import sendgrid_api_key

        sg = SendGridAPIClient(sendgrid_api_key())
        response = sg.send(message)

        return response.status_code

    except Exception as e:
        return str(e)
