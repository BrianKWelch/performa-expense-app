from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import base64
import streamlit as st

def send_email(subject: str, body: str, attachments: list[tuple[str, bytes]]) -> None:
    to_emails = [
        st.secrets["FINANCE_EMAIL"],
        st.secrets["APPROVER_EMAIL"],
        st.secrets["EMPLOYEE_ARCHIVE_EMAIL"],
    ]

    message = Mail(
        from_email=st.secrets["SENDER_EMAIL"],
        to_emails=to_emails,
        subject=subject,
        plain_text_content=body
    )

    for filename, file_bytes in attachments:
        encoded = base64.b64encode(file_bytes).decode("utf-8")
        attachment = Attachment(
            FileContent(encoded),
            FileName(filename),
            FileType("application/octet-stream"),
            Disposition("attachment")
        )
        message.add_attachment(attachment)

    sg = SendGridAPIClient(st.secrets["SENDGRID_API_KEY"])
    sg.send(message)
