from flask import Flask, render_template, request, flash
import smtplib
import re
import pandas as pd
from werkzeug.utils import secure_filename
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)
app.secret_key = "app_secret_key"

# Email credentials
SENDER_EMAIL = "contact@repixelx.com"
SENDER_PASSWORD = "Kriyon123/-kri"
SMTP_SERVER = "smtp.hostinger.com"

# File upload configuration
UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"xls", "xlsx", "jpg", "jpeg", "png", "pdf", "docx"}
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def is_valid_email(email):
    """Validate email format using regex."""
    email_regex = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(email_regex, email) is not None


def send_email(
    comp_email, comp_email_password, recipient, subject, body, attachment=None
):
    # Email credentials
    SENDER_EMAIL = comp_email
    SENDER_PASSWORD = comp_email_password
    SMTP_SERVER = "smtp.gmail.com"

    """Send an email to the recipient with or without an attachment."""
    try:
        # Create the email message container
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = recipient
        msg["Subject"] = subject

        # Attach the email body to the email
        msg.attach(MIMEText(body, "plain"))

        if attachment:
            # Attach the file
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={secure_filename(attachment.filename)}",
            )
            msg.attach(part)

        # Connect to SMTP server and send the email
        with smtplib.SMTP(SMTP_SERVER, 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            text = msg.as_string()
            server.sendmail(SENDER_EMAIL, recipient, text)

        return True
    except Exception as e:
        print(f"Error sending email to {recipient}: {e}")
        return False


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        comp_email = request.form.get("comp-email")
        comp_email_password = request.form.get("email-pass")
        subject = request.form.get("subject")
        body = request.form.get("body")
        file = request.files.get("file")
        attachment = request.files.get("attachment")

        if not file or file.filename == "":
            flash("Please upload a valid Excel file.", "danger")
            return render_template("index.html")

        if not (
            "." in file.filename
            and file.filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
        ):
            flash("Invalid file type. Please upload an Excel file.", "danger")
            return render_template("index.html")

        filepath = os.path.join(
            app.config["UPLOAD_FOLDER"], secure_filename(file.filename)
        )
        file.save(filepath)

        try:
            # Read emails from the uploaded Excel file
            df = pd.read_excel(filepath)
            if "Email" not in df.columns:
                flash(
                    "The uploaded file must contain a column named 'Email'.", "danger"
                )
                return render_template("index.html")

            emails = df["Email"].dropna().tolist()

            for email in emails:
                if is_valid_email(email):
                    if attachment:
                        success = send_email(
                            comp_email,
                            comp_email_password,
                            email,
                            subject,
                            body,
                            attachment,
                        )
                    else:
                        success = send_email(
                            comp_email, comp_email_password, email, subject, body
                        )
                    if success:
                        flash(f"Email sent to {email}", "success")
                    else:
                        flash(f"Failed to send email to {email}", "danger")
                else:
                    flash(f"Invalid email address: {email}", "danger")
        except Exception as e:
            flash(f"Error processing file: {str(e)}", "danger")
        finally:
            os.remove(filepath)  # Clean up the uploaded file

    return render_template("index.html")


