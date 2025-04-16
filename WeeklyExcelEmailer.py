import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Email configuration
SMTP_PORT = 465
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "sender@example.com")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL", "receiver@example.com")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.example.com")
REPORTS_DIR = os.getenv("REPORTS_DIR", "path/to/reports")

# Last week range
def get_last_week_range():
    today = datetime.now()
    days_to_last_monday = today.weekday()
    last_monday = today - timedelta(days=days_to_last_monday)
    last_friday = last_monday + timedelta(days=4)
    return last_monday.strftime("%m/%d/%Y"), last_friday.strftime("%m/%d/%Y")

# Date formatting for filename
def format_date_for_filename(date_str):
    date_obj = datetime.strptime(date_str, "%m/%d/%Y")
    return date_obj.strftime("%d%b%Y")

# Get date range
start_date, end_date = get_last_week_range()
start_str = format_date_for_filename(start_date)
end_str = format_date_for_filename(end_date)

# Email content
subject = f"Weekly Utilization Reports ({start_str}-{end_str})"
body = f"""\
Dear Team,

Attached are the CPU and disk utilization reports from {start_str} to {end_str}.

Regards,
Network Admin
"""

# Create MIME email object
message = MIMEMultipart()
message["From"] = SENDER_EMAIL
message["To"] = RECEIVER_EMAIL
message["Subject"] = subject
message.attach(MIMEText(body, "plain"))

# Get .xlsx files
xlsx_files = [f for f in os.listdir(REPORTS_DIR) if f.endswith('.xlsx') and f"{start_str}-{end_str}" in f]

# Attach files
for file_name in xlsx_files:
    file_path = os.path.join(REPORTS_DIR, file_name)
    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={file_name}",
            )
            message.attach(part)
        print(f"Attached: {file_name}")
    except FileNotFoundError:
        print(f"Error: File {file_path} not found.")
        continue

# Exit if no files
if not xlsx_files:
    print(f"Error: No .xlsx files found for {start_str}-{end_str}.")
    exit()

# Send email
try:
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
        server.sendmail(SENDER_EMAIL, [RECEIVER_EMAIL], message.as_string())
        print("Email sent successfully!")
except Exception as e:
    print(f"Failed to send email: {e}")