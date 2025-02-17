import json
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Load JSON data from a file
input_json_file = "cpu_usage_data_promql.json"  # Replace with your JSON file name
with open(input_json_file, "r") as file:
    data = json.load(file)

# Extract CPU busy data
cpu_busy_data = data["Quick CPU / Mem / Disk"]["Cpu Busy"]

# Create a new Excel workbook and sheet
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "CPU Usage"

# Add headers
sheet.append(["KPI", "Date and Time", "CPU Usage"])

# Write data to Excel with KPI column
for entry in cpu_busy_data:
    sheet.append(["CPU Busy", entry[0], entry[1]])

# Save the Excel file
output_file = "cpu_usage_with_kpi.xlsx"
workbook.save(output_file)
print(f"Output saved in {output_file} file.")

# Convert Excel data to HTML table for email body
html_table = "<table border='1' style='border-collapse: collapse;'>"
html_table += "<tr><th>KPI</th><th>Date and Time</th><th>CPU Usage</th></tr>"

for entry in cpu_busy_data:
    html_table += f"<tr><td>CPU Busy</td><td>{entry[0]}</td><td>{entry[1]}</td></tr>"

html_table += "</table>"

# Email configuration
sender_email = "amitesh.singh@coredge.io"  # Replace with your email
receiver_emails = [
    "yashwant.singh@coredge.io"
    # "gyan.bhatnagar@coredge.io"
]  # Add multiple recipient emails here 

# CC emails
cc_emails = [
    "amitesh.singh@coredge.io"
]

email_password = "xyz"  # Replace with your Outlook app password

subject = "CPU Usage Report"

# Create the email
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = ", ".join(receiver_emails)  # Join recipient list as a string of emails
message["Cc"] = ", ".join(cc_emails)  # Add CC recipients
message["Subject"] = subject

# Email body with HTML table
body = f"""
<p>Please find the attached CPU usage report in Excel format.</p>
<p>Below is the CPU usage data in table format:</p>
{html_table}
"""
message.attach(MIMEText(body, "html"))

# Attach the Excel file
with open(output_file, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={output_file}",
    )
    message.attach(part)

# Send the email
try:
    all_recipients = receiver_emails + cc_emails  # Combine "To" and "Cc" recipients
    with smtplib.SMTP("smtp.office365.com", 587) as server:  # Use port 587 for Outlook TLS
        server.starttls()  # Start TLS encryption
        server.login(sender_email, email_password)  # Login with Outlook email and App password
        server.sendmail(sender_email, all_recipients, message.as_string())  # Send to both "To" and "Cc"
    print("Email sent successfully to all recipients!")
except smtplib.SMTPAuthenticationError as e:
    print("Authentication Error: Check your username and password or App Password settings.")
    print(e)
except Exception as e:
    print("Error sending email:", e)
