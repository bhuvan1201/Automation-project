import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from collections import defaultdict

# Function to send reminder emails
def send_reminder_email(receiver_email, subject, message):
    # Email configuration
    sender_email = "bhuvang100@gmail.com"
    sender_password = "vemy qcdi zjhy ahjb"
    
    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    
    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)

# Load Excel file
wb = openpyxl.load_workbook(r'C:\Users\bhuva\OneDrive\Desktop\Python course\Automation\Daily monitering tasks - HR and Corps.xlsx')
sheet = wb.active

# Group servers by date
servers_by_date = defaultdict(list)

for row in sheet.iter_rows(min_row=2, values_only=True):
    if len(row) >= 3:
        server_name, application, date_str, *_ = row
        date_obj = datetime.strptime(str(date_str), '%Y-%m-%d %H:%M:%S')
        reminder_date = date_obj - timedelta(days=1)
        if reminder_date.date() == datetime.now().date():
            subject = f'Reminder: Patching for {server_name} tomorrow'
            message = f"Hi,\n\nThis is a reminder that the server {server_name} with the application {application} is scheduled for patching tomorrow, on {date_str}.\n\nRegards,\nYour Team"
            send_reminder_email('bhuvanm0806@gmail.com', subject, message)
    else:
        print(f"Skipping row: {row}. Expected at least 3 values.")
print("Reminder emails sent successfully.")
