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
wb = openpyxl.load_workbook(r'C:\Users\bhuva\OneDrive\Desktop\Python course\Automation-project\May-2024- Windows Server Patching-updated.xlsx')
sheet = wb.active

# Group servers by date
servers_by_date = defaultdict(list)

for row in sheet.iter_rows(min_row=2, values_only=True):
    if len(row) >= 4:
        server_name, maintenance_window, date_str, app_team_mail_id = row
        if isinstance(date_str, datetime):
            date_str = date_str.strftime('%d-%m-%Y')
        date_obj = datetime.strptime((date_str), '%d-%m-%Y')
        reminder_date = date_obj - timedelta(days=1)
        if reminder_date.date() == datetime.now().date():
            subject = f'Reminder: Windows patching for {server_name} of {app_team_mail_id} tomorrow'
            message = f"Hi,\n\nThis is a reminder about the maintenance window for the server {server_name} scheduled for tomorrow, {date_str}. The maintenance window is from {maintenance_window}.\n\nPlease ensure that the necessary actions are taken.\n\nRegards,\nYour Team"
            send_reminder_email("Bhuvan.Gowda@staples.com", subject, message)
    else:
        print(f"Skipping row: {row}. Expected at least 4 values.")
print("Reminder emails sent successfully.")
