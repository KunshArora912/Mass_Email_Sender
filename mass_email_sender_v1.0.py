import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import warnings

# Function to send email
def send_email(to_address, to_name, subject, body):
    # Set up SMTP server
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = ''  # Update with your Gmail email address
    app_password = ''  # Update with the generated App Password

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, app_password)

    # Create the email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = to_address
    message['Subject'] = subject

    # Replace the placeholder with the actual name
    personalized_body = body.replace('___', to_name)

    # Attach the personalized body of the email
    message.attach(MIMEText(personalized_body, 'plain'))

    # Send the email
    server.sendmail(sender_email, to_address, message.as_string())

    # Quit the server
    server.quit()

# Read Excel file
excel_file_path = ''  # Update with the path to your Excel file
column_name_email = 'Email'  # Update with the name of the column containing email addresses
column_name_name = 'Name'  # Update with the name of the column containing names

df = pd.read_excel(excel_file_path)

# Email content
email_subject = 'Test mail'

# Define email_body
email_body = """Dear ___:

My name is Kunsh Arora. 

Cordially,
Kunsh Arora
"""

# Send email to each address with personalized greeting
for index, row in pd.read_excel(excel_file_path).iterrows():
    email_address = row['Email']
    to_name = row['Name']
    send_email(email_address, to_name, email_subject, email_body)

print("Emails sent successfully.")
