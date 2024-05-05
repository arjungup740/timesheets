import os
import boto3
from botocore.exceptions import ClientError
from openpyxl import load_workbook
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def lambda_handler(event, context):
    # Load environment variables
    aws_access_key_id = os.environ.get('AWS_ACCESS_KEY_ID')
    aws_secret_access_key = os.environ.get('AWS_SECRET_ACCESS_KEY')

    # Retrieve secret from Secrets Manager
    secret = get_secret(aws_access_key_id, aws_secret_access_key)
    
    # Process Excel file
    file_path = 'Arjun-BOI-Timesheet-04_20_24.xlsx'
    final_file_path = process_excel(file_path)
    
    # Send email
    send_email(secret, final_file_path)

def get_secret(aws_access_key_id, aws_secret_access_key):
    secret_name = "secondary_email_creds"
    region_name = "us-west-2"

    # Create a Secrets Manager client
    session = boto3.session.Session(
        aws_access_key_id=aws_access_key_id,
        aws_secret_access_key=aws_secret_access_key,
    )
    client = session.client(
        service_name='secretsmanager',
        region_name=region_name
    )

    try:
        get_secret_value_response = client.get_secret_value(SecretId=secret_name)
    except ClientError as e:
        raise e

    secret = eval(get_secret_value_response['SecretString'])

    return secret

def process_excel(file_path):
    workbook = load_workbook(file_path)
    sheet = workbook.active

    today = datetime.today()
    days_since_saturday = today.weekday() - 5
    most_recent_saturday = today - timedelta(days=days_since_saturday)

    if today.weekday() == 4 and today.hour >= 16:
        saturday_to_use = most_recent_saturday
    else:
        saturday_to_use = most_recent_saturday - timedelta(weeks=1)

    sheet['B4'].value = saturday_to_use.strftime('%m/%d/%Y')
    final_file_path = file_path.replace("04_20_24", saturday_to_use.strftime('%m_%d_%y'))
    workbook.save(final_file_path)

    return final_file_path

def send_email(secret, file_path):
    your_email = secret['secondary_email']
    your_password = secret['secondary_email_app_pw']
    recipient_email = "arjungup745@gmail.com"

    subject = "Test Email with Attachment"
    body = "This is a test email with an attachment sent from Python."

    message = MIMEMultipart()
    message['From'] = your_email
    message['To'] = recipient_email
    message['Subject'] = subject

    message.attach(MIMEText(body, 'plain'))

    filename = os.path.basename(file_path)
    attachment = open(file_path, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {filename}")

    message.attach(part)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(your_email, your_password)
            server.sendmail(your_email, recipient_email, message.as_string())

        print("Email sent successfully!")
    except Exception as e:
        print(f"Error: {e}")

