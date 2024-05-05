from openpyxl import load_workbook
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os

import boto3
from botocore.exceptions import ClientError

load_dotenv()

aws_access_key_id = os.getenv('AWS_ACCESS_KEY_ID')
aws_secret_access_key = os.getenv('AWS_SECRET_ACCESS_KEY')

def get_secret():

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
        get_secret_value_response = client.get_secret_value(
            SecretId=secret_name
        )
    except ClientError as e:
        # For a list of exceptions thrown, see
        # https://docs.aws.amazon.com/secretsmanager/latest/apireference/API_GetSecretValue.html
        raise e

    secret = eval(get_secret_value_response['SecretString'])

    return secret

secret = get_secret()
# print('type secret = ', type(secret))
# print('secret = ', secret)

# Load the workbook and select the active worksheet
file_path = 'Arjun-BOI-Timesheet-04_20_24.xlsx'  
workbook = load_workbook(file_path)
sheet = workbook.active

today = datetime.today()
# Calculate the number of days to subtract to go back to the most recent Saturday
days_since_saturday = today.weekday() - 5  # Subtract the number of days since the last Saturday (5 is the index for Saturday in ISO)
# Calculate the date of the most recent Saturday
most_recent_saturday = today - timedelta(days=days_since_saturday)

if today.weekday() == 4 and today.hour >= 16:
    print("Today is Friday between 4pm and 11:59pm.")
    saturday_to_use = most_recent_saturday
    # saturday_to_use = today - timedelta(days=(today.weekday() + 2) % 7)
else:
    # Calculate the date of the saturday of the week we finished
    saturday_to_use = most_recent_saturday - timedelta(weeks=1)
    

# Update the cell B4 with the last Saturday date
sheet['B4'].value = saturday_to_use.strftime('%m/%d/%Y')  # Adjust the format if necessary

# Save the changes to the Excel file
final_file_path = file_path.replace("04_20_24", saturday_to_use.strftime('%m_%d_%y')) 
workbook.save(final_file_path)
print('wrote to workbook!')

############################################## Send the email

# Replace with your email and password
your_email = secret['secondary_email']
your_password = secret['secondary_email_app_pw']
# Replace with the recipient's email
recipient_email = "arjungup745@gmail.com"

# Email Subject and Body
subject = "Test Email with Attachment"
body = "This is a test email with an attachment sent from Python."

# File to attach
file_path = final_file_path

# Create the email
message = MIMEMultipart()
message['From'] = your_email
message['To'] = recipient_email
message['Subject'] = subject

# Attach the email body
message.attach(MIMEText(body, 'plain'))
print('got here')
# Attach the file
filename = os.path.basename(file_path)
attachment = open(file_path, "rb")

part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f"attachment; filename= {filename}")

message.attach(part)
print('got here 2')
# Send the email
try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(your_email, your_password)
        print('got here 3')
        server.sendmail(your_email, recipient_email, message.as_string())

    print("Email sent successfully!")
except Exception as e:
    print(f"Error: {e}")


###################### Steps
"""


* change the secrets

* run via lambda so that doesn't matter if your computer is on
* ultimately could have some logic that checks which emails you sent previously
* add some boto that pulls your costs


Ultimate dream -- send email asking "did you work 40 hours this week?" if yes, then send email to BOI timesheets
if no, then send filled out timesheet and say "please modify"

done
1) figure out date logic -- handle if you send previous email or not.... or just default to being careful and only sending one for the previous week unless it's Friday late pm the previous week
    if it is friday afternoon, send for last week. If it is anytime before Friday afternoon, send for the previous week. If it is Monday or any other day really, send for 2 saturdays ago, because presumably you missed it
* containerize so it just always works?

"""