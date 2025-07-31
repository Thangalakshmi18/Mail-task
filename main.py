import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Load Excel file
excel_path = 'Baxter database template.xlsx'
df = pd.read_excel(excel_path, header=1)
df.columns = df.columns.str.strip()

# Email setup
smtp_server = 'smtp.gmail.com'
smtp_port = 587
sender_email = 'thangalakshmi1802@gmail.com'
sender_password = 'djxl ldgh ermy hmcq'  # Use Gmail App Password

# Loop through each row
for index, row in df.iterrows():
    recipient_email = row.get('Email ID')
    if pd.isna(recipient_email):
        continue  # skip rows with no email

    try:
        cas_no = row.get('CAS No', 'N/A')
        product_code = row.get('PRODUCT CODE', 'N/A')
        item_description = row.get('Item Description', 'N/A')

        subject = f"Status of {product_code}"
        body = f"""
        Hi,

        This is to inform you that the status of your product
        Cas No: {cas_no}
        Product code: {product_code} 
        Item description {item_description}


        Thank you,
        Your Company
        """

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print(f"✅ Email sent to {recipient_email}")

    except Exception as e:
        print(f"❌ Failed to send email to {recipient_email}: {e}")
