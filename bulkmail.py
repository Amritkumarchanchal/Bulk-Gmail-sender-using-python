import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# change these as per use
your_email = "amritkrchanchal@gmail.com"
your_password = "upasfpogybtfgjadphd"

# Establish SMTP connection (Gmail Example)
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login(your_email, your_password)

# Read spreadsheet data
email_list = pd.read_excel('student.xlsx')

# Get names and emails
names = email_list['first_name']
emails = email_list['email']
passwords = email_list['password']

# Email Template (HTML with CSS)
html_template = """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f4f4; padding: 20px; }
        .container { max-width: 600px; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0px 0px 10px rgba(0,0,0,0.1); margin: auto; }
        .header { text-align: center; color: #007BFF; }
        .footer { font-size: 12px; color: gray; margin-top: 20px; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <h2 class="header">Welcome to Vicharanshala! 🎉</h2>
        <p>Hello <strong>{USER_NAME}</strong>,</p>
        <p>We are thrilled to have you on board! Below are your login credentials:</p>
        <p><strong>Email:</strong> {USER_EMAIL}</p>
        <p><strong>Password:</strong> {USER_PASSWORD}</p>
        <p><strong>Note:</strong> Please change your password after logging in for security purposes.</p>
        <p>Best Regards,</p>
        <p><strong>Vicharanshala Team</strong></p>
        <div class="footer">
            <p>Need help? Contact us at <a href="mailto:support@vicharanshala.com">support@vicharanshala.com</a></p>
        </div>
    </div>
</body>
</html>
"""

# Iterate through recipients and send emails
for i in range(len(emails)):
    name = names[i]
    email = emails[i]
    password = passwords[i]  # Fetch password from CSV
    
    # Customize HTML with user details
    personalized_html = html_template.replace("{USER_NAME}", name).replace("{USER_EMAIL}", email).replace("{USER_PASSWORD}", password)
    
    # Create Email
    msg = MIMEMultipart()
    msg["From"] = your_email
    msg["To"] = email
    msg["Subject"] = "Welcome to Vicharanshala – Your Login Credentials"
    msg.attach(MIMEText(personalized_html, "html"))

    # Send Email
    server.sendmail(your_email, email, msg.as_string())
    print(f"Email sent to {name} ({email}) with password: {password}")

# Close SMTP connection
server.quit()
print("All emails sent successfully!")