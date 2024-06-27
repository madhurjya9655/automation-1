import pandas as pd
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Read the Excel file
data = pd.read_excel("Students.xlsx")
print(type(data))

# Extract the email column
email_col = data.get("email")
list_of_emails = list(email_col)
print(list_of_emails)

try:
    # Connect to the SMTP server
    server = sm.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("bitumixkanika@gmail.com", "kedv sgdw brai ubez")

    from_ = "bitumixkanika@gmail.com"
    subject = "This is just a testing message"

    # Prepare the email content
    html = '''
    <html>
    <head></head>
    <body>
        <p>This is a test email. Please ignore.</p>
    </body>
    </html>
    '''

    # Send email to each recipient
    for to_ in list_of_emails:
        # Create a new MIME message for each recipient
        message = MIMEMultipart("alternative")
        message['Subject'] = subject
        message["From"] = from_
        message["To"] = to_

        # Attach the HTML content to the email
        mime_text = MIMEText(html, "html")
        message.attach(mime_text)

        # Send the email
        server.sendmail(from_, to_, message.as_string())

    # Close the server connection
    server.quit()

except Exception as e:
    print(f"An error occurred: {e}")
