import win32com.client as win32
import openpyxl
import time

# Path to your excel file
excel_file_path = ('email_addresses.xlsx')

# Outlook Application
outlook = win32.Dispatch("Outlook.Application")

# Open the Excel file
workbook = openpyxl.load_workbook(excel_file_path)
worksheet = workbook.active

# Get the email addresses from the Excel file
email_addresses = [cell.value for cell in worksheet["A"] if cell.value]

# Set the interval in seconds
interval = 10

# Loop through the email addresses
for email_address in email_addresses:
    # Create a new email
    mail = outlook.CreateItem(0)

    # Set the recipients
    mail.To = email_address

    # Set CC recipients (if any)
    cc_recipients = "email1@example.com, email2@example.com"
    mail.CC = cc_recipients

    # Set the subject and body
    mail.Subject = "Your email subject"
    mail.Body = "Your email body"

    # Display the email (optional)
    mail.Display(True)

    # Uncomment the line below if you want to send the email automatically
    # mail.Send()

    # Pause for specified interval
    time.sleep(interval)

# Close the Excel file
workbook.close()