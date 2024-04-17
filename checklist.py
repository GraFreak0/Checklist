import os
import datetime
import win32com.client as win32

def get_current_date():
    current_date = datetime.date.today()
    return current_date

current_date = get_current_date()
print("Current date:", current_date)

# Convert the date to a string in the desired format
current_date_str = current_date.strftime("%Y-%m-%d")

# Define the directory containing the file you want to send
directory_path = r'd:\My Documents\ijohnson\My Documents'

# Check if the directory exists
if not os.path.exists(directory_path):
    print(f"Directory '{directory_path}' does not exist.")
    exit()

# List files in the directory
files = os.listdir(directory_path)

# Check if there are files in the directory
if not files:
    print(f"No files found in '{directory_path}'.")
    exit()

# Choose the file to send (e.g., the first file in the list)
file_to_send = os.path.join(directory_path, r'd:\My Documents\ijohnson\My Documents\Enterprise_Report_Checklist.xlsx')

# Create an Outlook instance
outlook = win32.Dispatch('outlook.application')

# Create a new mail item
mail = outlook.CreateItem(0)  # 0 represents olMailItem, i.e., an email

# Set the email properties
mail.Subject = f'CHECKLIST_{current_date_str}'
mail.Body = 'Please find the attached file.'
mail.To = 'AllENG-ITRiskandCompliance@ecobank.com'
mail.Cc = 'AllENG-ITEnterpriseReport@ecobank.com'

# Attach the file
attachment = mail.Attachments.Add(file_to_send)

# Send the email
mail.Send()

print(f"Email sent with the attachment: {file_to_send}")