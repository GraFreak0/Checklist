import os
import datetime
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox

def send_email():
    # Get the current date
    current_date = datetime.date.today()
    current_date_str = current_date.strftime("%Y-%m-%d")

    # Define the directory containing the file you want to send
    directory_path = r'd:\My Documents\ijohnson\My Documents'

    # Define the file to send
    file_to_send = os.path.join(directory_path, 'Enterprise_Report_Checklist.xlsx')

    # Check if the file exists
    if not os.path.exists(file_to_send):
        messagebox.showerror("Error", "File not found.")
        return

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

    messagebox.showinfo("Success", f"Email sent with the attachment: {file_to_send}")

    # Update the label text to display the file path
    label_file.config(text=file_to_send)

# Create the GUI
root = tk.Tk()
root.title("Email Sender")

# Label for To and From emails
label_to = tk.Label(root, text="To: AllENG-ITRiskandCompliance@ecobank.com")
label_to.pack()
label_cc = tk.Label(root, text="Cc: AllENG-ITEnterpriseReport@ecobank.com")
label_cc.pack()

# Label for displaying the file path
file_to_send = os.path.join(r'd:\My Documents\ijohnson\My Documents', 'Enterprise_Report_Checklist.xlsx')
label_file = tk.Label(root, text=file_to_send)
label_file.pack()

# Automatically send the email when the program launches
send_email()

# Run the GUI
root.mainloop()