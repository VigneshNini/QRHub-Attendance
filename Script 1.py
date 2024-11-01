import qrcode
import smtplib
import tkinter as tk
from tkinter import Entry, Listbox, Scrollbar, filedialog
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import tempfile
import os
import pandas as pd
from datetime import datetime

successful_emails_count = 0

def send_email_with_qr(student_name, student_id, venue, mentioned_time, to_email, from_email, email_password, email_listbox):
    global successful_emails_count

    qr_data = f"Name: {student_name}\nStudent ID: {student_id}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_file:
        img.save(temp_file.name, format="PNG")

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = "Your QR Code"

    email_body = f"Dear {student_name},\n\nYour Attendance QR code is attached. Kindly come to the venue at {mentioned_time} in the {venue} and ensure you scan the QR code for attendance recording.\n\nBest regards, SRM IST"

    msg.attach(MIMEText(email_body, 'plain'))

    with open(temp_file.name, 'rb') as attachment:
        img = MIMEImage(attachment.read(), name=os.path.basename(temp_file.name))
        msg.attach(img)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, email_password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()

    email_listbox.insert("end", to_email)
    email_listbox.see("end")

    successful_emails_count += 1
    success_label.config(text=f"Successful Emails: {successful_emails_count}")

def send_qr_codes():
    global from_email, email_password, student_data, email_index, venue, mentioned_time

    from_email = email_entry.get()
    email_password = password_entry.get()
    venue = venue_entry.get()
    mentioned_time = time_entry.get()

    excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if not excel_file:
        return

    student_data = pd.read_excel(excel_file)
    success_message = "Emails sent successfully."

    email_listbox.insert(0, "Successful Emails")
    email_index = 0
    send_next_email()

def send_next_email():
    global email_index

    if email_index < len(student_data):
        student_name = student_data.iloc[email_index]["Name"]
        student_id = student_data.iloc[email_index]["Student ID"]
        student_email = student_data.iloc[email_index]["Email"]

        send_email_with_qr(student_name, student_id, venue, mentioned_time, student_email, from_email, email_password, email_listbox)

        email_index += 1

        root.after(1000, send_next_email)

root = tk.Tk()
root.title("QR Attendance DCC")

text_label = tk.Label(root, text="DCC Application to send attendance QR codes to students.", font=("Arial", 12))
text_label.pack(pady=10)

frame = tk.Frame(root)
frame.pack()

email_label = tk.Label(frame, text="Enter Your Email:", font=("Arial", 14))
email_label.pack(side="top")

email_entry = Entry(frame, font=("Arial", 12))
email_entry.pack(side="top")

password_label = tk.Label(frame, text="Enter Email Password:", font=("Arial", 14))
password_label.pack(side="top")

password_entry = Entry(frame, show="*", font=("Arial", 12))
password_entry.pack(side="top")

time_label = tk.Label(frame, text="Enter Event Time (hh:mm AM/PM):", font=("Arial", 14))
time_label.pack(side="top")

time_entry = Entry(frame, font=("Arial", 12))
time_entry.pack(side="top")

venue_label = tk.Label(frame, text="Enter Venue:", font=("Arial", 14))
venue_label.pack(side="top")

venue_entry = Entry(frame, font=("Arial", 12))
venue_entry.pack(side="top")

load_excel_button = tk.Button(root, text="Load Excel File", command=send_qr_codes, font=("Arial", 14))
load_excel_button.pack()

success_label = tk.Label(root, text="", font=("Arial", 12, "bold"), fg="green")
success_label.pack()

email_listbox = Listbox(root, selectmode="single", font=("Arial", 12), height=10)
email_listbox.pack()

scrollbar = Scrollbar(root)
scrollbar.pack(side="right", fill="y")

email_listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=email_listbox.yview)

student_data = None
email_index = 0

root.mainloop()
