import smtplib
import imaplib
import email
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
import tkinter as tk
from tkinter import scrolledtext, filedialog
import threading
import os

def send_email(smtp_server, port, username, password, recipient, subject, body, attachment_path=None):
    try:
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login(username, password)

        msg = MIMEMultipart()
        msg['From'] = username
        msg['To'] = recipient
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        if attachment_path:
            attachment_name = os.path.basename(attachment_path)
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {attachment_name}",
                )
                msg.attach(part)

        text = msg.as_string()

        server.sendmail(username, recipient, text)
        print(f"Sent email to {recipient}")

        server.quit()
    except Exception as e:
        print(f"Error sending email: {e}")

def fetch_email(imap_server, username, password, mailbox='inbox'):
    global last_seen_email_id
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(username, password)
        mail.select(mailbox)

        result, data = mail.search(None, 'FROM', f'"{recipient}"')
#        print("result=", result, " data=", data)
        mail_ids = data[0].split()

        if not mail_ids:
            return None

        latest_email_id = mail_ids[-1]

        if latest_email_id == last_seen_email_id:
            return None

        last_seen_email_id = latest_email_id

        result, data = mail.fetch(latest_email_id, '(RFC822)')
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        email_body = None
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    email_body = part.get_payload(decode=True).decode('utf-8')
                    break
        else:
            email_body = msg.get_payload(decode=True).decode('utf-8')

        return f"From: {msg['From']}\nSubject: {msg['Subject']}\n\n{email_body}"
    except Exception as e:
        print(f"Error receiving email: {e}")
        return None

def start_email_checker():
    while True:
        time.sleep(5)
        email_content = fetch_email(imap_server, username, password)
        if email_content:
            email_display.insert(tk.END, email_content + "\n\n")

def on_send_email():
    subject = subject_entry.get()
    body = body_entry.get("1.0", tk.END)
    attachment_path = attachment_path_var.get()
    send_email(smtp_server, port, username, password, recipient, subject, body, attachment_path)

def browse_file():
    filename = filedialog.askopenfilename()
    attachment_path_var.set(filename)

def clear_email_display():
    email_display.delete('1.0', tk.END)  # Clear the text box

smtp_server = 'smtp.gmail.com'
imap_server = 'imap.gmail.com'
port = 587
#username = input("Your email address: ")
username = "attackerSender@testmail.com"
password = getpass.getpass('Password: ')
#recipient = input("Target email address: ")
recipient = "victimRecipient@testmail.com"


last_seen_email_id = None

root = tk.Tk()
root.title("Email Client")

send_frame = tk.Frame(root)
send_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

tk.Label(send_frame, text="Subject:").pack(side=tk.LEFT)
subject_entry = tk.Entry(send_frame)
subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

tk.Label(send_frame, text="Body:").pack(side=tk.LEFT)
body_entry = tk.Text(send_frame, height=10)
body_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

attachment_frame = tk.Frame(root)
attachment_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

attachment_path_var = tk.StringVar()
attachment_entry = tk.Entry(attachment_frame, textvariable=attachment_path_var, width=40)
attachment_entry.pack(side=tk.LEFT)

browse_button = tk.Button(attachment_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.LEFT, padx=10)

send_button = tk.Button(send_frame, text="Send Email", command=on_send_email)
send_button.pack(side=tk.RIGHT, padx=10)

receive_frame = tk.Frame(root)
receive_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=10, pady=10)

email_display = scrolledtext.ScrolledText(receive_frame, height=15)
email_display.pack(fill=tk.BOTH, expand=True)

# Add Clear button to clear the text
clear_button = tk.Button(receive_frame, text="Clear", command=clear_email_display)
clear_button.pack(side=tk.BOTTOM, pady=5)

threading.Thread(target=start_email_checker, daemon=True).start()

root.mainloop()
