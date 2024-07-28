import tkinter as tk
from tkinter import messagebox
import threading
from mass_email_sender import send_emails_in_batches

def start_sending_emails():
    sheet_link = entry_sheet_link.get()
    subject = entry_subject.get()
    body = text_message.get("1.0", tk.END).strip()
    delay = int(entry_delay.get())
    batch_size = int(entry_batch_size.get())
    
    if not sheet_link or not subject or not body:
        messagebox.showerror("Input Error", "Please fill in all fields.")
        return
    
    threading.Thread(target=send_emails_in_batches, args=(sheet_link, subject, body, batch_size, delay)).start()

app = tk.Tk()
app.title("Email Sender")

tk.Label(app, text="Google Sheet Link:").grid(row=0, column=0, padx=10, pady=5)
entry_sheet_link = tk.Entry(app, width=50)
entry_sheet_link.grid(row=0, column=1, padx=10, pady=5)

tk.Label(app, text="Subject:").grid(row=1, column=0, padx=10, pady=5)
entry_subject = tk.Entry(app, width=50)
entry_subject.grid(row=1, column=1, padx=10, pady=5)

tk.Label(app, text="Message:").grid(row=2, column=0, padx=10, pady=5)
text_message = tk.Text(app, width=50, height=10)
text_message.grid(row=2, column=1, padx=10, pady=5)

tk.Label(app, text="Delay (seconds):").grid(row=3, column=0, padx=10, pady=5)
entry_delay = tk.Entry(app, width=20)
entry_delay.grid(row=3, column=1, padx=10, pady=5)

tk.Label(app, text="Batch Size:").grid(row=4, column=0, padx=10, pady=5)
entry_batch_size = tk.Entry(app, width=20)
entry_batch_size.grid(row=4, column=1, padx=10, pady=5)

send_button = tk.Button(app, text="Send Emails", command=start_sending_emails)
send_button.grid(row=5, column=1, pady=20)

app.mainloop()
