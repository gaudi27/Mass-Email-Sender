import tkinter as tk
from tkinter import messagebox, simpledialog
import threading
from mass_email_sender import send_emails_in_batches

# Function to insert hyperlink into message text
def insert_hyperlink():
    try:
        start_idx = text_message.index(tk.SEL_FIRST)
        end_idx = text_message.index(tk.SEL_LAST)
    except tk.TclError:
        messagebox.showerror("Selection Error", "Please select text to hyperlink.")
        return

    url = simpledialog.askstring("Insert Hyperlink", "Enter URL:")
    if url:
        selected_text = text_message.get(start_idx, end_idx)
        hyperlink = f'<a href="{url}">{selected_text}</a>'
        text_message.delete(start_idx, end_idx)
        text_message.insert(start_idx, hyperlink)

# Function to start sending emails in batches
def start_sending_emails():
    sheet_link = entry_sheet_link.get()
    subject = entry_subject.get()
    body = text_message.get("1.0", tk.END).strip()
    delay = int(entry_delay.get())
    batch_size = int(entry_batch_size.get())
    
    if not sheet_link or not subject or not body:
        messagebox.showerror("Input Error", "Please fill in all fields.")
        return

    # Start the email-sending process in a separate thread and update UI
    threading.Thread(target=send_emails_in_batches, args=(sheet_link, subject, body, batch_size, delay, update_chat_box)).start()

# Function to update the chat box with messages from the batch process
def update_chat_box(message):
    chat_box.configure(state=tk.NORMAL)
    chat_box.insert(tk.END, message + "\n")
    chat_box.see(tk.END)  # Auto-scroll to the bottom
    chat_box.configure(state=tk.DISABLED)

# Function to create rounded buttons
def create_rounded_button(text, command, parent):
    return tk.Button(parent, text=text, command=command, bg="#5A9", fg="white", relief="flat", font=("Arial", 12, "bold"), 
                     padx=10, pady=5, bd=0, activebackground="#58C", highlightthickness=0)

# Initialize the app window
app = tk.Tk()
app.title("Email Sender")
app.geometry("1200x600")
app.configure(bg="#F5F5F5")

# Grid configuration to allow resizing
app.grid_columnconfigure(1, weight=1)
app.grid_columnconfigure(2, weight=2)
app.grid_rowconfigure(3, weight=1)

# Chat box for batch sending updates
tk.Label(app, text="Batch Sending Updates:", font=("Arial", 14, "bold"), bg="#F5F5F5").grid(row=0, column=0, padx=10, pady=5, sticky='nw')
chat_box = tk.Text(app, width=40, height=25, relief="solid", borderwidth=1, state=tk.DISABLED, bg="#EAEAEA", wrap=tk.WORD)
chat_box.grid(row=1, column=0, rowspan=6, padx=10, pady=5, sticky="nsew")

# Title and instructions
tk.Label(app, text="Mass Email Sender", font=("Arial", 18, "bold"), bg="#F5F5F5", fg="#333").grid(row=0, column=1, columnspan=2, pady=(10, 20))

# Google Sheets link input
tk.Label(app, text="Google Sheet Link:", bg="#F5F5F5").grid(row=1, column=1, padx=10, pady=5, sticky='e')
entry_sheet_link = tk.Entry(app, width=50, relief="solid", borderwidth=1)
entry_sheet_link.grid(row=1, column=2, padx=10, pady=5, sticky="ew")

# (Set share permissions instructions)
tk.Label(app, text="(SET SHARE PERMISSIONS OF SHEET TO ANYONE WITH LINK AND KEEP COMPUTER RUNNING)", bg="#F5F5F5", fg="red").grid(row=2, column=1, columnspan=2, pady=(0, 10), sticky="w")

# Subject input
tk.Label(app, text="Subject:", bg="#F5F5F5").grid(row=3, column=1, padx=10, pady=5, sticky='e')
entry_subject = tk.Entry(app, width=50, relief="solid", borderwidth=1)
entry_subject.grid(row=3, column=2, padx=10, pady=5, sticky="ew")

# Message input (Text box)
tk.Label(app, text="Message:", bg="#F5F5F5").grid(row=4, column=1, padx=10, pady=5, sticky='ne')
text_message = tk.Text(app, width=50, height=10, relief="solid", borderwidth=1)
text_message.grid(row=4, column=2, padx=10, pady=5, sticky="nsew")

# Delay input
tk.Label(app, text="Delay (30s minimum):", bg="#F5F5F5").grid(row=5, column=1, padx=10, pady=5, sticky='e')
entry_delay = tk.Entry(app, width=20, relief="solid", borderwidth=1)
entry_delay.grid(row=5, column=2, padx=10, pady=5, sticky='w')

# Batch size input
tk.Label(app, text="Batch Size (30 maximum):", bg="#F5F5F5").grid(row=6, column=1, padx=10, pady=5, sticky='e')
entry_batch_size = tk.Entry(app, width=20, relief="solid", borderwidth=1)
entry_batch_size.grid(row=6, column=2, padx=10, pady=5, sticky='w')

# Insert Hyperlink button
hyperlink_button = create_rounded_button("Insert Hyperlink", insert_hyperlink, app)
hyperlink_button.grid(row=7, column=1, pady=20, padx=10, sticky="e")

# Send Emails button
send_button = create_rounded_button("Send Emails", start_sending_emails, app)
send_button.grid(row=7, column=2, pady=20, padx=10, sticky="w")

# Start the app loop
app.mainloop()
