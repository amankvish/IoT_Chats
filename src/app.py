import tkinter as tk
from tkinter import filedialog, messagebox
from automation import load_files, send_messages
import threading
import os

# Create the main window using Tkinter
root = tk.Tk()
root.title("WhatsApp Automation Application")
root.geometry("500x400")

# Global variables for stopping automation
stop_flag = False
automation_thread = None

# Set font styles
font_title = ("Arial", 16, "bold")
font_normal = ("Arial", 12)

# Labels and Buttons
title_label = tk.Label(root, text="WhatsApp Automation", font=font_title, fg="blue")
title_label.pack(pady=10)

# Variables to hold file paths
contacts_file = ""
message_template_file = ""
status_file = "message_status.xlsx"


def select_contacts_file():
    global contacts_file
    contacts_file = filedialog.askopenfilename(title="Select Contacts Excel File",
                                               filetypes=[("Excel Files", "*.xlsx")])
    if contacts_file:
        contacts_label.config(text=os.path.basename(contacts_file))


def select_template_file():
    global message_template_file
    message_template_file = filedialog.askopenfilename(title="Select Message Template (.docx)",
                                                       filetypes=[("Word Files", "*.docx")])
    if message_template_file:
        template_label.config(text=os.path.basename(message_template_file))


def send_messages_button():
    global automation_thread, stop_flag
    if not contacts_file or not message_template_file:
        messagebox.showerror("Error", "Please select both Contacts Excel and Message Template files.")
        return

    # Disable the send button while messages are being sent
    send_btn.config(state=tk.DISABLED)
    stop_btn.config(state=tk.NORMAL)

    stop_flag = False  # Reset the stop flag

    # Start the automation in a separate thread
    automation_thread = threading.Thread(target=run_automation)
    automation_thread.start()


def run_automation():
    global stop_flag
    df, message = load_files(contacts_file, message_template_file)
    success_count, failure_count = send_messages(df, message, status_file, stop_flag)

    if stop_flag:
        messagebox.showinfo("Stopped", "Automation was stopped by the user.")
    else:
        messagebox.showinfo("Completion", f"Task completed.\nSuccess: {success_count}\nFailed: {failure_count}")

    # Re-enable the send button after completion
    send_btn.config(state=tk.NORMAL)
    stop_btn.config(state=tk.DISABLED)


def stop_automation():
    global stop_flag
    stop_flag = True
    stop_btn.config(state=tk.DISABLED)


# Create buttons to select files
contacts_btn = tk.Button(root, text="Select Contacts Excel", command=select_contacts_file, font=font_normal)
contacts_btn.pack(pady=5)

contacts_label = tk.Label(root, text="No file selected", font=font_normal)
contacts_label.pack(pady=5)

template_btn = tk.Button(root, text="Select Message Template", command=select_template_file, font=font_normal)
template_btn.pack(pady=5)

template_label = tk.Label(root, text="No file selected", font=font_normal)
template_label.pack(pady=5)

# Button to send messages
send_btn = tk.Button(root, text="Send Messages", command=send_messages_button, font=font_normal, bg="green", fg="white")
send_btn.pack(pady=20)

# Button to stop automation
stop_btn = tk.Button(root, text="Stop Automation", command=stop_automation, font=font_normal, bg="red", fg="white")
stop_btn.pack(pady=5)
stop_btn.config(state=tk.DISABLED)

# Start the application
root.mainloop()
