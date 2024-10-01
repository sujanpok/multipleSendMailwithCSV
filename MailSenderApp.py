import os
import csv
import pandas as pd
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk, Menu
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipar
import threading

class MailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()  # Hide the main window initially
        self.root.title("Setup App")

        self.sender_email = None
        self.sender_password = None
        self.recipients_file = None
        self.recipients_list = []

        self.setup_logging()
        self.load_sender_credentials()

        # Open setup pop-up if credentials are not set
        if not self.sender_email or not self.sender_password:
            self.open_setup_popup()
        else:
            self.create_main_gui_elements()

    def setup_logging(self):
        logging.basicConfig(filename='email_log.txt', level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

    def create_main_gui_elements(self):
        # Main app section
        self.root.deiconify()  # Show the main window
        self.setup_menu()

        # Main GUI elements for sending emails
        ttk.Label(self.root, text="Subject:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        self.subject_entry = ttk.Entry(self.root, width=50)
        self.subject_entry.grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(self.root, text="Message:").grid(row=1, column=0, padx=10, pady=5, sticky='nw')
        self.message_text = tk.Text(self.root, width=50, height=10)
        self.message_text.grid(row=1, column=1, padx=10, pady=5)

        # Individual email input section
        self.individual_email_frame = ttk.Frame(self.root)
        self.individual_email_frame.grid(row=2, column=1, padx=10, pady=5, sticky='w')

        self.individual_email_entry = ttk.Entry(self.individual_email_frame, width=40)
        self.individual_email_entry.grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(self.individual_email_frame, text="+", command=self.add_individual_email).grid(row=0, column=1, padx=5, pady=5)

        self.email_listbox = tk.Listbox(self.root, width=60, height=5)
        self.email_listbox.grid(row=3, column=1, padx=10, pady=5)
        ttk.Button(self.root, text="Delete Selected Email", command=self.delete_selected_email).grid(row=4, column=1, padx=10, pady=5)

        # CSV file load button
        ttk.Button(self.root, text="Load Recipients CSV", command=self.load_recipients).grid(row=5, column=1, padx=10, pady=5, sticky='w')

        # Send emails button
        ttk.Button(self.root, text="Send Emails", command=self.send_emails).grid(row=6, column=1, padx=10, pady=10)

        # Loading indicator
        self.loading_label = ttk.Label(self.root, text="", foreground='red')
        self.loading_label.grid(row=7, column=1, padx=10, pady=10)

    def setup_menu(self):
        # Menu Bar
        menu_bar = Menu(self.root)
        self.root.config(menu=menu_bar)

        # Setup menu
        setup_menu = Menu(menu_bar, tearoff=0)
        setup_menu.add_command(label="Setup Email Credentials", command=self.open_setup_popup)
        menu_bar.add_cascade(label="Setup", menu=setup_menu)

        # Help menu
        help_menu = Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about_info)
        menu_bar.add_cascade(label="Help", menu=help_menu)

    def open_setup_popup(self):
        # Create a new window for setting up email and password
        self.setup_window = tk.Toplevel(self.root)
        self.setup_window.title("Setup Email Credentials")
        self.setup_window.geometry("500x200")  # Set a larger size for the pop-up

        ttk.Label(self.setup_window, text="Enter sender email:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.email_entry = ttk.Entry(self.setup_window, width=40)
        self.email_entry.grid(row=0, column=1, padx=10, pady=10)

        ttk.Label(self.setup_window, text="Enter Google App Password:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.password_entry = ttk.Entry(self.setup_window, show='*', width=40)
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        # Buttons
        ttk.Button(self.setup_window, text="Save Credentials", command=self.save_sender_credentials).grid(row=2, column=1, padx=10, pady=10, sticky='e')
        ttk.Button(self.setup_window, text="Skip", command=self.skip_setup).grid(row=2, column=0, padx=10, pady=10, sticky='w')

        # Load credentials to populate the fields
        self.populate_credentials()

    def skip_setup(self):
        self.setup_window.destroy()
        self.create_main_gui_elements()

    def load_sender_credentials(self):
        # Check if sender credentials are already saved
        if os.path.exists("sender_credentials.txt"):
            with open("sender_credentials.txt", "r") as file:
                self.sender_email = file.readline().strip()
                self.sender_password = file.readline().strip()

    def populate_credentials(self):
        # Populate email and password fields if they exist
        if self.sender_email and self.sender_password:
            self.email_entry.insert(0, self.sender_email)
            self.password_entry.insert(0, self.sender_password)

    def save_sender_credentials(self):
        self.sender_email = self.email_entry.get().strip()
        self.sender_password = self.password_entry.get().strip()

        if not self.sender_email or not self.sender_password:
            messagebox.showerror("Input Error", "Email and password cannot be empty.")
            return

        # Save credentials to file
        with open("sender_credentials.txt", "w") as file:
            file.write(f"{self.sender_email}\n{self.sender_password}")
        logging.info("Sender email credentials saved successfully.")
        messagebox.showinfo("Success", "Credentials saved successfully!")
        self.setup_window.destroy()
        self.create_main_gui_elements()

    def add_individual_email(self):
        if len(self.recipients_list) >= 5:
            messagebox.showwarning("Limit Reached", "You can only add up to 5 individual email addresses.")
            return

        email = self.individual_email_entry.get().strip()
        if email:
            self.recipients_list.append(email)
            self.email_listbox.insert(tk.END, email)
            self.individual_email_entry.delete(0, tk.END)
            logging.info(f"Added individual email: {email}")

    def delete_selected_email(self):
        selected_index = self.email_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            email = self.recipients_list.pop(index)
            self.email_listbox.delete(index)
            logging.info(f"Deleted individual email: {email}")

    def load_recipients(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.recipients_file = file_path
            logging.info(f"Loaded recipients file: {file_path}")

    def send_emails(self):
        threading.Thread(target=self.send_emails_thread).start()  # Start sending emails in a new thread

    def send_emails_thread(self):
        self.loading_label.config(text="Sending emails...")  # Show loading text
        try:
            subject = self.subject_entry.get()
            message = self.message_text.get("1.0", tk.END).strip()
            if not subject or not message:
                messagebox.showerror("Input Error", "Subject and message cannot be empty.")
                return

            recipients = self.recipients_list.copy()

            if self.recipients_file:
                recipients_df = pd.read_csv(self.recipients_file)
                if 'email' not in recipients_df.columns:
                    messagebox.showerror("CSV Error", "The CSV file must contain an 'email' column.")
                    return
                recipients += recipients_df['email'].tolist()

            if not recipients:
                messagebox.showerror("Recipients Error", "No recipients specified.")
                return

            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)

                for recipient_email in recipients:
                    msg = MIMEMultipart()
                    msg['From'] = self.sender_email
                    msg['To'] = recipient_email
                    msg['Subject'] = subject
                    msg.attach(MIMEText(message, 'plain'))

                    server.sendmail(self.sender_email, recipient_email, msg.as_string())
                    logging.info(f"Email sent to {recipient_email}")

            self.loading_label.config(text="Emails sent successfully!")  # Update loading label
            messagebox.showinfo("Success", "Emails sent successfully.")
        except Exception as e:
            logging.error(f"Error occurred: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.loading_label.config(text="")  # Clear loading text

    def show_about_info(self):
        messagebox.showinfo("About", "Setup App\nVersion 1.0\nDeveloped by YourName")

if __name__ == "__main__":
    root = tk.Tk()
    app = MailSenderApp(root)
    root.mainloop()
