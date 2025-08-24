import customtkinter as ctk
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
from email.mime.image import MIMEImage
import os
import threading
import queue

import tkinter.filedialog as fd
import json
from datetime import datetime

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class InvitationSenderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Invitation Sender")
        self.geometry("900x700")  # Made wider for two-column layout
        self.minsize(900, 700)    # Set minimum size to prevent elements from being hidden
        self.resizable(True, True) # Allow resizing for better usability

        self.excel_path = None
        self.invitees = None
        self.email_column = None
        self.name_column = None
        self.email_column_var = ctk.StringVar()
        self.name_column_var = ctk.StringVar()
        self.images_folder = os.getcwd()  # Default to current working directory
        
        # Initialize sent invitations tracking
        self.tracking_file = "sent_invitations.json"
        self.sent_invitations = self.load_sent_invitations()
        
        # Initialize selection tracking
        self.selected_invitees = {}  # Dictionary to track checkbox states
        
        self.create_widgets()
        
    def load_sent_invitations(self):
        """Load the record of sent invitations from JSON file"""
        if os.path.exists(self.tracking_file):
            try:
                with open(self.tracking_file, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                self.log("Warning: Tracking file corrupted, starting fresh.")
                return {}
        return {}
        
    def save_sent_invitations(self):
        """Save the record of sent invitations to JSON file"""
        with open(self.tracking_file, 'w') as f:
            json.dump(self.sent_invitations, f, indent=2)
            
    def was_invitation_sent(self, email, name):
        """Check if an invitation was already sent to this person"""
        key = f"{email}|{name}"
        return key in self.sent_invitations
        
    def mark_invitation_sent(self, email, name):
        """Mark an invitation as sent for this person"""
        key = f"{email}|{name}"
        self.sent_invitations[key] = {
            "email": email,
            "name": name,
            "sent_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.save_sent_invitations()

    def create_widgets(self):
        # Use a main frame to control layout and allow expansion
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title at the top
        self.label = ctk.CTkLabel(main_frame, text="Send Email Invitations", font=("Arial", 22))
        self.label.pack(pady=(0, 10))
        
        # Create two-column layout using a horizontal frame
        columns_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        columns_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Left column - Controls
        left_column = ctk.CTkFrame(columns_frame)
        left_column.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # Right column - Invitees list
        right_column = ctk.CTkFrame(columns_frame)
        right_column.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # === LEFT COLUMN CONTENT ===
        
        # Folder selection for invitation images
        folder_frame = ctk.CTkFrame(left_column)
        folder_frame.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(folder_frame, text="Images Folder:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
        folder_input_frame = ctk.CTkFrame(folder_frame, fg_color="transparent")
        folder_input_frame.pack(fill="x", padx=5, pady=(0, 5))
        self.folder_entry = ctk.CTkEntry(folder_input_frame, width=200)
        self.folder_entry.pack(side="left", padx=(0,5), fill="x", expand=True)
        self.folder_entry.insert(0, self.images_folder)
        self.folder_entry.configure(state="readonly")
        self.folder_btn = ctk.CTkButton(folder_input_frame, text="Browse", width=80, command=self.select_folder)
        self.folder_btn.pack(side="right")

        # Excel file section
        excel_frame = ctk.CTkFrame(left_column)
        excel_frame.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(excel_frame, text="Excel File:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
        self.open_btn = ctk.CTkButton(excel_frame, text="Open Excel File", command=self.open_excel)
        self.open_btn.pack(pady=5)
        self.status_label = ctk.CTkLabel(excel_frame, text="No file selected.", font=("Arial", 11))
        self.status_label.pack(pady=(0, 5))

        # Dropdowns for selecting email and name columns
        columns_section = ctk.CTkFrame(left_column)
        columns_section.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(columns_section, text="Column Mapping:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
        
        email_col_frame = ctk.CTkFrame(columns_section, fg_color="transparent")
        email_col_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(email_col_frame, text="Email Column:", width=100).pack(side="left")
        self.email_column_menu = ctk.CTkOptionMenu(email_col_frame, variable=self.email_column_var, values=[])
        self.email_column_menu.pack(side="right", fill="x", expand=True)
        
        name_col_frame = ctk.CTkFrame(columns_section, fg_color="transparent")
        name_col_frame.pack(fill="x", padx=5, pady=(2, 5))
        ctk.CTkLabel(name_col_frame, text="Name Column:", width=100).pack(side="left")
        self.name_column_menu = ctk.CTkOptionMenu(name_col_frame, variable=self.name_column_var, values=[])
        self.name_column_menu.pack(side="right", fill="x", expand=True)

        # Email credentials section
        email_creds_frame = ctk.CTkFrame(left_column)
        email_creds_frame.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(email_creds_frame, text="Email Credentials:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
        
        self.email_label = ctk.CTkLabel(email_creds_frame, text="Sender Email:")
        self.email_label.pack(pady=(5, 0), padx=5, anchor="w")
        self.email_entry = ctk.CTkEntry(email_creds_frame, width=250)
        self.email_entry.pack(padx=5, fill="x")

        self.pass_label = ctk.CTkLabel(email_creds_frame, text="App Password:")
        self.pass_label.pack(pady=(5, 0), padx=5, anchor="w")
        self.pass_entry = ctk.CTkEntry(email_creds_frame, show="*", width=250)
        self.pass_entry.pack(padx=5, pady=(0, 5), fill="x")

        # Log area
        log_frame = ctk.CTkFrame(left_column)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_label = ctk.CTkLabel(log_frame, text="Log:", font=("Arial", 12, "bold"))
        self.log_label.pack(pady=(5,0), anchor="w", padx=5)
        self.log_textbox = ctk.CTkTextbox(log_frame, wrap="word")
        self.log_textbox.pack(fill="both", expand=True, padx=5, pady=(0,5))
        self.log_textbox.configure(state="disabled")

        # === RIGHT COLUMN CONTENT ===
        
        # Invitees list with status
        ctk.CTkLabel(right_column, text="Invitees Status", font=("Arial", 16, "bold")).pack(pady=(10, 5))
        
        # Header with refresh button
        header_frame = ctk.CTkFrame(right_column)
        header_frame.pack(fill="x", padx=10, pady=(0, 5))
        self.refresh_btn = ctk.CTkButton(header_frame, text="Refresh", width=80, command=self.update_status_list)
        self.refresh_btn.pack(side="right", padx=5)
        
        # Selection buttons frame
        selection_frame = ctk.CTkFrame(right_column)
        selection_frame.pack(fill="x", padx=10, pady=(0, 5))
        self.select_all_btn = ctk.CTkButton(selection_frame, text="Select All", width=80, command=self.select_all_invitees)
        self.select_all_btn.pack(side="left", padx=2)
        self.select_none_btn = ctk.CTkButton(selection_frame, text="Select None", width=80, command=self.select_none_invitees)
        self.select_none_btn.pack(side="left", padx=2)
        self.select_unsent_btn = ctk.CTkButton(selection_frame, text="Select Unsent", width=90, command=self.select_unsent_invitees)
        self.select_unsent_btn.pack(side="left", padx=2)
        
        # Scrollable frame for invitees - takes up most of the right column
        self.scrollable_frame = ctk.CTkScrollableFrame(right_column)
        self.scrollable_frame.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        self.status_labels = {}  # Store labels for updating

        # Progress bar (hidden by default) - between list and send button
        self.progress_frame = ctk.CTkFrame(right_column)
        self.progress_frame.pack(pady=5, fill="x", padx=10)
        self.progress_frame.pack_forget()  # Hide initially
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(self.progress_frame, text="")
        self.progress_label.pack(pady=(0, 5))

        # Send button and result at bottom of right column
        send_frame = ctk.CTkFrame(right_column)
        send_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.send_btn = ctk.CTkButton(
            send_frame, 
            text="Send Invitations", 
            command=self.send_invitations, 
            state="disabled",
            height=40,
            font=("Arial", 14)
        )
        self.send_btn.pack(pady=5, fill="x", padx=10)

        self.result_label = ctk.CTkLabel(send_frame, text="", font=("Arial", 12))
        self.result_label.pack(pady=(0, 5))

    def select_folder(self):
        folder = fd.askdirectory(title="Select Invitation Images Folder")
        if folder:
            self.images_folder = folder
            self.folder_entry.configure(state="normal")
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, folder)
            self.folder_entry.configure(state="readonly")
            self.log(f"Selected images folder: {folder}")

    def log(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def select_all_invitees(self):
        """Select all invitees for sending"""
        for key in self.selected_invitees:
            self.selected_invitees[key].set(True)
        self.log("All invitees selected.")

    def select_none_invitees(self):
        """Deselect all invitees"""
        for key in self.selected_invitees:
            self.selected_invitees[key].set(False)
        self.log("All invitees deselected.")

    def select_unsent_invitees(self):
        """Select only invitees who haven't been sent invitations yet"""
        count = 0
        for key, checkbox_var in self.selected_invitees.items():
            email, name = key.split("|", 1)
            if not self.was_invitation_sent(email, name):
                checkbox_var.set(True)
                count += 1
            else:
                checkbox_var.set(False)
        self.log(f"Selected {count} unsent invitees.")

    def clear_status_list(self):
        """Clear all status labels and selection tracking"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.status_labels = {}
        self.selected_invitees = {}

    def update_status_list(self):
        """Update the status list with current invitees and checkboxes"""
        self.clear_status_list()
        if not hasattr(self, 'invitees') or self.invitees is None:
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            return

        for idx, row in self.invitees.iterrows():
            name = self.clean_name(str(row[name_col]).strip())
            email = str(row[email_col]).strip()
            key = f"{email}|{name}"
            
            # Create frame for this invitee
            frame = ctk.CTkFrame(self.scrollable_frame)
            frame.pack(fill="x", padx=2, pady=1)
            
            # Checkbox for selection
            checkbox_var = ctk.BooleanVar()
            checkbox = ctk.CTkCheckBox(frame, text="", variable=checkbox_var, width=20)
            checkbox.pack(side="left", padx=5)
            
            # Store checkbox variable for later use
            self.selected_invitees[key] = checkbox_var
            
            # Name and email
            info_text = f"{name} ({email})"
            ctk.CTkLabel(frame, text=info_text, anchor="w").pack(side="left", padx=5, fill="x", expand=True)
            
            # Status label
            status_label = ctk.CTkLabel(frame, text="", anchor="e", width=120)
            status_label.pack(side="right", padx=5)
            
            # Store label reference for updates
            self.status_labels[key] = status_label
            
            # Update status and default to select unsent invitees
            is_sent = self.was_invitation_sent(email, name)
            self.update_invitee_status(email, name)
            checkbox_var.set(not is_sent)  # Select unsent invitees by default

    def update_invitee_status(self, email, name):
        """Update the status display for a single invitee"""
        key = f"{email}|{name}"
        if key not in self.status_labels:
            return
            
        label = self.status_labels[key]
        if self.was_invitation_sent(email, name):
            sent_date = self.sent_invitations[key]["sent_date"]
            label.configure(text=f"Sent on {sent_date}", text_color="green")
        else:
            label.configure(text="Not sent", text_color="gray")

    def open_excel(self):
        file_path = fd.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.log(f"Opening Excel file: {file_path}")
                df = pd.read_excel(file_path)
                columns = list(df.columns)
                if not columns:
                    self.status_label.configure(text="Excel file has no columns.", text_color="red")
                    self.send_btn.configure(state="disabled")
                    self.email_column_menu.configure(values=[])
                    self.name_column_menu.configure(values=[])
                    self.log("Excel file has no columns.")
                    return
                # Set dropdowns for email and name columns
                self.email_column_menu.configure(values=columns)
                self.name_column_menu.configure(values=columns)
                # Try to auto-select likely columns
                email_guess = next((c for c in columns if 'email' in c.lower()), columns[0])
                name_guess = next((c for c in columns if 'name' in c.lower()), columns[0])
                self.email_column_var.set(email_guess)
                self.name_column_var.set(name_guess)
                self.invitees = df.dropna(subset=[email_guess, name_guess])
                self.excel_path = file_path
                self.status_label.configure(text=f"Loaded {len(self.invitees)} invitees.", text_color="green")
                self.send_btn.configure(state="normal")
                self.log(f"Loaded {len(self.invitees)} invitees from Excel.")
                self.update_status_list()
            except Exception as e:
                self.status_label.configure(text=f"Error: {e}", text_color="red")
                self.send_btn.configure(state="disabled")
                self.email_column_menu.configure(values=[])
                self.name_column_menu.configure(values=[])
                self.log(f"Error loading Excel: {e}")
        else:
            self.status_label.configure(text="No file selected.", text_color="gray")
            self.send_btn.configure(state="disabled")
            self.email_column_menu.configure(values=[])
            self.name_column_menu.configure(values=[])
            self.log("No file selected.")

    def clean_name(self, name):
        # Remove dots from middle initials and handle multiple spaces
        return ' '.join(part.replace('.', '') for part in name.split())

    def update_progress(self, current, total, message=""):
        """Update the progress bar and label"""
        progress = current / total if total > 0 else 0
        self.progress_bar.set(progress)
        self.progress_label.configure(text=message)

    def send_single_invitation(self, sender_email, sender_pass, name, recipient, img_filename):
        """Send a single invitation and return the result"""
        try:
            msg = EmailMessage()
            msg["Subject"] = "Invitation to the National Day and Armed Forces Day of the Republic of Korea"
            msg["From"] = sender_email
            msg["To"] = recipient
            cid = make_msgid(domain="xyz.com")
            msg.add_alternative(f"""\
            <html>
              <head>
                <style>
                  img {{ max-width: 900px; width: 100%; height: auto; }}
                </style>
              </head>
              <body>
                <img src=\"cid:{cid[1:-1]}\" style="width: 900px; max-width: 100%; height: auto;">
              </body>
            </html>
            """, subtype='html')

            with open(img_filename, 'rb') as img:
                img_data = img.read()
                image = MIMEImage(img_data, name=os.path.basename(img_filename))
                image.add_header('Content-ID', cid)
                html_part = None
                for part in msg.iter_parts():
                    if part.get_content_type() == 'text/html':
                        html_part = part
                        break
                if html_part is not None:
                    html_part.add_related(image)
                else:
                    return False, "Could not find HTML part to attach image"

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(sender_email, sender_pass)
                smtp.send_message(msg)
            return True, None
        except Exception as e:
            return False, str(e)

    def send_invitations_thread(self, sender_email, sender_pass, email_col, name_col):
        """Thread function for sending invitations"""
        sent_count = 0
        failed = []
        skipped = 0
        selected_count = 0
        
        # First, count selected invitees
        for idx, row in self.invitees.iterrows():
            name = self.clean_name(str(row[name_col]).strip())
            recipient = str(row[email_col]).strip()
            key = f"{recipient}|{name}"
            
            if key in self.selected_invitees and self.selected_invitees[key].get():
                selected_count += 1
        
        if selected_count == 0:
            self.after(0, self.log, "No invitees selected for sending.")
            self.after(0, self.finish_sending, 0, 0, [])
            return
        
        self.after(0, self.log, f"Starting to send {selected_count} selected invitations...")
        current_processed = 0
        
        for idx, row in self.invitees.iterrows():
            name = self.clean_name(str(row[name_col]).strip())
            recipient = str(row[email_col]).strip()
            key = f"{recipient}|{name}"
            
            # Skip if not selected
            if key not in self.selected_invitees or not self.selected_invitees[key].get():
                continue
                
            current_processed += 1
            
            # Update progress in the main thread
            self.after(0, self.update_progress, current_processed, selected_count, f"Processing: {name} ({recipient})")
            
            # Check if invitation was already sent
            if self.was_invitation_sent(recipient, name):
                self.after(0, self.log, f"[SKIPPED] Already sent to {name} ({recipient})")
                skipped += 1
                continue
                
            img_filename = os.path.join(self.images_folder, f"Invitation - {name}.png")
            if not os.path.exists(img_filename):
                failed.append((recipient, "Invitation image not found"))
                self.after(0, self.log, f"[{recipient}] Invitation image not found: {img_filename}")
                continue

            success, error = self.send_single_invitation(sender_email, sender_pass, name, recipient, img_filename)
            if success:
                sent_count += 1
                self.mark_invitation_sent(recipient, name)
                self.after(0, self.update_invitee_status, recipient, name)
                self.after(0, self.log, f"[{recipient}] Invitation sent successfully.")
            else:
                failed.append((recipient, error))
                self.after(0, self.log, f"[{recipient}] Failed to send: {error}")

        # Update final results in the main thread
        self.after(0, self.finish_sending, sent_count, skipped, failed)

    def finish_sending(self, sent_count, skipped, failed):
        """Update UI after sending is complete"""
        result_msg = f"Sent: {sent_count} invitations."
        if skipped:
            result_msg += f"\nSkipped (already sent): {skipped}"
        if failed:
            result_msg += f"\nFailed: {len(failed)}"
        
        self.result_label.configure(text=result_msg, text_color="green" if sent_count else "red")
        self.log(result_msg)
        
        # Re-enable the send button and hide progress
        self.send_btn.configure(state="normal")
        self.progress_frame.pack_forget()

    def send_invitations(self):
        sender_email = self.email_entry.get().strip()
        sender_pass = self.pass_entry.get().strip()
        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        
        if not sender_email or not sender_pass:
            self.result_label.configure(text="Enter sender email and app password.", text_color="red")
            self.log("Sender email or app password missing.")
            return
        if not email_col or not name_col:
            self.result_label.configure(text="Select email and name columns.", text_color="red")
            self.log("Email or name column not selected.")
            return

        # Show progress bar and disable send button
        self.progress_frame.pack(pady=5, fill="x", padx=10)
        self.send_btn.configure(state="disabled")
        self.log(f"Starting to send invitations from {sender_email}...")
        
        # Start sending thread
        threading.Thread(
            target=self.send_invitations_thread,
            args=(sender_email, sender_pass, email_col, name_col),
            daemon=True
        ).start()

if __name__ == "__main__":
    app = InvitationSenderApp()
    app.mainloop()