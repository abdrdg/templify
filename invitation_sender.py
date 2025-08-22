import customtkinter as ctk
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
from email.mime.image import MIMEImage
import os

import tkinter.filedialog as fd
import json
from datetime import datetime

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class InvitationSenderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Invitation Sender")
        self.geometry("500x650")  # Increased height to ensure all elements are visible
        self.minsize(500, 650)    # Set minimum size to prevent elements from being hidden
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
        main_frame.pack_propagate(False)

        self.label = ctk.CTkLabel(main_frame, text="Send Email Invitations", font=("Arial", 22))
        self.label.pack(pady=10)


        # Folder selection for invitation images
        folder_frame = ctk.CTkFrame(main_frame)
        folder_frame.pack(pady=5, fill="x")
        ctk.CTkLabel(folder_frame, text="Images Folder:").pack(side="left", padx=(0,5))
        self.folder_entry = ctk.CTkEntry(folder_frame, width=260)
        self.folder_entry.pack(side="left", padx=(0,5), fill="x", expand=True)
        self.folder_entry.insert(0, self.images_folder)
        self.folder_entry.configure(state="readonly")
        self.folder_btn = ctk.CTkButton(folder_frame, text="Browse", width=80, command=self.select_folder)
        self.folder_btn.pack(side="left")

        self.open_btn = ctk.CTkButton(main_frame, text="Open Excel File", command=self.open_excel)
        self.open_btn.pack(pady=5)
    

        self.status_label = ctk.CTkLabel(main_frame, text="No file selected.", font=("Arial", 14))
        self.status_label.pack(pady=5)

        # Dropdowns for selecting email and name columns
        self.column_frame = ctk.CTkFrame(main_frame)
        self.column_frame.pack(pady=5)
        ctk.CTkLabel(self.column_frame, text="Email Column:").pack(side="left", padx=(0,5))
        self.email_column_menu = ctk.CTkOptionMenu(self.column_frame, variable=self.email_column_var, values=[])
        self.email_column_menu.pack(side="left", padx=(0,15))
        ctk.CTkLabel(self.column_frame, text="Name Column:").pack(side="left", padx=(0,5))
        self.name_column_menu = ctk.CTkOptionMenu(self.column_frame, variable=self.name_column_var, values=[])
        self.name_column_menu.pack(side="left")

        self.email_label = ctk.CTkLabel(main_frame, text="Sender Email:")
        self.email_label.pack(pady=(10, 0))
        self.email_entry = ctk.CTkEntry(main_frame, width=300)
        self.email_entry.pack()

        self.pass_label = ctk.CTkLabel(main_frame, text="App Password:")
        self.pass_label.pack(pady=(5, 0))
        self.pass_entry = ctk.CTkEntry(main_frame, show="*", width=300)
        self.pass_entry.pack()

        # Invitees list with status
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.pack(fill="x", padx=2, pady=(10,5))
        
        # Header
        header_frame = ctk.CTkFrame(status_frame)
        header_frame.pack(fill="x", padx=5, pady=(5,0))
        ctk.CTkLabel(header_frame, text="Invitees Status", font=("Arial", 14, "bold")).pack(side="left")
        self.refresh_btn = ctk.CTkButton(header_frame, text="Refresh", width=80, command=self.update_status_list)
        self.refresh_btn.pack(side="right", padx=5)
        
        # Scrollable frame for invitees
        self.scrollable_frame = ctk.CTkScrollableFrame(status_frame, height=150)
        self.scrollable_frame.pack(fill="x", padx=5, pady=5)
        self.status_labels = {}  # Store labels for updating

        # Log area (scrollable textbox)
        self.log_label = ctk.CTkLabel(main_frame, text="Log:", font=("Arial", 12))
        self.log_label.pack(pady=(10,0), anchor="w")
        self.log_textbox = ctk.CTkTextbox(main_frame, height=100, width=450, wrap="word")
        self.log_textbox.pack(fill="x", padx=2, pady=(0,5))
        self.log_textbox.configure(state="disabled")

        # Create a bottom container frame for the send button and result label
        bottom_container = ctk.CTkFrame(main_frame, fg_color="transparent")
        bottom_container.pack(side="bottom", fill="x", pady=(5,10))

        # Create the send button with more prominence
        self.send_btn = ctk.CTkButton(
            bottom_container, 
            text="Send Invitations", 
            command=self.send_invitations, 
            state="disabled",
            height=40,  # Make button taller
            font=("Arial", 14)  # Larger font
        )
        self.send_btn.pack(pady=5, fill="x", padx=20)  # Add padding on sides

        # Result label below the send button
        self.result_label = ctk.CTkLabel(bottom_container, text="", font=("Arial", 12))
        self.result_label.pack(pady=5)

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

    def clear_status_list(self):
        """Clear all status labels"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.status_labels = {}

    def update_status_list(self):
        """Update the status list with current invitees"""
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
            
            # Create frame for this invitee
            frame = ctk.CTkFrame(self.scrollable_frame)
            frame.pack(fill="x", padx=2, pady=1)
            
            # Name and email
            info_text = f"{name} ({email})"
            ctk.CTkLabel(frame, text=info_text, anchor="w").pack(side="left", padx=5)
            
            # Status label
            status_label = ctk.CTkLabel(frame, text="", anchor="e")
            status_label.pack(side="right", padx=5)
            
            # Store label reference for updates
            self.status_labels[f"{email}|{name}"] = status_label
            
            # Update status
            self.update_invitee_status(email, name)

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

        sent_count = 0
        failed = []
        skipped = 0
        self.log(f"Starting to send invitations from {sender_email}...")
        for idx, row in self.invitees.iterrows():
            name = self.clean_name(str(row[name_col]).strip())
            recipient = str(row[email_col]).strip()
            
            # Check if invitation was already sent
            if self.was_invitation_sent(recipient, name):
                self.log(f"[SKIPPED] Already sent to {name} ({recipient})")
                skipped += 1
                continue
                
            img_filename = os.path.join(self.images_folder, f"Invitation - {name}.png")
            if not os.path.exists(img_filename):
                failed.append((recipient, "Invitation image not found"))
                self.log(f"[{recipient}] Invitation image not found: {img_filename}")
                continue
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
                    # Find the HTML part robustly
                    html_part = None
                    for part in msg.iter_parts():
                        if part.get_content_type() == 'text/html':
                            html_part = part
                            break
                    if html_part is not None:
                        html_part.add_related(image)
                    else:
                        self.log(f"[ERROR] Could not find HTML part to attach image for {recipient}")

                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(sender_email, sender_pass)
                    smtp.send_message(msg)
                sent_count += 1
                self.mark_invitation_sent(recipient, name)
                self.update_invitee_status(recipient, name)
                self.log(f"[{recipient}] Invitation sent successfully.")
            except Exception as e:
                failed.append((recipient, str(e)))
                self.log(f"[{recipient}] Failed to send: {e}")

        result_msg = f"Sent: {sent_count} invitations."
        if skipped:
            result_msg += f"\nSkipped (already sent): {skipped}"
        if failed:
            result_msg += f"\nFailed: {len(failed)}"
        self.result_label.configure(text=result_msg, text_color="green" if sent_count else "red")
        self.log(result_msg)

if __name__ == "__main__":
    app = InvitationSenderApp()
    app.mainloop()