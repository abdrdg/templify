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
        self.valid_email_invitees = {}  # Dictionary to track which invitees have valid emails
        
        # Cancel flag for sending process
        self.is_sending = False
        
        # Pagination for large datasets
        self.items_per_page = 100
        self.current_page = 0
        self.total_pages = 0
        
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
        left_column.pack(side="left", fill="both", expand=False, padx=(0, 5))
        
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
        
        # Pagination controls frame
        pagination_frame = ctk.CTkFrame(right_column)
        pagination_frame.pack(fill="x", padx=10, pady=(0, 5))
        
        # Pagination controls
        self.prev_page_btn = ctk.CTkButton(pagination_frame, text="← Prev", width=80, command=self.prev_page, state="disabled")
        self.prev_page_btn.pack(side="left", padx=5)
        
        self.page_label = ctk.CTkLabel(pagination_frame, text="Page 1 of 1")
        self.page_label.pack(side="left", padx=10)
        
        self.next_page_btn = ctk.CTkButton(pagination_frame, text="Next →", width=80, command=self.next_page, state="disabled")
        self.next_page_btn.pack(side="left", padx=5)
        
        self.refresh_btn = ctk.CTkButton(pagination_frame, text="Refresh", width=80, command=self.update_status_list)
        self.refresh_btn.pack(side="right", padx=5)
        
        # Selection and refresh buttons frame
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

        self.result_label = ctk.CTkLabel(send_frame, text="...", font=("Arial", 12))
        self.result_label.pack(pady=(0, 5))

    def prev_page(self):
        """Go to previous page"""
        if self.current_page > 0:
            self.current_page -= 1
            self.update_status_list()

    def next_page(self):
        """Go to next page"""
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.update_status_list()

    def update_pagination_controls(self):
        """Update pagination button states and labels"""
        if self.invitees is None or self.invitees.empty:
            self.total_pages = 0
            self.current_page = 0
        else:
            self.total_pages = (len(self.invitees) + self.items_per_page - 1) // self.items_per_page
            if self.current_page >= self.total_pages:
                self.current_page = max(0, self.total_pages - 1)
        
        # Update buttons
        self.prev_page_btn.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_page_btn.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")
        
        # Update label
        if self.total_pages > 0:
            start_item = self.current_page * self.items_per_page + 1
            end_item = min((self.current_page + 1) * self.items_per_page, len(self.invitees))
            self.page_label.configure(text=f"Items {start_item}-{end_item} of {len(self.invitees)} (Page {self.current_page + 1} of {self.total_pages})")
        else:
            self.page_label.configure(text="No items")

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
        """Select all invitees with valid emails for sending (across all pages)"""
        if not hasattr(self, 'invitees') or self.invitees is None:
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            return

        count = 0
        total_count = 0
        
        # Work with all invitees, not just visible ones
        for idx, row in self.invitees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            
            key = f"{email}|{name}"
            has_valid_email = self.is_valid_email(email)
            
            # Create checkbox variable if it doesn't exist
            if key not in self.selected_invitees:
                self.selected_invitees[key] = ctk.BooleanVar()
            if key not in self.valid_email_invitees:
                self.valid_email_invitees[key] = has_valid_email
                
            total_count += 1
            if has_valid_email:
                self.selected_invitees[key].set(True)
                count += 1
            else:
                self.selected_invitees[key].set(False)  # Ensure invalid emails stay deselected
                
        self.log(f"Selected {count} invitees with valid emails out of {total_count} total.")

    def select_none_invitees(self):
        """Deselect all invitees (across all pages)"""
        if not hasattr(self, 'invitees') or self.invitees is None:
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            return

        count = 0
        # Work with all invitees, not just visible ones
        for idx, row in self.invitees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            
            key = f"{email}|{name}"
            
            # Create checkbox variable if it doesn't exist
            if key not in self.selected_invitees:
                self.selected_invitees[key] = ctk.BooleanVar()
                
            self.selected_invitees[key].set(False)
            count += 1
            
        self.log(f"Deselected all {count} invitees.")

    def select_unsent_invitees(self):
        """Select only invitees with valid emails who haven't been sent invitations yet (across all pages)"""
        if not hasattr(self, 'invitees') or self.invitees is None:
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            return

        selected_count = 0
        total_count = 0
        
        # Work with all invitees, not just visible ones
        for idx, row in self.invitees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            
            key = f"{email}|{name}"
            has_valid_email = self.is_valid_email(email)
            
            # Create checkbox variable if it doesn't exist
            if key not in self.selected_invitees:
                self.selected_invitees[key] = ctk.BooleanVar()
            if key not in self.valid_email_invitees:
                self.valid_email_invitees[key] = has_valid_email
                
            total_count += 1
            
            # Only consider invitees with valid emails
            if has_valid_email:
                if not self.was_invitation_sent(email, name):
                    self.selected_invitees[key].set(True)
                    selected_count += 1
                else:
                    self.selected_invitees[key].set(False)
            else:
                self.selected_invitees[key].set(False)  # Ensure invalid emails stay deselected
                
        self.log(f"Selected {selected_count} unsent invitees with valid emails out of {total_count} total.")

    def clear_status_list(self):
        """Clear status widgets but preserve selection state"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.status_labels = {}
        # DON'T clear selected_invitees and valid_email_invitees - preserve selections across pages

    def reset_all_selections(self):
        """Completely reset all selections (used when loading new Excel file)"""
        self.selected_invitees = {}
        self.valid_email_invitees = {}

    def update_status_list(self):
        """Update the status list with current invitees and checkboxes - optimized with pagination"""
        self.clear_status_list()
        
        if not hasattr(self, 'invitees') or self.invitees is None:
            self.update_pagination_controls()
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            self.update_pagination_controls()
            return

        # Update pagination controls
        self.update_pagination_controls()
        
        # Calculate which items to show
        start_idx = self.current_page * self.items_per_page
        end_idx = min(start_idx + self.items_per_page, len(self.invitees))
        
        # Only create widgets for visible items
        visible_invitees = self.invitees.iloc[start_idx:end_idx]
        
        # Batch process invitee data first
        invitee_data = []
        for idx, row in visible_invitees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            
            # Check if email is valid
            has_valid_email = self.is_valid_email(email)
            is_sent = self.was_invitation_sent(email, name) if has_valid_email else False
            
            invitee_data.append({
                'idx': idx,
                'name': name,
                'email': email,
                'has_valid_email': has_valid_email,
                'is_sent': is_sent
            })
        
        # Now create UI elements in batch
        for data in invitee_data:
            self._create_invitee_widget(data)

    def _create_invitee_widget(self, data):
        """Create UI widget for a single invitee"""
        name = data['name']
        email = data['email'] 
        has_valid_email = data['has_valid_email']
        is_sent = data['is_sent']
        
        key = f"{email}|{name}"
        
        # Create frame for this invitee
        frame = ctk.CTkFrame(self.scrollable_frame)
        frame.pack(fill="x", padx=2, pady=1)
        
        # Checkbox for selection - disabled if no valid email
        checkbox_var = ctk.BooleanVar()
        checkbox = ctk.CTkCheckBox(
            frame, 
            text="", 
            variable=checkbox_var, 
            width=20,
            state="normal" if has_valid_email else "disabled"
        )
        checkbox.pack(side="left", padx=5)
        
        # Restore or set selection state
        if key in self.selected_invitees:
            # Restore previous selection state
            previous_state = self.selected_invitees[key].get()
            checkbox_var.set(previous_state)
        else:
            # Set default selection
            if not has_valid_email:
                checkbox_var.set(False)  # Don't select invalid emails
            else:
                checkbox_var.set(not is_sent)  # Select unsent invitees by default
        
        # Store checkbox variable and email validity for later use
        self.selected_invitees[key] = checkbox_var
        self.valid_email_invitees[key] = has_valid_email
        
        # Name and email display
        if has_valid_email:
            info_text = f"{name} ({email})"
            text_color = None  # Default color
        else:
            if not email or email.lower() in ['nan', 'none']:
                info_text = f"{name} (No email address)"
            else:
                info_text = f"{name} ({email} - Invalid email)"
            text_color = "gray"
        
        info_label = ctk.CTkLabel(frame, text=info_text, anchor="w", text_color=text_color)
        info_label.pack(side="left", padx=5, fill="x", expand=True)
        
        # Status label
        status_label = ctk.CTkLabel(frame, text="", anchor="e", width=120)
        status_label.pack(side="right", padx=5)
        
        # Store label reference for updates
        self.status_labels[key] = status_label
        
        # Update status
        if not has_valid_email:
            status_label.configure(text="Cannot send", text_color="red")
        elif is_sent:
            sent_date = self.sent_invitations[key]["sent_date"]
            status_label.configure(text=f"Sent ✓", text_color="green")
        else:
            status_label.configure(text="Not sent", text_color="gray")

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
                # Keep all rows, don't filter out missing data
                self.invitees = df
                total_invitees = len(df)
                # Count how many have valid emails
                valid_emails = 0
                for idx, row in df.iterrows():
                    email = str(row[email_guess]).strip() if pd.notna(row[email_guess]) else ""
                    if self.is_valid_email(email):
                        valid_emails += 1
                
                self.excel_path = file_path
                self.status_label.configure(
                    text=f"Loaded {total_invitees} invitees ({valid_emails} with valid emails).", 
                    text_color="green"
                )
                self.send_btn.configure(state="normal")
                self.log(f"Loaded {total_invitees} invitees from Excel ({valid_emails} with valid emails).")
                
                # Reset selections when loading new file
                self.reset_all_selections()
                self.current_page = 0  # Reset to first page
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
        # Clean the name and remove invalid filename characters
        cleaned_name = str(name).replace('\n', ' ')
        # Replace invalid Windows filename characters
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            cleaned_name = cleaned_name.replace(char, ' ')
        # Remove dots and normalize spaces
        return ' '.join(part.replace('.', '') for part in cleaned_name.split())

    def find_invitation_image(self, name):
        """Find invitation image file, trying different filename variations for backward compatibility"""
        # Try the current cleaned filename first
        cleaned_name = self.clean_name(name)
        primary_filename = os.path.join(self.images_folder, f"Invitation - {cleaned_name}.png")
        
        if os.path.exists(primary_filename):
            return primary_filename
        
        # Try various legacy processing variations for backward compatibility
        variations = [
            # Legacy processing (old get_filename logic)
            str(name).replace("\n", " ").replace(".", "").replace('"', "'").strip(),
            # Raw name with just quote replacement
            str(name).replace('"', "'"),
            # Raw name with no processing
            str(name),
            # Name with just space normalization (but keeping leading/trailing)
            str(name).replace("\n", " ").replace(".", "").replace('"', "'"),
            # Legacy with slash replacement (in case files were created with slash handling)
            str(name).replace("\n", " ").replace(".", "").replace('"', "'").replace("/", " ").replace("\\", " ").strip(),
        ]
        
        for variation in variations:
            # Try exact variation
            filename = os.path.join(self.images_folder, f"Invitation - {variation}.png")
            if os.path.exists(filename):
                return filename
                
            # Also try with potential extra space after dash (leading space in name)
            filename_extra_space = os.path.join(self.images_folder, f"Invitation -  {variation}.png")
            if os.path.exists(filename_extra_space):
                return filename_extra_space
        
        # Try to find any file that starts with "Invitation - " and contains parts of the name
        # This is a fuzzy matching approach for difficult cases
        if os.path.exists(self.images_folder):
            name_parts = cleaned_name.lower().split()
            if name_parts:
                try:
                    for filename in os.listdir(self.images_folder):
                        if filename.startswith("Invitation - ") and filename.endswith(".png"):
                            file_name_part = filename[13:-4].lower()  # Remove "Invitation - " and ".png"
                            # Check if all name parts are present in the filename
                            if all(part in file_name_part for part in name_parts):
                                full_path = os.path.join(self.images_folder, filename)
                                return full_path
                except OSError:
                    pass  # Handle permission errors gracefully
        
        # None found
        return None

    def is_valid_email(self, email):
        """Check if email address is valid (basic validation)"""
        if not email or email.lower() in ['nan', 'none', '']:
            return False
        return '@' in email and '.' in email and len(email) > 5

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
        
        # First, count selected invitees with valid emails
        for idx, row in self.invitees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            recipient = str(email_raw).strip()
            key = f"{recipient}|{name}"
            
            if key in self.selected_invitees and self.selected_invitees[key].get():
                if self.is_valid_email(recipient):
                    selected_count += 1
        
        if selected_count == 0:
            self.after(0, self.log, "No invitees selected for sending.")
            self.after(0, self.finish_sending, 0, 0, [])
            return
        
        self.after(0, self.log, f"Starting to send {selected_count} selected invitations...")
        current_processed = 0
        
        for idx, row in self.invitees.iterrows():
            # Check for cancellation
            if not self.is_sending:
                self.after(0, self.log, "Sending cancelled.")
                return
                
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            recipient = str(email_raw).strip()
            key = f"{recipient}|{name}"
            
            # Skip if not selected
            if key not in self.selected_invitees or not self.selected_invitees[key].get():
                continue
                
            # Skip if email is not valid
            if not self.is_valid_email(recipient):
                self.after(0, self.log, f"[SKIPPED] Invalid email for {name}: {recipient}")
                continue
                
            current_processed += 1
            
            # Update progress in the main thread
            self.after(0, self.update_progress, current_processed, selected_count, f"Processing: {name} ({recipient})")
            
            # Check if invitation was already sent
            if self.was_invitation_sent(recipient, name):
                self.after(0, self.log, f"[SKIPPED] Already sent to {name} ({recipient})")
                skipped += 1
                continue
                
            img_filename = self.find_invitation_image(name)
            if img_filename is None:
                # Try to provide helpful info about what files we looked for
                cleaned_name = self.clean_name(name)
                expected_filename = f"Invitation - {cleaned_name}.png"
                failed.append((recipient, "Invitation image not found"))
                self.after(0, self.log, f"[{recipient}] Invitation image not found. Expected: {expected_filename}")
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
        
        # Status refresh happens automatically via the wrapper's finally block

    def send_invitations(self):
        if self.is_sending:
            # Cancel sending
            self.is_sending = False
            self.log("Sending cancelled by user.")
            self.reset_send_button()
            return
            
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

        # Start sending process
        self.is_sending = True
        self.send_btn.configure(text="Cancel", fg_color="red")
        
        # Show progress bar
        self.progress_frame.pack(pady=5, fill="x", padx=10)
        self.log(f"Starting to send invitations from {sender_email}...")
        
        # Start sending thread
        threading.Thread(
            target=self._send_invitations_thread,
            args=(sender_email, sender_pass, email_col, name_col),
            daemon=True
        ).start()

    def reset_send_button(self):
        """Reset the send button to its original state"""
        self.is_sending = False
        self.send_btn.configure(text="Send Invitations", fg_color=["#1f538d", "#14375e"])
        self.progress_frame.pack_forget()  # Hide progress bar

    def _send_invitations_thread(self, sender_email, sender_pass, email_col, name_col):
        try:
            self.send_invitations_thread(sender_email, sender_pass, email_col, name_col)
        finally:
            # Always reset the button when sending ends
            self.after(0, self.reset_send_button)

if __name__ == "__main__":
    app = InvitationSenderApp()
    app.mainloop()