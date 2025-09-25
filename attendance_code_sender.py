import customtkinter as ctk
import pandas as pd
import tkinter.filedialog as fd
import smtplib
import json
import os
import threading
import re
from datetime import datetime
from email.message import EmailMessage

# Set the appearance mode and color theme
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class AttendanceCodeSenderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Code Sender")
        self.geometry("900x900")  # Made wider for two-column layout
        self.minsize(900, 900)    # Set minimum size to prevent elements from being hidden
        self.resizable(True, True) # Allow resizing for better usability

        self.excel_path = None
        self.attendees = None
        self.email_column = None
        self.name_column = None
        self.code_column = None
        self.email_column_var = ctk.StringVar()
        self.name_column_var = ctk.StringVar()
        self.code_column_var = ctk.StringVar()
        
        # Initialize sent codes tracking
        self.tracking_file = "sent_attendance_codes.json"
        self.sent_codes = self.load_sent_codes()
        
        # Initialize template saving
        self.template_file = "email_templates.json"
        
        # Initialize selection tracking
        self.selected_attendees = {}  # Dictionary to track checkbox states
        self.valid_email_attendees = {}  # Dictionary to track which attendees have valid emails
        
        # Cancel flag for sending process
        self.is_sending = False
        
        # Pagination for large datasets
        self.items_per_page = 100
        self.current_page = 0
        self.total_pages = 0
        
        self.create_widgets()
        
    def destroy(self):
        """Override destroy to save template before closing"""
        try:
            # Auto-save the current template before closing
            if hasattr(self, 'body_textbox') and hasattr(self, 'subject_entry'):
                self.save_current_template()
        except Exception as e:
            print(f"Warning: Could not auto-save template: {e}")
        finally:
            super().destroy()
        
    def load_sent_codes(self):
        """Load the record of sent attendance codes from JSON file"""
        if os.path.exists(self.tracking_file):
            try:
                with open(self.tracking_file, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                self.log("Warning: Tracking file corrupted, starting fresh.")
                return {}
        return {}
        
    def save_sent_codes(self):
        """Save the record of sent attendance codes to JSON file"""
        with open(self.tracking_file, 'w') as f:
            json.dump(self.sent_codes, f, indent=2)
            
    def was_code_sent(self, email, name):
        """Check if an attendance code was already sent to this person"""
        key = f"{email}|{name}"
        return key in self.sent_codes
        
    def mark_code_sent(self, email, name, code):
        """Mark an attendance code as sent for this person"""
        key = f"{email}|{name}"
        self.sent_codes[key] = {
            "email": email,
            "name": name,
            "code": code,
            "sent_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.save_sent_codes()
        
    def get_sent_code(self, email, name):
        """Get the attendance code that was sent to this person"""
        key = f"{email}|{name}"
        if key in self.sent_codes:
            return self.sent_codes[key]["code"]
        return None

    def format_code(self, code_value):
        """Format code as 3-digit padded string"""
        try:
            # Convert to int first to handle any float values from Excel
            num = int(float(str(code_value)))
            return f"{num:03d}"
        except (ValueError, TypeError):
            return None

    def load_email_templates(self):
        """Load email templates from JSON file"""
        if os.path.exists(self.template_file):
            try:
                with open(self.template_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, UnicodeDecodeError):
                self.log("Warning: Template file corrupted, using defaults.")
                return {}
        return {}

    def save_email_templates(self, templates):
        """Save email templates to JSON file"""
        try:
            with open(self.template_file, 'w', encoding='utf-8') as f:
                json.dump(templates, f, indent=2, ensure_ascii=False)
        except Exception as e:
            self.log(f"Warning: Could not save templates: {e}")

    def save_current_template(self):
        """Save the current email template"""
        if hasattr(self, 'body_textbox') and hasattr(self, 'subject_entry'):
            current_body = self.body_textbox.get("1.0", "end-1c").strip()
            current_subject = self.subject_entry.get().strip()
            
            templates = self.load_email_templates()
            templates['attendance_subject'] = current_subject
            templates['attendance_body'] = current_body
            self.save_email_templates(templates)
            self.log("Email template saved successfully!")

    def reset_template(self):
        """Reset the email template to default"""
        default_subject = "Your Attendance Code - Korean Embassy National Day"
        
        default_body = """Greetings from the Embassy of the Republic of Korea.

We wish to convey our sincere appreciation for your kind confirmation to attend the Reception in celebration of the National Day and Armed Forces Day of the Republic of Korea.

Date and time: 01 October 2025, 6:30pm

Venue: Grand Ballroom, Grand Hyatt Manila, 8th Avenue, Corner 35th St, Taguig, 1634 Metro Manila

Attire: Business

For security and registration purposes, may we respectfully request that you present your 3-digit invitation code upon entry to the event venue.

<h3>[Code]</h3>

Should you require any assistance or  further information, please do not hesitate to email us at rokphilamb@mofa.or.kr.

We look forward to welcoming you soon.

With highest regards,
Embassy of the Republic of Korea"""
        
        if hasattr(self, 'body_textbox'):
            self.body_textbox.delete("1.0", "end")
            self.body_textbox.insert("1.0", default_body)
            
        if hasattr(self, 'subject_entry'):
            self.subject_entry.delete(0, "end")
            self.subject_entry.insert(0, default_subject)
            
        self.log("Template reset to default.")

    def create_widgets(self):
        # Use a main frame to control layout and allow expansion
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title at the top
        self.label = ctk.CTkLabel(main_frame, text="Send Attendance Codes", font=("Arial", 22))
        self.label.pack(pady=(0, 10))
        
        # Create two-column layout using a horizontal frame
        columns_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        columns_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Left column - Controls
        left_column = ctk.CTkFrame(columns_frame)
        left_column.pack(side="left", fill="both", expand=False, padx=(0, 5))
        
        # Right column - Attendees list
        right_column = ctk.CTkFrame(columns_frame)
        right_column.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # === LEFT COLUMN CONTENT ===
        
        # Excel file section
        excel_frame = ctk.CTkFrame(left_column)
        excel_frame.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(excel_frame, text="Excel File:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
        self.open_btn = ctk.CTkButton(excel_frame, text="Open Excel File", command=self.open_excel)
        self.open_btn.pack(pady=5)
        self.status_label = ctk.CTkLabel(excel_frame, text="No file selected.", font=("Arial", 11))
        self.status_label.pack(pady=(0, 5))

        # Dropdowns for selecting columns
        columns_section = ctk.CTkFrame(left_column)
        columns_section.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(columns_section, text="Column Mapping:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
        
        email_col_frame = ctk.CTkFrame(columns_section, fg_color="transparent")
        email_col_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(email_col_frame, text="Email Column:", width=100).pack(side="left")
        self.email_column_menu = ctk.CTkOptionMenu(email_col_frame, variable=self.email_column_var, values=[], command=self.on_column_change)
        self.email_column_menu.pack(side="right", fill="x", expand=True)
        
        name_col_frame = ctk.CTkFrame(columns_section, fg_color="transparent")
        name_col_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(name_col_frame, text="Name Column:", width=100).pack(side="left")
        self.name_column_menu = ctk.CTkOptionMenu(name_col_frame, variable=self.name_column_var, values=[], command=self.on_column_change)
        self.name_column_menu.pack(side="right", fill="x", expand=True)
        
        code_col_frame = ctk.CTkFrame(columns_section, fg_color="transparent")
        code_col_frame.pack(fill="x", padx=5, pady=(2, 5))
        ctk.CTkLabel(code_col_frame, text="Code Column:", width=100).pack(side="left")
        self.code_column_menu = ctk.CTkOptionMenu(code_col_frame, variable=self.code_column_var, values=[], command=self.on_column_change)
        self.code_column_menu.pack(side="right", fill="x", expand=True)

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

        # Email template section
        template_frame = ctk.CTkFrame(left_column)
        template_frame.pack(pady=5, fill="both", expand=False, padx=10)
        
        # Header with title and template buttons
        header_frame = ctk.CTkFrame(template_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=5, pady=(5, 0))
        
        ctk.CTkLabel(header_frame, text="Email Template:", font=("Arial", 12, "bold")).pack(side="left", anchor="w")
        
        # Button frame for template buttons
        button_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        button_frame.pack(side="right")
        
        self.reset_template_btn = ctk.CTkButton(
            button_frame, 
            text="Reset", 
            width=60, 
            height=24,
            font=("Arial", 10),
            command=self.reset_template
        )
        self.reset_template_btn.pack(side="right", padx=(0, 5))
        
        self.save_template_btn = ctk.CTkButton(
            button_frame, 
            text="Save Template", 
            width=100, 
            height=24,
            font=("Arial", 10),
            command=self.save_current_template
        )
        self.save_template_btn.pack(side="right")
        
        # Subject line section
        subject_frame = ctk.CTkFrame(template_frame, fg_color="transparent")
        subject_frame.pack(fill="x", padx=5, pady=(5, 5))
        
        ctk.CTkLabel(subject_frame, text="Subject:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        
        self.subject_entry = ctk.CTkEntry(subject_frame, placeholder_text="Email subject line")
        self.subject_entry.pack(side="right", fill="x", expand=True)
        
        # Email body section
        ctk.CTkLabel(template_frame, text="Email Body:", font=("Arial", 11, "bold")).pack(anchor="w", padx=5, pady=(5, 0))
        
        # HTML formatting help label
        help_text = "HTML formatting: <b>bold</b>, <i>italic</i>, <u>underline</u>, <a href='url'>links</a>, <br> line breaks, <p>paragraphs</p>"
        ctk.CTkLabel(template_frame, text=help_text, font=("Arial", 9), text_color="gray").pack(anchor="w", padx=5, pady=(0, 5))
        
        # Info label for placeholders
        info_text = "Use [Name] for attendee name and [Code] for the 3-digit code"
        ctk.CTkLabel(template_frame, text=info_text, font=("Arial", 9), text_color="gray").pack(anchor="w", padx=5, pady=(0, 5))
        
        self.body_textbox = ctk.CTkTextbox(template_frame, height=120, wrap="word")
        self.body_textbox.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        
        # Load saved template or use default
        templates = self.load_email_templates()
        
        if 'attendance_subject' in templates and templates['attendance_subject'].strip():
            # Use saved subject
            saved_subject = templates['attendance_subject']
            self.subject_entry.insert(0, saved_subject)
        else:
            # Set default subject
            default_subject = "Your Attendance Code - Korean Embassy National Day"
            self.subject_entry.insert(0, default_subject)
        
        if 'attendance_body' in templates and templates['attendance_body'].strip():
            # Use saved template
            saved_body = templates['attendance_body']
            self.body_textbox.insert("1.0", saved_body)
        else:
            # Set default email body
            default_body = """Dear [Name],

Your attendance code for the Korean Embassy National Day and Armed Forces Day Reception is:

**[Code]**

Please present this code upon arrival at the venue for verification.

Event Details:
Date: Wednesday, 01 October 2025
Time: 6:30 PM
Venue: Grand Ballroom, Grand Hyatt Manila, Taguig City
Attire: Business Formal

We look forward to welcoming you to this special celebration.

Best regards,
Embassy of the Republic of Korea"""
            
            self.body_textbox.insert("1.0", default_body)

        # Log area
        log_frame = ctk.CTkFrame(left_column)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_label = ctk.CTkLabel(log_frame, text="Log:", font=("Arial", 12, "bold"))
        self.log_label.pack(pady=(5,0), anchor="w", padx=5)
        self.log_textbox = ctk.CTkTextbox(log_frame, wrap="word")
        self.log_textbox.pack(fill="both", expand=True, padx=5, pady=(0,5))
        self.log_textbox.configure(state="disabled")

        # === RIGHT COLUMN CONTENT ===
        
        # Attendees list with status
        ctk.CTkLabel(right_column, text="Attendees Status", font=("Arial", 16, "bold")).pack(pady=(10, 5))
        
        # Pagination controls frame
        pagination_frame = ctk.CTkFrame(right_column)
        pagination_frame.pack(fill="x", padx=10, pady=(0, 5))
        
        # Items per page control
        items_per_page_frame = ctk.CTkFrame(pagination_frame)
        items_per_page_frame.pack(side="left", padx=5)
        
        ctk.CTkLabel(items_per_page_frame, text="Items per page:").pack(side="left", padx=2)
        self.items_per_page_entry = ctk.CTkEntry(items_per_page_frame, width=60)
        self.items_per_page_entry.pack(side="left", padx=2)
        self.items_per_page_entry.insert(0, str(self.items_per_page))
        self.items_per_page_entry.bind("<Return>", self.update_items_per_page)
        self.items_per_page_entry.bind("<FocusOut>", self.update_items_per_page)
        
        # Pagination controls
        self.prev_page_btn = ctk.CTkButton(pagination_frame, text="← Prev", width=80, command=self.prev_page, state="disabled")
        self.prev_page_btn.pack(side="left", padx=5)
        
        self.page_label = ctk.CTkLabel(pagination_frame, text="Page 1 of 1")
        self.page_label.pack(side="left", padx=10)
        
        self.next_page_btn = ctk.CTkButton(pagination_frame, text="Next →", width=80, command=self.next_page, state="disabled")
        self.next_page_btn.pack(side="left", padx=5)
        
        self.refresh_btn = ctk.CTkButton(pagination_frame, text="Refresh", width=80, command=self.update_status_list)
        self.refresh_btn.pack(side="right", padx=5)
        
        # Selection buttons frame
        selection_frame = ctk.CTkFrame(right_column)
        selection_frame.pack(fill="x", padx=10, pady=(0, 5))
        self.select_all_btn = ctk.CTkButton(selection_frame, text="Select All", width=80, command=self.select_all_attendees)
        self.select_all_btn.pack(side="left", padx=2)
        self.select_none_btn = ctk.CTkButton(selection_frame, text="Select None", width=80, command=self.select_none_attendees)
        self.select_none_btn.pack(side="left", padx=2)
        self.select_unsent_btn = ctk.CTkButton(selection_frame, text="Select Unsent", width=90, command=self.select_unsent_attendees)
        self.select_unsent_btn.pack(side="left", padx=2)
        
        # Scrollable frame for attendees - takes up most of the right column
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
            text="Send Attendance Codes", 
            command=self.send_codes, 
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
        if self.attendees is None or self.attendees.empty:
            self.total_pages = 0
            self.current_page = 0
        else:
            self.total_pages = (len(self.attendees) + self.items_per_page - 1) // self.items_per_page
            if self.current_page >= self.total_pages:
                self.current_page = max(0, self.total_pages - 1)
        
        # Update buttons
        self.prev_page_btn.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_page_btn.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")
        
        # Update label
        if self.total_pages > 0:
            start_item = self.current_page * self.items_per_page + 1
            end_item = min((self.current_page + 1) * self.items_per_page, len(self.attendees))
            self.page_label.configure(text=f"Items {start_item}-{end_item} of {len(self.attendees)} (Page {self.current_page + 1} of {self.total_pages})")
        else:
            self.page_label.configure(text="No items")

    def update_items_per_page(self, event=None):
        """Update items per page based on user input"""
        try:
            value = self.items_per_page_entry.get().strip()
            if not value:
                return
            
            new_items_per_page = int(value)
            
            # Validate range (minimum 1, maximum 1000)
            if new_items_per_page < 1:
                new_items_per_page = 1
                self.items_per_page_entry.delete(0, 'end')
                self.items_per_page_entry.insert(0, "1")
            elif new_items_per_page > 1000:
                new_items_per_page = 1000
                self.items_per_page_entry.delete(0, 'end')
                self.items_per_page_entry.insert(0, "1000")
            
            # Update items per page and refresh if changed
            if new_items_per_page != self.items_per_page:
                self.items_per_page = new_items_per_page
                self.current_page = 0  # Reset to first page
                self.update_status_list()
                
        except ValueError:
            # Reset to current value if invalid input
            self.items_per_page_entry.delete(0, 'end')
            self.items_per_page_entry.insert(0, str(self.items_per_page))

    def log(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def on_column_change(self, selected_value):
        """Handle column mapping change - refresh the attendee list"""
        if hasattr(self, 'attendees') and self.attendees is not None:
            # Reset to first page when column mapping changes
            self.current_page = 0
            # Reset selections when column mapping changes
            self.reset_all_selections()
            # Refresh the status list with new column mapping
            self.update_status_list()
            self.log(f"Column mapping updated. Refreshed attendee list.")

    def select_all_attendees(self):
        """Select all attendees with valid emails for sending (across all pages)"""
        if not hasattr(self, 'attendees') or self.attendees is None:
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            return

        count = 0
        total_count = 0
        
        # Work with all attendees, not just visible ones
        for idx, row in self.attendees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            
            key = f"{email}|{name}"
            has_valid_email = self.is_valid_email(email)
            
            # Create checkbox variable if it doesn't exist
            if key not in self.selected_attendees:
                self.selected_attendees[key] = ctk.BooleanVar()
            if key not in self.valid_email_attendees:
                self.valid_email_attendees[key] = has_valid_email
                
            total_count += 1
            if has_valid_email:
                self.selected_attendees[key].set(True)
                count += 1
            else:
                self.selected_attendees[key].set(False)  # Ensure invalid emails stay deselected
                
        self.log(f"Selected {count} attendees with valid emails out of {total_count} total.")

    def select_none_attendees(self):
        """Deselect all attendees (across all pages)"""
        if not hasattr(self, 'attendees') or self.attendees is None:
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            return

        count = 0
        # Work with all attendees, not just visible ones
        for idx, row in self.attendees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            
            key = f"{email}|{name}"
            
            # Create checkbox variable if it doesn't exist
            if key not in self.selected_attendees:
                self.selected_attendees[key] = ctk.BooleanVar()
                
            self.selected_attendees[key].set(False)
            count += 1
            
        self.log(f"Deselected all {count} attendees.")

    def select_unsent_attendees(self):
        """Select only attendees with valid emails who haven't been sent codes yet (across all pages)"""
        if not hasattr(self, 'attendees') or self.attendees is None:
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        if not email_col or not name_col:
            return

        selected_count = 0
        total_count = 0
        
        # Work with all attendees, not just visible ones
        for idx, row in self.attendees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            
            key = f"{email}|{name}"
            has_valid_email = self.is_valid_email(email)
            
            # Create checkbox variable if it doesn't exist
            if key not in self.selected_attendees:
                self.selected_attendees[key] = ctk.BooleanVar()
            if key not in self.valid_email_attendees:
                self.valid_email_attendees[key] = has_valid_email
                
            total_count += 1
            
            # Only consider attendees with valid emails
            if has_valid_email:
                # Check if code was already sent
                if not self.was_code_sent(email, name):
                    self.selected_attendees[key].set(True)
                    selected_count += 1
                else:
                    self.selected_attendees[key].set(False)
            else:
                self.selected_attendees[key].set(False)  # Ensure invalid emails stay deselected
                
        self.log(f"Selected {selected_count} unsent attendance codes with valid emails out of {total_count} total.")

    def clean_name(self, name):
        # Clean the name and remove invalid filename characters
        cleaned_name = str(name).replace('\n', ' ')
        # Replace invalid Windows filename characters
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            cleaned_name = cleaned_name.replace(char, ' ')
        # Remove dots and normalize spaces
        return ' '.join(part.replace('.', '') for part in cleaned_name.split())

    def is_valid_email(self, email):
        """Check if email address is valid (basic validation)"""
        if not email or email.lower() in ['nan', 'none', '']:
            return False
        return '@' in email and '.' in email and len(email) > 5

    def clear_status_list(self):
        """Clear status widgets but preserve selection state"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.status_labels = {}
        # DON'T clear selected_attendees and valid_email_attendees - preserve selections across pages

    def reset_all_selections(self):
        """Completely reset all selections (used when loading new Excel file)"""
        self.selected_attendees = {}
        self.valid_email_attendees = {}

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
                    self.code_column_menu.configure(values=[])
                    self.log("Excel file has no columns.")
                    return
                
                # Set dropdowns for email, name and code columns
                self.email_column_menu.configure(values=columns)
                self.name_column_menu.configure(values=columns)
                self.code_column_menu.configure(values=columns)
                
                # Try to auto-select likely columns
                email_guess = next((c for c in columns if 'email' in c.lower()), columns[0])
                name_guess = next((c for c in columns if 'name' in c.lower()), columns[0])
                code_guess = next((c for c in columns if 'code' in c.lower() or 'id' in c.lower() or 'number' in c.lower()), columns[0])
                
                self.email_column_var.set(email_guess)
                self.name_column_var.set(name_guess)
                self.code_column_var.set(code_guess)
                
                # Keep all rows, don't filter out missing data
                self.attendees = df
                total_attendees = len(df)
                
                # Count how many have valid emails and codes
                valid_entries = 0
                for idx, row in df.iterrows():
                    email = str(row[email_guess]).strip() if pd.notna(row[email_guess]) else ""
                    code = row[code_guess] if pd.notna(row[code_guess]) else None
                    if self.is_valid_email(email) and self.format_code(code) is not None:
                        valid_entries += 1
                
                self.excel_path = file_path
                self.status_label.configure(
                    text=f"Loaded {total_attendees} attendees ({valid_entries} with valid emails and codes).", 
                    text_color="green"
                )
                self.send_btn.configure(state="normal")
                self.log(f"Loaded {total_attendees} attendees from Excel ({valid_entries} with valid emails and codes).")
                
                # Reset selections when loading new file
                self.reset_all_selections()
                self.current_page = 0  # Reset to first page
                self.update_status_list()
                
            except Exception as e:
                self.status_label.configure(text=f"Error: {e}", text_color="red")
                self.send_btn.configure(state="disabled")
                self.email_column_menu.configure(values=[])
                self.name_column_menu.configure(values=[])
                self.code_column_menu.configure(values=[])
                self.log(f"Error loading Excel: {e}")
        else:
            self.status_label.configure(text="No file selected.", text_color="gray")
            self.send_btn.configure(state="disabled")
            self.email_column_menu.configure(values=[])
            self.name_column_menu.configure(values=[])
            self.code_column_menu.configure(values=[])
            self.log("No file selected.")

    def update_status_list(self):
        """Update the status list with current attendees and checkboxes - optimized with pagination"""
        self.clear_status_list()
        
        if not hasattr(self, 'attendees') or self.attendees is None:
            self.update_pagination_controls()
            return

        email_col = self.email_column_var.get()
        name_col = self.name_column_var.get()
        code_col = self.code_column_var.get()
        if not email_col or not name_col or not code_col:
            self.update_pagination_controls()
            return

        # Update pagination controls
        self.update_pagination_controls()
        
        # Calculate which items to show
        start_idx = self.current_page * self.items_per_page
        end_idx = min(start_idx + self.items_per_page, len(self.attendees))
        
        # Only create widgets for visible items
        visible_attendees = self.attendees.iloc[start_idx:end_idx]
        
        # Batch process attendee data first
        attendee_data = []
        for idx, row in visible_attendees.iterrows():
            name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
            name = self.clean_name(str(name_raw).strip())
            email_raw = row[email_col] if pd.notna(row[email_col]) else ""
            email = str(email_raw).strip()
            code_raw = row[code_col] if pd.notna(row[code_col]) else None
            formatted_code = self.format_code(code_raw)
            
            # Check if email is valid and code is valid
            has_valid_email = self.is_valid_email(email)
            has_valid_code = formatted_code is not None
            was_sent = self.was_code_sent(email, name)
            sent_code = self.get_sent_code(email, name)
            
            attendee_data.append({
                'idx': idx,
                'name': name,
                'email': email,
                'formatted_code': formatted_code,
                'has_valid_email': has_valid_email,
                'has_valid_code': has_valid_code,
                'was_sent': was_sent,
                'sent_code': sent_code
            })
        
        # Now create UI elements in batch
        for data in attendee_data:
            self._create_attendee_widget(data)

    def _create_attendee_widget(self, data):
        """Create UI widget for a single attendee"""
        name = data['name']
        email = data['email'] 
        formatted_code = data['formatted_code']
        has_valid_email = data['has_valid_email']
        has_valid_code = data['has_valid_code']
        was_sent = data['was_sent']
        sent_code = data['sent_code']
        
        key = f"{email}|{name}"
        
        # Create frame for this attendee
        frame = ctk.CTkFrame(self.scrollable_frame)
        frame.pack(fill="x", padx=2, pady=1)
        
        # Checkbox for selection - disabled if no valid email or code
        can_send = has_valid_email and has_valid_code
        checkbox_var = ctk.BooleanVar()
        checkbox = ctk.CTkCheckBox(
            frame, 
            text="", 
            variable=checkbox_var, 
            width=20,
            state="normal" if can_send else "disabled"
        )
        checkbox.pack(side="left", padx=5)
        
        # Restore or set selection state
        if key in self.selected_attendees:
            # Restore previous selection state
            previous_state = self.selected_attendees[key].get()
            checkbox_var.set(previous_state)
        else:
            # Set default selection - select valid entries that haven't been sent
            checkbox_var.set(can_send and not was_sent)
        
        # Store checkbox variable and email validity for later use
        self.selected_attendees[key] = checkbox_var
        self.valid_email_attendees[key] = has_valid_email
        
        # Name and email display
        if can_send:
            code_display = sent_code if was_sent else formatted_code
            info_text = f"{name} ({email}) - Code: {code_display}"
            text_color = None  # Default color
        else:
            if not has_valid_email:
                info_text = f"{name} ({email if email else 'No email'} - Invalid email)"
            elif not has_valid_code:
                info_text = f"{name} ({email}) - Invalid code"
            else:
                info_text = f"{name} ({email}) - Cannot send"
            text_color = "gray"
        
        info_label = ctk.CTkLabel(frame, text=info_text, anchor="w", text_color=text_color)
        info_label.pack(side="left", padx=5, fill="x", expand=True)
        
        # Status label
        status_label = ctk.CTkLabel(frame, text="", anchor="e", width=100)
        status_label.pack(side="right", padx=5)
        
        # Store label reference for updates
        self.status_labels[key] = status_label
        
        # Update status
        if not can_send:
            status_label.configure(text="Cannot send", text_color="red")
        elif was_sent:
            status_label.configure(text="Code sent", text_color="green")
        else:
            status_label.configure(text="Not sent", text_color="gray")

    def update_attendee_status(self, email, name):
        """Update the status display for a single attendee"""
        key = f"{email}|{name}"
        if key not in self.status_labels:
            return
            
        label = self.status_labels[key]
        was_sent = self.was_code_sent(email, name)
        
        if was_sent:
            label.configure(text="Code sent", text_color="green")
        else:
            label.configure(text="Not sent", text_color="gray")

    def update_progress(self, current, total, message=""):
        """Update the progress bar and label"""
        progress = current / total if total > 0 else 0
        self.progress_bar.set(progress)
        self.progress_label.configure(text=message)

    def send_single_code(self, sender_email, sender_pass, name, recipient, code, subject, body):
        """Send a single attendance code and return the result"""
        try:
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = sender_email
            msg["To"] = recipient
            
            # Replace placeholders in the body
            personalized_body = body.replace("[Name]", name).replace("[Code]", code)
            
            # Check if the body contains HTML tags
            has_html = any(tag in personalized_body.lower() for tag in ['<b>', '<i>', '<u>', '<strong>', '<em>', '<a', '<br', '<p>', '<div>'])
            
            # Create plain text version first
            plain_text = personalized_body
            if has_html:
                # Remove HTML tags for plain text version
                plain_text = re.sub(r'<[^>]+>', '', personalized_body)
                plain_text = plain_text.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
            
            # Set the plain text content first
            msg.set_content(plain_text)
            
            # Add HTML version if HTML tags are present
            if has_html:
                # Convert line breaks to <br> tags for proper HTML display
                html_content = personalized_body.replace('\n', '<br>')
                msg.add_alternative(f"""
                <html>
                  <body>
                    {html_content}
                  </body>
                </html>
                """, subtype='html')
            else:
                # Even for plain text, add HTML version with line breaks converted
                html_body = personalized_body.replace('\n', '<br>')
                msg.add_alternative(f"""
                <html>
                  <body>
                    {html_body}
                  </body>
                </html>
                """, subtype='html')

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(sender_email, sender_pass)
                smtp.send_message(msg)
            return True, None
        except Exception as e:
            return False, str(e)

    def send_codes(self):
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
        code_col = self.code_column_var.get()
        
        if not sender_email or not sender_pass:
            self.result_label.configure(text="Enter sender email and app password.", text_color="red")
            self.log("Sender email or app password missing.")
            return
        if not email_col or not name_col or not code_col:
            self.result_label.configure(text="Select email, name and code columns.", text_color="red")
            self.log("Email, name or code column not selected.")
            return
        
        # Validate email template
        subject = self.subject_entry.get().strip()
        body = self.body_textbox.get("1.0", "end-1c").strip()
        
        if not subject:
            self.result_label.configure(text="Email subject is required.", text_color="red")
            self.log("Email subject is empty!")
            return
            
        if not body:
            self.result_label.configure(text="Email body is required.", text_color="red")
            self.log("Email body is empty!")
            return

        # Start sending process
        self.is_sending = True
        self.send_btn.configure(text="Cancel", fg_color="red")
        
        # Show progress bar
        self.progress_frame.pack(pady=5, fill="x", padx=10)
        self.log(f"Starting to send attendance codes from {sender_email}...")
        
        # Start sending thread
        threading.Thread(
            target=self._send_codes_thread,
            args=(sender_email, sender_pass, email_col, name_col, code_col, subject, body),
            daemon=True
        ).start()

    def reset_send_button(self):
        """Reset the send button to its original state"""
        self.is_sending = False
        self.send_btn.configure(text="Send Attendance Codes", fg_color=["#1f538d", "#14375e"])
        self.progress_frame.pack_forget()  # Hide progress bar

    def _send_codes_thread(self, sender_email, sender_pass, email_col, name_col, code_col, subject, body):
        try:
            sent_count = 0
            failed = []
            skipped = 0
            selected_count = 0
            
            # First, count selected attendees with valid emails and codes
            for idx, row in self.attendees.iterrows():
                name_raw = row[name_col] if pd.notna(row[name_col]) else "Unknown"
                name = self.clean_name(str(name_raw).strip())
                email_raw = row[email_col] if pd.notna(row[email_col]) else ""
                recipient = str(email_raw).strip()
                key = f"{recipient}|{name}"
                
                if key in self.selected_attendees and self.selected_attendees[key].get():
                    if self.is_valid_email(recipient):
                        code_raw = row[code_col] if pd.notna(row[code_col]) else None
                        if self.format_code(code_raw) is not None:
                            selected_count += 1
            
            if selected_count == 0:
                self.after(0, self.log, "No attendees selected for sending codes.")
                self.after(0, self.finish_sending, 0, 0, [])
                return
            
            self.after(0, self.log, f"Starting to send {selected_count} selected attendance codes...")
            current_processed = 0
            
            for idx, row in self.attendees.iterrows():
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
                if key not in self.selected_attendees or not self.selected_attendees[key].get():
                    continue
                    
                # Skip if email is not valid
                if not self.is_valid_email(recipient):
                    self.after(0, self.log, f"[SKIPPED] Invalid email for {name}: {recipient}")
                    continue
                
                # Get and format code
                code_raw = row[code_col] if pd.notna(row[code_col]) else None
                formatted_code = self.format_code(code_raw)
                
                if formatted_code is None:
                    self.after(0, self.log, f"[SKIPPED] Invalid code for {name}: {code_raw}")
                    continue
                    
                current_processed += 1
                
                # Update progress in the main thread
                self.after(0, self.update_progress, current_processed, selected_count, f"Processing: {name} ({recipient})")
                
                # Send the code
                success, error = self.send_single_code(sender_email, sender_pass, name, recipient, formatted_code, subject, body)
                    
                if success:
                    sent_count += 1
                    self.mark_code_sent(recipient, name, formatted_code)
                    self.after(0, self.update_attendee_status, recipient, name)
                    self.after(0, self.log, f"[{recipient}] Attendance code sent successfully: {formatted_code}")
                else:
                    failed.append((recipient, error))
                    self.after(0, self.log, f"[{recipient}] Failed to send: {error}")

            # Update final results in the main thread
            self.after(0, self.finish_sending, sent_count, skipped, failed)
            
        finally:
            self.after(0, self.reset_send_button)
            self.after(0, self.update_status_list)

    def finish_sending(self, sent_count, skipped, failed):
        """Update UI after sending is complete"""
        result_msg = f"Sent: {sent_count} attendance codes."
        if skipped:
            result_msg += f"\nSkipped (already sent): {skipped}"
        if failed:
            result_msg += f"\nFailed: {len(failed)}"
        
        self.result_label.configure(text=result_msg, text_color="green" if sent_count else "red")
        self.log(result_msg)

if __name__ == "__main__":
    app = AttendanceCodeSenderApp()
    app.mainloop()