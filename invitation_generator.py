
# Standard library imports
import sys
import os
import urllib.request
import zipfile
import re
import threading
import json
from datetime import datetime

# Third-party imports
from pdf2image import convert_from_path
from docx2pdf import convert as docx2pdf_convert
from docxtpl import DocxTemplate
import openpyxl
import pandas as pd

# Attendee class for OOP
class Attendee:
	def __init__(self, data_dict):
		self.data = data_dict

	def get_context(self, mapping):
		# mapping: {placeholder: excel_column}
		context = {}
		for ph, col in mapping.items():
			value = self.data.get(col, "")
			# Convert None to empty string and strip whitespace
			if value is None or str(value).strip() == "":
				context[ph] = ""
			else:
				context[ph] = str(value).strip()
		return context

	def get_filename(self):
		# Use Name or fallback to first column
		name = self.data.get("Name") or list(self.data.values())[0]
		return str(name).replace("\n", " ").replace(".", "").replace('"', "'")


# Modern GUI for invitation generation
import customtkinter as ctk
from tkinter import filedialog, messagebox




# Ensure Poppler is available for pdf2image (Windows only)
def ensure_poppler():
	"""
	Download and extract Poppler for Windows if not already present.
	Returns the bin path containing pdftoppm.exe for pdf2image.
	"""
	if sys.platform != "win32":
		return None
	# Explicitly check the expected path
	poppler_bin = os.path.join(os.path.dirname(__file__), "poppler", "poppler-23.11.0", "Library", "bin")
	pdftoppm_path = os.path.join(poppler_bin, "pdftoppm.exe")
	if os.path.exists(pdftoppm_path):
		return poppler_bin
	# Fallback: search all subfolders for pdftoppm.exe
	poppler_dir = os.path.join(os.path.dirname(__file__), "poppler")
	for root, dirs, files in os.walk(poppler_dir):
		if "pdftoppm.exe" in files:
			return root
	# Download and extract Poppler if not found
	url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v23.11.0-0/Release-23.11.0-0.zip"
	zip_path = os.path.join(poppler_dir, "poppler.zip")
	os.makedirs(poppler_dir, exist_ok=True)
	print("Downloading Poppler...")
	urllib.request.urlretrieve(url, zip_path)
	print("Extracting Poppler...")
	with zipfile.ZipFile(zip_path, 'r') as zip_ref:
		zip_ref.extractall(poppler_dir)
	os.remove(zip_path)
	# Find the extracted folder
	for root, dirs, files in os.walk(poppler_dir):
		if "pdftoppm.exe" in files:
			return root
	return None

class InvitationGeneratorApp(ctk.CTk):

	def __init__(self):
		super().__init__()
		self.title("Invitation Generator")
		self.geometry("1000x700")  # Made wider for two-column layout
		ctk.set_appearance_mode("System")
		ctk.set_default_color_theme("blue")

		# File paths
		self.template_path = ctk.StringVar()
		self.excel_path = ctk.StringVar()
		self.output_folder = ctk.StringVar(value=os.path.abspath("output"))

		# Placeholders and column mapping
		self.placeholders = []
		self.excel_columns = []
		self.mapping_vars = {}

		# Initialize generation tracking
		self.tracking_file = "generated_invitations.json"
		self.generated_invitations = self.load_generated_invitations()

		# Initialize selection tracking for invitees
		self.selected_invitees = {}  # Dictionary to track checkbox states
		self.invitees = None  # Will hold the DataFrame of invitees

		# UI Elements
		self.create_widgets()

	def create_widgets(self):
		# Title at the top
		ctk.CTkLabel(self, text="Invitation Generator", font=("Arial", 22)).pack(pady=(10, 0))
		
		# Create two-column layout using a horizontal frame
		columns_frame = ctk.CTkFrame(self, fg_color="transparent")
		columns_frame.pack(fill="both", expand=True, padx=10, pady=(10, 10))
		
		# Left column - Controls
		left_column = ctk.CTkFrame(columns_frame)
		left_column.pack(side="left", fill="both", expand=True, padx=(0, 5))
		
		# Right column - Invitees list
		right_column = ctk.CTkFrame(columns_frame)
		right_column.pack(side="right", fill="both", expand=True, padx=(5, 0))

		# === LEFT COLUMN CONTENT ===
		
		# Template selection
		template_frame = ctk.CTkFrame(left_column)
		template_frame.pack(pady=5, fill="x", padx=10)
		ctk.CTkLabel(template_frame, text="1. Select DOCX Template:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
		template_input_frame = ctk.CTkFrame(template_frame, fg_color="transparent")
		template_input_frame.pack(fill="x", padx=5, pady=(0, 5))
		ctk.CTkEntry(template_input_frame, textvariable=self.template_path, width=300, state="readonly").pack(side="left", padx=(0,5), fill="x", expand=True)
		ctk.CTkButton(template_input_frame, text="Browse", command=self.select_template, width=80).pack(side="right")

		# Excel selection
		excel_frame = ctk.CTkFrame(left_column)
		excel_frame.pack(pady=5, fill="x", padx=10)
		ctk.CTkLabel(excel_frame, text="2. Select Excel File:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
		excel_input_frame = ctk.CTkFrame(excel_frame, fg_color="transparent")
		excel_input_frame.pack(fill="x", padx=5, pady=(0, 5))
		ctk.CTkEntry(excel_input_frame, textvariable=self.excel_path, width=300, state="readonly").pack(side="left", padx=(0,5), fill="x", expand=True)
		ctk.CTkButton(excel_input_frame, text="Browse", command=self.select_excel, width=80).pack(side="right")

		# Mapping area
		mapping_section = ctk.CTkFrame(left_column)
		mapping_section.pack(pady=5, fill="x", padx=10)
		ctk.CTkLabel(mapping_section, text="3. Map Columns to Placeholders:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
		self.mapping_dropdowns_frame = ctk.CTkFrame(mapping_section)
		self.mapping_dropdowns_frame.pack(fill="x", padx=5, pady=(0, 5))

		# Output folder
		output_frame = ctk.CTkFrame(left_column)
		output_frame.pack(pady=5, fill="x", padx=10)
		ctk.CTkLabel(output_frame, text="4. Output Folder:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
		output_input_frame = ctk.CTkFrame(output_frame, fg_color="transparent")
		output_input_frame.pack(fill="x", padx=5, pady=(0, 5))
		ctk.CTkEntry(output_input_frame, textvariable=self.output_folder, width=300, state="readonly").pack(side="left", padx=(0,5), fill="x", expand=True)
		ctk.CTkButton(output_input_frame, text="Change", command=self.select_output_folder, width=80).pack(side="right")

		# Progress bar
		progress_frame = ctk.CTkFrame(left_column)
		progress_frame.pack(pady=5, fill="x", padx=10)
		ctk.CTkLabel(progress_frame, text="Progress:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
		self.progress = ctk.CTkProgressBar(progress_frame)
		self.progress.pack(fill="x", padx=5, pady=(0, 5))
		self.progress.set(0)

		# Status log
		log_frame = ctk.CTkFrame(left_column)
		log_frame.pack(fill="both", expand=True, padx=10, pady=5)
		ctk.CTkLabel(log_frame, text="Status Log:", font=("Arial", 12, "bold")).pack(anchor="w", padx=5)
		self.log_text = ctk.CTkTextbox(log_frame, state="disabled")
		self.log_text.pack(fill="both", expand=True, padx=5, pady=(0, 5))

		# === RIGHT COLUMN CONTENT ===
		
		# Invitees list with status
		ctk.CTkLabel(right_column, text="Invitees Selection", font=("Arial", 16, "bold")).pack(pady=(10, 5))
		
		# Header with refresh button
		header_frame = ctk.CTkFrame(right_column)
		header_frame.pack(fill="x", padx=10, pady=(0, 5))
		self.refresh_btn = ctk.CTkButton(header_frame, text="Refresh", width=80, command=self.update_invitees_list)
		self.refresh_btn.pack(side="right", padx=5)
		
		# Selection buttons frame
		selection_frame = ctk.CTkFrame(right_column)
		selection_frame.pack(fill="x", padx=10, pady=(0, 5))
		self.select_all_btn = ctk.CTkButton(selection_frame, text="Select All", width=80, command=self.select_all_invitees)
		self.select_all_btn.pack(side="left", padx=2)
		self.select_none_btn = ctk.CTkButton(selection_frame, text="Select None", width=80, command=self.select_none_invitees)
		self.select_none_btn.pack(side="left", padx=2)
		self.select_ungenerated_btn = ctk.CTkButton(selection_frame, text="Select New", width=90, command=self.select_ungenerated_invitees)
		self.select_ungenerated_btn.pack(side="left", padx=2)
		
		# Scrollable frame for invitees
		self.invitees_scrollable_frame = ctk.CTkScrollableFrame(right_column)
		self.invitees_scrollable_frame.pack(fill="both", expand=True, padx=10, pady=(0, 5))
		self.invitee_labels = {}  # Store labels for updating
		
		# Generate button at bottom of right column
		generate_frame = ctk.CTkFrame(right_column)
		generate_frame.pack(fill="x", padx=10, pady=(0, 10))
		
		self.generate_btn = ctk.CTkButton(
			generate_frame, 
			text="Generate Invitations", 
			command=self.generate_invitations,
			height=40,
			font=("Arial", 14)
		)
		self.generate_btn.pack(pady=5, fill="x", padx=10)

	def select_template(self):
		path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
		if path:
			self.template_path.set(path)
			self.log(f"Selected template: {path}")
			self.placeholders = self.extract_placeholders(path)
			self.update_mapping_dropdowns()

	def extract_placeholders(self, docx_path):
		"""
		Extract placeholders from all XML files in the .docx archive, including those split across runs.
		Returns a list of unique placeholder names.
		"""
		import zipfile
		import xml.etree.ElementTree as ET
		found = set()
		with zipfile.ZipFile(docx_path) as docx_zip:
			for file in docx_zip.namelist():
				if file.endswith('.xml'):
					with docx_zip.open(file) as xml_file:
						try:
							xml = xml_file.read().decode('utf-8')
						except Exception:
							continue
						# Join all <w:t> text nodes for robust placeholder extraction
						try:
							root = ET.fromstring(xml)
							texts = []
							for elem in root.iter():
								# Word text node
								if elem.tag.endswith('}t'):
									texts.append(elem.text or '')
							joined_text = ''.join(texts)
							found.update(re.findall(r'{{\s*(\w+)\s*}}', joined_text))
						except Exception:
							# Fallback: regex on raw xml
							found.update(re.findall(r'{{\s*(\w+)\s*}}', xml))
		self.log(f"Found placeholders: {', '.join(found)}")
		return list(found)

	def select_excel(self):
		path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
		if path:
			self.excel_path.set(path)
			self.log(f"Selected Excel: {path}")
			self.excel_columns = self.extract_excel_columns(path)
			self.load_invitees(path)
			self.update_mapping_dropdowns()
			self.update_invitees_list()

	def load_invitees(self, excel_path):
		"""Load invitees from Excel file"""
		try:
			import pandas as pd
			self.invitees = pd.read_excel(excel_path)
			self.log(f"Loaded {len(self.invitees)} invitees from Excel.")
		except Exception as e:
			self.log(f"Error loading invitees: {e}")
			self.invitees = None

	def extract_excel_columns(self, excel_path):
		wb = openpyxl.load_workbook(excel_path)
		ws = wb.active
		columns = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
		self.log(f"Excel columns: {', '.join([str(c) for c in columns if c])}")
		return [str(c) for c in columns if c]

	def update_mapping_dropdowns(self):
		# Clear previous
		for widget in self.mapping_dropdowns_frame.winfo_children():
			widget.destroy()
		self.mapping_vars = {}
		if not self.placeholders or not self.excel_columns:
			return
		for ph in self.placeholders:
			frame = ctk.CTkFrame(self.mapping_dropdowns_frame)
			frame.pack(fill="x", pady=2)
			ctk.CTkLabel(frame, text=f"{ph}:", width=120).pack(side="left")
			var = ctk.StringVar()
			dropdown = ctk.CTkOptionMenu(frame, variable=var, values=self.excel_columns)
			dropdown.pack(side="left", padx=5)
			self.mapping_vars[ph] = var

	def select_output_folder(self):
		path = filedialog.askdirectory()
		if path:
			self.output_folder.set(path)
			self.log(f"Output folder set to: {path}")

	def update_invitees_list(self):
		"""Update the invitees list with checkboxes"""
		self.clear_invitees_list()
		if self.invitees is None or self.invitees.empty:
			return

		for idx, row in self.invitees.iterrows():
			# Convert row to dict format that Attendee expects
			data = row.to_dict()
			data = {str(k): v for k, v in data.items()}
			
			# Create Attendee object to get consistent filename
			attendee = Attendee(data)
			filename = attendee.get_filename()
			
			# Also get a display name (original, unprocessed)
			name_columns = ['Name', 'name', 'Full Name', 'full_name']
			display_name = None
			for col in name_columns:
				if col in row and pd.notna(row[col]):
					display_name = str(row[col]).strip()
					break
			
			if not display_name:
				# Fallback to first non-null column
				for col_name, value in row.items():
					if pd.notna(value) and str(value).strip():
						display_name = str(value).strip()
						break
				if not display_name:
					display_name = f"Row {idx + 1}"
			
			# Use the processed filename for the key (consistent with tracking)
			key = f"{idx}|{filename}"
			
			# Create frame for this invitee
			frame = ctk.CTkFrame(self.invitees_scrollable_frame)
			frame.pack(fill="x", padx=2, pady=1)
			
			# Checkbox for selection
			checkbox_var = ctk.BooleanVar()
			checkbox = ctk.CTkCheckBox(frame, text="", variable=checkbox_var, width=20)
			checkbox.pack(side="left", padx=5)
			
			# Store checkbox variable for later use
			self.selected_invitees[key] = checkbox_var
			
			# Name display (show the original display name)
			info_label = ctk.CTkLabel(frame, text=display_name, anchor="w")
			info_label.pack(side="left", padx=5, fill="x", expand=True)
			
			# Status label
			status_label = ctk.CTkLabel(frame, text="", anchor="e", width=120)
			status_label.pack(side="right", padx=5)
			
			# Store label reference for updates
			self.invitee_labels[key] = status_label
			
			# Update status using the processed filename (consistent with tracking)
			is_generated = self.was_invitation_generated(filename)
			if is_generated:
				status_label.configure(text="Generated", text_color="green")
				checkbox_var.set(False)  # Don't select already generated
			else:
				status_label.configure(text="Not generated", text_color="gray")
				checkbox_var.set(True)  # Select ungenerated by default

	def clear_invitees_list(self):
		"""Clear all invitee widgets and selection tracking"""
		for widget in self.invitees_scrollable_frame.winfo_children():
			widget.destroy()
		self.invitee_labels = {}
		self.selected_invitees = {}

	def select_all_invitees(self):
		"""Select all invitees for generation"""
		count = 0
		for key, checkbox_var in self.selected_invitees.items():
			checkbox_var.set(True)
			count += 1
		self.log(f"Selected all {count} invitees.")

	def select_none_invitees(self):
		"""Deselect all invitees"""
		for key in self.selected_invitees:
			self.selected_invitees[key].set(False)
		self.log("All invitees deselected.")

	def select_ungenerated_invitees(self):
		"""Select only invitees who haven't been generated yet"""
		count = 0
		for key, checkbox_var in self.selected_invitees.items():
			_, filename = key.split("|", 1)  # Now using processed filename
			if not self.was_invitation_generated(filename):
				checkbox_var.set(True)
				count += 1
			else:
				checkbox_var.set(False)
		self.log(f"Selected {count} ungenerated invitees.")

	def log(self, message):
		# Schedule the log update on the main thread
		self.after(0, self._log_update, message)
	
	def _log_update(self, message):
		self.log_text.configure(state="normal")
		self.log_text.insert("end", message + "\n")
		self.log_text.see("end")
		self.log_text.configure(state="disabled")

	def load_generated_invitations(self):
		"""Load the record of generated invitations from JSON file"""
		if os.path.exists(self.tracking_file):
			try:
				with open(self.tracking_file, 'r') as f:
					return json.load(f)
			except json.JSONDecodeError:
				self.log("Warning: Generation tracking file corrupted, starting fresh.")
				return {}
		return {}

	def save_generated_invitations(self):
		"""Save the record of generated invitations to JSON file"""
		try:
			with open(self.tracking_file, 'w') as f:
				json.dump(self.generated_invitations, f, indent=2)
		except Exception as e:
			self.log(f"Warning: Could not save generation tracking: {e}")

	def was_invitation_generated(self, name):
		"""Check if invitation was already generated for this person"""
		return name in self.generated_invitations

	def mark_invitation_generated(self, name, output_folder):
		"""Mark invitation as generated for this person"""
		self.generated_invitations[name] = {
			"generated_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
			"output_folder": output_folder
		}
		self.save_generated_invitations()

	def generate_invitations(self):
		# Start generation in a separate thread to keep UI responsive
		thread = threading.Thread(target=self._generate_invitations_thread, daemon=True)
		thread.start()

	def _generate_invitations_thread(self):
		self.log("Generation started...")
		template_path = self.template_path.get()
		excel_path = self.excel_path.get()
		output_folder = self.output_folder.get()
		if not template_path or not excel_path or not output_folder:
			self.log("Please select a template, Excel file, and output folder.")
			return

		# Build mapping from dropdowns
		mapping = {ph: var.get() for ph, var in self.mapping_vars.items()}
		if not all(mapping.values()):
			self.log("Please map all placeholders to Excel columns.")
			return

		# Check if we have invitees loaded and selected
		if self.invitees is None or self.invitees.empty:
			self.log("No invitees data loaded.")
			return

		# Count selected invitees
		selected_count = 0
		selected_indices = []
		for key, checkbox_var in self.selected_invitees.items():
			if checkbox_var.get():
				idx, name = key.split("|", 1)
				selected_indices.append(int(idx))
				selected_count += 1

		if selected_count == 0:
			self.log("No invitees selected for generation.")
			return

		self.log(f"Starting to generate {selected_count} selected invitations...")

		# Prepare output folder
		os.makedirs(output_folder, exist_ok=True)

		# Ensure Poppler is available for pdf2image
		poppler_path = None
		if sys.platform == "win32":
			poppler_path = ensure_poppler()

		# Generate invitations for selected invitees only
		generated_count = 0
		current_processed = 0

		for idx in selected_indices:
			current_processed += 1
			row = self.invitees.iloc[idx]
			data = row.to_dict()
			
			# Convert data keys to strings for consistency
			data = {str(k): v for k, v in data.items()}
			
			attendee = Attendee(data)
			context = attendee.get_context(mapping)
			filename = attendee.get_filename()
			
			try:
				doc = DocxTemplate(template_path)
				doc.render(context)
				out_docx = os.path.join(output_folder, f"Invitation - {filename}.docx")
				out_pdf = os.path.join(output_folder, f"Invitation - {filename}.pdf")
				out_png = os.path.join(output_folder, f"Invitation - {filename}.png")
				doc.save(out_docx)
				self.log(f"Saved: {out_docx}")
				
				# Convert DOCX to PDF
				pdf_success = False
				try:
					docx2pdf_convert(out_docx, output_folder)
					self.log(f"PDF created: {out_pdf}")
					pdf_success = True
				except Exception as e:
					self.log(f"PDF conversion failed: {e}")
					out_pdf = None
				
				# Convert PDF to PNG (first page)
				png_success = False
				if out_pdf and os.path.exists(out_pdf):
					try:
						images = convert_from_path(out_pdf, dpi=200, fmt='png', poppler_path=poppler_path)
						if images:
							images[0].save(out_png, 'PNG')
							self.log(f"PNG created: {out_png}")
							png_success = True
					except Exception as e:
						self.log(f"PNG conversion failed: {e}")
				
				# Mark as generated only if at least the DOCX was created successfully
				self.mark_invitation_generated(filename, output_folder)
				generated_count += 1
				
				# Update the status in the list
				key = f"{idx}|{filename}"
				if key in self.invitee_labels:
					self.after(0, self.update_invitee_status, key, True)
				
			except Exception as e:
				self.log(f"Error for {filename}: {e}")
			
			# Update progress on main thread
			self.after(0, self.progress.set, current_processed / selected_count)

		self.log(f"Generation complete. Generated: {generated_count} invitations")
		
		# Refresh the invitees list to show updated statuses
		self.after(0, self.update_invitees_list)

	def update_invitee_status(self, key, is_generated):
		"""Update the status display for a single invitee"""
		if key in self.invitee_labels:
			label = self.invitee_labels[key]
			if is_generated:
				label.configure(text="Generated", text_color="green")
			else:
				label.configure(text="Not generated", text_color="gray")

if __name__ == "__main__":
	app = InvitationGeneratorApp()
	app.mainloop()
