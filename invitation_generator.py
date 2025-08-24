
# Standard library imports
import sys
import os
import urllib.request
import zipfile
import re
import threading

# Third-party imports
from pdf2image import convert_from_path
from docx2pdf import convert as docx2pdf_convert
from docxtpl import DocxTemplate
import openpyxl

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
		self.geometry("700x600")
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

		# UI Elements
		self.create_widgets()

	def create_widgets(self):
		# Template selection
		ctk.CTkLabel(self, text="1. Select DOCX Template:").pack(anchor="w", padx=20, pady=(20, 0))
		frame1 = ctk.CTkFrame(self)
		frame1.pack(fill="x", padx=20)
		ctk.CTkEntry(frame1, textvariable=self.template_path, width=400, state="readonly").pack(side="left", padx=(0,10))
		ctk.CTkButton(frame1, text="Browse", command=self.select_template).pack(side="left")

		# Excel selection
		ctk.CTkLabel(self, text="2. Select Excel File:").pack(anchor="w", padx=20, pady=(20, 0))
		frame2 = ctk.CTkFrame(self)
		frame2.pack(fill="x", padx=20)
		ctk.CTkEntry(frame2, textvariable=self.excel_path, width=400, state="readonly").pack(side="left", padx=(0,10))
		ctk.CTkButton(frame2, text="Browse", command=self.select_excel).pack(side="left")

		# Mapping area
		self.mapping_frame = ctk.CTkFrame(self)
		self.mapping_frame.pack(fill="x", padx=20, pady=(20,0))
		ctk.CTkLabel(self.mapping_frame, text="3. Map Excel Columns to Template Placeholders:").pack(anchor="w")
		self.mapping_dropdowns_frame = ctk.CTkFrame(self.mapping_frame)
		self.mapping_dropdowns_frame.pack(fill="x", pady=(5,0))

		# Output folder
		ctk.CTkLabel(self, text="4. Output Folder:").pack(anchor="w", padx=20, pady=(20, 0))
		frame3 = ctk.CTkFrame(self)
		frame3.pack(fill="x", padx=20)
		ctk.CTkEntry(frame3, textvariable=self.output_folder, width=400, state="readonly").pack(side="left", padx=(0,10))
		ctk.CTkButton(frame3, text="Change", command=self.select_output_folder).pack(side="left")

		# Generate button
		ctk.CTkButton(self, text="Generate Invitations", command=self.generate_invitations, width=200).pack(pady=(30,10))

		# Progress bar
		self.progress = ctk.CTkProgressBar(self)
		self.progress.pack(fill="x", padx=20, pady=(0,10))
		self.progress.set(0)

		# Status log
		ctk.CTkLabel(self, text="Status Log:").pack(anchor="w", padx=20)
		self.log_text = ctk.CTkTextbox(self, height=120, state="disabled")
		self.log_text.pack(fill="both", expand=True, padx=20, pady=(0,20))

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
			self.update_mapping_dropdowns()

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

	def log(self, message):
		# Schedule the log update on the main thread
		self.after(0, self._log_update, message)
	
	def _log_update(self, message):
		self.log_text.configure(state="normal")
		self.log_text.insert("end", message + "\n")
		self.log_text.see("end")
		self.log_text.configure(state="disabled")

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

		# Read Excel rows
		wb = openpyxl.load_workbook(excel_path)
		ws = wb.active
		rows = list(ws.iter_rows(min_row=2, values_only=True))
		columns = [str(cell.value) for cell in ws[1]]
		total = len(rows)
		if total == 0:
			self.log("No data rows found in Excel.")
			return

		# Prepare output folder
		os.makedirs(output_folder, exist_ok=True)

		# Ensure Poppler is available for pdf2image
		poppler_path = None
		if sys.platform == "win32":
			poppler_path = ensure_poppler()

		# Generate invitations
		for idx, row in enumerate(rows, 1):
			data = dict(zip(columns, row))
			attendee = Attendee(data)
			context = attendee.get_context(mapping)
			try:
				doc = DocxTemplate(template_path)
				doc.render(context)
				filename = attendee.get_filename()
				out_docx = os.path.join(output_folder, f"Invitation - {filename}.docx")
				out_pdf = os.path.join(output_folder, f"Invitation - {filename}.pdf")
				out_png = os.path.join(output_folder, f"Invitation - {filename}.png")
				doc.save(out_docx)
				self.log(f"Saved: {out_docx}")
				# Convert DOCX to PDF
				try:
					docx2pdf_convert(out_docx, output_folder)
					self.log(f"PDF created: {out_pdf}")
				except Exception as e:
					self.log(f"PDF conversion failed: {e}")
					out_pdf = None
				# Convert PDF to PNG (first page)
				if out_pdf and os.path.exists(out_pdf):
					try:
						images = convert_from_path(out_pdf, dpi=200, fmt='png', poppler_path=poppler_path)
						if images:
							images[0].save(out_png, 'PNG')
							self.log(f"PNG created: {out_png}")
					except Exception as e:
						self.log(f"PNG conversion failed: {e}")
			except Exception as e:
				self.log(f"Error for row {idx}: {e}")
			# Update progress on main thread
			self.after(0, self.progress.set, idx / total)

		self.log("Generation complete.")

if __name__ == "__main__":
	app = InvitationGeneratorApp()
	app.mainloop()
