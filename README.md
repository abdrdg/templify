# Templify

Templify is a set of Windows desktop tools for managing event guest lists, generating personalized invitations, and sending event emails. The applications use CustomTkinter graphical interfaces and Excel files as the primary input format.

## Features

- Generate and persist unique six-character guest access codes.
- Process Excel guest lists and preserve existing codes when the same name and email are processed again.
- Export processed guest data to formatted Excel or JSON files.
- Generate personalized invitations from a DOCX template.
- Replace DOCX placeholders with values from Excel columns.
- Export invitations as DOCX, PDF, and PNG files.
- Send invitation emails or reminder emails in bulk through an SMTP server.
- Send three-digit attendance codes from an Excel file through an SMTP server.
- Select individual guests, all guests, unsent guests, or newly generated invitations.
- Track generated invitations, sent emails, and sent attendance codes to avoid accidental duplicates.
- Edit and save reusable email subject and body templates.
- View progress and status logs, with pagination for larger guest lists.

## Tools

### Guest List Processor

File: `process_guest_list.py`

Use this tool to:

1. Open an Excel guest list.
2. Specify the name and email columns.
3. Generate unique six-character codes for new guests.
4. Keep existing codes for guests already in the local data store.
5. Export the results to Excel or JSON.
6. Send codes by email using SMTP.
7. View guest statistics and mark guests as checked in through the underlying processor API.

The expected default Excel columns are `Name` and `Email`, although the column names can be changed in the interface.

### Invitation Generator

File: `invitation_generator.py`

Use this tool to create personalized invitations:

1. Select a DOCX template.
2. Select an Excel file.
3. Map template placeholders to Excel columns.
4. Choose an output folder.
5. Select invitees and generate their files.

Placeholders should use the `docxtpl` format, for example:

```text
{{ name }}
{{ organization }}
{{ invitation_code }}
```

The application detects placeholders in the template and creates a mapping control for each one. Each selected invitee can produce files named like:

```text
Invitation - Guest Name.docx
Invitation - Guest Name.pdf
Invitation - Guest Name.png
```

PDF and PNG generation requires Microsoft Word for DOCX-to-PDF conversion and Poppler for PDF-to-image conversion. On Windows, the application attempts to download Poppler automatically when it is not found.

### Invitation Sender

File: `invitation_sender.py`

Use this tool to email invitations or reminders:

1. Select the folder containing generated invitation files.
2. Open the Excel guest list.
3. Map the name and email columns.
4. Enter SMTP credentials.
5. Choose `Invitation Email` or `Reminder Email`.
6. Select recipients and send.

Invitation and reminder emails embed or attach the matching PNG invitation from the selected images/files folder. Reminder messages support HTML content such as links, bold text, paragraphs, and line breaks.

### Attendance Code Sender

File: `attendance_code_sender.py`

Use this tool to send three-digit attendance codes from an Excel file:

1. Open the Excel file.
2. Map the email, name, and code columns.
3. Enter the sender email and app password.
4. Edit or reset the subject and email body if needed.
5. Select recipients and send.

Codes are converted to three-digit values, so a value such as `7` is sent as `007`. The email body supports `[Name]` and `[Code]` placeholders.

## Requirements

- Windows is recommended because the batch launchers and DOCX/PDF workflow target Windows.
- Python 3.7 or newer.
- Microsoft Word is required for PDF conversion in the Invitation Generator.
- An SMTP account that permits authenticated SMTP and, where required, an app password.
- An Excel workbook containing the relevant guest data.
- A DOCX template for invitation generation.

Python packages are listed in `requirements.txt`, including CustomTkinter, pandas, openpyxl, docxtpl, docx2pdf, and pdf2image.

## Installation and Running

The easiest option on Windows is to double-click one of the batch files. Each launcher creates or reuses a local virtual environment in `venv`, installs or updates the packages in `requirements.txt`, and starts the selected application.

- `run_invitation_generator.bat` - start the Invitation Generator.
- `run_invitation_sender.bat` - start the Invitation Sender.
- `run_attendance_code_sender.bat` - start the Attendance Code Sender.

The Guest List Processor does not currently have a batch launcher. Run it from PowerShell in the project folder:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python process_guest_list.py
```

If PowerShell blocks script activation, run the program with the virtual-environment interpreter directly:

```powershell
.\venv\Scripts\python.exe process_guest_list.py
```

The batch files can also be started from PowerShell:

```powershell
.\run_invitation_generator.bat
.\run_invitation_sender.bat
.\run_attendance_code_sender.bat
```

## Typical Workflow

1. Run the Guest List Processor and process the original Excel guest list.
2. Export the processed guest list, including generated codes.
3. Use the Invitation Generator with a DOCX template and the guest list to create invitation files.
4. Use the Invitation Sender to email invitations or reminders with the generated files.
5. Use the Attendance Code Sender to email three-digit codes when required.
6. Keep the generated JSON tracking files with the applications so previous activity remains available.

## Local Data and Tracking Files

The applications create local files in the directory from which they are run:

- `guest_codes.json` - persistent guest and code data.
- `guest_tracking.json` - guest email tracking data used by the Guest List Processor.
- `generated_invitations.json` - invitation generation history.
- `sent_invitations.json` - invitation and reminder send history.
- `sent_attendance_codes.json` - attendance-code send history.
- `email_templates.json` - saved invitation, reminder, and attendance email templates.
- `output\` - default output folder for generated invitations.
- `poppler\` - downloaded Poppler files when automatic installation is used.

These files may contain names, email addresses, event information, and message history. Back them up securely and do not commit them to a public repository.

## SMTP Notes

- Use an app password instead of your normal mailbox password when the provider requires it.
- The applications use STARTTLS and default to port `587` in the Guest List Processor.
- Confirm the SMTP server, port, account permissions, and provider sending limits before sending to a large list.
- Test with one recipient first.
- Credentials are entered into the GUI and are not intended to be stored in the project files.

## Troubleshooting

- **Python is not found:** Install Python and enable `Add Python to PATH`, then rerun the batch file.
- **Package import errors:** Activate the project virtual environment and run `python -m pip install -r requirements.txt`.
- **PDF or PNG files are not generated:** Confirm Microsoft Word is installed and that Poppler is available. The Invitation Generator logs Poppler setup progress.
- **Recipients are not selectable:** Check that the selected email column contains valid email addresses.
- **Previously processed guests are skipped:** Review the relevant tracking JSON file or use the application's selection controls to choose recipients explicitly.
- **Email sending fails:** Verify SMTP settings and use an app password if required by the email provider.
