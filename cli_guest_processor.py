#!/usr/bin/env python3
"""
Command Line Interface for Guest List Processor
"""

import argparse
import sys
import os
from process_guest_list import GuestListProcessor

def main():
    parser = argparse.ArgumentParser(description='Process guest list and generate unique codes')
    parser.add_argument('excel_file', help='Path to Excel file with guest list')
    parser.add_argument('--name-column', default='Name', help='Name of the column containing guest names (default: Name)')
    parser.add_argument('--email-column', default='Email', help='Name of the column containing guest emails (default: Email)')
    parser.add_argument('--export-excel', help='Export results to Excel file')
    parser.add_argument('--export-json', help='Export results to JSON file')
    parser.add_argument('--send-emails', action='store_true', help='Send codes via email (requires email config)')
    parser.add_argument('--smtp-server', default='smtp.gmail.com', help='SMTP server (default: smtp.gmail.com)')
    parser.add_argument('--smtp-port', type=int, default=587, help='SMTP port (default: 587)')
    parser.add_argument('--email-user', help='Sender email address')
    parser.add_argument('--email-password', help='Sender email password')
    parser.add_argument('--stats', action='store_true', help='Show current statistics')
    parser.add_argument('--verify-code', help='Verify a specific code and show guest info')
    
    args = parser.parse_args()
    
    # Initialize processor
    processor = GuestListProcessor()
    
    # Show stats only
    if args.stats:
        stats = processor.get_stats()
        print("\n=== Guest List Statistics ===")
        print(f"Total Guests: {stats['total_guests']}")
        print(f"Unique Codes Generated: {stats['unique_codes_generated']}")
        print(f"Emails Sent: {stats['emails_sent']}")
        print(f"Emails Pending: {stats['emails_pending']}")
        return
    
    # Verify code only
    if args.verify_code:
        guest = processor.verify_code(args.verify_code)
        if guest:
            print(f"\n=== Code Verification ===")
            print(f"Code: {args.verify_code}")
            print(f"Name: {guest['name']}")
            print(f"Email: {guest['email']}")
            print(f"Email Sent: {guest.get('email_sent', False)}")
            print(f"Date Generated: {guest['date_generated']}")
        else:
            print(f"Code '{args.verify_code}' not found!")
        return
    
    # Check if Excel file exists
    if not os.path.exists(args.excel_file):
        print(f"Error: Excel file '{args.excel_file}' not found!")
        sys.exit(1)
    
    try:
        # Process Excel file
        print(f"Processing Excel file: {args.excel_file}")
        results = processor.process_excel_file(
            args.excel_file, 
            args.name_column, 
            args.email_column
        )
        
        print("\n=== Processing Results ===")
        print(f"Total Guests: {results['total_guests']}")
        print(f"New Codes Generated: {results['new_codes_generated']}")
        print(f"Existing Codes: {results['existing_codes']}")
        print(f"Invalid Entries: {results['invalid_entries']}")
        
        # Export to Excel if requested
        if args.export_excel:
            print(f"\nExporting to Excel: {args.export_excel}")
            processor.export_to_excel(args.export_excel)
            print("Excel export completed!")
        
        # Export to JSON if requested
        if args.export_json:
            print(f"\nExporting to JSON: {args.export_json}")
            processor.export_to_json(args.export_json)
            print("JSON export completed!")
        
        # Send emails if requested
        if args.send_emails:
            if not all([args.email_user, args.email_password]):
                print("Error: Email sending requires --email-user and --email-password")
                sys.exit(1)
            
            print("\nSending emails...")
            
            def progress_callback(sent, total, guest_name):
                print(f"Sending... {sent}/{total} - {guest_name}")
            
            email_results = processor.send_codes_by_email(
                args.smtp_server,
                args.smtp_port,
                args.email_user,
                args.email_password,
                progress_callback=progress_callback
            )
            
            print("\n=== Email Results ===")
            print(f"Emails Sent: {email_results['sent']}")
            print(f"Failed: {email_results['failed']}")
            print(f"Total Processed: {email_results['total_processed']}")
        
        # Show final stats
        stats = processor.get_stats()
        print("\n=== Final Statistics ===")
        print(f"Total Guests: {stats['total_guests']}")
        print(f"Emails Sent: {stats['emails_sent']}")
        print(f"Emails Pending: {stats['emails_pending']}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
