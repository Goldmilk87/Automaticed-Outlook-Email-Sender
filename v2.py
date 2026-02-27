import pandas as pd
import win32com.client as win32
from tkinter import Tk, filedialog
import mammoth
import re
import os
import chardet
import argparse


def clean_word_empty_paragraphs(val: str) -> str:
    # Remove paragraphs that only contain &nbsp; inside <o:p>
    pattern = r'<p[^>]*>\s*<span[^>]*>\s*<o:p>\s*&nbsp;\s*</o:p>\s*</span>\s*</p>'
    cleaned = re.sub(pattern, '', val, flags=re.IGNORECASE)
    return cleaned


def select_files():
    Tk().withdraw()  # hide the root window

    # Select Excel file
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not excel_file:
        print("No Excel file selected. Exiting.")
        exit()

    # Select Word file
    template_file = filedialog.askopenfilename(
        title="Select Template File (Word or HTML)",
        filetypes=[("Word files", "*.docx"), ("HTML files", "*.htm *.html")]
    )
    if not template_file:
        print("No template file selected. Exiting.")
        exit()

    # Read Excel data
    df = pd.read_excel(excel_file)

    ext = os.path.splitext(template_file)[1].lower()
    if ext == ".docx":
        with open(template_file, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            template_content = result.value  # HTML content as string
    elif ext in [".htm", ".html"]:
        with open(template_file, "rb") as f:
            raw_data = f.read()

        # Detect encoding
        detected = chardet.detect(raw_data)
        encoding = detected['encoding'] if detected['encoding'] else 'utf-8'

        # Decode
        template_content = raw_data.decode(encoding, errors='replace')
    else:
        print("Unsupported template file type. Exiting.")
        exit()

    attach_files = filedialog.askopenfilenames(
        title="Select Attachment Files (Optional)",
        filetypes=[("All files", "*.*")]
    )

    # If nothing selected, set to None
    if not attach_files:
        attach_files = None

    return df, template_content, attach_files


def select_outlook_account():
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    accounts = [acc.DisplayName for acc in namespace.Accounts]

    print("Select the email account to send from:")
    for i, acc in enumerate(accounts, 1):
        print(f"{i}: {acc}")

    while True:
        try:
            choice = int(input("Enter the number of the account: "))
            if 1 <= choice <= len(accounts):
                selected_account = accounts[choice - 1]
                break
            else:
                print(f"Please enter a number between 1 and {len(accounts)}")
        except ValueError:
            print("Invalid input. Enter a number.")

    # Return the Outlook account object
    for acc in namespace.Accounts:
        if acc.DisplayName == selected_account:
            return acc, outlook

    return None, outlook


def evaluate_condition(row, condition_str):
    """
    Evaluate a condition string against a row of data.
    Examples:
    - "Round == 2"
    - "Status == 'Active'"
    - "Round == 2 and Department == 'CS'"
    """
    if not condition_str or condition_str.lower() == 'all':
        return True

    try:
        # Create a safe evaluation environment with row data
        safe_dict = {key: value for key, value in row.items()}
        # Add common operators
        safe_dict.update({
            '__builtins__': {},
            'str': str,
            'int': int,
            'float': float,
            'bool': bool
        })

        # Evaluate the condition
        result = eval(condition_str, safe_dict)
        return bool(result)
    except Exception as e:
        print(f"Error evaluating condition '{condition_str}': {e}")
        print(f"Available columns: {list(row.keys())}")
        return False


def send_emails(df, html_template, account_to_use, outlook, attach_files, condition_str, cc_list, subject_template):
    sent_count = 0
    total_rows = len(df)

    for index, row in df.iterrows():
        # Check if row meets the condition
        if not evaluate_condition(row, condition_str):
            continue

        html_body = html_template
        # Replace placeholders in HTML body
        for key, value in row.items():
            placeholder = f"{{{{{key}}}}}"  # {{ColumnName}} in HTML
            html_body = html_body.replace(placeholder, str(value))
        html_body = f"""
        <div style="font-family:Calibri, sans-serif; font-size:11pt;">
        {html_body}
        </div>
        """

        # Prepare CC list
        final_cc_list = cc_list.copy()
        # Add CC from row data if exists and not empty
        row_cc = row.get("CC", "")
        if row_cc and str(row_cc).strip():
            final_cc_list.append(str(row_cc).strip())

        # Prepare Subject
        final_subject = subject_template
        # Replace placeholders in subject
        for key, value in row.items():
            placeholder = f"{{{{{key}}}}}"
            final_subject = final_subject.replace(placeholder, str(value))

        mail = outlook.CreateItem(0)
        # Send from selected account
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account_to_use))
        mail.Display()  # use mail.Send() to send automatically
        signature = mail.HTMLBody

        # Prepend your custom content before the signature
        mail.HTMLBody = html_body + clean_word_empty_paragraphs(signature)
        mail.To = row.get("Email", "")  # assumes column "Email" exists

        # Set CC only if there are addresses
        if final_cc_list and any(cc.strip() for cc in final_cc_list):
            mail.cc = ";".join([cc for cc in final_cc_list if cc.strip()])

        if attach_files:
            for file in attach_files:
                mail.Attachments.Add(file)
        mail.Subject = final_subject

        sent_count += 1
        print(f"Processed email {sent_count} for {row.get('Email', 'Unknown')} with subject: '{final_subject}'")

    print(f"All emails processed. {sent_count} emails sent out of {total_rows} total rows.")


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Send bulk emails using Excel data and template",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python script.py --condition "Round == 2" --subject "Meeting Reminder"
  python script.py -c "Status == 'Active'" --cc "admin@company.com" -s "Urgent Update"
  python script.py --condition "all" --subject "Hello {{Name}}" # Subject can use placeholders from Excel
  python script.py  # Interactive mode (will prompt for condition, CC, and Subject)
        """
    )

    parser.add_argument(
        '--condition', '-c',
        type=str,
        help='Condition to filter rows for sending emails (e.g., "Round == 2", "Status == \'Active\'", "all" for all rows)'
    )

    parser.add_argument(
        '--cc',
        type=str,
        help='Comma-separated list of CC email addresses (e.g., "admin@company.com,hr@company.com")'
    )

    parser.add_argument(
        '--subject', '-s',
        type=str,
        help='Subject of the email. Can include placeholders like "{{Name}}" from Excel columns.'
    )

    return parser.parse_args()


def get_user_inputs():
    """Get condition, CC list, and subject from user input if not provided via command line"""
    print("\nEmail Sending Configuration:")
    print("=" * 50)

    # Get condition
    print("\nEnter the condition for sending emails:")
    print("Examples:")
    print("  - Round == 2")
    print("  - Status == 'Active'")
    print("  - Round == 2 and Department == 'CS'")
    print("  - all (to send to all rows)")
    condition = input("Condition (or 'all' for all rows): ").strip()

    if not condition:
        condition = "all"

    # Get CC list
    print("\nEnter CC email addresses (comma-separated, or press Enter for none):")
    cc_input = input("CC addresses: ").strip()

    if cc_input:
        cc_list = [email.strip() for email in cc_input.split(",") if email.strip()]
    else:
        cc_list = []

    # Get Subject
    print("\nEnter the email subject:")
    print("  (You can use placeholders like {{Name}}, {{Course}} from your Excel columns)")
    subject = input("Subject: ").strip()
    if not subject:
        subject = "No Subject Provided"  # Default subject if none given

    return condition, cc_list, subject


def main():
    args = parse_arguments()

    # Get files and account
    df, html_template, attach_files = select_files()
    account_to_use, outlook = select_outlook_account()

    # Get condition, CC list, and subject
    if args.condition is not None or args.cc is not None or args.subject is not None:
        # Command line mode
        condition = args.condition if args.condition else "all"
        cc_list = args.cc.split(",") if args.cc else []
        cc_list = [email.strip() for email in cc_list if email.strip()]
        subject = args.subject if args.subject else "Invitation: Letter of Evaluation for Dr. Claudia Fernandes Nisa"  # Default subject
    else:
        # Interactive mode
        condition, cc_list, subject = get_user_inputs()

    print(f"\nConfiguration:")
    print(f"Condition: {condition}")
    print(f"CC List: {cc_list}")
    print(f"Subject Template: {subject}")
    print(f"Total rows in Excel: {len(df)}")

    # Show preview of which rows will be processed
    if condition.lower() != 'all':
        matching_rows = []
        for index, row in df.iterrows():
            if evaluate_condition(row, condition):
                matching_rows.append(index)
        print(f"Rows that match condition: {len(matching_rows)}")
        if matching_rows:
            print(f"Row indices: {matching_rows[:10]}{'...' if len(matching_rows) > 10 else ''}")

    confirm = input("\nProceed with sending emails? (y/N): ").strip().lower()
    if confirm != 'y':
        print("Operation cancelled.")
        return

    send_emails(df, html_template, account_to_use, outlook, attach_files, condition, cc_list, subject)


if __name__ == "__main__":
    main()
