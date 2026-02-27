import pandas as pd
import win32com.client as win32
from tkinter import Tk, filedialog
from docx import Document
import mammoth
import re
import os
import chardet

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

def send_emails(df, html_template, account_to_use, outlook, attach_files):
    cc_dlist = ["dku_tenure_and_promotion@dukekunshan.edu.cn"]
    for index, row in df.iterrows():
        html_body = html_template
        for key, value in row.items():
            placeholder = f"{{{{{key}}}}}"  # {{ColumnName}} in HTML
            html_body = html_body.replace(placeholder, str(value))
        html_body = f"""
        <div style="font-family:Calibri, sans-serif; font-size:11pt;">
        {html_body}
        </div>
        """
        #print(html_body)
        cc_list = cc_dlist.copy()
        cc_list.append(row.get("CC", ""))

        #print(row.get("Email", ""))
        #print(f"TA Position Offer - {row.get('Course', '')} (Spring 2026 S4)")
        #print(";".join(cc_list))
        #print(html_body)
        if row.get("Round", -1) == 2:
            mail = outlook.CreateItem(0)
            # Send from selected account
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account_to_use))
            mail.Display()  # use mail.Send() to send automatically
            signature = mail.HTMLBody
            print(signature)
            # Prepend your custom content before the signature
            mail.HTMLBody = html_body + clean_word_empty_paragraphs(signature)
            mail.To = row.get("Email", "")  # assumes column "Email" exists
            mail.cc = ";".join(cc_list)
            if attach_files:
                for file in attach_files:
                    mail.Attachments.Add(file)
            mail.Subject = f"Invitation: Letter of Evaluation for Dr. Claudia Fernandes Nisa"
    print("All emails processed.")

def main():
    df, html_template, attach_files = select_files()
    account_to_use, outlook = select_outlook_account()
    send_emails(df, html_template, account_to_use, outlook, attach_files)

if __name__ == "__main__":
    main()
