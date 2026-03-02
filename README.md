## Overview

**Email-Sender** is a Windows-based automation tool for sending personalized emails through Microsoft Outlook. It uses:

- **Excel** as the contact list  
- **Word or HTML** as the email template  
- Optional **attachments**  
- Placeholder-based personalization  

The executable file (.exe) is located in the `dist` folder. It may be run directly or repackaged if needed.



## Core Functionality

- Reads recipient data from an Excel file  
- Replaces placeholders in a Word or HTML template  
- Generates individualized emails  
- Supports per-recipient CC and global CC  
- Allows conditional sending  
- Opens generated emails as drafts in Outlook before sending  

---

## System Restrictions

- Windows operating system only  
- All files (Excel, Word, HTML) must be closed before running  
- Images cannot be embedded in the email body  



## Template Format

Use placeholders in the following format:

`{{Column Name}}`

Each placeholder must exactly match a column name in the Excel sheet.

Example:  
`Dear {{FirstName}},`



## Excel Sheet Requirements

- First row: Column headers  
- A column named exactly:  
  - `Email` (required)  
  - `CC` (required only if using per-recipient CC)

Each row represents one recipient.



## Execution Steps

1. Run the `.exe` file  
2. Select the Excel contact list  
3. Select a Word or HTML template  
   - HTML saved from Word preserves richer formatting  
4. (Optional) Select an attachment file  
5. Choose the sender email address  
6. (Optional) Set a sending condition  
7. (Optional) Enter global CC addresses  
8. Enter the email subject  
9. Confirm details and enter `y` to proceed  
10. Emails open as drafts in Outlook  



## AI Usage Disclosure

Artificial intelligence tools assisted with:

- Code suggestions  
- Refactoring  
- Debugging  
- Documentation drafting  

All generated content was reviewed, modified, and validated before inclusion.
