# Email-Sender

Email-Sender automates the email sending process through Word, HTML as Email content input, Excel as a list of contacts
including recevier's email address and CC's email address. The placeholder in Word or HTHL will be replaced by the value in the Excel sheet to generate personalize email. This email sender script also allow user to select an attachment
file to send with.

The .exe file is in the dist folder, feel free to run it as Python or repackage it if feeling unsafe.

## Restriction
- The script can only run on window operating system.
- The file needs to be closed when using

## Instruction
### Email Template Format
Place "{{Column Name}}" as a placeholder, where the script will replace it with the value in the Excel Sheet.

### Running
1. Run the exe
2. Select the contact list as an Excel Sheet
3. Select a word or HTML template(HTML template saves in Word has more details in the email.)
4. (Optional) Select attachment file
5. Select the Email Address to send from
6. (Optional) Enter the condition for sending an email
7. (Optional) Enter the list of CC email addresses that would apply to all email
8. Enter the Email Subject
9. Check the information and enter y to send
10. The draft of the email will show in the window of Outlook

## Other
Artificial intelligence tools were used to assist in developing this project. 
AI support included code suggestions, refactoring assistance, documentation drafting, 
and debugging guidance. All generated content was reviewed, modified, and validated 
by the author before inclusion.
