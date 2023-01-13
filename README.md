# README.md

## Overview
This program is a tool for sending emails to a list of recipients stored in a Google Sheets document. It uses the `Email` class from the `emailClass.py` file to handle the creation and sending of emails, and the `fileMaster.py` file to handle file management tasks such as moving sent emails to an archive directory. 

## Requirements
- A Google Sheets document with a list of email addresses in the first column
- A valid Outlook account
- The Outlook application installed on your device and connected to the Outlook account of your choice.
- The Microsoft Word application installed
- Python 3 with the `win32com` library installed

## Usage
1. Clone or download the repository onto your local machine.
2. In the `fileMaster.py` file, update the `SUBDIRECTORY` variable to the desired location where sent emails will be archived.
3. In the `emailClass.py` file, update the `signature` variable with the desired signature for the emails. Signatures can be created within Outlook or through HTML. The folder containing all your signatures need to be in the same folder as the rest of the project.
4. Create a new instance of the `Email` class with the desired recipients, subject, and message body. 
5. If you would like to send the email immediately, pass in `True` for the `send` argument when initializing the `Email` class. If you would like to save the email as a draft, pass in `False` (default) or omit the `send` argument.

Example:
```python
from emailClass import Email

recipients = "example1@gmail.com, example2@gmail.com"
subject = "Test Email"
message = "Hello World!"
email = Email(recipients, subject, message, send=True)
```
## Note
The `Email` class uses the BCC field for the recipients by default. If you wish to use the To field, pass in `False` for the `bcc` argument when initializing the class.

Also, the program uses win32com library, which allows communication with Outlook, this is only available on windows operating system.
