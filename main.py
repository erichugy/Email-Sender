from emailClass import Email
from getRecipients import recipients
from wordToHtml import word_to_html
from fileMaster import extract_filename, cwd, move_files_to_subdirectory, get_file_paths

# Global Variables
## Inputs
WORD_PATH = r"C:\Users\Eric Huang\Desktop\Coding\Email Sender\Old-Emails\2023-02-07\msg_2023-02-07.docx"
SUBJECT = "Checking-in with Eric: Event-full"
PATH_TO_ATTACHMENTS = r"C:\Users\Eric Huang\Desktop\Coding\Email Sender\Old-Emails\2023-02-07\attachments"



## Default Inputs
SIGNATURE_PATH = cwd + r"\Signatures\Uni Student (eric.huang5@mail.mcgill.ca).htm"

## Calculated Global Variables
FILENAME = extract_filename(WORD_PATH)


#----------------------------------
# Step 1: Convert word file of email to html

word_to_html(WORD_PATH, FILENAME)


#----------------------------------
# Step 2: Read HTMLs as strings
attempts = 0
import time
while attempts < 10:
    try:
        with open(f'{FILENAME}.htm', 'r') as file:
            html = file.read()
        break
    except FileNotFoundError:
        time.sleep(5)
        print("Attempt " + attempts)
        attempts += 1


with open(SIGNATURE_PATH, "r") as file:
    signature = file.read()


#----------------------------------
# Step 3: Create email object
email = Email(
    recipients = recipients,
    subject = SUBJECT,
    msg = html,
    signature = signature,
    attachments = get_file_paths(PATH_TO_ATTACHMENTS)
    )

#move completed htm file (email) into the same folder as the word doc
move_files_to_subdirectory([f"{FILENAME}.htm"])
print("Done")