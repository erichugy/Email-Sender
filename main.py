from emailClass import Email
from getRecipients import recipients
from wordToHtml import word_to_html
from fileMaster import extract_filename, cwd, move_files_to_subdirectory

# Global Variables
## Inputs
WORD_PATH = r"C:\Users\Eric Huang\Desktop\Coding\Email Sender\msg_2023-01-09.docx"
SUBJECT = "Monthly Mail: The First of Many"



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

email = Email(
    recipients = recipients,
    subject = SUBJECT,
    msg = html,
    signature = signature
)


move_files_to_subdirectory([f"{FILENAME}.htm"])
print("Done")