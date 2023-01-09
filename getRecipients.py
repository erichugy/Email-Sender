import gspread
import pandas as pd
from fileMaster import cwd


#Pathing
ROOT = cwd
CREDENTIALS_FILENAME = "bot-credentials.json"
PATH_TO_CREDENTIALS = f"{ROOT}\\{CREDENTIALS_FILENAME}"

#Filenames
SHEETNAME = "Email Subscribers"
TABNAME = "Emails"


# Get File
client = gspread.service_account(filename=PATH_TO_CREDENTIALS)

# Open the Google Sheets document and get the specified sheet
spreadsheet = client.open(SHEETNAME)
worksheet = spreadsheet.get_worksheet(0)  # Get the first worksheet


# Read all the rows in the sheet as a list of dictionaries
data = worksheet.get_all_records()

# Turn data into dataframe
df = pd.DataFrame.from_dict(data)

# Drop Emtpy Rows
df = df[df.Email != ""] #Dropna won't work cuz Sheets turns everything into text

# Drop Check == False, since check means send them an email
df = df[df.Check == "TRUE"]


# Get Emails

def emails_to_string(emails: pd.Series) -> str:
    """Converts a pandas series of emails to a string with each email separated by a semicolon.

    Args:
        emails: A pandas series containing the emails.

    Returns:
        A string with each email separated by a semicolon.
    """
    return "; ".join(emails)

emails = df["Email"]
recipients = emails_to_string(emails)