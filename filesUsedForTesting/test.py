import win32com.client as win32
import os
cwd = os.path.abspath(os.getcwd())

#Create Outlook Application Instance 
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
# mail.To = 'fireshot2002@gmail.com; eric.huang3143@gmail.com'
mail.BCC = 'fireshot2002@gmail.com; eric.huang3143@gmail.com'

mail.Subject = 'Subject'
with open('test.htm', 'r') as file:
    html = file.read()
with open(cwd + r"\Signatures\Uni Student (eric.huang5@mail.mcgill.ca).htm", "r") as file:
    signature = file.read()
# signature = outlook.Session.Accounts.Item(1).DeliveryStore.GetDefaultFolder(6).Items.Item(1).Body
mail.HTMLBody = html + "<p></p>"*2 + signature



# send_time = dt.datetime(2023, 1, 9, 3, 5, 0)
# mail.DeferredDeliveryTime = send_time

# mail.Body = 'Message body'
mail.Send()