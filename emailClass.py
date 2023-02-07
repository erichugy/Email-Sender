from typing import List
class Email:
    """Representation of an email.

    An Email has recipients, a subject line, and a message body. It can be in plain text or HTML format.
    An Email can be sent immediately or saved as a draft. An optional signature can be added to the email if it is in HTML format.
    """

    import win32com.client as win32

    # Create Outlook Application Instance for whole class
    outlook = win32.Dispatch("Outlook.Application")

    def __init__(self, 
                recipients: str, 
                subject: str, 
                msg: str, 
                html: bool = True, 
                signature: str = None, 
                bcc: bool = True, 
                send: bool = False,
                attachments :List[str] = [""]
                ) -> None:
        """Initializes an Email instance with recipients, a subject line, a message body, and optional arguments for the format of the email, a signature, and whether to send the email or save it as a draft.

        Args:
            recipients: The recipients of the email as a string.
            subject: The subject line of the email as a string.
            msg: The message body of the email as a string.
            html: Whether the message body is in HTML format (default is True).
            signature: The HTML of the signature to be added to the email (default is None).
            bcc: Whether to use the BCC field or the To field for the recipients (default is True).
            send: Whether to send the email immediately or save it as a draft (default is False).
            attachments : Attachments to add. Need to add the whole path for each attachment as elements of the list

        Returns:
            None
        """

        self.mail = Email.outlook.CreateItem(0)
        self.subject = self.mail.Subject = subject
        self.html = html
        self.signature = "<p></p>" * 2 + signature # <p></p> is to add newline
        self.bcc = bcc

        #Attachements Section
        if not attachments or not (len(attachments) >= 1 and not attachments[0]):
            for i in attachments:
                self.mail.Attachments.Add(i)


        # Main Section
        # Getter Setter Variables
        self.to = recipients
        self.body = msg

        if not send:
            self.save_draft()
        else:
            self.send()

    
    #-Getter/Setter Section-----

    # Message Recipients
    @property
    def to(self): #Getter
        return self._to
    
    @to.setter #Setter
    def to(self, recipients):
        self._to = recipients
        if self.bcc:
            self.mail.BCC = self._to
        else:
            self.mail.To = self._to

    # Message Body
    @property
    def body(self): #Getter
        return self._body
    
    @body.setter #Setter
    def body(self, msg):
        self._body = msg

        if self.html:
            self.mail.HTMLBody = self._body + self.signature
        else:
            self.mail.Body = self._body

    #--------------------------

    def save_draft(self):
        self.mail.Save()
        
    def send(self):
        self.mail.Send()

    def __str__(self) -> str:
        lines = [
            f"To: {self.to}",
            f"Subject Line: {self.subject}",
            f"Email Body (below): \n{self.body}"
        ]
        return "\n".join(lines)

