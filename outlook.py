import win32com.client

class Outlook:
    def __init__(self, folder):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6).Folders[folder].Items

    def get_mail(self, mail):
        subject = mail.Subject
        htmlbody = mail.HTMLbody
        transfer = mail.SenderName
        receiver = mail.To
        received_time = mail.ReceivedTime
        attachments = mail.attachments

        return subject, htmlbody, transfer, receiver, received_time, attachments
    
    def get_attachment(self, attachment):
        cid = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")