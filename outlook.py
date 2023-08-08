import win32com.client
import os

class Outlook:
    def __init__(self, folder, path):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6).Folders[folder].Items
        self.attachments_path = path

    def get_mail(self, mail):
        ret = {"Subject": mail.Subject,
                "HTMLbody": mail.HTMLbody,
                "Transfer": mail.SenderName,
                "Transfer_Email": mail.SenderEmailAddress,
                "Receiver": mail.To,
                "Received_Time": mail.ReceivedTime,
                "Recipients": mail.Recipients,
                "Attachments": mail.Attachments}
        return ret
    
    def get_attachment(self, title, attachment):
        cid = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
        filename = "{}_{}".format(cid, attachment.FileName)
        path = "{}/{}".format(self.attachments_path, title)
        ret = {"FileName": attachment.FileName,
                "Path": "{}/{}".format(path, filename),
                "CID": cid, }
        return ret
    
    def save_attachment(self, title, attachment, cid):
        path = "{}/{}".format(self.attachments_path, title)
        if (not os.path.exists(path)):
            os.mkdir(path)
        filename = "{}_{}".format(cid, attachment.FileName)
        path = os.path.abspath("{}/{}".format(path, filename))
        attachment.SaveAsFile(path)
        return filename