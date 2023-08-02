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
                "Receiver": mail.To,
                "Received_Time": mail.ReceivedTime,
                "Attachments": mail.Attachments}
        return ret
    
    def get_attachment(self, title, attachment):
        path = "{}/{}".format(self.attachments_path, title)
        ret = {"FileName": attachment.FileName,
                "Path": "{}/{}".format(path, attachment.FileName),
                "CID": attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")}
        return ret
    
    def save_attachment(self, title, attachment):
        path = "{}/{}".format(self.attachments_path, title)
        if (not os.path.exists(path)):
            os.mkdir(path)
        path = os.path.abspath("{}/{}".format(path, attachment.FileName))
        attachment.SaveAsFile(path)