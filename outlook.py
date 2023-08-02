import win32com.client

class Outlook:
    def __init__(self, folder):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6).Folders[folder].Items