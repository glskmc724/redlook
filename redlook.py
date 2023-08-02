import settings
import outlook
import redmine
import base64
from bs4 import BeautifulSoup

class Redlook:
    def __init__(self):
        self.redlook_settings = settings.Configuration(filename = "redlook/redlook.conf")
        self.outlook_folder = self.redlook_settings.outlook["outlook_folder"]
        self.redmine_address = self.redlook_settings.redmine["redmine_address"]
        self.redmine_username = self.redlook_settings.redmine["redmine_username"]
        self.redmine_password = self.redlook_settings.redmine["redmine_password"]
        self.Outlook = outlook.Outlook(self.outlook_folder, self.redlook_settings.outlook["outlook_attachments_path"])
        self.Redmine = redmine.redmine(self.redmine_address, self.redmine_username, self.redmine_password, project = "iron")
        self.Inbox = self.Outlook.inbox
        self.mails = dict()

    def do(self):
        for mail in self.Inbox:
            mail_items = self.Outlook.get_mail(mail)
            title = base64.b64encode(mail_items["Subject"].encode()).decode()
            for attachment in mail_items["Attachments"]:
                attachment_items = self.Outlook.get_attachment(attachment)
                self.Outlook.save_attachment(title = title, attachment = attachment)
            self.mails[title] = mail_items
            break

if __name__ == "__main__":
    redlook = Redlook()
    redlook.do()