import settings
import outlook
import redmine
from bs4 import BeautifulSoup

class Redlook:
    def __init__(self):
        self.redlook_settings = settings.Configuration(filename = "redlook/redlook.conf")
        self.outlook_folder = self.redlook_settings.outlook["outlook_folder"]
        self.redmine_address = self.redlook_settings.redmine["redmine_address"]
        self.redmine_username = self.redlook_settings.redmine["redmine_username"]
        self.redmine_password = self.redlook_settings.redmine["redmine_password"]
        self.Outlook = outlook.Outlook(self.outlook_folder)
        self.Redmine = redmine.redmine(self.redmine_address, self.redmine_username, self.redmine_password)
        self.Inbox = self.Outlook.inbox


if __name__ == "__main__":
    redlook = Redlook()