import settings
import outlook
import redmine
import base64
import hashlib
import redminelib
import time
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
        #self.mails = dict()

    def get_file_id(self, filename, attachments):
        for attachment in attachments:
            if (attachment.filename == filename):
                return attachment.id

    def do(self):
        for mail in self.Inbox:
            mail_items = self.Outlook.get_mail(mail)
            print("{}".format(mail_items["Subject"]))
            #title = base64.b64encode(mail_items["Subject"].encode()).decode()
            title = hashlib.md5(mail_items["Subject"].encode()).hexdigest()
            redmine_uploads = list()
            attachments = dict()
            for attachment in mail_items["Attachments"]:
                attachment_items = self.Outlook.get_attachment(title = title, attachment = attachment)
                self.Outlook.save_attachment(title = title, attachment = attachment)
                path = attachment_items["Path"]
                cid = attachment_items["CID"]
                filename = attachment_items["FileName"]
                redmine_upload = {"path": path, "cid": cid, "filename": filename}
                redmine_uploads.append(redmine_upload)
                attachments[cid] = attachment_items["FileName"]
            #self.mails[title] = mail_items
            htmlbody = "{{html\n" + mail_items["HTMLbody"] + "\n}}"
            wiki_page = None
            try:
                wiki_page = self.Redmine.create_wiki_page(mail_items["Subject"], htmlbody, "Wiki", redmine_uploads)
            except redminelib.exceptions.ValidationError as e:
                print("Exception error: {}, skipped".format(e))
            soup = BeautifulSoup(mail_items["HTMLbody"], "html.parser")
            if (wiki_page == None):
                wiki_page = self.Redmine.get_wiki_page(mail_items["Subject"])
            for img in soup.find_all("img"):
                src = img["src"].replace("cid:", "")
                filename = attachments[src]
                id = self.get_file_id(filename, list(wiki_page.attachments))
                img["src"] = "/redmine/attachments/download/{}/{}".format(id, filename)
            wiki_page.text = "{{html\n" + str(soup) + "\n}}"
            wiki_page.save()
            
if __name__ == "__main__":
    redlook = Redlook()
    redlook.do()