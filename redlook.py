import settings
import outlook
import redmine
import base64
import hashlib
import redminelib
import time
from bs4 import BeautifulSoup, Comment

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
        return False

    def do(self):
        self.Inbox.Sort("[ReceivedTime]", False)
        for mail in self.Inbox:
            mail_items = self.Outlook.get_mail(mail)
            #print("title:{}".format(mail_items["Subject"]))
            subject = mail_items["Subject"]
            subject = subject.replace("RE: ", "").replace("Re: ", "").replace("rE: ", "").replace("re: ", "")
            subject = subject.replace("FW: ", "").replace("Fw: ", "").replace("fW: ", "").replace("fw: ", "")
            subject = subject.replace("RE:", "").replace("Re:", "").replace("rE:", "").replace("re:", "")
            subject = subject.replace("FW:", "").replace("Fw:", "").replace("fW:", "").replace("fw:", "")
            subject = subject[:32]
            #print("Conv ID: {}, subject: {}, Receive date: {}".format(mail.ConversationID, subject, mail_items["Received_Time"]))
            #title = base64.b64encode(mail_items["Subject"].encode()).decode()
            title = hashlib.md5(mail_items["Subject"].encode()).hexdigest()
            redmine_uploads = list()
            attachments = dict()
            for attachment in mail_items["Attachments"]:
                attachment_items = self.Outlook.get_attachment(title = title, attachment = attachment)
                path = attachment_items["Path"]
                cid = attachment_items["CID"]
                filename = self.Outlook.save_attachment(title = title, attachment = attachment, cid = cid)
                redmine_upload = {"path": path, "cid": cid, "filename": filename}
                redmine_uploads.append(redmine_upload)
                attachments[cid] = filename
            #self.mails[title] = mail_items
            soup = BeautifulSoup(mail_items["HTMLbody"], "html.parser")
            soup.html.unwrap()
            soup.head.extract()
            soup.body.unwrap()
            htmlbody = "{{html\n" + str(soup) + "\n}}"
            try:
                wiki_page = self.Redmine.get_wiki_page(subject)
            except:
                wiki_page = self.Redmine.create_wiki_page(subject, htmlbody, "Email", redmine_uploads)

            uploads = []
            for upload in redmine_uploads:
                if (self.get_file_id(upload["filename"], list(wiki_page.attachments)) == False):
                    uploads.append(upload)

            wiki_page.uploads = uploads
            wiki_page.text = "a"
            wiki_page.save()
            
            """
            <div id="appendonsend"></div>
            <hr style="display:inline-block;width:98%" tabindex="-1"/>
            <div dir="ltr" id="divRplyFwdMsg">
            <font color="#000000" face="Calibri, sans-serif" style="font-size:11pt">
            <b>보낸 사람:</b>username &lt;email&gt;<br/>
            <b>보낸 날짜:</b><br/>
            <b>받는 사람:</b><br/>
            <b>참조:</b> <br/>
            <b>제목:</b> </font>
            <div> </div>
            </div>
            """
            
            wiki_header =  "<div id=\"appendonsend\"></div>"
            wiki_header += "<hr style=\"display:inline-block;width:98%\" tabindex=\"-1\"/>"
            wiki_header += "<div dir=\"ltr\" id = \"divRplyFwdMsg\">"
            wiki_header += "<font color=\"#000000\" face=\"Calibri, sans-serif\" style=\"font-size:11pt\">"
            if (mail.SenderEmailType == "EX"):
                sender = mail.Sender.GetExchangeUser().PrimarySmtpAddress
                wiki_header += "<b>보낸 사람:</b>{} &lt;{}&gt;<br/>".format(mail_items["Transfer"], sender)
            else:
                wiki_header += "<b>보낸 사람:</b>{} &lt;{}&gt;<br/>".format(mail_items["Transfer"], mail_items["Transfer_Email"])
            wiki_header += "<b>보낸 날짜:</b>{} <br/>".format(mail_items["Received_Time"])
            wiki_header += "<b>받는 사람:</b>"
            for recipient in mail_items["Recipients"]:
                if (recipient.Type != 2):
                    email_address = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                    wiki_header += "{} &lt;{}&gt; ".format(recipient, email_address)
            wiki_header += "<br/>"
            wiki_header += "<b>참조:</b>"
            for recipient in mail_items["Recipients"]:
                if (recipient.Type == 2):
                    email_address = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                    wiki_header += "{} &lt;{}&gt; ".format(recipient, email_address)
            wiki_header += "<br/>"
            wiki_header += "<b>제목:</b>{} </font>".format(mail_items["Subject"])
            wiki_header += "<div> </div>"
            wiki_header += "</div>"
            wiki_header += "<br/>"

            wiki_page = self.Redmine.get_wiki_page(subject)
            for img in soup.find_all("img"):
                src = img["src"].replace("cid:", "")
                filename = attachments[src]
                id = self.get_file_id(filename, list(wiki_page.attachments))
                img["src"] = "/redmine/attachments/download/{}/{}".format(id, filename)
            wiki_page.text = "{{html\n" + wiki_header + str(soup) + "\n}}"
            wiki_page.save()

            
if __name__ == "__main__":
    redlook = Redlook()
    redlook.do()