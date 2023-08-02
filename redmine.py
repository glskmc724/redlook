from redminelib import Redmine

class redmine:
    def __init__(self, host, user, pw, project):
        self.redmine = Redmine("http://{}/redmine".format(host), username = user, password = pw)
        self.project = self.redmine.project.get(project)
        
    def create_wiki_page(self, title, html, parent_title, uploads):
        return self.redmine.wiki_page.create(project_id = "iron", title = title, text = html, parent_title = parent_title, uploads = uploads)
    
    def get_wiki_page(self, title):
        return self.redmine.wiki_page.get(title, project_id = "iron")