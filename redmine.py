from redminelib import Redmine

class redmine:
    def __init__(self, host, user, pw, project):
        self.redmine = Redmine("http://{}/redmine".format(host), username = user, password = pw)
        self.project = self.redmine.project.get(project)