REDMINE_CONFIGURATION = ["redmine_address", "redmine_username", "redmine_password"]
OUTLOOK_CONFIGURATION = ["outlook_folder", "outlook_attachments_path"]
CONFIGURATIONS = [REDMINE_CONFIGURATION, OUTLOOK_CONFIGURATION]

class Configuration:
    redmine = dict()
    outlook = dict()
    configuration_list = [redmine, outlook]
    configuration_name = ["redmine_", "outlook_"]
    def __init__(self, redmine_address = "", redmine_username = "", redmine_password = "", outlook_folder = "", outlook_attachments_path = "", filename = ""):
        if (filename == ""):
            self.redmine["redmine_address"] = redmine_address
            self.redmine["redmine_username"] = redmine_username
            self.redmine["redmine_password"] = redmine_password
            self.outlook["outlook_folder"] = outlook_folder
            self.outlook["outlook_attachments_path"] = outlook_attachments_path
        else:
            conf_file = open(filename, mode = "r")
            confs = conf_file.readlines()

            for conf in confs:
                if (conf[0] == "#"):
                    continue
                else:
                    idx = 0
                    for configuration in CONFIGURATIONS:
                        for item in configuration:
                            item_len = len(item)
                            if (conf[:item_len] == item):
                                value = conf.split("=")[1].replace(self.configuration_name[idx], "").replace("\n", "").replace("\"", "")
                                self.configuration_list[idx][item] = value
                        idx += 1