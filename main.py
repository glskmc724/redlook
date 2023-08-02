import settings
import outlook

redlook_settings = settings.Configuration(filename = "redlook.conf")
Outlook = outlook.Outlook(redlook_settings.outlook["outlook_folder"])
