import settings

redlook_settings = settings.Configuration(filename = "redlook.conf")

print(redlook_settings.redmine["redmine_address"])
print(redlook_settings.redmine["redmine_username"])
print(redlook_settings.redmine["redmine_password"])
print(redlook_settings.outlook["outlook_folder"])