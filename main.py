# import ui.app as app
#
# app.start()
import base.auth as auth
id = '1Rnw8YZZSy5svtdlxxVRD22IvyhjOM0TRfIj1zZBcxd8'
print(auth.drive_service.files().get(fileId=id).execute())
