# from sheet2api import Sheet2APIClient


from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
site_url = "https://qedge2020-my.sharepoint.com/:x:/r/personal/idsbgoffice5_qedge2020_onmicrosoft_com/Documents/Material%20and%20Build/Daily%20Update%20Format%20Master%20List.xlsx?d=w304f04022cb847fdb59a3d6e1e1a3fb4&csf=1&web=1&e=Cpj4PC"
# ctx = ClientContext(site_url).with_credentials(UserCredential("{idsbgectest@outlook.com}", "{Gdstgdsf114141!}"))
# web = ctx.web.get().execute_query()
# print("Web title: {0}".format(web.properties['Title']))
password = "Gdstgdsf114141!"
userName = "idsbgectest@outlook.com"
user_credentials = UserCredential(userName,password)
ctx = ClientContext(site_url).with_credentials(user_credentials)
web = ctx.web
ctx.load(web)
ctx.execute_query()
print(web.url)

