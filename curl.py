import pycurl
import io
import requests
from requests.auth import HTTPBasicAuth
import json
from pathlib import Path

#https://doc.aidacloud.com/it/aida/input-aida-api
key="5UAB816FDCBYC1BG1MKDG691VD8GKMUUQ0J3XUM16PGRDJ1ZYPJC7SNIJMWYS53T0IXEUCBHG6J4XBY09XASWQEQ8OSM4RHI1AP4523RJNLO6TPO6UCOWB8243058694"
file_ids = ''
headers={'Username': 'abc@gmail.com', 'apikey':'5UAB816FDCBYC1BG1MKDG691VD8GKMUUQ0J3XUM16PGRDJ1ZYPJC7SNIJMWYS53T0IXEUCBHG6J4XBY09XASWQEQ8OSM4RHI1AP4523RJNLO6TPO6UCOWB8243058694'}
# Upload file

f = open('C:\HTS\da_importare\done\GK169RD - carr -26.04.2024 - 808788.pdf', 'rb')

files = {"file": ("C:\HTS\da_importare\done\GK169RD - carr -26.04.2024 - 808788.pdf", f)}

resp = requests.post("https://api.aidacloud.com/api/v1/upload/direct/5UAB816FDCBYC1BG1MKDG691VD8GKMUUQ0J3XUM16PGRDJ1ZYPJC7SNIJMWYS53T0IXEUCBHG6J4XBY09XASWQEQ8OSM4RHI1AP4523RJNLO6TPO6UCOWB8243058694", files=files, headers=headers )
print (resp.text)
print ("status code " + str(resp.status_code))

if resp.status_code == 200:
    print ("Success")
    data = json.loads(resp.text)
    file_ids = data['file_ids']
    print (file_ids)
else:
    print ("Failure")

'''
curl -X POST -H "Content-Type: multipart/form-data" -F "file=@provainvio_carr.pdf" "https://api.aidacloud.com/api/v1/upload/direct/5UAB816FDCBYC1BG1MKDG691VD8GKMUUQ0J3XUM16PGRDJ1ZYPJC7SNIJMWYS53T0IXEUCBHG6J4XBY09XASWQEQ8OSM4RHI1AP4523RJNLO6TPO6UCOWB8243058694"

response = io.StringIO()
c = pycurl.Curl()
c.setopt(c.URL, 'https://api.aidacloud.com/api/v1/upload/direct/' +  key)

c.setopt(c.HTTPPOST, [
    ('fileupload', (
        # upload the contents of this file
        c.FORM_FILE, "C:\HTS\da_importare\done\GK169RD - carr -26.04.2024 - 808788.pdf",
    )),
])
c.setopt(c.WRITEFUNCTION, response.write)
c.perform()
print(response.getvalue())
response.close()
c.close()


'''