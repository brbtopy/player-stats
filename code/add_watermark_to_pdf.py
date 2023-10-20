# api code to generate watermark image on pdf
#LIVE API KEY = pdf_live_09Jw3fmsDx53azf2r0TEjwhxqlqriaVRV3TGCHtl539

import requests
import json

instructions = {
  'parts': [
    {
      'file': 'document'
    }
  ],
  'actions': [
    {
      'type': 'watermark',
      'image': 'logo',
      'width': '70%',
      "opacity": 0.2
    }
  ]
}

response = requests.request(
  'POST',
  'https://api.pspdfkit.com/build',
  headers = {
    'Authorization': 'Bearer pdf_live_09Jw3fmsDx53azf2r0TEjwhxqlqriaVRV3TGCHtl539'
  },
  files = {
    'document': open('Indv. Player Profile.pdf', 'rb'),
    'logo': open('logo.png', 'rb')
  },
  data = {
    'instructions': json.dumps(instructions)
  },
  stream = True
)

if response.ok:
  with open('Indv. Player Profile.pdf', 'wb') as fd:
    for chunk in response.iter_content(chunk_size=8096):
      fd.write(chunk)
else:
  print(response.text)
  exit()
