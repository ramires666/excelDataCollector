import requests

url = "https://mon.tranco.kz/wialon/ajax.html?svc=report/exec_report&sid=49915d17c6dd955e98bcbbd1ee51fb96"

headers = {
    "accept": "*/*",
    "accept-encoding": "gzip, deflate, br, zstd",
    "accept-language": "en-US,en;q=0.9,ru;q=0.8",
    "content-type": "application/x-www-form-urlencoded",
    "cookie": "lang=ru; gr=1; sessions=49915d17c6dd955e98bcbbd1ee51fb96",
    "origin": "https://mon.tranco.kz",
    "referer": "https://mon.tranco.kz/wialon/post.html",
    "sec-ch-ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A_Brand";v="24"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
}

data = {
    "params": '''{"reportResourceId":3331,
               "reportTemplateId":6,
               "reportTemplate":"",
               "reportObjectId":3353,
               "reportObjectSecId":0,
               "interval":
                   {"flags":16777216,
                    "from":1704049200,
                    "to":1706727599},
               "remoteExec":1,
               "reportObjectIdList":[]
               }''',
    "sid": "49915d17c6dd955e98bcbbd1ee51fb96"
}


# data = {
#     "params": '{"tableIndex":0,"config":{"type":"range","data":{"from":0,"to":47,"level":0,"unitInfo":1}}}',
#     "sid": "49915d17c6dd955e98bcbbd1ee51fb96"
# }

# Send the POST request
response = requests.post(url, headers=headers, data=data)

# Print the response details
print("Response status code:", response.status_code)
print("Response body:", response.text)
