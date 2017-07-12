import requests

local_proxy = '127.0.0.1:8118'
http_proxy = {'http': local_proxy,
              'https': local_proxy}

current_ip = requests.get(url='http://icanhazip.com/',
                          proxies=http_proxy,
                          verify=False)
