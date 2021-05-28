import requests
import json

headers = {'Content-Type': 'application/json; chearset=utf-8'}

# shops = ['동대문(아)','대구(아)','가든(아)','송도(아)','김포(아)','가산(아)','목동(백)','울산(백)',
# '미아(백)','디큐브(백)','천호(백)','킨텍스(백)','판교(백)','동구(백)','본점(백)','무역(백)','신촌(백)','부산(백)','중동(백)','충청(백)','대구(백)']
shops = ['무역(백)','동구(백)','판교(백)']
url = 'http://10.103.200.51:8081/api/task'

for shop in shops:
    data = {'name': '네이버_고객부담배송비_{}'.format(shop), 'month': '*', 'week': '', 'day':'25'}
    requests.post(url, data=json.dumps(data), headers=headers)
