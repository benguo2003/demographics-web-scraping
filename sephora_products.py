import requests
from bs4 import BeautifulSoup
import openpyxl
import json
from requests_html import HTMLSession

wb = openpyxl.load_workbook("C:\\Users\\ben.guo\\Desktop\\sephora.xlsx")

sheet = wb["Bath & Body"]

r = requests.get('https://www.res-x.com/ws/r2/Resonance.aspx'
                 '?appid=sephora01&tk=534940934506231&ss=5840'
                 '31261737188&sg=1&pg=res22071109950148660567'
                 '587&vr=5.10x&bx=true&sc=content2_rr&ev=&ei='
                 '&no=20&language=ENGLISH&categoryid=cat140014'
                 '&page=1&ccb=Sephora.certona&ur=https%3A%2F%2'
                 'Fwww.sephora.com%2Fbeauty%2Fbest-selling-bath'
                 '-body-products&plk=&rf=https%3A%2F%2Fwww.sephora.com%2Fshop%2Fbath-body')

json_file = json.loads(r.text[16:-2])
for i in range(0, 20):
    location = i + 2
    sheet[f'A{location}'].value = json_file['resonance']['schemes'][0]['items'][i]['display_name']

for i in range(0, 20):
    location = i + 2
    sheet[f'B{location}'].value = json_file['resonance']['schemes'][0]['items'][i]['brand_name']

for i in range(0, 20):
    location = i + 2
    sheet[f'C{location}'].value = json_file['resonance']['schemes'][0]['items'][i]['skus'][0]['list_price']

for i in range(0, 20):
    location = i + 2
    sheet[f'D{location}'].value = json_file['resonance']['schemes'][0]['items'][i]['default_sku_id']

wb.save("C:\\Users\\ben.guo\\Desktop\\sephora.xlsx")


