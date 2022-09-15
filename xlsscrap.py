import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.reader.excel import load_workbook

book = 'ada.org_final_.xlsx'
book_loaded = load_workbook(book)
workbook = pd.ExcelFile(book)
df = workbook.parse(0, header=0)
dl = df['Domain'].values.tolist()
for i in range(len(dl[0:5])):
    url = 'https://' + dl[i]
    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 '
                      'Safari/537.36'}
    request = requests.get(url, headers=headers)
    status = request.status_code
    soup = BeautifulSoup(request.text, 'html.parser')
    ti = soup.find('title').text
    hh = soup.findAll('h1')
    sheets = book_loaded.sheetnames
    sheet1 = book_loaded[sheets[0]]
    sheet1.cell(row=(i+2), column=3).value = ti
    for a in range(len(hh)):
        sheet1.cell(row=(i+2), column=4+a).value = hh[a].text
    book_loaded.save('ada.org_final_.xlsx')



