import requests
import xlwt
from bs4 import BeautifulSoup
from user_agent import generate_user_agent

'''
page_response = requests.get("https://www.ozon.ru/category/pelenki-i-podguzniki-dlya-zhivotnyh-12414/", headers=headers)
soup = BeautifulSoup(page_response.text,'html.parser')

page = requests.get("https://www.ozon.ru/category/pelenki-i-podguzniki-dlya-zhivotnyh-12414/", proxies = {'http':"175.144.198.226"})

soup = BeautifulSoup(page.text,'html.parser')
print(soup)
print(page.status_code)

response = requests.get("https://www.ozon.ru/category/pelenki-i-podguzniki-dlya-zhivotnyh-12414/")
print(response.content.decode('cp1251',errors='ignore'))
'''

book = xlwt.Workbook(encoding="utf-8")
sheet = book.add_sheet("Sheet1")
'''
print("page =")
links = input()
'''
print("Элементы через запятую")
MassElem = input().split(',')
print("user name")
username = input()
print("projectname")
projectname = input()

counter = 1
sheet.write(0,0,'Животное')
sheet.write(0, 1, 'Товар')
sheet.write(0, 2, 'Ссылка на товар')
sheet.write(0, 3, 'Ссылка на фотографию')
sheet.write(0, 4, 'Бренд')
sheet.write(0, 5, 'Тип')
sheet.write(0, 6, 'Тип2')
sheet.write(0, 7, 'Описание')
sheet.write(0, 8, 'Цена')
sheet.write(0, 9, 'Наличие')
sheet.write(0, 10, 'Характеристики')
'''
headers = {'User-Agent': generate_user_agent(device_type="desktop", os=('mac', 'linux'))}
page = requests.get("https://www.ozon.ru/category/obuv-dlya-sobak-12332/?page=1", headers=headers)
print(page.cookies)

session = requests.Session()
print(session.get("https://www.ozon.ru/category/obuv-dlya-sobak-12332/?page=1", headers=headers))

soup = BeautifulSoup(page.text, 'html.parser')
all_data = soup.find_all('div', {'class': "a0s9 a0t0"})
print(all_data)
'''

print('text file name')
filename = input()
page = open(filename + '.txt','r',encoding="utf-8")

while True:
    soup = BeautifulSoup(page, 'html.parser')
    all_data = soup.find_all('div', {'class': MassElem[0]})
    if all_data == []:
        break
    for elem in all_data:
        headers = {'User-Agent': generate_user_agent(device_type="desktop", os=('mac', 'linux'))}
        link = elem.find('a', {'class': MassElem[1]})
        link3 = 'https://www.ozon.ru' + link['href']
        sheet.write(counter, 1, link.text)
        sheet.write(counter, 2, link3)
        suggestion = ''
        pricesale = elem.find('span', {'class': MassElem[2]})
        price = elem.find('span', {'class': MassElem[3]})
        try:
            sheet.write(counter, 8, pricesale.text)
        except:
            try:
                sheet.write(counter, 8, price.text)
            except:
                sheet.write(counter, 8, '')
        page2 = requests.get(link3, headers=headers)
        if page2.status_code == 200:
            print('Элементов готово : ' + str(counter))
            soup2 = BeautifulSoup(page2.text, 'html.parser')
            data4 = soup2.find('div', {'class': 'gallery'})
            curl = data4['coverimage']
            sheet.write(counter, 3, curl)
            description = soup2.find('div', {'class': MassElem[4]})
            try:
                sheet.write(counter, 7, description.text)
            except:
                sheet.write(counter, 7, '')
            leftm = []
            rightm =[]
            suggestionsl = soup2.find_all('div', {'class': MassElem[5]})
            suggestionsr = soup2.find_all('div', {'class': MassElem[6]})
            for elem in suggestionsl:
                leftm.append(elem.text)
            for elem in suggestionsr:
                rightm.append(elem.text)
            for i in range(len(suggestionsr)):
                if leftm[i] not in suggestion:
                    if 'Тип' in leftm[i] or leftm[i] == 'Предназначено для' or 'Вид' in leftm[i] :
                        continue
                    suggestion = suggestion + ' ' + leftm[i] + ' ' + rightm[i]
            sheet.write(counter,10, suggestion)
            try:
                breand = rightm[leftm.index('Бренд')]
                sheet.write(counter, 4, breand)
            except:
                sheet.write(counter, 4, '')
        sheet.write(counter,9, 'В наличии')
        counter +=1
    break
'''
pages = 1
while True:
  headers = {'User-Agent': generate_user_agent(device_type="desktop", os=('mac', 'linux'))}
  page_link = links +  "?layout_container=sold_out&layout_page_index=1&sold_out_page=" + str(pages)
  page = requests.get(page_link, headers=headers)
  if page.status_code == 200:
      soup = BeautifulSoup(page.text, 'html.parser')
      all_data = soup.find_all('div', {'class': MassElem[0]})
      if all_data == []:
          break
      for elem in all_data:
          link = elem.find('a', {'class': MassElem[7]})
          link3 = 'https://www.ozon.ru' + link['href']
          sheet.write(counter, 1, link.text)
          sheet.write(counter, 2, link3)
          suggestion = ''
          pricesale = elem.find('span', {'class': MassElem[2]})
          price = elem.find('span', {'class': MassElem[3]})
          try:
              sheet.write(counter, 8, pricesale.text)
          except:
              try:
                  sheet.write(counter, 8, price.text)
              except:
                  sheet.write(counter, 8, '')
          page2 = requests.get(link3, headers=headers)
          if page2.status_code == 200:
              print('Элементов готово : ' + str(counter))
              soup2 = BeautifulSoup(page2.text, 'html.parser')
              data4 = soup2.find('div', {'class': 'gallery'})
              curl = data4['coverimage']
              sheet.write(counter, 3, curl)
              description = soup2.find('div', {'class': MassElem[4]})
              try:
                  sheet.write(counter, 7, description.text)
              except:
                  sheet.write(counter, 7, '')
              leftm = []
              rightm = []
              suggestionsl = soup2.find_all('div', {'class': MassElem[5]})
              suggestionsr = soup2.find_all('div', {'class': MassElem[6]})
              for elem in suggestionsl:
                  leftm.append(elem.text)
              for elem in suggestionsr:
                  rightm.append(elem.text)
              for i in range(len(suggestionsr)):
                  if leftm[i] not in suggestion:
                      if 'Тип' in leftm[i] or leftm[i] == 'Предназначено для' or 'Вид' in leftm[i]:
                          continue
                      suggestion = suggestion + ' ' + leftm[i] + ' ' + rightm[i]
              sheet.write(counter,10, suggestion)
              try:
                  breand = rightm[leftm.index('Бренд')]
                  sheet.write(counter, 4, breand)
              except:
                  sheet.write(counter, 4, '')
          sheet.write(counter,9, 'Нет в наличии')
          counter += 1
      pages+=1
'''
print("save")
book.save("C:/Users/"+  username + "/Desktop/" + projectname + ".xls")