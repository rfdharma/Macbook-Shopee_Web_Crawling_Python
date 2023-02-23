from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time

opsi = webdriver.ChromeOptions()
opsi.add_argument('--headless')
servis = Service('chromedriver.exe')
driver = webdriver.Chrome(service=servis, options=opsi)

shopee_link = "https://shopee.co.id/search?keyword=macbook"
driver.set_window_size(1300, 800)
driver.get(shopee_link)

rentang = 500
for i in range(1, 8):
    akhir = rentang * i
    perintah = "window.scrollTo(0,"+str(akhir)+")"
    driver.execute_script(perintah)
    print("loading ke-"+str(i))
    time.sleep(1)

time.sleep(5)

contents = driver.page_source
driver.quit()

data = BeautifulSoup(contents, "html.parser")
# print(data.encode("utf-8"))
list_nama, list_harga, list_link, list_terjual, list_lokasi = [], [], [], [], []
i = 1
base_url = 'https://shopee.co.id'
for area in data.find_all('div', class_="col-xs-2-4 shopee-search-item-result__item"):
    print(i)
    nama = area.find('div', class_="ie3A+n bM+7UW Cve6sh").get_text()
    harga = area.find('span', class_="ZEgDH9").get_text()
    link = base_url + area.find('a')['href']
    terjual = area.find('div', class_="r6HknA uEPGHT")
    lokasi = area.find('div', class_="zGGwiV").get_text()
    if terjual != None:
        terjual = terjual.get_text()
    list_nama.append(nama)
    list_harga.append(harga)
    list_link.append(link)
    list_terjual.append(terjual)
    list_lokasi.append(lokasi)
    i += 1
    print("-----------------")

df = pd.DataFrame({
    'Nama': list_nama,
    'Harga': list_harga,
    "Link": list_link,
    "Terjual": list_terjual
})

writer = pd.ExcelWriter("Macbook_Shoope.xlsx")
df.to_excel(writer, 'Sheet1', index=False)
writer.save()
