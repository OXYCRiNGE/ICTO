import requests
from bs4 import BeautifulSoup
import speedtest
from datetime import datetime
import os.path
import time


path = os.path.abspath(os.curdir)
check_file = os.path.exists(f'{path}/parse.txt')

if not check_file:
    my_file = open("parse.txt", "w+")



# Перевод байт в мегабайты
def humansize(nbytes):
    suffixes = ['B', 'KB', 'MB', 'GB', 'TB', 'PB']
    i = 0
    while nbytes >= 1024 and i < len(suffixes)-1:
        nbytes /= 1024.
        i += 1
    f = ('%.2f' % nbytes).rstrip('0').rstrip('.')
    return '%s %s' % (f, suffixes[i])


print(f'Путь к отчету: {path}\parse.txt')
while True:
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    # print(current_time)

    try:
        st = speedtest.Speedtest()
        speed = st.download()
        spped_h = humansize(speed)
        # print(spped_h)
    
    
        link = "https://sch5-schel.edumsko.ru"

        second_to_connect = requests.get(f"{link}/food").elapsed.total_seconds()
        # print(second_to_connect)

        responce = requests.get(f"{link}/food").text
        soup = BeautifulSoup(responce, 'lxml')
        block = soup.find('div', class_ = 'ListItems')
        xlsx = block.find('a').get('href')

        link_to_download = f"{link}{xlsx}"
        print(xlsx)
        name = xlsx.lstrip('/food/')
        print(name)

        res = requests.get(link_to_download)
        # print(res.status_code)
        size_of_xlsx = int(res.headers['content-length'])
        # print(humansize(size_of_xlsx))

        res = size_of_xlsx/speed

        # print(f"{res} секунд")

        my_file = open("parse.txt", "a+", encoding="utf-8")
        my_file.write(f"Время запуска: {current_time}, скорость загрузки: {spped_h}, время загрузки файла: {res} секунд \n")
        my_file.close()
        print(f'{current_time} Данные записаны.')

    except Exception as ex:
        my_file = open("parse.txt", "a+", encoding="utf-8")
        my_file.write(f"{current_time} Не удалось выполнить подключение к сайту или отсутствует подключение к интернету. \n")
        my_file.close()
        print(f'{current_time} Не удалось выполнить подключение к сайту или отсутствует подключение к интернету.')
        
    print("Обновление данных через 5 минут")
    time.sleep(300)