#Подключение необходимых библиотек
import sys
# import bs4
import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox
from urllib.parse import urljoin
import pandas
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from tqdm import tqdm
output_dir = "output"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
def success_message(): #Функция вывода окна об успешном считываниии данных
    root_success = tk.Tk()
    root_success.withdraw()
    root_success.title("Информация")
    root_success.attributes("-topmost", True)
    msg = "Считывание данных успешно завершено!"
    messagebox.showinfo("Информация", msg)
    root_success.quit()
    root_success.destroy()
    sys.exit()
def failed_message(code): #Функция вывода окна об ошибке при считывании данных
    root_failure = tk.Tk()
    root_failure.withdraw()
    root_failure.title("Информация")
    root_failure.attributes("-topmost", True)
    msg="Считывание данных не удалось"
    messagebox.showinfo("Информация", msg+". "+f'Ошибка {str(code)}')
    root_failure.quit()
    root_failure.destroy()
    sys.exit()
def connection_failed_message(): #Функция вывода окна об отсутсвии подключения к Интернету
    root_failure = tk.Tk()
    root_failure.withdraw()
    root_failure.title("Информация")
    root_failure.attributes("-topmost", True)
    msg="Необходимо подключение к интернету"
    messagebox.showinfo("Информация", msg)
    root_failure.quit()
    root_failure.destroy()
    sys.exit()
#Распознование семестра при считывании данных
current_date= datetime.now()
autumn_start=datetime(current_date.year,9,1)
autumn_end=datetime(current_date.year+1,1,31)
spring_start=datetime(current_date.year,2,1)
spring_end=datetime(current_date.year,8,31)
semestr=""
if autumn_start<=current_date<= autumn_end:
    semestr="Осенний семестр"
elif spring_start<=current_date<=spring_end:
    semestr="Весенний семестр"
url= "https://institutes.mai.ru/control/staff-305/?layout=table&index=0"
timeout = 1
#Проверка на наличие интернет-соединения
try:
    requests.head(url, timeout=timeout)
except requests.ConnectionError:
    connection_failed_message()
response = requests.get(url, headers={'User-Agent': UserAgent().chrome})
print('Для успешного считывания данных с сайта необходимо стабильное Интернет-соединение \n Не закрывайте это окно до конца считывания!')
if response.status_code!= 200: #Проверка на создание успешного запроса
    failed_message(response.status_code)
else:
    bs = BeautifulSoup(response.text,"lxml")
    table=bs.find("table")
    rows=table.find_all("tr")
    links=[]
    for row in rows:
        cols=row.find_all("td")
        for col in cols:
            link=col.find('a')
            if link and "href" in link.attrs:
                href=link["href"]
                full_url = urljoin (url, href)
                links.append(full_url) #Сбор гипер-ссылок по преподавателям
    total_links=len(links)
    progressbar=tqdm(total=total_links, desc="Идет считывание данных с сайта", file=sys.stdout, unit="link") #Создание шкалы прогресса
    for link in links: #Цикл считывания основной информации о преподавателе
        url2 = link + f'&week={1}'
        response2 = requests.get(url2,headers={'User-Agent': UserAgent().chrome})
        #print(response2.status_code)
        if response2.status_code !=200: #Проверка на создание успешного запроса
            failed_message(response2.status_code)
        bs2 = BeautifulSoup(response2.text, "lxml")
        name = bs2.find("h1")
        prepod = name.text.strip()
        #print(name.text.replace("\t", "") + ":")
        vse = []
        counts = {}
        neds=bs2.find("div", class_="row pb-5").find_all("span", class_="badge bg-soft-dark text-dark ms-2")
        for p in range(1, len(neds)+1): #Цикл, собирающий информацию о дисциплинах понедельно
            url1=link+f'&week={p}'
            response1 = requests.get(url1,headers={'User-Agent': UserAgent().chrome})
            if response1.status_code != 200: #Проверка на создание успешного запроса
                failed_message(response1.status_code)
            disc = []
            #print(response1.status_code)
            bs1 = BeautifulSoup(response1.text, "lxml")
            try:
                strems = bs1.find("ul", class_="step mb-5").find_all("div", class_="mb-4")
            except:
                continue
            for strem in strems:
                es = strem.find_all('p', class_="mb-2 fw-semi-bold text-dark")
                for e in es:
                    e = e.text.strip().replace("\n", "").replace("\t", "").replace("ЛК", ";ЛК;").replace("ПЗ",
                                                                                                        ";ПЗ;").replace("ЛР",
                                                                                                                       ";ЛР;")
                    res = strem.find('ul', class_="list-inline list-separator text-body small").find_all('li',
                                                                                                        class_="list-inline-item")
                    groupa = []
                    for re in res:
                        tys = re.find_all("a", class_="text-body")
                        for ty in tys:
                            groupa.append(ty.text)
                    vse.append(e + " " + " ".join(map(str, groupa)))
            alsofinaly = []
            for i in vse:
                if i not in alsofinaly: #Создание словаря с названиями дисциплин и суммарным количеством часов
                    alsofinaly.append(i)
                    counts[i] = 2
                else:
                    counts[i] += 2
        results = [(key+';', str(value)+';') for key, value in counts.items()] #Создание массива для записи дисциплин и часов
        results.sort()
        exel = {
            "Дисциплина" : [],
            "Тип занятия" : [],
            "Группы" : [],
            'Часы' : []
        }
        for result in results: #Цикл, записывающий название дисциплин, количество часов, тип занятий и номера групп, а также вывод информации в формате .xlsx
            stroka=result[0]
            hour=result[1]
            stroka=stroka.split(";")
            exel['Дисциплина'].append((stroka[0]))
            exel['Тип занятия'].append((stroka[1]))
            exel['Группы'].append((stroka[2]))
            exel["Часы"].append(hour)
        df=pandas.DataFrame(exel)
        file_path = os.path.join(output_dir, f'{prepod},{semestr}.xlsx')
        df.to_excel(file_path, index=False)
        progressbar.update(1) #Обновление шкалы прогерсса
success_message() # Вывод окна об успешном считывании данных