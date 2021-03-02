from tkinter import *
from tkinter import ttk
from tkinter import messagebox as mb
from bs4 import BeautifulSoup
from selenium import webdriver
from tkinter import filedialog as fd
import pandas as pd
import random
import time

#######################
#######################
#Блок интрерфейса
root = Tk()
root.title = 'ФГИС parser'
root.geometry('1024x600+40+40')

tab_control = ttk.Notebook(root)
tab_control.pack(expand=1, fill='both')

# создаем саму вкладку
tab_1 = ttk.Frame(tab_control)
# называем первую вкладку
tab_control.add(tab_1, text='Ручной поиск')

# создаем саму вкладку
tab_2 = ttk.Frame(tab_control)
# называем первую вкладку
tab_control.add(tab_2, text='Автоматический поиск')

# Текстовое окно для вкладки ручного поиска
text = Text(tab_1)
text.place(x=5, y=5, height=520, width=650)

# Текстовое окно для вкладки автоматического поиска
text2 = Text(tab_2)
text2.place(x=5, y=5, height=450, width=650)

label_org_of_verify = ttk.Label(
    tab_1, text='Организация производившая поверку')
label_org_of_verify.place(x=670, y=10)

entry_org_of_verify = ttk.Entry(tab_1, width=30)
entry_org_of_verify.place(x=670, y=40)
entry_org_of_verify.insert(0, str('ФБУ "ТЮМЕНСКИЙ ЦСМ"'))

label_name_of_equip = ttk.Label(tab_1, text='Наименование типа СИ')
label_name_of_equip.place(x=670, y=90)

entry_name_of_equip = ttk.Entry(tab_1, width=30)
entry_name_of_equip.place(x=670, y=120)

label_number_of_equip = ttk.Label(tab_1, text='Заводской номер СИ')
label_number_of_equip.place(x=670, y=170)

entry_number_of_equip = ttk.Entry(tab_1, width=30)
entry_number_of_equip.place(x=670, y=200)


########################################
########################################
# запуск selenium и движка Google Chrome
# option = webdriver.ChromeOptions()
# option.add_argument('headless')
# driver = webdriver.Chrome(options=option)
# # переменая для функции search
results = ''
# # переменая для Класса Parser
date_2 = []


#########################################
#########################################
#Логика
class Parser():
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    driver = webdriver.Chrome(options=option)
    raw_html = ''
    html = ''
    results = []

    def __init__(self, url):
        self.url = url

    def get_html(self):
        try:
            self.driver.get(self.url)
        except:
            mb.showinfo("Внимание", "Похоже сайт как всегда не работает :(")
            
    def pre_parse(self):
        self.html = BeautifulSoup(self.driver.page_source, 'html.parser')

    def parse(self):
        self.pre_parse()
        try:
            if self.html.find('div', {"class": "spinner-container"}):
                # print('spinner')
                time.sleep(random.randint(3,5))
                self.parse()
            else:
                table_body = self.html.find('tbody')
                rows = table_body.find_all('tr')
                for row in rows:
                    cols = row.find_all('td')
                    cols = [ele.text.strip() for ele in cols]
                    date_2.append([ele for ele in cols if ele])
                # print(date_2)
        except AttributeError:
            # print('AttributeError')
            time.sleep(random.randint(3,5))
            self.parse()


    def run(self):
        self.get_html()
        self.parse()
        # self.driver.close()
        return date_2

    def run_2(self):
        self.get_html()
        self.parse()
        return date_2


def search():
    equip_edited = entry_org_of_verify.get().replace(
        '"', '%22').replace('"', '%22').replace(' ', '%20')
    param = {'organization': equip_edited, 'name_of_equip': entry_name_of_equip.get(
    ), 'number_of_equip': entry_number_of_equip.get()}
    # URL = 'https://fgis.gost.ru/fundmetrology/cm/results?filter_org_title={0}&filter_mi_mititle={1}&filter_mi_number={2}&activeYear=2020'.format(
    #     param['organization'], param['name_of_equip'], param['number_of_equip'])
    URL = 'https://fgis.gost.ru/fundmetrology/cm/results?filter_org_title={0}&filter_mi_mititle={1}&filter_mi_number={2}'.format(
        param['organization'], param['name_of_equip'], param['number_of_equip'])
    # print(URL)
    data = Parser(URL)
    results = data.run()
    # print(results)
    if results!=[]:
        for item in range(len(results)):
            if results[item][0] == entry_org_of_verify.get():
                # print('Поверка от:' + results[item][6] + " ДО:" + results[item][7])
                line_of_results = '\n'+entry_org_of_verify.get() + " " + results[item][2] + " " + results[item][3] + " " + "\nЗаводской номер: " + results[item][5] + \
                    '\nДата поверки: ' + results[item][6] + "\nПоверка действительна до: " + \
                    results[item][7] + "\nНомер свидетельства: " + \
                    str(results[item][8]) + '\n'
                text.insert(END, line_of_results)
    else:
        line = 'нет данных для '+ param['name_of_equip'] +' '+param['number_of_equip']+'\n'
        text.insert(END, line)
    date_2.clear()
    # data.driver.close()




def open():
    try:
        file_name = fd.askopenfilename()
        data = pd.read_excel(file_name)
        # text2.insert(1.0,data)
        global data_1
        data_1 = data
    except FileNotFoundError or AssertionError:
        mb.showinfo("Внимание", "Файл не загружен")


def search_from_excel():
    res_list = []
    data = data_1
    org = data['организация'].tolist()
    name = data['наименование си'].tolist()
    number = data['заводской си'].tolist()
    for i in range(len(org)):
        semi_res = []
        semi_res.append(org[i])
        semi_res.append(str(name[i]))
        semi_res.append(str(number[i]))
        res_list.append(semi_res)
    # print(res_list)
    part_for_progress=100/(len(res_list))
    if my_progress['value'] !=0:
        my_progress['value']=0
    for item in range(len(res_list)):
        # print(res_list[item][0])
        URL = 'https://fgis.gost.ru/fundmetrology/cm/results?filter_org_title={0}&filter_mi_mititle={1}&filter_mi_number={2}'.format(
            res_list[item][0], res_list[item][1], res_list[item][2])
        data = Parser(URL)
        results = data.run_2()
        # results = data.run()
        my_progress['value'] +=part_for_progress
        root.update()
        # print(results)
        if results!=[]:
            for item in range(len(results)):
                if results[item][0] == res_list[item][0]:
                    # print('Поверка от:' + results[item]
                    #       [6] + " ДО:" + results[item][7])
                    line_of_results = '\n'+results[item][0] + " " + results[item][2] + " " + results[item][3] + " " + "\nЗаводской номер: " + results[item][5] + \
                        '\nДата поверки: ' + results[item][6] + "\nПоверка действительна до: " + \
                        results[item][7] + "\nНомер свидетельства: " + \
                        str(results[item][8]) + '\n'
                    text2.insert(END, line_of_results)
            root.update()
        else:
            line = '\n'+'нет данных для '+ res_list[item][1] +' '+res_list[item][2]+' '+res_list[item][0]+'\n'
            text2.insert(END, line)
            continue
        date_2.clear()
    # data.driver.close()


#######################
#######################
#Блок интрерфейса

#Кнопка поиска и загрузки файла в приложение
bnt_download = ttk.Button(tab_2, text='Загрузка файла', width=20, command=open)
bnt_download.place(x=670, y=10)

my_progress=ttk.Progressbar(tab_2, orient=HORIZONTAL, length=650, mode='determinate')
my_progress.place(x=5,y=470)

#Кнопка поиска ручного
btn_search = Button(tab_1, text='Найти', command=search)
btn_search.place(x=670, y=250, height=50, width=150)


# Кнопка поиска и загрузки файла в автоматическом режиме
btn_convert = Button(tab_2, text='Найти', command=search_from_excel)
btn_convert.place(x=670, y=60, height=50, width=150)


root.mainloop()
