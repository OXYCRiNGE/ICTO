import pandas as pd
import requests
from bs4 import BeautifulSoup
from pyzabbix.api import ZabbixAPI
import warnings
import tkinter as tk
import tkinter.filedialog as fd
from natsort import natsorted
from math import trunc
import os

warnings.filterwarnings("ignore", category=UserWarning, module='bs4')
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        btn_file = tk.Button(self, text="Выбрать файл",
                             command=self.choose_file)
        btn_dir = tk.Button(self, text="Выбрать папку",
                            command=self.choose_directory)
        btn_file.pack(padx=60, pady=10)
        btn_dir.pack(padx=60, pady=10)

    def choose_file(self):
        filetypes = (("Таблица excel", "*.xlsx"),
                     ("Любой", "*"))
        self.withdraw()
        self.update()
        filename = fd.askopenfilename(title="Открыть файл", initialdir="E:\Загрузки\\Реестр от ОО",
                                      filetypes=filetypes)
        return filename

    def choose_directory(self):
        self.withdraw()
        self.update()
        directory = fd.askdirectory(
            title="Открыть папку", initialdir="D:\Загрузки")
        return directory


class ZabbixController:
    def __init__(self, url, user, password):
        self.url = url
        self.user = user
        self.password = password
        self.zabbix = ZabbixAPI(server=self.url)
        self.zabbix.login(user=self.user, password=self.password)
        print("Connected to Zabbix API Version %s" % self.zabbix.api_version())

    def logout(self):
        self.zabbix.user.logout()

    def request(self, name, params):
        try:
            return self.zabbix.do_request(name, params)
        except:
            print(f"{bcolors.WARNING}[WARNING]{bcolors.ENDC} Request failed")
            log.write(f"[WARNING] Request failed\n")
            return Exception

    def get_address(self, address):
        payload = {
            'geocode': address,
            'results': 1
        }
        r = requests.get('https://geocode-maps.yandex.ru/1.x/?apikey=APIKEY',
                         params=payload)
        soup = BeautifulSoup(r.text, 'html.parser')
        parts = soup.pos.text.split()
        parts[1], parts[0] = parts[0], parts[1]
        return parts

    def change_host(self, host, value):
        for l in range(len(value)):
            if pd.isnull(value[l]):
                value[l] = ''
            if isinstance(value[l], float):
                value[l] = int(value[l])
        value = list(map(str, value))
        query = {
            'inventory_mode': 1,
            'inventory': {  # https://www.zabbix.com/documentation/6.0/ru/manual/api/reference/host/object
                'vendor': value[0],
                'poc_1_name': value[1],
                'alias': value[2],
                'os_full': value[3],
                'hardware': value[4].strip('/r/n'),
                'serialno_b': value[5],
                'site_address_b': value[6],
                'site_address_c': value[7],
                'type': value[8],
                'date_hw_purchase': value[9],
                'notes': value[10],
                'location': host['result'][0]['inventory']['location'],
                'location_lat': host['result'][0]['inventory']['location_lat'],
                'location_lon': host['result'][0]['inventory']['location_lon']
            }
        }
        self.request('host.update', {
                     "hostid": host['result'][0]['hostid'], **query})
        
    def del_info(self, host):
        query = {
            'inventory_mode': 1,
            'inventory': {  # https://www.zabbix.com/documentation/6.0/ru/manual/api/reference/host/object
                'vendor': '',
                'poc_1_name': '',
                'alias': '',
                'os_full': '',
                'hardware': '',
                'serialno_b': '',
                'site_address_b': '',
                'site_address_c': '',
                'type': '',
                'date_hw_purchase': '',
                'notes': ''
            }
        }
        self.request('host.update', {
                     "hostid": host['result'][0]['hostid'], **query})
       
    def add_info(self, data_new, data_df, inn, organization, subsubdir_name):
        
        # hostname = input('Введите hostname: ')
        hostname = subsubdir_name
        if hostname == '':
            exit()
        linux_list = ['lin', 'лин', 'alt', 'альт', 'red', 'ред']

        count = 0
        host = dict()
        for k in range(len(data_new[data_df[0]])):
            try:
                value = []
                value.append(inn.strip())
                value.append(organization.strip())

                for i in range(1, len(data_df)):
                    str(value.append(data_new[data_df[i]][count]))
                # count = count + 1
                os = None
                if not pd.isnull(value[3]):
                    os = value[3].lower()

                if data_new[data_df[0]][count].dtype == 'float64':
                    number = trunc(data_new[data_df[0]][count])
                else:
                    number = data_new[data_df[0]][count]
                number = str(number)
            

                found_linux = False
                for item in linux_list:
                    if os.find(item) != -1:
                        found_linux = True
                        break

                if found_linux:
                    host = self.request(
                            'host.get', {"selectInventory": ["location", 'location_lat', 'location_lon'], 'filter': {'host': hostname + "-linux_" + number}})
                    if host['result'] != []:
                        self.change_host(host, value)
                        print(f"{bcolors.OKGREEN}[SUCCESS]{bcolors.ENDC} {data_new[data_df[0]][count]} Updates {host['result'][0]['name']}")
                        log.write(f"[SUCCESS] {data_new[data_df[0]][count]} Updates {host['result'][0]['name']}\n")
                    else:
                        print(f"{bcolors.OKBLUE}[INFO]{bcolors.ENDC} {data_new[data_df[0]][count]} Not enough linux host")
                        log.write(f"[INFO] {data_new[data_df[0]][count]} Not enough linux host\n")
                elif os.find("win") != -1 or os.find("вин") != -1:
                    host = self.request(
                            'host.get', {"selectInventory": ["location", 'location_lat', 'location_lon'], 'filter': {'host': hostname + "-windows_" + number}})
                    if host['result'] != []:
                        self.change_host(host, value)
                        print(f"{bcolors.OKGREEN}[SUCCESS]{bcolors.ENDC} {data_new[data_df[0]][count]} Updates {host['result'][0]['name']}")
                        log.write(f"[SUCCESS] {data_new[data_df[0]][count]} Updates {host['result'][0]['name']}\n")
                    else:
                        print(f"{bcolors.OKBLUE}[INFO]{bcolors.ENDC} {data_new[data_df[0]][count]} Not enough windows host")
                        log.write(f"[INFO] {data_new[data_df[0]][count]} Not enough windows host\n")
                else:
                    print(f"{bcolors.FAIL}[ERROR]{bcolors.ENDC} {data_new[data_df[0]][count]} Invalid OS name in table")
                    log.write(f"[ERROR] {data_new[data_df[0]][count]} Invalid OS name in table\n")
                count = count + 1
            except:
                print(f"{bcolors.FAIL}[ERROR]{bcolors.ENDC} {data_new[data_df[0]][count]} This row in table is empty")
                log.write(f"[ERROR] {data_new[data_df[0]][count]} This row in table is empty\n")
                count = count + 1

    def set_adsress(self, data):
        for i in range(len(data['hostname'])):
            hostname = data['hostname'][i]
            hostList = []
            hostsWindows = controller.request('host.get', {"output": ["host"], "sortfield": "host",  "search": {
                "host": hostname + "-windows"
            }})
            for j in range(len(hostsWindows['result'])):
                hostList.append(hostsWindows['result'][j]['host'])
            hostsLinux = controller.request('host.get', {"output": ["host"], "sortfield": "host",  "search": {
                "host": hostname + "-linux"
            }})
            for h in range(len(hostsLinux['result'])):
                hostList.append(hostsLinux['result'][h]['host'])
            if len(hostList) > 0:
                print(f"{bcolors.OKBLUE}[INFO]{bcolors.ENDC} Address: {bcolors.OKBLUE}{data['address'][i]}{bcolors.ENDC}")
            else: 
                print(f"{bcolors.FAIL}[ERROR]{bcolors.ENDC} Hostname {bcolors.BOLD}{data['hostname'][i]}{bcolors.ENDC} dosn't exist")
                continue
            for k in range(0, len(hostList)):
                try:
                    if pd.isnull(data['inn'][i]):
                        inn = ""
                    else:
                        inn = trunc(data['inn'][i])
                    query = {
                        'inventory_mode': 1,
                        'inventory': {  # https://www.zabbix.com/documentation/6.0/ru/manual/api/reference/host/object
                            'location': data['address'][i],
                            'location_lat': data['latitude'][i], #ширина
                            'location_lon': data['longitude'][i],  #долгота
                            'vendor': inn
                        }
                    }
                    host = self.request(
                        'host.get', {'filter': {'host': hostList[k]}})
                    self.request('host.update', {
                                "hostid": host['result'][0]['hostid'], **query})
                    print(f"{bcolors.OKGREEN}[SUCCESS]{bcolors.ENDC} Updated {hostList[k]}")
                except:
                    print(f"{bcolors.FAIL}[ERROR]{bcolors.ENDC} Failed to update {hostList[k]}")


def get_subdirectories(path):
    subdirectories = []
    for entry in os.scandir(path):
        if entry.is_dir():
            subdirectories.append(entry.path)
    return subdirectories

def add_reestr_files_in_directory(controller):
    app = App()
    app.attributes("-topmost",True)

    path = app.choose_directory()
    
    app.destroy()
    for filename in os.listdir(path):
        # проверяем, что это файл
        if os.path.isfile(os.path.join(path, filename)):
            # отделяем имя файла от его расширения
            base_name, _ = os.path.splitext(filename)
            print(base_name)
            try:
                data_old = pd.read_excel(data_old, skiprows=0)
            except Exception as invalid_table:
                print(f'Неверный формат таблицы: {invalid_table}')
                log.write(f'Неверный формат таблицы: {invalid_table}\n')
                continue
            data_df = pd.DataFrame(data_old)
            data_df = data_df.columns.tolist()
            try:
                inn = data_df[0].split(':')
                organization = data_old[data_df[0]][0].split(':')
            except Exception as missing_colon:
                print(f'Отсутствует ":" в ИИН или ОО: {missing_colon}')
                log.write(f'Отсутствует ":" в ИИН или ОО: {missing_colon}\n')
                continue
            del data_old
            data_df = None
            data_new = pd.read_excel(data_new, skiprows=2)
            data_df = pd.DataFrame(data_new)
            data_df = data_df.columns.tolist()
            # Начало удаления инвентаризации для всех хостов школы
            hosts_list = []
            hostsWindows = controller.request('host.get', {"output": ["host"], "sortfield": "host",  "search": {
                "host": base_name + "-windows"
            }})
            for i in range(len(hostsWindows['result'])):
                hosts_list.append(hostsWindows['result'][i]['host'])
            hostsLinux = controller.request('host.get', {"output": ["host"], "sortfield": "host",  "search": {
                "host": base_name + "-linux"
            }})
            for i in range(len(hostsLinux['result'])):
                hosts_list.append(hostsLinux['result'][i]['host'])
                          
            host = None
            for i in range(len(hosts_list)):
                host = controller.request(
                            'host.get', {"selectInventory": ["location", 'location_lat', 'location_lon'], 'filter': {'host': hosts_list[i]}})
                controller.del_info(host)
            print(f"{bcolors.BOLD}[INFO] {base_name} cleared{bcolors.ENDC}")
            log.write(f"[INFO] {base_name} cleared\n")
            del hostsWindows, hostsLinux, hosts_list, host
            # Конец
            # time.sleep(500)
            controller.add_info(data_new, data_df, inn[1], organization[1], base_name)
            # time.sleep(500)

def add_reestr_folder_in_folder(controller):
    app = App()
    app.attributes("-topmost",True)

    path = app.choose_directory()

    app.destroy()

    # Замените "путь_к_папке" на фактический путь к вашей папке
    folder_path = path

    subdirectories = get_subdirectories(folder_path)

    largest_file_path = None
    largest_size = 0

    for subdir in subdirectories:
        for subsubdir in get_subdirectories(subdir):
            subsubdir_name = os.path.basename(subsubdir)
            files = os.listdir(subsubdir)
            for file in files:
                file_path = os.path.join(subsubdir, file)
                file_size = os.path.getsize(file_path)
                # Если текущий файл имеет больший размер, обновляем значения наибольшего файла
                if file_size > largest_size:
                    largest_file_path = file_path
                    largest_size = file_size
            if largest_file_path:
                # Делайте что-то с наибольшим файлом
                print("Подкаталог:", subsubdir_name)
                log.write(f"Подкаталог: {subsubdir_name}\n")
                #----------------------------------------------------
                data_old = data_new = largest_file_path
                try:
                    data_old = pd.read_excel(data_old, skiprows=0)
                except Exception as invalid_table:
                    print(f'Неверный формат таблицы: {invalid_table}')
                    log.write(f'Неверный формат таблицы: {invalid_table}\n')
                    continue
                data_df = pd.DataFrame(data_old)
                data_df = data_df.columns.tolist()
                #----------------------------------------------------
                try:
                    # Проверяем, является ли значение inn целым числом
                    if isinstance(data_df[0], int):
                        raise ValueError("ИИН должен быть строкой, а не целым числом")

                    inn = data_df[0].split(':')
                    if len(inn) != 2:
                        raise ValueError("Отсутствует ':' в ИИН")

                    organization = data_old[data_df[0]][0].split(':')
                    if len(organization) != 2:
                        raise ValueError("Отсутствует ':' в ОО")

                except ValueError as error_message:
                    print(f"{error_message}")
                    log.write(f"{error_message}\n")
                    continue

                del data_old
                data_df = None

                data_new = pd.read_excel(data_new, skiprows=2)
                data_df = pd.DataFrame(data_new)
                data_df = data_df.columns.tolist()

                # # Начало удаления инвентаризации для всех хостов школы
                # hosts_list = []
                # hostsWindows = controller.request('host.get', {"output": ["host"], "sortfield": "host",  "search": {
                #     "host": subsubdir_name + "-windows"
                # }})
                # for i in range(len(hostsWindows['result'])):
                #     hosts_list.append(hostsWindows['result'][i]['host'])

                # hostsLinux = controller.request('host.get', {"output": ["host"], "sortfield": "host",  "search": {
                #     "host": subsubdir_name + "-linux"
                # }})

                # for i in range(len(hostsLinux['result'])):
                #     hosts_list.append(hostsLinux['result'][i]['host'])
                              
                # host = None
                # for i in range(len(hosts_list)):
                #     host = controller.request(
                #                 'host.get', {"selectInventory": ["location", 'location_lat', 'location_lon'], 'filter': {'host': hosts_list[i]}})
                #     controller.del_info(host)
                # print(f"{bcolors.BOLD}[INFO] {subsubdir_name} cleared{bcolors.ENDC}")
                # log.write(f"[INFO] {subsubdir_name} cleared\n")
                # del hostsWindows, hostsLinux, hosts_list, host
                # # Конец
                # time.sleep(500)
                controller.add_info(data_new, data_df, inn[1], organization[1], subsubdir_name)
                # time.sleep(500)
            # Сбрасываем значения для следующей итерации
            largest_file_path = None
            largest_size = 0

def add_adress(controller):
    data = pd.read_excel(r'D:\\ИЦТО\\updateAllHosts\\хостнейм-координаты000 (2).xlsx', sheet_name='местоположение')
    controller.set_adsress(data)


if __name__ == "__main__":
    # -------------------------------- РАСКОМЕНТИТЬ -------------------------------------------------
    log = open('log.txt', 'a+', encoding='utf-8')
    controller = ZabbixController(URL, LOGIN, PASSWORD)
    add_reestr_folder_in_folder(controller)
    # add_reestr_files_in_directory(controller)
    log.close()
    controller.logout()
    # -----------------------------------------------------------------------------------------------

    # controller = ZabbixController(URL, LOGIN, PASSWORD)
    # add_adress(controller)
    # controller.logout()