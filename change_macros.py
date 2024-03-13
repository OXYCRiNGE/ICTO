import pandas as pd
from pyzabbix import ZabbixAPI


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

class Connector():

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
        except Exception as e:
            print(f"{bcolors.FAIL}[ERROR]{bcolors.ENDC} Request failed")
            log.write(f"[ERROR] Request failed\n")
            raise e

if __name__ == "__main__":
    log = open(r'D:\\ИЦТО\\macros\\log.txt', 'a+', encoding='utf-8')
    connect = Connector(URL, LOGIN, PASSWORD)

    data = pd.read_excel(r'D:\\ИЦТО\\macros\\ввод данных по ЦОС в заббикс.xlsx')

    host_list = ["windows", "linux", "sber", "laptop", "desktop", "workstation"]
    for i in range(len(data)):
        try:
            groupid = connect.request('hostgroup.get', {'output':['groupid'], 'filter': {'name':[data['name'][i]]}})

            # print(i)
            print(f"GROUP: {data['name'][i]}")
            log.write(f"GROUP: {data['name'][i]}\n")

            if len(groupid['result']) == 0 or len(groupid['result'][0]) == 0:
                log.write(f"[ERROR] Нет такой группы\n")
                raise ValueError(f'{bcolors.FAIL}[ERROR]{bcolors.ENDC} Нет такой группы')

            host = connect.request('host.get', {'output':['hostid', 'name'], 'groupids':groupid['result'][0], "search": {"host": host_list}, "excludeSearch": True})

            if len(host['result']) == 0:
                log.write(f"[ERROR] Нет нужного хоста\n")
                raise ValueError(f'{bcolors.FAIL}[ERROR]{bcolors.ENDC} Нет нужного хоста')
            print(host['result'][0]['hostid'])

            a = connect.request('usermacro.create', {'hostid':host['result'][0]['hostid'], "macro":"{$COS_IN_GO}", "value": str(data['total'][i])})
            print(f"{bcolors.OKGREEN}[SUCCESS]{bcolors.ENDC} {host['result'][0]['name']} macros create")
            log.write(f"[SUCCESS] {host['result'][0]['name']} macros create\n")

        except ValueError as v:
            print(v)
        except Exception as e:  # Ловим общее исключение
            print(f"{bcolors.FAIL}[ERROR] Request Exception: {e}{bcolors.ENDC}")
            log.write(f"[ERROR] Request Exception: {e}\n")
            
    log.close()
    connect.logout()