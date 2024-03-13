import pandas as pd
from pyzabbix import ZabbixAPI
import time


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


connect = Connector(URL, LOGIN, PASSWORD)

hosts = connect.zabbix.host.get(groupids=[5])
db_file = pd.read_excel('dbsMainMaxim.xlsx')

# log = open('log.txt', 'w')

success_count = 0
for i in range(0, len(hosts)):
    try:
        is_sber = False
        host_name = hosts[i]["name"]
        if host_name.find('sber') != -1:
            is_sber = True

        host = host_name.split('-')
        q = 0
        while (host[q+1].find('windows') == -1) or (host[q+1].find('linux') == -1) or (host[q+1].find('sber') == -1):
            if (host[q+1].find('windows') != -1) or (host[q+1].find('linux') != -1) or (host[q+1].find('sber') != -1):
                host.pop(1)
                break
            else:
                host[0] = '-'.join([host[q], host[q+1]])
                host.pop(1)
        q = None

        for j in range(len(db_file['hostname'])):
            
            if host[0] == db_file["hostname"][j]:
                if is_sber:
                    print(f"{host_name} Сбер пропущен")
                    continue
                else:
                    query = {"groupid": 39}, {"groupid": int(db_file["groupgoid"][j])}, {"groupid": int(db_file["groupnameid"][j])}

                connect.zabbix.host.update(hostid=hosts[i]["hostid"], groups=query)
                print(f"{bcolors.OKGREEN}[SUCCESS]{bcolors.ENDC} {host_name} updated")
                success_count = success_count + 1
                time.sleep(0.500)
                # log.write(f"[SUCCESS] {host_name} updated" + '\n')
                break
            if host[0] != db_file["hostname"][j] and j == (len(db_file['hostname'])-1):
                print(f"{bcolors.OKBLUE}[INFO]{bcolors.ENDC} {host_name} doesn't exist in the excel")
                time.sleep(0.500)
                # log.write(f"[INFO] {host_name} doesn't exist in the excel" + '\n')
    except:
        print(f"{bcolors.FAIL}[ERROR]{bcolors.ENDC} Failed to update {host_name}")
        time.sleep(0.500)
        # log.write(f"[ERROR] Failed to update {host_name}" + '\n')

print(success_count)
connect.logout()