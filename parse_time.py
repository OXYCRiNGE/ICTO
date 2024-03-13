import os
import tkinter as tk
from tkinter import filedialog as fd
import psycopg2 as pg
import pandas as pd
import datetime
from math import trunc


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        btn_file = tk.Button(self, text="Выбрать файл", command=self.choose_file)
        btn_dir = tk.Button(self, text="Выбрать папку", command=self.choose_directory)
        btn_file.pack(padx=60, pady=10)
        btn_dir.pack(padx=60, pady=10)

    def choose_file(self, msg):
        filetypes = (("Таблица excel", "*.xlsx"), ("Любой", "*"))
        self.withdraw()
        self.update()
        initial_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        filename = fd.askopenfilename(
            title=msg, initialdir=initial_dir, filetypes=filetypes
        )
        self.deiconify()  # Restore the window after file dialog is closed
        return filename

    def choose_directory(self):
        self.withdraw()
        self.update()
        initial_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        directory = fd.askdirectory(title="Открыть папку", initialdir=initial_dir)
        self.deiconify()  # Restore the window after directory dialog is closed
        return directory


def create_days_and_data(unixtime):
    days = {}

    data_list = {
        "ГО": [],
        "ГО / Наименование школы": [],
        "Хостнейм": [],
        "ИНН": [],
        "Тип закупок": [],
    }
    for i in range(len(unixtime)):
        timestamp = pd.Timestamp(unixtime["день"][i])
        date = timestamp.strftime("%d.%m.%Y")
        days[date] = [trunc(unixtime["юникстаймНач"][i]), unixtime["юникстаймКон"][i]]
        data_list[date] = []
    return [days, data_list]


def cos_hosts(data, cur, data_list):
    count = 0
    for i in range(len(data["grouid"])):
        query = f"""
        select hosts.hostid, hosts.host, hosts_groups.groupid, hstgrp.name, items.itemid, items.key_

        from items

        inner join hosts on hosts.hostid = items.hostid
        inner join hosts_groups on hosts.hostid = hosts_groups.hostid
        inner join hstgrp on hosts_groups.groupid = hstgrp.groupid
        where items.key_ = 'system.uptime' and hosts_groups.groupid = {data['grouid'][i]}
        order by hosts.hostid
        """
        cur.execute(query)
        itemid_col = cur.fetchall()
        for itemid in itemid_col:
            query = (
                f"SELECT vendor, type FROM host_inventory WHERE hostid = {itemid[0]}"
            )
            cur.execute(query)
            type_col = cur.fetchall()
            row_count = cur.rowcount
            if row_count == 0:
                continue
            else:
                for type in type_col:
                    if "цос" in type[1].lower():
                        data_list["ИНН"].append(type[0])
                        data_list["Тип закупок"].append(type[1])
                        query = f"""
                        select hosts.hostid, hosts.host, hosts_groups.groupid, hstgrp.name
                        from hosts
                        inner join hosts_groups on hosts.hostid = hosts_groups.hostid
                        inner join hstgrp on hosts_groups.groupid = hstgrp.groupid
                        where hosts.hostid = {itemid[0]}
                        """
                        cur.execute(query)
                        name_of_host_col = cur.fetchall()
                        data_list["ГО"].append(name_of_host_col[-2][3])
                        data_list["ГО / Наименование школы"].append(
                            name_of_host_col[-1][3]
                        )
                        data_list["Хостнейм"].append(name_of_host_col[-1][1])
                        for day in days:
                            query = f"""
                            SELECT max(value)
                            FROM history_uint
                            where itemid = {itemid[4]} and clock > {days[day][0]} and clock < {days[day][1]}
                            """
                            cur.execute(query)
                            max_col = cur.fetchall()
                            row_count = cur.rowcount
                            if row_count == 0:
                                data_list[day].append("")
                                continue
                            else:
                                for max in max_col:
                                    if max[0] is None:
                                        unix_time = 0
                                    else:
                                        unix_time = int(max[0])

                                    dt_object = datetime.datetime.fromtimestamp(
                                        unix_time
                                    )
                                    dt_object = dt_object + datetime.timedelta(hours=-3)
                                    formatted_time = dt_object.strftime("%H:%M:%S")
                                    data_list[day].append(formatted_time)
                                    print(count)
                                    count = count + 1

                    else:
                        continue

    return data_list


def all_hosts(data, cur, data_list):
    count = 0
    for i in range(len(data["grouid"])):
        query = f"""
        select hosts.hostid, hosts.host, hosts_groups.groupid, hstgrp.name, items.itemid, items.key_

        from items

        inner join hosts on hosts.hostid = items.hostid
        inner join hosts_groups on hosts.hostid = hosts_groups.hostid
        inner join hstgrp on hosts_groups.groupid = hstgrp.groupid
        where items.key_ = 'system.uptime' and hosts_groups.groupid = {data['grouid'][i]}
        order by hosts.hostid
        """
        cur.execute(query)
        itemid_col = cur.fetchall()
        for itemid in itemid_col:
            query = f"""
            select hosts.hostid, hosts.host, hosts_groups.groupid, hstgrp.name
            from hosts
            inner join hosts_groups on hosts.hostid = hosts_groups.hostid
            inner join hstgrp on hosts_groups.groupid = hstgrp.groupid
            where hosts.hostid = {itemid[0]}
            """
            cur.execute(query)
            name_of_host_col = cur.fetchall()
            data_list["ГО"].append(name_of_host_col[-2][3])
            data_list["ГО / Наименование школы"].append(name_of_host_col[-1][3])
            data_list["Хостнейм"].append(name_of_host_col[-1][1])

            query = (
                f"SELECT vendor, type FROM host_inventory WHERE hostid = {itemid[0]}"
            )
            cur.execute(query)
            type_col = cur.fetchall()
            row_count = cur.rowcount
            if row_count == 0:
                data_list["ИНН"].append("")
                data_list["Тип закупок"].append("")
            else:
                for type in type_col:
                    data_list["ИНН"].append(type[0])
                    data_list["Тип закупок"].append(type[1])
            for day in days:
                query = f"""
                SELECT max(value)
                FROM history_uint
                where itemid = {itemid[4]} and clock > {days[day][0]} and clock < {days[day][1]};
                """
                cur.execute(query)
                max_col = cur.fetchall()
                row_count = cur.rowcount
                if row_count == 0:
                    data_list[day].append("")
                    continue
                else:
                    for max in max_col:
                        if max[0] is None:
                            unix_time = 0
                        else:
                            unix_time = int(max[0])

                        dt_object = datetime.datetime.fromtimestamp(unix_time)
                        dt_object = dt_object + datetime.timedelta(hours=-3)
                        formatted_time = dt_object.strftime("%H:%M:%S")
                        data_list[day].append(formatted_time)
                        print(count)
                        count = count + 1
    return data_list


app = App()

msg = "Выберете файл с ГО"
print(msg)
data = app.choose_file(msg)
data = pd.read_excel(data)

msg = "Выберете файл с юникстаймами"
print(msg)
unixtime = app.choose_file(msg)

app.destroy()

unixtime = pd.read_excel(unixtime)

data_and_days = create_days_and_data(unixtime)
del unixtime

days = data_and_days[0]
data_list = data_and_days[1]
del data_and_days

name = input("Введите название для будующей таблицы (без расширения): ")

try:
    conn = pg.connect(
    database=DATABASE,
    user=USER,
    password=PASSWORD,
    host=HOST,
    port=PORT
    )
except Exception as ex:
    print(f"Не удалось подключиться к бд: {ex}")

conn.autocommit = False
cur = conn.cursor()

while True:
    choice = input(
        """
1 - Только ЦОС
2 - Все хосты

"""
    )
    if choice == "1":
        print("Вы выбрали 1. Выполняется действие 1.")
        data_list = cos_hosts(data, cur, data_list)
        break
    elif choice == "2":
        print("Вы выбрали 2. Выполняется действие 2.")
        data_list = all_hosts(data, cur, data_list)
        break


df = pd.DataFrame(data_list)
file_name = f"{name}.xlsx"
df.to_excel(file_name, index=False)
print(f"Файл {file_name} сохранен")
input("Нажмите любую кнопку, чтобы выйти...")

cur.close()
conn.close()
