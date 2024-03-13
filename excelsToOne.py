import pandas as pd
import tkinter as tk
import tkinter.filedialog as fd
import os
import os.path


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
        filename = fd.askopenfilename(title="Открыть файл", initialdir="C:\\",
                                      filetypes=filetypes)
        return filename

    def choose_directory(self):
        self.withdraw()
        self.update()
        directory = fd.askdirectory(
            title="Открыть папку", initialdir="C:\\")
        return directory


def collect_files(dirname):  # вернуть коллекцию всех файлов из папки
    lst = os.listdir(dirname)
    return lst


if __name__ == "__main__":
    app = App()
    app.attributes("-topmost", True)

    folder = app.choose_directory()
    excel_list = collect_files(folder)

    excel_readed_list = []

    for i in range(len(excel_list)):
        try:
            path = f'{folder}/{excel_list[i]}'
            data = pd.read_excel(path, skiprows=1)
            data.columns = data.columns.str.lower()
            excel_readed_list.append(data)
        except:
            print(f'{excel_list[i]} не может быть прочитан. Решение: Преобразуйте таблицу в .xlsx формат')

    excel_merged = pd.concat(excel_readed_list)
    writer = pd.ExcelWriter('моя школа.xlsx')
    excel_merged.to_excel(writer, sheet_name='данные', index=False, startcol=0)
    writer.close()
    input("Нажмите enter, чтобы выйти...")