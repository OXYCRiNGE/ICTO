# import cv2
import pytesseract
# from pytesseract import Output
from PIL import Image
from pdf2image import pdfinfo_from_path, convert_from_path
import os, os.path
import pandas as pd
import time


def delete_files(dirname): #удалить файлы из папки
    lst = os.listdir(dirname)
    for l in lst:
        os.remove(f'{dirname}/{l}')

def find_count_files(dirname): #найти коли-во файлов в папке
    lst = os.listdir(dirname)
    number_files = len(lst)
    return number_files

def collect_files(dirname): # вернуть коллекцию всех файлов из папки
    lst = os.listdir(dirname)
    return lst

def rename_files(dirname, new_name): #переименовать файлы, тк convert_from_path не воспринимает кириллицу
    i = 1
    lst = os.listdir(dirname)
    for l in lst:
        os.rename(f'{dirname}/{l}', f'{dirname}/{new_name}{i}.pdf')
        i += 1
    lst = os.listdir(dirname)

def remove_transparency(im, bg_colour=(255, 255, 255)): # убрать альфа-канал

    # Only process if image has transparency (http://stackoverflow.com/a/1963146)
    if im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info):

        # Need to convert to RGBA if LA format due to a bug in PIL (http://stackoverflow.com/a/1963146)
        alpha = im.convert('RGBA').split()[-1]

        # Create a new background image of our matt color.
        # Must be RGBA because paste requires both images have the same format
        # (http://stackoverflow.com/a/8720632  and  http://stackoverflow.com/a/9459208)
        bg = Image.new("RGBA", im.size, bg_colour + (255,))
        bg.paste(im, mask=alpha)
        return bg

    else:
        return im


pytesseract.pytesseract.tesseract_cmd = r'C:\\Tesseract-OCR\\tesseract.exe'
tessdata_dir_config = r'--tessdata-dir "C:\\Tesseract-OCR\\tessdata"'

dirname = 'E:\ицто документы\ИКТ_компетентность учителя'
new_name = 'IKTcompetencies'
print(f'Started at: {time.ctime()}')
# rename_files(dirname, new_name)
numlist = [6, 9, 12, 14, 16, 20, 21, 22, 24, 45]
# for i in range(0, find_count_files(dirname)):
for i in range(0, len(numlist)):
# i = 5

    # list_pdf = collect_files(dirname)
    
    # path_to_pdf = f'{dirname}\{list_pdf[i]}'
    path_to_pdf = f'{dirname}\{new_name}{numlist[i]}.pdf'
    # number = list_pdf[i][15:-4]
    number = f'{new_name}{numlist[i]}.pdf'[15:-4]
    delete_files('img')
        # pages = convert_from_path(f'{dirname}\{list_pdf[i-1]}', 500,  poppler_path=r'C:\\poppler-23.01.0\\Library\bin') #Не читает файлы с русским названием
    # dirname = 'E:\ицто документы\Цифровые технологии в помощь учителю'
    # new_name = 'helpToTeachers'
    # for i in range(find_count_files(dirname)):
    # rename_files(dirname, new_name)
    # list_pdf = collect_files(dirname)
    # pages = convert_from_path(pdf_file, 500,  poppler_path=r'C:\\poppler-23.01.0\\Library\bin') #Не читает файлы с русским названием
    # pages = convert_from_path(list_pdf[i] 'E:\ицто документы\Цифровые технологии в помощь учителю', 500,  poppler_path=r'C:\\poppler-23.01.0\\Library\bin') #Не читает файлы с русским названием
    # ///////////////////////////////////////////////////////////////
    info = pdfinfo_from_path(path_to_pdf, userpw=None, poppler_path=r'C:\\poppler-23.01.0\\Library\bin')
    maxPages = info["Pages"]
    for page in range(0, maxPages):
        pages = convert_from_path(path_to_pdf, dpi=500, first_page=page, last_page=page+1, poppler_path=r'C:\\poppler-23.01.0\\Library\bin')
        if page != 0:
            im_rotate = pages[1].rotate(90, expand=True)
        else:
            im_rotate = pages[0].rotate(90, expand=True)
        im_rotate.save('img/out' + str(page) + '.jpg', 'JPEG')
    # for h in range(len(pages)):
    #     pages[h].save('img/out' + str(h) + '.jpg', 'JPEG')
    red_num = []
    black_num = []
    fio = []
    for k in range(0, find_count_files('img')):
        imgcv = Image.open("img/out"+str(k)+".jpg")
        thresh = 175
        fn = lambda x : 255 if x > thresh else 0
        imgcv = imgcv.convert('L').point(fn, mode='1')
        imgcv = remove_transparency(imgcv)
        # imgcv.show()
        x_start = [980, 819, 3083]
        y_start = [1523, 2304, 548]
        x_end = [2695, 2519, 5038]
        y_end = [2003, 2674, 1138]
        # x_start = [1191, 1019, 3083]
        # y_start = [1663, 2404, 628]
        # x_end = [1895, 2019, 5038]
        # y_end = [1803, 2574, 938]
        for w in range(3):
            itog = imgcv.crop((x_start[w], y_start[w], x_end[w], y_end[w]))
            # itog.show()
            itog.save("tmp/temp"+str(w)+".jpg", "JPEG")
        for j in range(3):
            result = pytesseract.image_to_string('tmp/temp'+str(j)+'.jpg',  lang='rus')
            result = result.strip()
            result = result.replace('"', ' ')
            result = result.replace('\n', ' ')
            match j:
                case 0:
                    red_num.append(result)
                case 1:
                    black_num.append(result)
                case 2:
                    fio.append(result)
        # print(k)
    # print(len(fio))
    # print(len(red_num))
    # print(len(black_num))
    # delete_files('img')
    full_result = {'fio': fio, 'red_num': red_num, 'black_num': black_num}
    df = pd.DataFrame(full_result)
    df.to_csv(f'ispravlenie\{new_name}{number}.csv',sep=';',  header=False, index=False)
    print(f'New iteration at: {time.ctime()}')
    #     print(result)
    # print('-----------------------------')
    # result = pytesseract.image_to_string(imgcv,  lang='rus')
    # result = result.split("\n")
    # print(result)