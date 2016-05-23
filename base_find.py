"""
Что хотим??
Подготовка данных:
1) Программа скачивает таблицы по списку адресов, в зафиксированную папку на компьютере.
2) Запрашивается Школа-Фамилия-Имя-Очество человека
3) Записывается список дата-адрес-экзамен

"""
""" Основная часть"""
"""Задачи
6) Замена регулярного выражения на функцию из библиотеки -- no done, nothing
Запись:
7) Правильная ширина колонок в итоговой таблице
Слив двух таблиц
8)  Сортировка таблицы имени по дате

"""


import codecs
import datetime
import requests
import xlrd
import xlwt
from time import time
import os
from Send_meil import send_mail, send

"""Никуда без декораторов!!!"""
def with_time_printing(f):
    def decorated(*args):
        start = time()
        res = f(*args)
        print("Executing %s took %d ms" % (f.__name__, (time() - start) * 1000))
        return res
    return decorated


"""Чтение страниц rckoi.Обработка кода страницы rckoi."""



def read_html_cod(url_ege,url_gia):
    ege_11 = requests.get(url_ege)
    name_11 = 'exams.txt'
    f = open(name_11,'w')
    f.write(ege_11.text)
    gia_9 = requests.get(url_gia)
    f = open(name_11, 'a')
    f.write(gia_9.text)
    f.close()
    return name_11


def list_of_link(cod_text):
    find_str = 'Список сотрудников'
    f = open(cod_text,'r')
    list_find_str = [line.strip() for line in f if find_str in line.strip()]
    f.close()
    list_true_find_str = [i[i.index(find_str):] for i in list_find_str]
    first_piece = 'http://rcoi.mcko.ru'
    list_link_as_str = [first_piece + seckond_piece.split('href="')[1][:-6] for seckond_piece in list_true_find_str]  ## всем костылям костыль
    return list_link_as_str


@with_time_printing
def save_table(link_as_str):
    name = link_as_str[-16:-1]  ##в лучших традициях костылей
    p = requests.get(link_as_str)
    out = open(name, "wb")
    out.write(p.content)
    out.close()
    return out.name


def slice_school(table_of_data,name_school):
    """срез записией относящихся к школе name_school в таблице table_of_data"""
    rb = xlrd.open_workbook(table_of_data)
    sheet = rb.sheet_by_index(0)
    all_len = sheet.nrows
    our_school = [sheet.row_values(i) for i in range(1,all_len) if name_school  in str(sheet.row_slice(i)[7])]
    return our_school


def name_exam(table_of_data):
    imp = table_of_data[-5]
    name =''
    if imp =='e':
        name = 'ЕГЭ_11'
    elif imp =='o':
        name='ОГЭ_9'
    else:
        name = 'я не знаю'
    return name


def write_slice_data_of_school(table_of_data,name_school,data_of_exam):
    font0 = xlwt.Font()  # непонятно
    font0.name = 'Times New Roman'
    font0.bold = True

    style0 = xlwt.XFStyle()  # не понятно
    style0.font = font0

    name_table_of_data_school = name_school+'_'+data_of_exam +'time_of_create'+str(time())+'.xls'
    our_school = slice_school(table_of_data,name_school)
    len_1553 = len(our_school)
    wb = xlwt.Workbook()
    ws = wb.add_sheet(data_of_exam)
    important_number_of_column = [3,4,5,6]
    len_of_header = 1
    name_of_exam = name_exam(table_of_data)
    for i in range(len_1553):
        for j in important_number_of_column:
            if j == 3:
                record = write_name_school(str(our_school[i][j]))
            else:
                record = str(our_school[i][j])
            ws.write(i+len_of_header,j-3+2,record,style0)
        ws.write(i+len_of_header,0,write_data_exam(data_of_exam),style0)
        ws.write(i+len_of_header,1,day_of_week(data_of_exam),style0)
        ws.write(i+len_of_header,len(important_number_of_column)+2,name_of_exam,style0)

    wb.save(name_table_of_data_school)
    try:
        no_empty(name_table_of_data_school)
        return name_table_of_data_school
    except FileNotFoundError:
        print(name_table_of_data_school)






def no_empty(file):
    rb = xlrd.open_workbook(file)
    sheet = rb.sheet_by_index(0)
    proverca = [sheet.row_values(1),len(sheet.row_values(1))]
#no_empty('1553_05_26time_of_create1464010588.134745.xls')
#no_empty('1553_05_26time_of_create1464010592.596.xls')


@with_time_printing
def save_and_write_all_table(cod_text):
    list_link_as_str = list_of_link(cod_text)
    delete_file(cod_text)
    list_name_table_school = []
    for link_as_str in list_link_as_str:
        name_of_table = save_table(link_as_str)
        try:
            no_empty(name_of_table)
            table_of_data = os.path.join(os.path.abspath(os.path.dirname(__file__)), name_of_table)
            name_school = '1553'
            data_of_exam = name_of_table[:5]
            list_name_table_school.append(write_slice_data_of_school(table_of_data, name_school, data_of_exam))
            delete_file(name_of_table)
        except IndexError:
            delete_file(name_of_table)
    return list_name_table_school


"""таблица под имя"""


def delete_file(name_file):
    way = os.path.join(os.path.abspath(os.path.dirname(__file__)), name_file)
    os.remove(way)


def write_name_school(name_of_school):
    very_long_name = 'Государственное бюджетное общеобразовательное учреждение города Москвы'
    if very_long_name in name_of_school:
        name_of_school = name_of_school[len(very_long_name):]
    return name_of_school


def write_data_exam(data):
    month = data[:2]
    day = data[3:5]
    return day+'.'+month


def day_of_week(data):
    data_new = datetime.date(2016, int(data[:2]),int(data[3:5]))
    week = ['понедельник','вторник','среда','четверг','пятница','суббота','воскресенье']
    ind = datetime.datetime.weekday(data_new)
    return week[ind]


def sort_array_for_data(array):
    number_of_data = 0
    data = array[number_of_data]
    data_to_comp = data[3:]+data[:2]
    return data_to_comp


def find_name_str(name_person,list_name_table_school):
    list_str_of_person = []
    for table_of_data in list_name_table_school:
        try:
            no_empty(table_of_data)
            rb = xlrd.open_workbook(table_of_data)
            sheet = rb.sheet_by_index(0)
            all_len = sheet.nrows
            record = [sheet.row_values(i) for i in range(1, all_len) if name_person in str(sheet.row_slice(i)[5])]
            list_str_of_person = list_str_of_person + record
        except IndexError:
            None

    return list_str_of_person


def list_of_person_records(name_person,cod_text,list_name_table_school):
    list_of_person = find_name_str(name_person, list_name_table_school)
    list_of_person.sort(key = sort_array_for_data)
    return list_of_person


def write_person_table(name_person,list_str_of_person):
    if 'Гадас Роман' in name_person:
        print(list_str_of_person)
    try:
        len_list_person = len(list_str_of_person)
        width = len(list_str_of_person[0])
        font0 = xlwt.Font()  # непонятно
        font0.name = 'Times New Roman'
        font0.bold = True

        wb = xlwt.Workbook()
        ws = wb.add_sheet(name_person)

        len_of_header = 1
        style0 = xlwt.XFStyle()  # не понятно
        style0.font = font0

        name_table_of_person = name_person+ '.xls'
        header = ['дата','день недели','организация','адрес организации','должность','имя преподавателя','когда','тип экзамена']
        for i,element in enumerate(header):
            ws.write(0,i,element)

        for i in range(len_list_person):
            for j in range(width):
                ws.write(i + len_of_header,j,str(list_str_of_person[i][j]),style0)

        ws.col(1).width = 5000
        ws.col(2).width = 5000
        ws.col(3).width = 15000
        ws.col(4).width = 10000
        ws.col(5).width = 10000
        wb.save(name_table_of_person)
    except IndexError:
        return name_person


def write_lists_of_persons_records(list_of_persons,cod_text):
    list_name_table_school = save_and_write_all_table(cod_text)
    bad_person = []
    for name_person in list_of_persons:
        print(name_person)
        rec = list_of_person_records(name_person,cod_text,list_name_table_school)
        bad_person.append(write_person_table(name_person, rec))
    for table in list_name_table_school:
        print(table)
        delete_file(table)
    return bad_person


"""не нужно"""
def exist_person_file(name_person):
    way = os.path.join(os.path.abspath(os.path.dirname(__file__)), name_person+'.xls')
    return os.path.exists(way)


def list_of_persons_and_mails(file):
    f = codecs.open(file, "r", "utf-8")
    list_per= [line.strip() for line in f if len(line.strip())!=0]
    f.close()
    persons_and_mails = [i.split(',') for i in list_per]
    persons = [i[0] for i in persons_and_mails]
    persons[0] = persons[0][1:]
    mails = [i[1] for i in persons_and_mails]
    return persons,mails


def main():
    file = 'сотрудники 1553.txt'
    list_persons,list_mails = list_of_persons_and_mails(file)
    url_ege = 'http://rcoi.mcko.ru/index.php?option=com_content&view=article&id=898&Itemid=197'
    url_gia = 'http://rcoi.mcko.ru/index.php?option=com_content&view=article&id=1033&Itemid=211'
    cod_text = read_html_cod(url_ege,url_gia)
    bad_man = write_lists_of_persons_records(list_persons, cod_text)
    bad_man=[i for i in bad_man if i is not None]
    index_bad_man = [list_persons.index(i) for i in list_persons if i in bad_man]
    list_mails = [i for i in list_mails if not list_mails.index(i) in index_bad_man ]
    list_of_files = [[i + '.xls'] for i in list_persons]
    list_mails_bad_man = [i for i in list_mails if list_mails.index(i) in index_bad_man ]
    list_mails = [[i] for i in list_mails]
    from_email = 'n.k.kruglyak@gmail.com'

    subject = ''
    for send_to, files in zip(list_mails,list_of_files):
        text = 'Ваше расписание экзаменов.'
        send_mail(from_email, send_to, subject, text, files, server='smtp.gmail.com')
    for to_email in list_mails_bad_man:
        text = 'Вы не участвуете в проведении экзаменов'
        send(text, subject, from_email, to_email, host='smtp.gmail.com')



main()

"""
day_of_week(data)
 Обработка списка имен!!!!!!!!"""
