from openpyxl import Workbook
import openpyxl as op
from tqdm import trange  # для отображения прогресса в консоли
from progress.bar import IncrementalBar  # для отображения прогресса в консоли


def read_excel(filename: str):  # функция принимающая имя эксель таблицы
    wb = op.load_workbook(filename, data_only=True)
    sheet = wb.active
    
    information = []  # список информации о студентах
    for i in range(1, sheet.max_column + 1): information.append(sheet.cell(row=1, column=i).value)
    for student_number in trange(2, sheet.max_row + 1):  # бежим по каждой строчке
        column_names = []  # список имён столбцов
        column_students = []  # список значений студента в этих столбцах
        for i in range(1, sheet.max_column + 1):  # бежим по каждому столбцу
            column_students.append(sheet.cell(row=student_number, column=i).value)  # добавляем в список значений студента текущее значение студента
        information.append(dict(zip(column_names, column_students)))  # создаем из списка имён столбцов и список значений студента словарик  и добавляем каждый словарик в список information
    return information  # вовзращаем список из словарей

def del_None(information):
    tmp=[]
    for elem in information:
        tmp.append({key:val for key,val in elem.items() if val != None  and val!=''})
    return tmp

def sum(list1, list2):  # функция сложения двух списков из двух эксель таблиц
    list1=uniq_list(list1)
    list2=uniq_list(list2)
    
    bar = IncrementalBar( "Countdown", max=len(list1) )  # для отображения прогресса в консоли
    tmp2=list2
    for elem1 in list1:  # для каждого элемента из первой таблицы
        for elem2 in list2:  # для каждого элемента из второйтаблицы
            # теперь проверяем есть ли среди двух таблиц повторяющийся студент
            if ( elem1.get("Дата рождения").lower() == elem2.get("Дата рождения").lower()):  # сначала проверяем равны ли даты рождения,
                if (elem1.get("Фамилия").lower() == elem2.get("Фамилия").lower()):  # если даты равны то затем фамилии,
                    if elem1.get("Отчество").lower() == elem2.get("Отчество").lower():
                        if ( elem1.get("Имя").lower() == elem2.get("Имя").lower()):  # ну если и имена равны то это явно один и тот же человек
                            value = set(elem2) - set(elem1)  # смотрим разницу между словарями. получитс список ключей второго словаря, которых нет в первом
                            for new_key in list(value):  # для каждого такого ключа добавляем занчение  в первый словарь
                                elem1[new_key] = elem2.get(new_key)
                            tmp2.remove(elem2) #удаляем значения одинаковых студентов из копии списка2
        bar.next()  # для отображения прогресса в консоли

    return list1+tmp2  # возвращаем дополненный первый список


def uniq_list(information):
    list1=[]
    for elem in information:
        tmp={key:val for key,val in elem.items() if val != None  and val!=''}
        if (len(tmp))!=0:list1.append(tmp) # удаляем значения у словарей где None 

    for i in range (len(list1)-1):
        for j in range (1,len(list1)-1):
            if list1[i].get("Дата рождения")==list1[j].get("Дата рождения"):
                if list1[i].get("Имя")==list1[j].get("Имя"):
                    if list1[i].get("Фамилия")== list1[j].get("Фамилия"):
                        if list1[i].get("Отчество")== list1[j].get("Отчество"):
                            if len(list1[i])<=len(list1[j]):
                                value = set(list1[j]) - set(list1[i])  # смотрим разницу между словарями. получитс список ключей второго словаря, которых нет в первом
                                for new_key in list(value):  # для каждого такого ключа добавляем занчение  в первый словарь
                                    list1[i][new_key] = list1[j].get(new_key)
                            list1.pop(j)

    return list1


def read_officer(filename: str): # 44 столбца заранее известных

    information = []  # список информации о студентах
    column_names = [
        "ФГОО ВО в котором обучается студент",#1
        "ФГОО ВО при которой создан ВУЦ",#2
        "ВУС",#3
        "ОВУ",#4
        "ФГОС",#5
        "Программа военной подготовки",#6
        "Год зачисления в ВУЗ",#7
        "Год начала обучения в ВУЦ",#8
        "Месяц нала обучения в ВУЦ",#9
        "Срок обучения (месяцев)",#10
        "Год окончания ВУЦ",#11
        "Месяц окончания обучения в ВУЦ",#12
        "Отметка о завершении военной подготовки",#13
        "Год окончания ВУЗа",#14
        "Месяц окончания обучения в ВУЗе",#15
        "Личный номер",#16
        "Фамилия",#17
        "Имя",#18
        "Отчество",#19
        "Дата рождения",#20
        "Место рождения",#21
        "Национальность",#22
        "Пол",#23
        "СНИЛС",#24
        "ИНН",#25
        "серия1",#26
        "серия2",#27
        "Номер",#28
        "Выдан",#29
        "Код подразделения",#30
        "Кем выдан",#31
        "Адрес регистрации",#32
        "Номер телефона",#33
        "Семейное положение",#34
        "Количество детей",#35
        "Военный комиссариат (по месту жительства)",#36
        "Военный комиссариат (в который будет направлено личное дело)",#37
        "По призыву",#38
        "По контракту",#39
        "Примечание",#40
        "Статус",	#41
        "Приказ о зачислении",	#42
        "Приказ о отчислении",	#43
        "Причина отчисления"    #44																							
    ]
    
    wb = op.load_workbook(filename, data_only=True)
    sheet = wb.active
   
    
    for student_number in trange (5,sheet.max_row + 1):
        column_students = []  # список значений студента в этих столбцах
        for i in range(1,45):column_students.append(sheet.cell(row=student_number, column=i).value )
        information.append(dict(zip(column_names, column_students)))  # создаем из списка имён столбцов и список значений студента словарик  и добавляем каждый словарик в список information  
    return information

if __name__ == "__main__":
     #read_officer читает
    #СОЕДИНЯЕТ содинение+ проверка на уникальность 
    #записывает в файл

    # list1 = read_excel("Солдаты их БД ВУЦ.xlsx")
    # print(list1)
     with open("tmp.txt", mode="w+") as file:
        tmp=read_officer('офиц.xlsx')
        
    

    # list1=uniq_list(read_excel( 'офиц.xlsx'))
    # with open("tmp.txt", mode="w+") as file:
    #     file.write(str(list1))

    #print(list1)
    # list1=[{"Фамилия":'Морозов',"Имя":"Сергей","Отчество":"Сергеевич","Дата рождения":"06-03-2003","ср балл":"1222222222222222222",}]
    # list2=[{"Фамилия":'Морозов',"Имя":"Сергей","Отчество":"Сергеевич","Дата рождения":"06-03-2003","Оценка":"5",},{"Фамилия":'Морозов',"Имя":"Сергей","Отчество":"Сергеевич","Дата рождения":"06-03-2003","Оценка":"5","ср балл":"1222222222222222222"}]
    # print(uniq_list(list2))