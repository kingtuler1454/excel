from openpyxl import Workbook
import openpyxl as op
from tqdm import trange  # для отображения прогресса в консоли
from progress.bar import IncrementalBar  # для отображения прогресса в консоли


def read_excel(filename: str):  # функция принимающая имя эксель таблицы
    wb = op.load_workbook(filename, data_only=True)
    sheet = wb.active
    information = []  # список информации о студентах

    for student_number in trange(2, sheet.max_row + 1):  # бежим по каждой строчке
        column_names = []  # список имён столбцов
        column_students = []  # список значений студента в этих столбцах
        for i in range(1, sheet.max_column + 1):  # бежим по каждому столбцу
            column_names.append(
                sheet.cell(row=1, column=i).value
            )  # добавляем в список имен столбцов текущее имя столбца
            column_students.append(
                sheet.cell(row=student_number, column=i).value
            )  # добавляем в список значений студента текущее значение студента
        information.append(
            dict(zip(column_names, column_students))
        )  # создаем из списка имён столбцов и список значений студента словарик  и добавляем каждый словарик в список information
    return information  # вовзращаем список из словарей


def sum(list1, list2):  # функция сложения двух списков из двух эксель таблиц
    bar = IncrementalBar(
        "Countdown", max=len(list1)
    )  # для отображения прогресса в консоли

    for elem1 in list1:  # для каждого элемента из первой таблицы
        for elem2 in list2:  # для каждого элемента из второйтаблицы
            # теперь проверяем есть ли среди двух таблиц повторяющийся студент
            if (
                elem1.get("Дата рождения").lower() == elem2.get("Дата рождения").lower()
            ):  # сначала проверяем равны ли даты рождения,
                if (
                    elem1.get("Фамилия").lower() == elem2.get("Фамилия").lower()
                ):  # если даты равны то затем фамилии,
                    if elem1.get("Отчество").lower() == elem2.get("Отчество").lower():
                        if (
                            elem1.get("Имя").lower() == elem2.get("Имя").lower()
                        ):  # ну если и имена равны то это явно один и тот же человек
                            value = set(elem2) - set(
                                elem1
                            )  # смотрим разницу между словарями. получитс список ключей второго словаря, которых нет в первом
                            for new_key in list(
                                value
                            ):  # для каждого такого ключа добавляем занчение  в первый словарь
                                elem1[new_key] = elem2.get(new_key)
                                print(elem1[new_key])
        bar.next()  # для отображения прогресса в консоли
    return list1  # возвращаем дополненный первый список


if __name__ == "__main__":
    list1 = read_excel("Солдаты их БД ВУЦ.xlsx")
    print(list1)
    with open("tmp.txt", mode="w+") as file:
        file.write(str(list1))

    # list2=read_excel( 'Из Паруса сведения.xlsx')
# print(list2)
# list1=[{"Фамилия":'Морозов',"Имя":"Сергей","Отчество":"Сергеевич","Дата рождения":"06-03-2003","ср балл":"1222222222222222222",}]
# list2=[{"Фамилия":'Морозов',"Имя":"Сергей","Отчество":"Сергеевич","Дата рождения":"06-03-2003","Оценка":"5",}]
# sum(list1,list2)