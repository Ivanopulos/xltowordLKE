# -*- coding: utf-8 -*-
from docx import Document
import re

def update_word_table(doc_path, data_var, units_var):
    # Разделение данных на списки
    data_list = data_var.split('; ')
    units_list = units_var.split('; ')

    # Открытие документа
    doc = Document(doc_path)

    # Доступ к таблице
    table = doc.tables[1]

    # Добавление данных в таблицу
    for idx, data in enumerate(data_list):
        key, value = data.split(' - ')
        cells = table.add_row().cells
        cells[0].text = key
        cells[1].text = value
        cells[2].text = value
        cells[3].text = units_list[idx]
        cells[5].text = "Территориальный орган Федеральной службы государственной статистики по РЕГИОН"

    # Сохранение документа
    doc.save(doc_path)

data_var = 'Чз - численность населения в возрасте 3-79 лет, занимающегося физической культурой и спортом, в соответствии с данными федерального статистического наблюдения по форме N 1-ФК "Сведения о физической культуре и спорте" тысяч (человек); Чн - численность населения в возрасте 3-79 лет за отчетный год (человек). Источник данных - Единая межведомственная информационно-статистическая система; Чнп - численность населения в возрасте 3-79 лет, имеющего противопоказания и ограничения для занятий физической культурой и спортом, согласно формам статистического наблюдения, за отчетный год.'
doc_path = "1.docx"
units_var = data_var
# зачистка от шума
units_var = re.sub(r'Источник данных[^;]*', r'', units_var)
# очевидные по шаблону комментария
units_var = re.sub(r'[^;]*процент[оы]?[в]?([;\.])', r'процент\1', units_var)
units_var = re.sub(r'[^;]*лет([;\.])', r'лет\1', units_var)
units_var = re.sub(r'[^;]*\bтыс(яч)?\b[^;]{,5}человек[^;]{,7}', r'тысъъчеловек', units_var)
units_var = re.sub(r'[^;][^ъ]человек[^;]{,7}', r'человек', units_var)

print(units_var)
# ориентирующие по вхождению в текст
units_var = re.sub(r'[^;]*числен[^;]*', r'человек', units_var)
units_var = re.sub(r'[^;]*число[^;]*ших[^;]*', r'человек', units_var)
units_var = re.sub(r'[^;]*[дД]оля[^;]*', r'процент', units_var)
# финишный разделитель

print(units_var)


# substrings = data_var.split(';')
# Создаем список для результатов
# result = []
#
# # Проверяем каждую подстроку
# for s in substrings:
#     if "численность" in s or "количество" in s:
#         result.append("человек")
#     elif "доля" in s:
#         result.append("проценты")
#     else:
#         result.append("человек")  # по условиям, если ни одно из слов не найдено, то дефолтное значение "человек"
#
# # Объединяем результаты в одну строку с разделителем ";"
# units_var = "; ".join(result) + "."

# update_word_table(doc_path, data_var, units_var)