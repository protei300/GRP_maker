import pandas as pd
from docxtpl import DocxTemplate
import os
from num2words import num2words
import openpyxl
import numpy as np
from datetime import datetime
import locale
import tqdm
import re
import math
from settings import *

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')



learning_form_translator = {
    "очно-заочная": "очно-заочной",
    "очная": "очной",
    "заочная": "заочной",
}

MONTHS_TRANSLATOR = {
    1: 'января',
    2: 'февраля',
    3: 'марта',
    4: 'апреля',
    5: 'мая',
    6: 'июня',
    7: 'июля',
    8: 'августа',
    9: 'сентября',
    10: 'октября',
    11: 'ноября',
    12: 'декабря',
}


def translate_month(date):
    return f"«{date.day}» {MONTHS_TRANSLATOR[date.month]} {date.year}"

def render_gpd(context_list):
    '''
    Функция по генерации ГПД на основе данных файла контекст

    :param context:
    :return:
    Генерирует файлы в папке ГПД с коротким именем исполнителя
    '''

    print(f"[*] Начинаю создание файлов ГПД")

    for root, dirs, files in os.walk(os.path.join(RESULT_DIR, 'ГПД')):
        for f in files:
            os.unlink(os.path.join(root, f))

    for context in tqdm.tqdm(context_list):
        filename_to_save = context['short_executor_name'].split(' ', maxsplit=1)
        filename_to_save = f"{filename_to_save[-1]} {filename_to_save[0]}.docx"

        doc = DocxTemplate(os.path.join(TEMPLATES_DIR, CONTRACT_TEMPLATE))
        doc.render(context)
        doc.save(os.path.join(RESULT_DIR,'ГПД',filename_to_save))
    print("")
    print(f"[+] Закончил создание файлов ГПД")

def render_justification(context_list):
    print(f"[*] Начинаю создание файлов Справок-обоснований")

    for root, dirs, files in os.walk(os.path.join(RESULT_DIR, 'Справки')):
        for f in files:
            os.unlink(os.path.join(root, f))


    for context in tqdm.tqdm(context_list):
        filename_to_save = context['short_executor_name'].split(' ', maxsplit=1)
        filename_to_save = f"{filename_to_save[-1]} {filename_to_save[0]}-справка.docx"

        doc = DocxTemplate(os.path.join(TEMPLATES_DIR, REFERENCE_TEMPLATE))
        doc.render(context)
        doc.save(os.path.join(RESULT_DIR, 'Справки', filename_to_save))

    print("")
    print(f"[+] Закончил создание файлов Справок-обоснований")

def get_dataframe(filename_to_read='Данные для договоров внебюджет, очка 2 сем 2020-2021.xlsx'):
    '''
    Считываем датафрейм с учетом скрытых строк (их отбрасываем), переименовываем столбцы
    в необходимые нам теги
    :param filename_to_read: название файла с данными договоров
    :return:
    '''

    print(f"[*] Подгружаю файл {filename_to_read} в датафрейм")

    filename = os.path.join(TEMPLATES_DATA_DIR,
                                    filename_to_read)


    wb = openpyxl.load_workbook(filename)
    ws = wb['Лист1']

    rows_to_skip = []

    for rowNum,rowDimension in ws.row_dimensions.items():
      if rowDimension.hidden == True:
         rows_to_skip.append(rowNum-1)


    df = pd.read_excel(os.path.join(TEMPLATES_DATA_DIR,
                                    filename_to_read),
                        engine='openpyxl',
                        skiprows=rows_to_skip
                   )

    df['Денег в текущем'] = df['Денег в текущем'].fillna(0)
    df['Денег в текущем'] = df['Денег в текущем'].astype('int32')

    print(f"[*] Датафрейм загружен")

    return df

def make_context(df, start_date):
    '''
    Создаем контекст для заполнения ГПД
    Есть следующие поля для подстановки
    short_executor_name - Короткое ФИО
    long_executor_name - Полное ФИО
    power_of_attorney - данные доверенности
    total - полная сумма договора
    total_words - полная сумма договора прописью
    text_before_table - текст перед таблицей
    money_this_year - всего в этом году
    money_this_year_words - всего в этом году прописью
    money_next_year - всего в след году
    money_next_year_words - всего в след году прописью
    executor_address - реквизты исполнителя
    total_hours - всего часов
    ending_date - дата окончания окозания услуг
    agreement_ending_date - дата окончания действия договора
    program_codes - поле с указание ОПОП
    learning_form - форма обучения
    all_disciplines - все дисциплины преподаваемые преподавателем

    Таблица:
    number - Номер строки
    discipline - название дисциплины
    group_number - номер группы
    services - оказываемые услуги
    hours - часов на данную дисциплину
    hours_price - стоимость часа
    total_for_service - общая стоимость за услугу

    :param df:
    :return:
    '''
    print(f"[*] Начинаю создание данных для заполнения шаблонов ГПД")

    context_list = []
    task_cols = df.columns[df.columns.str.contains(pat = r'Дисциплина/|Форма обучения|Перечень услуг|'
                                              'Объем услуг|Цена за 1 ак. час|Всего[\.\d]*$')]

    for _,row in df.iterrows():
        learning_forms = []
        context ={
            'long_executor_name': row['ФИОисполнителя'],
            'short_executor_name': row['Краткое ФИО исполнителя'],
            'power_of_attorney': row['Доверенность проректора'],
            'text_before_table': row['Текст перед таблицей'],
            'start_date': start_date,
            'ending_date': translate_month(row['Дата окончания оказ услуги']),
            'agreement_ending_date': translate_month(row['Срок действия договора']),
            'executor_address': row['Адрес исполнителя'],
            'total_hours': row['Всего часов'],
            'total': row['Всего денег'],
            'total_words': num2words(row['Всего денег'], lang='ru'),
            'this_year': datetime.now().year,
            'next_year': datetime.now().year+1,
            'program_codes': row['ОП ВО'],


        }
        if not np.isnan(row['Денег в текущем']):
            context['money_this_year'] =  row['Денег в текущем']
            context['money_this_year_words'] =  num2words(row['Денег в текущем'], lang='ru')
        else:
            context['money_this_year'] = 0
            context['money_this_year_words'] = num2words(0, lang='ru')
        if not np.isnan(row['Денег в следующем']):
            if row['Денег в следующем'] != 0:
                context['money_next_year'] = row['Денег в следующем']
                context['money_next_year_words'] =  num2words(row['Денег в следующем'], lang='ru')

        context['tbl_contents'] = []

        tasks = row[task_cols]


        tasks = tasks.values.reshape(4,-1)
        for i,task in enumerate(tasks):
            if type(task[0]) == str:
                context['tbl_contents'].append(
                    {'number': i+1,
                     'discipline': task[0],
                     'group_number': task[1],
                     'services': task[2],
                     'hours': int(task[3]) if math.modf(task[3])[0]==0 else round(task[3],1),
                     'hour_price': int(task[4]) if math.modf(task[4])[0]==0 else round(task[4],1),
                     'total_for_service': int(task[5]) if math.modf(task[5])[0]==0 else round(task[5],1),
                     }
                )
                l_form = re.split(r'[, ]', task[1], maxsplit=1)[0].lower()
                learning_forms.append(learning_form_translator[l_form])

        learning_forms = list(set(learning_forms))
        all_disciplines = [ tbl_row['discipline'] for tbl_row in context['tbl_contents']]
        all_disciplines = ', '.join(all_disciplines)

        context['learning_form'] = ', '.join(learning_forms)
        context['all_disciplines'] = all_disciplines
        context_list.append(context)


    print(f"[*] Было создано {len(context_list)} контекстов для заполнения шаблонов ГПД")
    return context_list


if __name__ == '__main__':
    df = get_dataframe(EXCEL_FILE)
    context = make_context(df, START_DATE)
    render_justification(context)
    render_gpd(context)

