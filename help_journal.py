import os
import pandas as pd
import warnings

warnings.simplefilter('ignore')

BASE_LINK_JOURNAL = 'https://dnevnik.mos.ru/webteacher/study-process/grade-journals/'
LINK_JOURNALS = os.path.abspath(os.getcwd()) + '\\journals'
LINK_HOME_WORKS = os.path.abspath(os.getcwd()) + '\\homeworks'

ktp = []
links = []
hw_journals = []
hw_real_dict = {}
hw_journal_dict = {}

def get_class_subj_sheet(cell_value, n_parallel):
    if str(n_parallel) in cell_value:
        index_parallel = cell_value.find(str(n_parallel))
        name_subj = cell_value[0:index_parallel - 1]
        index_whitespace = cell_value.find(' ', index_parallel)
        name_class = cell_value[index_parallel:index_whitespace]
        if 'НДО' in cell_value:
            name_class += cell_value[index_whitespace:index_whitespace + 20]
    else:
        name_subj = cell_value
        name_class = ''
    return name_class, name_subj

def get_no_ktp(current_sheet, name_class, name_subj):
    string_sheet = 0
    for date_journal in current_sheet['Дата']:
        if not pd.isna(date_journal) and isinstance(date_journal, str):
            if len(date_journal) == 5 and current_sheet['Тема'][string_sheet] == 'Без темы':
                ktp.append([name_class, name_subj, date_journal, current_sheet['Тема'][string_sheet]])
        string_sheet += 1
    return

def get_home_works(sheet_hw, parallel_hw):
    string_sheet = 0
    for group in sheet_hw['Группа']:
        date_set = sheet_hw['Даты'][string_sheet]
        index_date = date_set.find('Задано на: ') + 11
        group_class_subj = get_class_subj_sheet(group, parallel_hw)
        hw_key = group_class_subj[0] + group_class_subj[1]
        if hw_key not in hw_real_dict:
            hw_real_dict[hw_key] = set()
        hw_real_dict[hw_key].add(date_set[index_date: index_date + 5])
        string_sheet += 1


def get_home_works_journal(current_sheet, name_class, name_subj):
    string_sheet = 0
    for date_journal in current_sheet['Дата']:
        if not pd.isna(date_journal) and isinstance(date_journal, str):
            if len(date_journal) == 5 and current_sheet['Домашнее задание'][string_sheet] == 'не задано':
                hw_journal_key = name_class + name_subj
                if hw_journal_key not in hw_journal_dict:
                    hw_journal_dict[hw_journal_key] = set()
                hw_journal_dict[hw_journal_key].add(date_journal)
        string_sheet += 1
    return

parallel = input('Введите номер параллели: ')
links_answer = input('Нужно генерировать ссылки на журналы? (y/n)')
if links_answer == 'y':
    links_answer = True
else:
    links_answer = False
ktp_answer = input('Нужно проверять КТП? (y/n)')
if ktp_answer == 'y':
    ktp_answer = True
else:
    ktp_answer = False
hw_answer = input('Нужно проверять отсутствие домашнего задания? (y/n)')
if hw_answer == 'y':
    hw_answer = True
    files_home_works = os.listdir(LINK_HOME_WORKS)
    for file_home_works in files_home_works:
        print('... ', file_home_works)
        table_home_works = pd.ExcelFile(LINK_HOME_WORKS + '\\' + file_home_works)
        sheet_home_works = table_home_works.parse(sheet_id=0)
        get_home_works(sheet_home_works, parallel)
else:
    hw_answer = False
files_journals = os.listdir(LINK_JOURNALS)
for file_journal in files_journals:
    print('... ', file_journal)
    table_journal = pd.ExcelFile(LINK_JOURNALS + '\\' + file_journal)
    names_sheet = table_journal.sheet_names
    if links_answer:
        links.append(['', '', ''])
    if ktp_answer:
        ktp.append(['', '', ''])
    if hw_answer:
        hw_journals.append(['', ''])
    for i in names_sheet:
        sheet = table_journal.parse(sheet_name=i)
        names_class_subj = get_class_subj_sheet(sheet['Дата'][39], parallel)
        if links_answer:
            links.append([names_class_subj[0], names_class_subj[1], BASE_LINK_JOURNAL + i[-7:]])
        if ktp_answer:
            get_no_ktp(sheet, names_class_subj[0], names_class_subj[1])
        if hw_answer:
            get_home_works_journal(sheet, names_class_subj[0], names_class_subj[1])
if links_answer:
    q = pd.DataFrame(links, columns=['Класс', 'Предмет', 'Ссылка на журнал'])
    q.to_excel(r'mydata.xlsx')
    print('Файл mydata.xlsx обновлен')
if ktp_answer:
    ktp_frame = pd.DataFrame(ktp, columns=['Класс', 'Предмет', 'Дата', 'Тема'])
    ktp_frame.to_excel(r'mydata_ktp.xlsx')
    print('Файл mydata_ktp.xlsx обновлен')
if hw_answer:
    for current_group in hw_journal_dict:
        if current_group in hw_real_dict:
            no_hw_dates = hw_journal_dict[current_group] - hw_real_dict[current_group]
        else:
            no_hw_dates = hw_journal_dict[current_group]
        hw_journals.append([current_group, no_hw_dates])
    hw_frame = pd.DataFrame(hw_journals, columns=['Класс, Предмет', 'Даты без домашнего задания'])
    hw_frame.to_excel(r'mydata_hw.xlsx')
    print('Файл mydata_hw.xlsx обновлен')
