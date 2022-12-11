import os
import datetime
import openpyxl
import pandas as pd
import warnings

warnings.simplefilter('ignore')

BASE_LINK_JOURNAL = 'https://dnevnik.mos.ru/webteacher/study-process/grade-journals/'
LINK_JOURNALS = os.path.abspath(os.getcwd()) + '\\journals'
LINK_HOME_WORKS = os.path.abspath(os.getcwd()) + '\\homeworks'

ktp = []
ktp_dict = {}
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

def nice_date(dates_set, date_begin, date_end):
    dates_list = list(dates_set)
    dates_list.sort()
    dates_nice = []
    for date in dates_list:
        if (date >= date_begin) and (date <= date_end):
            dates_nice.append(str(date)[8:] + '.' + str(date)[5:7])
    return dates_nice

def to_date(date_str):
    if int(date_str[3:]) > 8:
        return datetime.date(2022, int(date_str[3:]), int(date_str[0:2]))
    return datetime.date(2023, int(date_str[3:]), int(date_str[0:2]))

def get_no_ktp(current_sheet, name_class, name_subj, d1, d2):
    string_sheet = 0
    ktp_key = ''
    for date_journal in current_sheet['Дата']:
        if not pd.isna(date_journal) and isinstance(date_journal, str):
            if len(date_journal) == 5 and current_sheet['Тема'][string_sheet] == 'Без темы':
                ktp_key = name_class + name_subj
                if ktp_key not in ktp_dict:
                    ktp_dict[ktp_key] = set()
                ktp_dict[ktp_key].add(to_date(date_journal))
        string_sheet += 1
    if ktp_key:
        for ktp_group in ktp_dict:
            no_ktp_dates = nice_date(ktp_dict[ktp_group], d1, d2)
        if len(no_ktp_dates) > 0:
            ktp.append([name_class, name_subj, no_ktp_dates])
    return

def get_home_works(sheet_hw, parallel_hw):
    string_sheet = 0
    for group in sheet_hw['Группа']:
        date_set = sheet_hw['Даты'][string_sheet]
        index_date = date_set.find('Задано на: ') + 11
        group_class_subj = get_class_subj_sheet(group, parallel_hw)
        hw_key = group_class_subj[0] + group_class_subj[1]
        date_hw = date_set[index_date: index_date + 5]
        if hw_key not in hw_real_dict:
            hw_real_dict[hw_key] = set()
        hw_real_dict[hw_key].add(to_date(date_hw))
        string_sheet += 1


def get_home_works_journal(current_sheet, name_class, name_subj):
    string_sheet = 0
    for date_journal in current_sheet['Дата']:
        if not pd.isna(date_journal) and isinstance(date_journal, str):
            if len(date_journal) == 5 and current_sheet['Домашнее задание'][string_sheet] == 'не задано':
                hw_journal_key = name_class + name_subj
                if hw_journal_key not in hw_journal_dict:
                    hw_journal_dict[hw_journal_key] = set()
                hw_journal_dict[hw_journal_key].add(to_date(date_journal))
        string_sheet += 1
    return

def question_dates(question_date):
    period = []
    print(question_date)
    date1_str = input('Начало периода: ')
    period.append(datetime.date(int(date1_str[6:]), int(date1_str[3:5]), int(date1_str[0:2])))
    date2_str = input('Конец периода: ')
    period.append(datetime.date(int(date2_str[6:]), int(date2_str[3:5]), int(date2_str[0:2])))
    return period[0], period[1]

def question_yes_no(question):
    answer = input(question)
    if answer == 'y':
        return True
    return False

def new_file(name_file):
    new_book = openpyxl.Workbook()
    new_book_link = link_files + '\\' + name_file
    try:
        new_book.save(new_book_link)
    except FileExistsError:
        pass
    return new_book_link

parallel = input('Введите номер параллели: ')
links_answer = question_yes_no('Сгенерировать ссылки на журналы? (y/n)')
ktp_answer = question_yes_no('Проверить КТП? (y/n)')
if ktp_answer:
    date1_ktp, date2_ktp = question_dates('Введите период для проверки КТП. Например: 01.09.2022')
hw_answer = question_yes_no('Проверить отсутствие домашнего задания? (y/n)')
if hw_answer:
    date1_hw, date2_hw = question_dates('Введите период для проверки домашнего задания. Например: 01.09.2022')
    files_home_works = os.listdir(LINK_HOME_WORKS)
    for file_home_works in files_home_works:
        print('... ', file_home_works)
        table_home_works = pd.ExcelFile(LINK_HOME_WORKS + '\\' + file_home_works)
        sheet_home_works = table_home_works.parse(sheet_id=0)
        get_home_works(sheet_home_works, parallel)
if links_answer or ktp_answer or hw_answer:
    link_files = os.path.abspath(os.getcwd()) + '\\' + str(datetime.date.today()) + '_' + str(parallel)
    try:
        os.mkdir(link_files)
    except FileExistsError:
        pass
    files_journals = os.listdir(LINK_JOURNALS)
    for file_journal in files_journals:
        print('... ', file_journal)
        table_journal = pd.ExcelFile(LINK_JOURNALS + '\\' + file_journal)
        names_sheet = table_journal.sheet_names
        if links_answer:
            links.append(['', '', ''])
        for i in names_sheet:
            sheet = table_journal.parse(sheet_name=i)
            names_class_subj = get_class_subj_sheet(sheet['Дата'][39], parallel)
            if links_answer:
                links.append([names_class_subj[0], names_class_subj[1], BASE_LINK_JOURNAL + i[-7:]])
            if ktp_answer:
                get_no_ktp(sheet, names_class_subj[0], names_class_subj[1], date1_ktp, date2_ktp)
            if hw_answer:
                get_home_works_journal(sheet, names_class_subj[0], names_class_subj[1])
if links_answer:
    links_frame = pd.DataFrame(links, columns=['Класс', 'Предмет', 'Ссылка на журнал'])
    links_frame.to_excel(new_file('links_journals.xlsx'))
    print('Файл links_journals.xlsx обновлен')
if ktp_answer:
    ktp_frame = pd.DataFrame(ktp, columns=['Класс', 'Предмет', 'Даты уроков без КТП'])
    ktp_frame.to_excel(new_file('ktp.xlsx'))
    print('Файл ktp.xlsx обновлен')
if hw_answer:
    for current_group in hw_journal_dict:
        if current_group in hw_real_dict:
            no_hw_dates_set = hw_journal_dict[current_group] - hw_real_dict[current_group]
        else:
            no_hw_dates_set = hw_journal_dict[current_group]
        no_hw_dates = nice_date(no_hw_dates_set, date1_hw, date2_hw)
        if len(no_hw_dates) > 0:
            hw_journals.append([current_group, no_hw_dates])
    hw_frame = pd.DataFrame(hw_journals, columns=['Класс, Предмет', 'Даты без домашнего задания'])
    hw_frame.to_excel(new_file('no_hw.xlsx'))
    print('Файл no_hw.xlsx обновлен')
