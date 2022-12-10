Программа выводит в файлы отчеты: ссылки на журналы в новой версии ЭЖД, уроки без тем (отсутствие КТП), уроки, где не было задано домашнее задание.
Во время исполнения программы можно выбрать любой набор модулей функционала.
Результаты работы программы основаны на анализе отчетов ЭЖД в одной параллели.
Программу необходимо запускать в виртуальном окружении Python с установленными пакетами os, pandas, warnings.
Рядом с файлом программы (в одном каталоге) должны быть каталоги homeworks, journals, файлы mydata.xlsx, mydata_ktp.xlsx, mydata_hw.xlsx (обычный пустой файл xlsx, если файл не пустой – он будет обновлен).
Подготовка:
1)	Скачать все журналы параллели (старая версия ЭЖД, Журналы, скачать все журналы доступные в параллели). Все журналы скопировать в каталог journals.
2)	Для анализа домашних заданий нужно скачать отчет: Дополнительно – Администрирование – Все ДЗ, выбрать справа номер параллели, нажать Применить, скачать отчет. 
Открыть скаченный отчет в программе Excel, удалить первые строки с датами до шапки таблицы. Шапку таблицы оставить, сохранить.
Отчет скопировать в папку homeworks.
3)	Установить необходимые пакеты Python (os, pandas, warnings).
4)	Во время выполнения программы файлы mydata.xlsx, mydata_ktp.xlsx, mydata_hw.xlsx должны быть закрыты.
