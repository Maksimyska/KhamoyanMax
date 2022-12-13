import csv
import re
import os
from functools import cmp_to_key
from prettytable import PrettyTable
from prettytable import ALL


def normalize_input_info(input_info):
    """Обработка данных, входящих от пользователя

    Args:
        input_info (list[str | bool | list[str] | list[int]]): Данные, входящие от пользователя

    Returns:
        string: Результаты обработки
    """
    table_fields = [
        "Название",
        "Описание",
        "Навыки",
        "Опыт работы",
        "Премиум-вакансия",
        "Компания",
        "Оклад",
        "Название региона",
        "Дата публикации вакансии",
        "Идентификатор валюты оклада",
        "None"
    ]
    if os.stat(input_info[0]).st_size == 0:
        return "Пустой файл"
    if input_info[1] == '':
        input_info[1] = "None: None"
    temp = input_info[1].find(': ')
    if temp == -1:
        return "Формат ввода некорректен"
    input_info[1] = [input_info[1][:temp], input_info[1][temp + 2:]]
    if input_info[1][0] not in table_fields:
        return "Параметр поиска некорректен"
    if input_info[2] == '':
        input_info[2] = '№'
    elif input_info[2] not in table_fields:
        return "Параметр сортировки некорректен"
    if input_info[3] == 'Да':
        input_info[3] = True
    elif input_info[3] == 'Нет' or input_info[3] == '':
        input_info[3] = False
    else:
        return "Порядок сортировки задан некорректно"
    input_info[4] = list(map(lambda x: 0 if x == '' else int(x) - 1, input_info[4].split(' ')))
    if len(input_info[4]) == 1:
        input_info[4].append(10000)
    input_info[5] = input_info[5].split(', ')
    if input_info[5][0] == '':
        input_info[5] = table_fields[:9]
    input_info[5].insert(0, '№')
    return "Нормализация прошла успешно"


def csv_reader(file_name):
    """Чтение csv файла
    Args:
        file_name (str): Имя csv файла
    Returns:
        title, info: Результат чтения
    """
    with open(file_name, encoding="utf-8-sig") as file:
        reader = [x for x in csv.reader(file)]
        title = reader.pop(0)
        titleCount = len(title)
        info = list(filter(lambda data: '' not in data and len(data) == titleCount, reader))
    return title, info


def csv_filter(title, info):
    """Преобразование данных из csv файла в список словарей

        Args:
            title (list[str]): Названия столбцов
            info (list[list[str]]): Основные данные csv файла

        Returns:
            list[dict[str,str]]: Список строк в виде словарей
        """

    def normalize_csv_file(info_cell):
        """Нормализация данных. Удаление лишних элементов

            Args:
                info_cell (str): Ячейка csv файла

            Returns:
                str: Нормализованная ячейка csv файла
            """
        temp_info = "__temp__".join(info_cell.split("\n"))
        temp_info = re.sub(r"<[^<>]*>", "", temp_info)
        temp_info = re.sub(r"\s+", " ", temp_info)
        return str.strip(temp_info)

    info_dictionaries = []
    for info_row in info:
        info_dictionary = {}
        for i in range(len(title)):
            info_dictionary[title[i]] = normalize_csv_file(info_row[i])
        info_dictionaries.append(info_dictionary)
    return info_dictionaries


def info_formatter(vacancies):
    """Нормализация данных в вакансиях

    Args:
        vacancies (list[Vacancy] | Vacancy): Список вакансий

    Returns:
        list[Vacancy] | Vacancy: Результат форматирования
    """

    def formatter_string_number(str_num):
        """Устранение дробных разделителей в строковом числе

        Args:
            str_num (str): Число для нормализации

        Returns:
            str: Результат форматирования числа
        """
        num = int(str_num if str_num.find('.') == -1 else str_num[:len(str_num) - 2])
        str_num_reverse = str(num)[::-1]
        return ' '.join(str_num_reverse[i:i + 3] for i in range(0, len(str_num_reverse), 3))[::-1]

    def formatter_experience_id(new_info_dictionary, value):
        """Преобразование оклада, написаного на английском, к русскому виду

        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Английский вариант написания
        """
        new_info_dictionary["Опыт работы"] = dic_experience[value]

    def formatter_salary_from(new_info_dictionary, value):
        """Преобразование нижней линии оклада в нормированный вид
        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Значение оклада
        """
        new_info_dictionary['Оклад'] = formatter_string_number(value)

    def formatter_salary_to(new_info_dictionary, value):
        """Преобразование верхней линии оклада в нормированный вид

        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Значение оклада
        """
        new_info_dictionary['Оклад'] = f"{new_info_dictionary['Оклад']} - {formatter_string_number(value)}"

    def formatter_salary_currency(new_info_dictionary, value):
        """Преобразование валюты, написаной на английском, к русскому виду

        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Английский вариант написания
        """
        new_info_dictionary["Оклад"] = f"{new_info_dictionary['Оклад']} ({dic_currency[value]})" \
                                       f" ({new_info_dictionary['salary_currency']})"

    def formatter_salary_gross(new_info_dictionary, value):
        """Преобразование значения вычета налогов, написаного на английском, к русскому виду
        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Английский вариант написания
        """
        new_info_dictionary['salary_currency'] = 'Без вычета налогов' if value == 'True' \
            else 'С вычетом налогов' if value == 'False' else value

    def formatter_published_at(new_info_dictionary, value):
        """Преобразование времени выкладывания вакансии в формат ДД:ММ:ГГ и запись в словарь результата

        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Значение времени выкладывания вакансии для форматирования
        """
        new_info_dictionary["Дата публикации вакансии"] = f"{value}#{value[8:10]}.{value[5:7]}.{value[0:4]}"

    def formatter_premium(new_info_dictionary, value):
        """Преобразование значения премиум вакансии, написаного на английском, к русскому виду

        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Английский вариант написания
        """
        new_info_dictionary["Премиум-вакансия"] = 'Да' if value == 'True' else 'Нет'

    def formatter_key_skills(new_info_dictionary, value):
        """Преобразование скиллов, написанных на английском, к русскому виду

        Args:
            new_info_dictionary (dict[str,str | int]): Новый список словарей для результата общего метода
            value (str): Английский вариант написания
        """
        value = value.replace("__temp__", '\n')
        new_info_dictionary["Количество навыков"] = value.count('\n') + 1
        new_info_dictionary["Навыки"] = f"{value[0:100]}..." if len(value) > 100 else value

    def formatter_standard_field_value(new_info_dictionary, value, key):
        """Сокращение строки до первых 100 символов
        Args:
            new_info_dictionary (dict[str,str]): Новый список словарей для результата общего метода
            value (str): Строка для форматирования
            key (str): Ключ для записи в список словарей нового значения
        """
        new_info_dictionary[dic_naming[key]] = f"{value[0:100]}..." if len(value) > 100 else value

    dic_naming = {
        "name": "Название",
        "description": "Описание",
        "employer_name": "Компания",
        "area_name": "Название региона"
    }
    dic_experience = {
        "noExperience": "Нет опыта",
        "between1And3": "От 1 года до 3 лет",
        "between3And6": "От 3 до 6 лет",
        "moreThan6": "Более 6 лет"
    }
    dic_currency = {
        "AZN": "Манаты",
        "BYR": "Белорусские рубли",
        "EUR": "Евро",
        "GEL": "Грузинский лари",
        "KGS": "Киргизский сом",
        "KZT": "Тенге",
        "RUR": "Рубли",
        "UAH": "Гривны",
        "USD": "Доллары",
        "UZS": "Узбекский сум"
    }
    dic_func = {
        "experience_id": formatter_experience_id,
        "salary_from": formatter_salary_from,
        "salary_to": formatter_salary_to,
        "salary_currency": formatter_salary_currency,
        "salary_gross": formatter_salary_gross,
        "published_at": formatter_published_at,
        "premium": formatter_premium,
        "key_skills": formatter_key_skills,
        "name": formatter_standard_field_value,
        "description": formatter_standard_field_value,
        "employer_name": formatter_standard_field_value,
        "area_name": formatter_standard_field_value
    }

    formatted_info_dictionaries = []
    for info_dictionary in vacancies:
        formatted_info_dictionary = {}
        for item_key, item_value in info_dictionary.items():
            dic_func[item_key](formatted_info_dictionary, item_value, item_key)
        if 'salary_currency' in formatted_info_dictionary:
            formatted_info_dictionary.pop('salary_currency')
        formatted_info_dictionaries.append(formatted_info_dictionary)
    return formatted_info_dictionaries


def info_filter(info_dictionaries, filtering_parameter):
    """Фильтрация списка словарей, соответствующих строкам csv файла
        Args:
            info_dictionaries (list[dict[str,str]]): Список словарей для фильтрации
            filtering_parameter (list[str,str]): Параметр фильтрации
        Returns:
            list[dict[str,str]]: Результат фильтрации
    """

    def filter_verbatim(dic, field_value_should):
        """Лексикографическое сравнивание значения из словаря с требуемым значением
        Args:
            dic (dict[str,str]): Словарь, представляющий собой одну строку csv файла
            field_value_should (tuple[str, str]): Ключ словаря для фильтрации и соответсвующее требуемое значение
        Returns:
            string: Результат сравнения
        """
        return dic[field_value_should[0]] == field_value_should[1]

    def filter_key_skills(dic, field_value_should):
        """Уникальное сравнивание значения скиллов из словаря с требуемым значением
        Args:
            dic (dict[str,str]): Словарь, представляющий собой одну строку csv файла
            field_value_should (tuple[str, str]): Ключ словаря для фильтрации и соответсвующее требуемое значение
        Returns:
            string: Результат сравнения
        """
        value_should = field_value_should[1].split(', ')
        dic_values = dic['Навыки'].replace(', ', '\n').replace('...', '\n').split('\n')
        return all(list(map(lambda value_should: value_should in dic_values, value_should)))

    def filter_salary(dic, field_value_should):
        """Уникальное сравнивание значения оклада из словаря с требуемым значением
        Args:
            dic (dict[str,str]): Словарь, представляющий собой одну строку csv файла
            field_value_should (tuple[str, str]): Ключ словаря для фильтрации и соответсвующее требуемое значение
        Returns:
            string: Результат сравнения
        """
        dic_value = dic["Оклад"]
        salary_area = dic_value[:dic_value.find('(')].replace(' ', '').split('-')
        return int(salary_area[0]) <= int(field_value_should[1]) <= int(salary_area[1])

    def filter_salary_currency(dic, field_value_should):
        """Уникальное сравнивание значения валюты оклада из словаря с требуемым значением
        Args:
            dic (dict[str,str]): Словарь, представляющий собой одну строку csv файла
            field_value_should (tuple[str, str]): Ключ словаря для фильтрации и соответсвующее требуемое значение
        Returns:
            string: Результат сравнения
        """
        dic_value = dic["Оклад"]
        temp = dic_value[dic_value.find('(') + 1:dic_value.find(')')]
        return temp == field_value_should[1]

    def filter_published_at(dic, field_value_should):
        """Уникальное сравнивание значения времени публикации из словаря с требуемым значением
        Args:
            dic (dict[str,str]): Словарь, представляющий собой одну строку csv файла
            field_value_should (tuple[str, str]): Ключ словаря для фильтрации и соответсвующее требуемое значение
        Returns:
            string: Результат сравнения
        """
        dic_value = dic["Дата публикации вакансии"]
        return dic_value[dic_value.find('#') + 1:] == field_value_should[1]

    dic_filter = {"Название": filter_verbatim, "Описание": filter_verbatim, "Навыки": filter_key_skills,
                  "Опыт работы": filter_verbatim, "Премиум-вакансия": filter_verbatim, "Компания": filter_verbatim,
                  "Оклад": filter_salary, "Дата публикации вакансии": filter_published_at,
                  "Идентификатор валюты оклада": filter_salary_currency, "Название региона": filter_verbatim}

    return list(filter(lambda info_dictionary:
                       filtering_parameter[0] == "None" or
                       dic_filter[filtering_parameter[0]](info_dictionary, filtering_parameter), info_dictionaries))


def info_sorter(info_dictionaries, sort_field, reverse_sort):
    """Сортировка списка словарей, представляющих собой строки файла csv формата
        Args:
            info_dictionaries (list[dict[str,str]]): Данные для сортировки
            sort_field (str): Параметр сортировки
            reverse_sort (bool): Сортировать ли в обратном порядке
        Returns:
            list[dict[str,str]]: Результат сортировки
    """

    def lexcographic_sorter(row1, row2):
        """Лексикографическое сравнение одного словаря с другим по параметру сортировки
        Args:
            row1 (dict[str,str]): Первый словаря для сравнения
            row2 (dict[str,str]): Первый словаря для сравнения
        Returns:
            int: Результат сравнения
        """
        return 1 if row1[sort_field] >= row2[sort_field] else -1

    def key_skills_sorter(row1, row2):
        """Сравнение одного словаря с другим по количеству навыков
        Args:
            row1 (dict[str,str]): Первый словаря для сравнения
            row2 (dict[str,str]): Первый словаря для сравнения
        Returns:
            int: Результат сравнения
        """
        (row1_len, row2_len) = list(map(lambda row: row["Количество навыков"], (row1, row2)))
        return row1_len - row2_len

    def experience_sorter(row1, row2):
        """Сравнение одного словаря с другим по количеству требуемых лет опыта
        Args:
            row1 (dict[str,str]): Первый словаря для сравнения
            row2 (dict[str,str]): Первый словаря для сравнения
        Returns:
            int: Результат сравнения
        """

        def find_first_num(row_value):
            """Получение первого числа в строке
            Args:
                row_value (str): Строка для поиска
            Returns:
                int: Результат сравнения
            """
            row_num = list(filter(lambda char: char.isdigit(), row_value))
            return int(row_num[0]) if len(row_num) > 0 else 0

        (row1_num, row2_num) = list(map(lambda row: find_first_num(row["Опыт работы"]), (row1, row2)))
        return row1_num - row2_num

    def salary_sorter(row1, row2):
        """Сравнение одного словаря с другим по окладу
        Args:
            row1 (dict[str,str]): Первый словаря для сравнения
            row2 (dict[str,str]): Первый словаря для сравнения
        Returns:
            int: Результат сравнения
        """

        def salary_process(row):
            """Вычисление среднего значения оклада
            Args:
                row (dict[str,str]): Словарь, представляющий собой строку csv файла
            Returns:
                int: Среднее оклада
            """
            dic_currency_to_rub = {
                "Манаты": 35.68,
                "Белорусские рубли": 23.91,
                "Евро": 59.90,
                "Грузинский лари": 21.74,
                "Киргизский сом": 0.76,
                "Тенге": 0.13,
                "Рубли": 1,
                "Гривны": 1.64,
                "Доллары": 60.66,
                "Узбекский сум": 0.0055
            }
            string_nums = row[sort_field].split(' - ')
            salary_currency = row[sort_field][row[sort_field].find('(') + 1:row[sort_field].find(')')]
            string_nums[1] = string_nums[1][:string_nums[1].find(' (')]
            nums = list(map(lambda string_num: int(string_num.replace(' ', '')), string_nums))
            rub_nums = list(map(lambda num: num * dic_currency_to_rub[salary_currency], nums))
            return sum(rub_nums) / 2

        (row1_salary, row2_salary) = list(map(lambda row: salary_process(row), (row1, row2)))
        return row1_salary - row2_salary

    dic_sorter = {
        "Название": lexcographic_sorter,
        "Описание": lexcographic_sorter,
        "Навыки": key_skills_sorter,
        "Опыт работы": experience_sorter,
        "Премиум-вакансия": lexcographic_sorter,
        "Компания": lexcographic_sorter,
        "Оклад": salary_sorter,
        "Название региона": lexcographic_sorter,
        "Дата публикации вакансии": lexcographic_sorter
    }

    info_dictionaries.sort(key=cmp_to_key(dic_sorter[sort_field]), reverse=reverse_sort)

    return info_dictionaries


def print_vacancies(info_dictionaries, start_end_nums, table_fields):
    """Печать талицы с вакансиями

    Args:
        info_dictionaries (list[dict[str,str]]): Список словарей
        start_end_nums (list[int, int]): Диапозон номеров вакансий
        table_fields (list[str]): Название столбцов для вывода в таблицу
    """
    info_table = PrettyTable([
        "Название",
        "Описание",
        "Навыки",
        "Опыт работы",
        "Премиум-вакансия",
        "Компания",
        "Оклад",
        "Название региона",
        "Дата публикации вакансии"
    ])
    for info_dictionary in info_dictionaries:
        values = list(map(lambda key: info_dictionary[key], info_dictionary))
        values.pop(2)
        info_table.add_row(values)
    published_at_data = list(filter(lambda x: x != '', info_table.get_string(
        fields=["Дата публикации вакансии"], border=False, header=False).replace(' ', '').split('\n')))
    info_table.del_column("Дата публикации вакансии")
    info_table.add_column("Дата публикации вакансии", list(map(lambda x: x[x.find('#') + 1:], published_at_data)))
    info_table.add_autoindex('№')
    info_table.hrules = ALL
    info_table.align = 'l'
    info_table.max_width = 20
    print(info_table.get_string(start=start_end_nums[0], end=start_end_nums[1], fields=table_fields))


######################################################################################################################


def get_vacancies():
    """Получение информации с csv файла в виде таблицы с вакансиями на основе вводимых пользователем данных

    """
    input_requests = [
        "Введите название файла: ",
        "Введите параметр фильтрации: ",
        "Введите параметр сортировки: ",
        "Обратный порядок сортировки (Да / Нет): ",
        "Введите диапазон вывода: ",
        "Введите требуемые столбцы: "
    ]
    input_info: list[str | bool | list[str] | list[int]] = [input(input_request) for input_request in input_requests]
    normalize_result = normalize_input_info(input_info)
    if normalize_result != "Нормализация прошла успешно":
        return normalize_result
    (headers, info) = csv_reader(input_info[0])
    filtered_csv_data = csv_filter(headers, info)
    if len(filtered_csv_data) == 0:
        return "Нет данных"
    formatted_info = info_formatter(filtered_csv_data)
    filtered_info = info_filter(formatted_info, input_info[1])
    if len(filtered_info) == 0:
        return "Ничего не найдено"
    if input_info[2] != '№':
        filtered_info = info_sorter(filtered_info, input_info[2], input_info[3])
    print_vacancies(filtered_info, input_info[4], input_info[5])
