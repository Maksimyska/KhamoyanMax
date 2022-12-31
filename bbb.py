import csv
import operator
import multiprocessing
import pdfkit
import numpy as np
import matplotlib.pyplot as plt
from os import listdir, stat
from os.path import isfile, join
from functools import reduce, cmp_to_key
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
from jinja2 import Environment, FileSystemLoader


class Vacancy:
    """Класс Вакансий

    Attributes:
        name (str): Имя вакансии
        description (str):  Описание вакансии
        key_skills (list[str]): Скиллы, необходимые вакансии
        experience_id (str): Опыт, необходимый для вакансии
        premium (str): Премиум вакансия или нет
        employer_name (str): Имя компании
        salary (Salary): Оклад
        area_name (str): Местонахождение
        published_at (str): Время публикации вакансии
    """

    def __init__(self, name, description, key_skills, experience_id, premium,
                 employer_name, salary, area_name, published_at):
        """Инициализирует объект Vacancy

        Args:
            name (str):
            description (str | None):
            key_skills (list[str] | None):
            experience_id (str | None):
            premium (str | None):
            employer_name (str | None):
            salary (Salary):
            area_name (str):
            published_at (str):
        """
        self.name = name
        self.description = description
        self.key_skills = key_skills
        self.experience_id = experience_id
        self.premium = premium
        self.employer_name = employer_name
        self.salary = salary
        self.area_name = area_name
        self.published_at = published_at


class Salary:
    """Класс для представления оклада

    Attributes:
        salary_from (str):
        salary_to (str):
        salary_gross (str):
        salary_currency (str):
    """

    def __init__(self, salary_from, salary_to, salary_gross, salary_currency):
        """Инициализирует объект Salary

        Args:
            salary_from (str):
            salary_to (str):
            salary_gross (any):
            salary_currency (str):
        """
        self.salary_from = salary_from
        self.salary_to = salary_to
        self.salary_gross = salary_gross
        self.salary_currency = salary_currency

    def currency_to_rur(self):
        """Перевод в верхнии и нижние вилки оклада

        Returns:
            list[int,int]: Верхняя и нижняя вилки оклада в рублях
        """
        currency_to_rub = {
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

        return list(map(lambda x: int(x.replace(' ', '')) * currency_to_rub[self.salary_currency],
                        (self.salary_from, self.salary_to)))

    def get_salary(self):
        """Получение среднего значения оклада

        Returns:
            float: Среднее значение оклада
        """
        return sum(self.currency_to_rur()) / 2


class DataSet:
    """Класс обработчик csv файла

    Attributes:
        file_name (str): Имя csv файла
        vacancies_objects (list[Vacancy]): Список вакансий полученных из csv файла
    """

    def __init__(self, file_name):
        """Инициализирует объект DataSet

        Args:
            file_name (str): Название файла
        """
        self.file_name = file_name
        self.headers = []

    def split_csv_by_year(self):
        """Разделение csv файла по годам.

        """
        years_info = {}
        with open(self.file_name, encoding="utf-8-sig") as f:
            for row in [x for x in csv.reader(f)]:
                if len(self.headers) == 0:
                    self.headers = row
                    continue
                if '' in row or len(row) != len(self.headers):
                    continue
                year = row[-1][0:4]
                if year in years_info:
                    years_info[year].append(row)
                else:
                    years_info[year] = [row]
            f.close()
        for year, info in years_info.items():
            with open(f"years/{year}.csv", mode="w", encoding='utf-8-sig') as csv_year:
                file_writer = csv.writer(csv_year, delimiter=",", lineterminator="\r")
                file_writer.writerow(self.headers)
                file_writer.writerows(info)
                csv_year.close()

    def get_year_file_names(self, folder_path):
        """Получение названия файлов из определённой папки

        Args:
            folder_path (str): Название папки, из которой нужно брать имена файлов

        Returns:
            list[str]: Названия файлов
        """
        return [f for f in listdir(folder_path) if isfile(join(folder_path, f))]

    def get_vacancies_from_file(self, csv_year_file_name):
        """Чтение информации из csv файла определённого года и запись в список списков, в котором каждому внутреннему
            списку соответствует одна строка из файла

        Args:
            csv_year_file_name (str): Название csv файла определённого года

        Returns:
            list[list[str]]: Форматированный список вакансий
        """
        info = self._read_csv(csv_year_file_name)[1:]
        return self._create_vacancies(info)

    def _read_csv(self, file_name):
        """Чтение информации из csv файла я запись в список списков, в котором каждому внутреннему списку
            соответствует одна строка из файла

        Args:
            file_name (str): Название csv файла

        Returns:
            list[list[str]]: Форматированный список вакансий
        """
        with open(f"years/{file_name}", encoding="utf-8-sig") as f:
            reader_info = [x for x in csv.reader(f)]
            reader_info.pop(0)
            f.close()
        return reader_info

    def _create_vacancies(self, info):
        """Преобразование данных из csv файла в список вакансий, в котором каждой вакансии соответствует одна строка
            из файла (для статистики)

        Args:
            info (list[list[str]]): Строки csv файла

        Returns:
            list[Vacancy]: Форматированный список вакансий
        """
        return [Vacancy(info_row[0], None, None, None, None, None, Salary(info_row[1], info_row[2], None, info_row[3]),
                        info_row[4], info_row[5]) for info_row in info]


class InputConnect:
    """Класс для работы над списком Vacancy: полное форматирование, нахождения необходимых вакансий

    """
    def info_formatter(self, vacancies):
        """Нормализация данных в вакансиях

        Args:
            vacancies (list[Vacancy]): Список вакансий

        Returns:
            list[Vacancy]: Результат форматирования
        """
        def formatter_string_number(str_num):
            """Устранение дробных разделителей в строковом числе

            Args:
                str_num (str): Число для нормализации

            Returns:
                str: Результат форматирования числа
            """
            return str_num if str_num.find('.') == -1 else str_num[:len(str_num) - 2]

        def formatter_salary(attr_value):
            """Преобразование оклада в нормированный вид

            Args:
                attr_value (Salary): Объект оклада

            Returns:
                Salary: Результат форматирования оклада
            """
            salary_from = formatter_string_number(attr_value.salary_from)
            salary_to = formatter_string_number(attr_value.salary_to)
            salary_currency = dic_currency[attr_value.salary_currency]
            return Salary(salary_from, salary_to, None, salary_currency)

        def formatter_published_at(attr_value):
            """Получение года из строки, содержащей дату

            Args:
                attr_value (str): Значение времени публикации вакансии

            Returns:
                str: Год публикации
            """
            return attr_value[0:4]

        def format_vacancy(vacancy):
            """Форматирование аттрибутов вакансии

            Args:
                vacancy (Vacancy): Вакансия

            Returns:
                Vacancy: Вакансия с отформатированными аттрибутами
            """
            setattr(vacancy, "salary", formatter_salary(getattr(vacancy, "salary")))
            setattr(vacancy, "published_at", formatter_published_at(getattr(vacancy, "published_at")))
            return vacancy

        dic_currency = {"AZN": "Манаты", "BYR": "Белорусские рубли", "EUR": "Евро",
                        "GEL": "Грузинский лари", "KGS": "Киргизский сом", "KZT": "Тенге", "RUR": "Рубли",
                        "UAH": "Гривны", "USD": "Доллары", "UZS": "Узбекский сум"}

        # for vacancy in vacancies:
        #     setattr(vacancy, "salary", formatter_salary(getattr(vacancy, "salary")))
        #     setattr(vacancy, "published_at", formatter_published_at(getattr(vacancy, "published_at")))

        return [format_vacancy(vacancy) for vacancy in vacancies]

    def year_info_finder(self, vacancies, finder_parameter, year_str):
        """Формирование информации по годам о вакансиях: уровень зарплат по годам, уровень зарплат по годам для
            выбранной вакансии, количество вакансий по годам, количество вакансий по годам для выбранной вакансии,
            уровень зарплат по городам, количество вакансий по городам, общее количество вакансий

        Args:
            vacancies (list[Vacancy]): Список вакансий
            finder_parameter (str): Название вакансии в качестве параметра фильтрации
            year_str (str): Год

        Returns:
            tuple[ dict[int: tuple[int, int]], dict[int: tuple[int, int]], dict[int: int], dict[int: int] ]: Группа
                списков
        """
        year = int(year_str)
        salaries_year_level, selected_salary_year_level, vacancies_year_count, selected_vacancy_year_count, = \
            {}, {}, {}, {}
        for vacancy in vacancies:
            salary = vacancy.salary.get_salary()
            previous_value = salaries_year_level.get(year, (0, 0))
            salaries_year_level[year] = (previous_value[0] + salary, previous_value[1] + 1)
            vacancies_year_count[year] = vacancies_year_count.get(year, 0) + 1
            if finder_parameter in vacancy.name:
                previous_value = selected_salary_year_level.get(year, (0, 0))
                selected_salary_year_level[year] = (previous_value[0] + salary, previous_value[1] + 1)
                selected_vacancy_year_count[year] = selected_vacancy_year_count.get(year, 0) + 1
        return self._year_info_calculating(salaries_year_level, selected_salary_year_level, vacancies_year_count,
                                           selected_vacancy_year_count)

    def city_info_finder(self, vacancies):
        """Формирование информации по годам о вакансиях: уровень зарплат по годам, уровень зарплат по годам для
            выбранной вакансии, количество вакансий по годам, количество вакансий по годам для выбранной вакансии,
            уровень зарплат по городам, количество вакансий по городам, общее количество вакансий

        Args:
            vacancies (list[Vacancy]): Список вакансий

        Returns:
            tuple[ dict[str: tuple[int, int]], dict[str: int] ]: Группа списков
        """
        salaries_city_level, vacancies_city_count = {}, {}
        for vacancy in vacancies:
            salary = vacancy.salary.get_salary()
            previous_value = salaries_city_level.get(vacancy.area_name, (0, 0))
            salaries_city_level[vacancy.area_name] = (previous_value[0] + salary, previous_value[1] + 1)
            vacancies_city_count[vacancy.area_name] = vacancies_city_count.get(vacancy.area_name, 0) + 1
        return self._city_info_calculating(salaries_city_level, vacancies_city_count, len(vacancies))

    def _year_info_calculating(self, salaries_year_level, selected_salary_year_level, vacancies_year_count,
                               selected_vacancy_year_count):
        """Окончательное форматирование словарей, фильтрация, сортировка, выборка первого десятка для некоторых

        Args:
            salaries_year_level (dict[int: tuple[int, int]]): Уровень зарплат по годам
            selected_salary_year_level (dict[int: tuple[int, int]]): Уровень зарплат по годам для выбранной вакансии
            vacancies_year_count (dict[int: int]): Количество вакансий по годам
            selected_vacancy_year_count (dict[int: int]): Количество вакансий по годам для выбранной вакансии

        Returns:
            tuple[ dict[int: tuple[int, int]], dict[int: tuple[int, int]], dict[int: int], dict[int: int] ]:
                Уровень зарплат по годам, Уровень зарплат по годам для выбранной вакансии, Количество вакансий по годам,
                Количество вакансий по годам для выбранной вакансии
        """

        (salaries_year_level, selected_salary_year_level) = [
                {dict_pair[0]: int(dict_pair[1][0] / dict_pair[1][1]) if dict_pair[1][1] != 0 else int(dict_pair[1][0])
                    for dict_pair in dictionary.items()}
                for dictionary in (salaries_year_level, selected_salary_year_level)]

        # (salaries_year_level, selected_salary_year_level) = \
        #     list(map(lambda dictionary:
        #              dict(map(lambda dict_pair:
        #                       (dict_pair[0], int(dict_pair[1][0] / dict_pair[1][1])
        #                       if dict_pair[1][1] != 0
        #                       else int(dict_pair[1][0])), dictionary.items())),
        #              (salaries_year_level, selected_salary_year_level)))
        return salaries_year_level, selected_salary_year_level, vacancies_year_count, selected_vacancy_year_count

    @staticmethod
    def _city_info_calculating(salaries_city_level, vacancies_city_count, vacancies_count):
        """Окончательное форматирование словарей, фильтрация, сортировка, выборка первого десятка для некоторых

        Args:
            salaries_city_level (dict[str: tuple[int, int]]): Уровень зарплат по городам
            vacancies_city_count (dict[str: int]): Количество вакансий по городам
            vacancies_count (int): Общее количество вакансий

        Returns:
            tuple[ dict[str: tuple[int, int]], dict[str: int] ]: Уровень зарплат по городам, Количество вакансий
                по городам
        """
        def sort_dict(dictionary):
            """Сортировка словаря лексикографически

            Args:
                dictionary (dict[str: int]): Словарь для сортировки

            Returns:
                tuple[dict[str: int], dict[str: int]]: Отсортированный словарь
            """
            dict_pairs = [(key, value) for key, value in dictionary.items()]
            dict_pairs.sort(key=cmp_to_key(lambda x, y: -1 if x[1] <= y[1] else 1))
            return dict(dict_pairs)

        vacancies_city_count = {dict_pair[0]: float(f"{dict_pair[1] / vacancies_count:.4f}")
                                for dict_pair in vacancies_city_count.items()}
        vacancies_city_count = {dict_pair[0]: dict_pair[1] for dict_pair in vacancies_city_count.items() if dict_pair[1] >= 0.01}
        vacancies_city_count = sort_dict(vacancies_city_count)
        vacancies_city_count = {dict_pair[0]: f"{round(dict_pair[1] * 100, 2)}%" for dict_pair in vacancies_city_count.items()}
        salaries_city_level = {dict_pair[0]:
                                int(dict_pair[1][0] / dict_pair[1][1]) if dict_pair[1][1] != 0 else int(dict_pair[1][0])
                               for dict_pair in salaries_city_level.items()}
        salaries_city_level = {dict_pair[0]: dict_pair[1] for dict_pair in salaries_city_level.items() if dict_pair[0] in vacancies_city_count}
        salaries_city_level = sort_dict(salaries_city_level)


        # vacancies_city_count = dict(map(lambda dict_pair: (dict_pair[0],
        #                             float(f"{dict_pair[1] / vacancies_count:.4f}")), vacancies_city_count.items()))
        # vacancies_city_count = dict(filter(lambda dict_pair: dict_pair[1] >= 0.01, vacancies_city_count.items()))
        # vacancies_city_count = sort_dict(vacancies_city_count)
        # vacancies_city_count = dict(map(lambda dict_pair: (dict_pair[0], f"{round(dict_pair[1] * 100, 2)}%"),
        #                                 vacancies_city_count.items()))
        # salaries_city_level = dict(map(lambda dict_pair: (dict_pair[0], int(dict_pair[1][0] / dict_pair[1][1])
        #                                if dict_pair[1][1] != 0 else int(dict_pair[1][0])), salaries_city_level.items()))
        # salaries_city_level = dict(filter(lambda dict_pair: dict_pair[0] in vacancies_city_count,
        #                                   salaries_city_level.items()))
        # salaries_city_level = sort_dict(salaries_city_level)

        vacancies_city_count = {k: vacancies_city_count[k] for k in list(vacancies_city_count)[-10:][::-1]}
        salaries_city_level = {k: salaries_city_level[k] for k in list(salaries_city_level)[-10:][::-1]}
        return salaries_city_level, vacancies_city_count


class Report:
    """Класс для генерации файлов по анализу статистики: графиков, excel таблиц, общего pdf-файла

    Attributes:
        salaries_year_level (dict[int: tuple[int, int]]): Уровень зарплат по годам
        selected_salary_year_level (dict[int: tuple[int, int]]): Уровень зарплат по годам для выбранной вакансии
        vacancies_year_count (dict[int: int]): Количество вакансий по годам
        selected_vacancy_year_count (dict[int: int]): Количество вакансий по годам для выбранной вакансии
        salaries_city_level (dict[str: tuple[int, int]]): Уровень зарплат по городам
        vacancies_city_count (dict[str: int]): Количество вакансий по городам
    """
    def __init__(self, vacancy_info):
        """Инициализация объекта Report

        Args:
            vacancy_info (tuple[dict[int: tuple[int, int]], dict[int: tuple[int, int]], dict[int: int],
             dict[int: int], dict[str: tuple[int, int]], dict[str: int]]):
             Все словари созданные методом info_finder класса Input_Connect
        """
        self.salaries_year_level = vacancy_info[0]
        self.vacancies_year_count = vacancy_info[1]
        self.selected_salary_year_level = vacancy_info[2]
        self.selected_vacancy_year_count = vacancy_info[3]
        self.salaries_city_level = vacancy_info[4]
        self.vacancies_city_count = vacancy_info[5]

    def print_statistics(self):
        """Выводит на печать все статистику

        """
        print("Динамика уровня зарплат по годам:", self.salaries_year_level)
        print("Динамика количества вакансий по годам:", self.vacancies_year_count)
        print("Динамика уровня зарплат по годам для выбранной профессии:", self.selected_salary_year_level)
        print("Динамика количества вакансий по годам для выбранной профессии:", self.selected_vacancy_year_count)
        print("Уровень зарплат по городам (в порядке убывания):", self.salaries_city_level)
        print("Доля вакансий по городам (в порядке убывания):", self.vacancies_city_count)

    def generate_excel(self, vacancy_name):
        """Создание excel-файла основываясь на словарях аттрибутов объекта Report

        Args:
            vacancy_name (str): Название выбранной вакансии
        """
        wb = Workbook()
        stats_by_year = wb.worksheets[0]
        stats_by_year.title = "Cтатистика по годам"
        stats_by_city = wb.create_sheet("Cтатистика по городам")

        stats_by_year.append(["Год", "Средняя зарплата", f"Средняя зарплата - {vacancy_name}",
                              "Количество вакансий", f"Количество вакансий - {vacancy_name}"])
        for i, year in enumerate(self.salaries_year_level.keys(), 2):
            stats_by_year.cell(row=i, column=1, value=year)
            for j, dictionary in enumerate((self.salaries_year_level, self.vacancies_year_count,
                                            self.selected_salary_year_level, self.selected_vacancy_year_count), 2):
                stats_by_year.cell(row=i, column=j, value=dictionary[year])

        stats_by_city.append(["Город", "Уровень зарплат", "", "Город", "Доля вакансий"])
        for i, city in enumerate(self.salaries_city_level.keys(), 2):
            stats_by_city.cell(row=i, column=1, value=city)
            stats_by_city.cell(row=i, column=2, value=self.salaries_city_level[city])
        for i, city in enumerate(self.vacancies_city_count.keys(), 2):
            stats_by_city.cell(row=i, column=4, value=city)
            stats_by_city.cell(row=i, column=5, value=self.vacancies_city_count[city])

        self._slylize_wb(wb)
        wb.save('report.xlsx')

    @staticmethod
    def _slylize_wb(wb):
        """Стилизация рабочего листа excel-файла

        Args:
            wb (Workbook): Excel-лист
        """
        bold_font = Font(bold=True)
        thin = Side(border_style="thin", color="000000")
        outline = Border(top=thin, left=thin, right=thin, bottom=thin)
        for worksheet in wb.worksheets:
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value) if cell.value is not None else "") for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = length + 3
            for cell in worksheet[1]:
                cell.font = bold_font
            for column in tuple(worksheet.columns):
                if column[1].value is None:
                    continue
                for cell in column:
                    cell.border = outline

    def generate_image(self, vacancy_name):
        """Создание графиков основываясь на словарях аттрибутов объекта Report

        Args:
            vacancy_name (str): Название выбранной вакансии
        """
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 7.5), layout='constrained')
        self._generate_salary_year_levels_graph(ax1, vacancy_name)
        self._generate_vacancy_year_count_graph(ax2, vacancy_name)
        self._generate_salary_city_levels_graph(ax3)
        self._generate_vacancy_city_count_graph(ax4)
        plt.savefig('graph.png')

    def _generate_salary_year_levels_graph(self, ax, vacancy_name):
        """Создание графика уровня зарплат по годам

        Args:
            ax (Ax): Объект графика
            vacancy_name (str): Название выбранной вакансии
        """
        ax_labels = self.salaries_year_level.keys()
        x = np.arange(len(ax_labels))
        width = 0.35
        ax.bar(x - width / 2, self.salaries_year_level.values(), width, label='Средняя з/п')
        ax.bar(x + width / 2, self.selected_salary_year_level.values(), width, label=f'З/п {vacancy_name}')
        ax.set_xticks(x, ax_labels, fontsize=8, rotation=90, ha='right')
        ax.set_title("Уровень зарплат по годам")
        ax.yaxis.grid(True)
        ax.legend(fontsize=8, loc='upper left')

    def _generate_vacancy_year_count_graph(self, ax, vacancy_name):
        """Создание графика количества вакансий по годам

        Args:
            ax (Ax): Объект графика
            vacancy_name (str): Название выбранной вакансии
        """
        ax_labels = self.vacancies_year_count.keys()
        x = np.arange(len(ax_labels))
        width = 0.35
        ax.bar(x - width / 2, self.vacancies_year_count.values(), width, label='Количество вакансий')
        ax.bar(x + width / 2, self.selected_vacancy_year_count.values(), label=f'Количество вакансий {vacancy_name}')
        ax.set_xticks(x, ax_labels, fontsize=8, rotation=90, ha='right')
        ax.set_title("Количество вакансий по годам")
        ax.yaxis.grid(True)
        ax.legend(fontsize=8, loc='upper left')

    def _generate_salary_city_levels_graph(self, ax):
        """Создание графика уровня зарплат по городам

        Args:
            ax (Ax): Объект графика
        """
        ax_labels = self.salaries_city_level.keys()
        y_pos = np.arange(len(ax_labels))
        ax.barh(y_pos, self.salaries_city_level.values(), align='center')
        ax.set_yticks(y_pos, fontsize=8, labels=ax_labels)
        ax.invert_yaxis()
        ax.set_title("Уровень зарплат по городам")

    def _generate_vacancy_city_count_graph(self, ax):
        """Создание графика количества вакансий по городам

        Args:
            ax (Ax): Объект графика
        """
        ax_labels, values = list(self.vacancies_city_count.keys()), self.vacancies_city_count.values()
        ax_labels.append('Другие')
        values = list(map(lambda value: float(value[:-1]), values))
        values.append(100 - sum(values))
        ax.pie(values, labels=ax_labels)
        ax.set_title("Доля вакансий по городам")

    def generate_pdf(self, vacancy_name):
        """Создание pdf-файла основываясь на словарях аттрибутов объекта Report

        Args:
            vacancy_name (str): Название выбранной вакансии
        """
        headers1, headers2, headers3 = (["Год", "Средняя зарплата", f"Средняя зарплата - {vacancy_name}",
                                        "Количество вакансий", f"Количество вакансий - {vacancy_name}"],
                                        ["Город", "Уровень зарплат"], ["Город", "Доля вакансий"])
        rows1 = list(map(lambda year: [year] + [dictionary[year] for dictionary in
                                       (self.salaries_year_level, self.vacancies_year_count,
                                        self.selected_salary_year_level, self.selected_vacancy_year_count)]
                         , self.salaries_year_level.keys()))
        rows2 = list(map(lambda city: [city, self.salaries_city_level[city]], self.salaries_city_level.keys()))
        rows3 = list(map(lambda city: [city, self.vacancies_city_count[city]], self.vacancies_city_count.keys()))

        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template("pdf_template.html")
        pdf_template = template.render(graph_name='graph.png',
                                       vacancy_name=vacancy_name, headers1=headers1, headers2=headers2,
                                       headers3=headers3, rows1=rows1, rows2=rows2, rows3=rows3)
        config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')
        options = {'enable-local-file-access': None}
        pdfkit.from_string(pdf_template, 'report.pdf', options=options, configuration=config)


######################################################################################################################


class Consumer(multiprocessing.Process):
    """Служит для представления одного процесса, который берёт одну задачу из очереди задач и после выполнения кладёт
        результат в очереди результатов

    Attributes:
        task_queue (multiprocessing.JoinableQueue): Очередь задач
        vacancies_info (multiprocessing.Queue): Очередь результатов для информации о вакансиях
        statistics_info (multiprocessing.Queue): Очередь результатов для статистики
    """
    def __init__(self, task_queue, vacancies_info, statistics_info):
        """Инициализация объекта Consumer

        Args:
            task_queue (multiprocessing.JoinableQueue): Очередь задач
            vacancies_info (multiprocessing.Queue): Очередь результатов для информации о вакансиях
            statistics_info (multiprocessing.Queue): Очередь результатов для статистики
        """
        multiprocessing.Process.__init__(self)
        self.task_queue = task_queue
        self.vacancies_info = vacancies_info
        self.statistics_info = statistics_info

    def run(self):
        """Выполняет одну задачу, полученную из списка задач, и сохраняет результат в соответствующих
            очередях результатов

        """
        while True:
            temp_task = self.task_queue.get()

            if temp_task is None:
                self.task_queue.task_done()
                break

            answer = temp_task.process()
            self.task_queue.task_done()
            self.vacancies_info.put(answer[0])
            self.statistics_info.put(answer[1])

class Task():
    """Представляет собой одну задачу для выполнения процессом Consumer; составляет статистику о вакансиях
        по определённому году

    Attributes:
        file_name (str): Название файла, из которого нужно брать данные
        data_set (DadaSet): Объект DadaSet для анализа данных
        input_connect (InputConnect): Объект InputConnect для форматирования и составления статистики по данным
    """
    def __init__(self, file_name, data_set, input_connect):
        """Инициализирует одн объект класс Task

        Args:
            file_name (str): Название файла, из которого нужно брать данные
            data_set (DadaSet): Объект DadaSet для анализа данных
            input_connect (InputConnect): Объект InputConnect для форматирования и составления статистики по данным
        """
        self.file_name = file_name
        self.data_set = data_set
        self.input_connect = input_connect

    def process(self):
        """Служит списком команд, которые нужно будет выполнять процессу Consumer

        Returns:
            tuple[list[Vacancy], tuple[dict[int: tuple[int, int]], dict[int: tuple[int, int]],
             dict[int: int], dict[int: int]]: Список вакансий за соответствующий год и словари, содержащие статистику
        """
        vacancies_for_year = self.data_set.get_vacancies_from_file(self.file_name)
        formatted_info = self.input_connect.info_formatter(vacancies_for_year)
        return formatted_info, self.input_connect.year_info_finder(formatted_info, "Программист", self.file_name[:-4])


def get_statistics():
    """Получение информации с csv файла и создание графиков, таблиц и общего pdf-файл со статистикой
        на основе вводимых пользователем данных

    """
    def concat_dictionaries_in_tuples(tuples):
        """Используя группы словарей соединить каждый i-тый словарь

        Args:
            tuples (list[tuple[dict[int: int], dict[int: int], dict[int: int], dict[int: int]]]): Группы словарей для слияния

        Returns:
            tuple[dict[int: int], dict[int: int], dict[int: int], dict[int: int]]: Словари после слияния
        """
        (dict1, dict2, dict3, dict4) = {}, {}, {}, {}
        for dictionaries in tuples:
            dict1 = dict1 | dictionaries[0]
            dict2 = dict2 | dictionaries[1]
            dict3 = dict3 | dictionaries[2]
            dict4 = dict4 | dictionaries[3]
        return dict1, dict2, dict3, dict4

    def sort_dict_by_key(dictionary):
        """Сортировка словаря лексикографически по ключу

        Args:
            dictionary (dict[str: int]): Словарь для сортировки

        Returns:
            dict[str: int]: Отсортированный словарь
        """
        dict_pairs = [(key, value) for key, value in dictionary.items()]
        dict_pairs.sort(key=cmp_to_key(lambda x, y: -1 if x[0] <= y[0] else 1))
        return dict(dict_pairs)

    input_requests = ["Введите название файла: ", "Введите название профессии: "]
    # input_info = [input(input_request) for input_request in input_requests]
    input_info = ["vacancies_by_year.csv", "Программист"]
    if stat(input_info[0]).st_size == 0:
        print("Пустой файл")
        return

    data_set = DataSet(input_info[0])
    input_connect = InputConnect()
    # data_set.split_csv_by_year()
    year_file_names = data_set.get_year_file_names("years")

    tasks = multiprocessing.JoinableQueue()
    vacancies_info = multiprocessing.Queue()
    statistics_info = multiprocessing.Queue()
    consumers_count = multiprocessing.cpu_count()
    consumers = [Consumer(tasks, vacancies_info, statistics_info) for _ in range(consumers_count)]
    for consumer in consumers:
        consumer.start()
    for file_name in year_file_names:
        tasks.put(Task(file_name, data_set, input_connect))
    for _ in range(consumers_count):
        tasks.put(None)

    tasks.join()
    full_vacancies_info = [vacancies_info.get() for _ in range(len(year_file_names))]
    years_statistics = concat_dictionaries_in_tuples([statistics_info.get() for _ in range(len(year_file_names))])
    years_statistics = tuple(sort_dict_by_key(dictionary) for dictionary in years_statistics)
    cities_statistics = input_connect.city_info_finder(reduce(operator.concat, full_vacancies_info))
    report = Report(reduce(operator.concat, [years_statistics, cities_statistics]))
    report.print_statistics()
    # report.generate_excel(input_info[1])
    # report.generate_image(input_info[1])
    # report.generate_pdf(input_info[1])
