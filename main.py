from vacancy import get_vacancies
from statistics import get_statistics
from statistics import HHruApiConnect


def main_function():
    """Выбор типа анализа данных из csv-файла

    """
    main_input_request = "Выберите тип вывода: "
    # main_input_info = input(main_input_request)
    main_input_info = "Статистика"
    if main_input_info != "Вакансии" and main_input_info != "Статистика":
        print("Введён неправильный тип вывода")
        return
    # if main_input_info == "Вакансии":
    #     get_vacancies()
    # else:
    get_statistics()


def test():
    hh = HHruApiConnect()
    hh.save_vacancy_data_for_past_day()


if __name__ == '__main__':
    main_function()
