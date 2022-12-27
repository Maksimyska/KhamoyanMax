from statistics import get_statistics
from vacancy import get_vacancies


def main_function():
    """Выбор типа анализа данных из csv-файла
    """
    main_input_request = "Выберите тип вывода: "
    main_input_info = input(main_input_request)
    # input_info = ["Вакансии"]
    if main_input_info != "Вакансии" and main_input_info != "Статистика":
        print("Введён неправильный тип вывода")
        return
    if main_input_info == "Вакансии":
        get_vacancies()
    else:
        get_statistics()
    # get_statistics()


if __name__ == '__main__':
    main_function()
