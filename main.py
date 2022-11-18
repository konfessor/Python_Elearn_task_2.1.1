currency_to_rub = {
    "AZN": 35.68,
    "BYR": 23.91,
    "EUR": 59.90,
    "GEL": 21.74,
    "KGS": 0.76,
    "KZT": 0.13,
    "RUR": 1,
    "UAH": 1.64,
    "USD": 60.66,
    "UZS": 0.0055,
}

print("Введите название файла: ", end="")
file_name = input()
print("Введите название профессии: ", end="")
profession_name = input()

data = DataSet(file_name).vacancies_objects
if (data == []):
    print("Нет данных")
else:
    result = InputConect(data, profession_name)

    salary_by_year = result.get_salary_by_year()
    vacancies_by_year = result.get_vacancies_by_year()
    salary_by_year_for_profession = result.get_salary_by_year_for_profession()
    vacancies_by_year_for_profession = result.get_vacancies_by_year_for_profession()
    salary_by_city = result.get_salary_by_city()
    vacancies_by_city = result.get_vacancies_by_city()

    Report(salary_by_year, vacancies_by_year, salary_by_year_for_profession,
           vacancies_by_year_for_profession, salary_by_city, vacancies_by_city).generate_excel()
