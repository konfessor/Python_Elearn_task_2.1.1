from data_set import DataSet
from input_conect import InputConect
from report import Report

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

    print(f"Динамика уровня зарплат по годам: {salary_by_year}")
    print(f"Динамика количества вакансий по годам: {vacancies_by_year}")
    print(f"Динамика уровня зарплат по годам для выбранной профессии: {salary_by_year_for_profession}")
    print(f"Динамика количества вакансий по годам для выбранной профессии: {vacancies_by_year_for_profession}")
    print(f"Уровень зарплат по городам (в порядке убывания): {salary_by_city}")
    print(f"Доля вакансий по городам (в порядке убывания): {vacancies_by_city}")

    report = Report(salary_by_year, vacancies_by_year, salary_by_year_for_profession,
                    vacancies_by_year_for_profession, salary_by_city, vacancies_by_city, profession_name)
    report.generate_excel()
    report.generate_image()
    report.generate_pdf()
