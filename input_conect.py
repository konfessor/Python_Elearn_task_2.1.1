from math import floor

class InputConect:
    def __init__(self, data: list, profession_name: str):
        self.__data = data
        self.__profession_name = profession_name
        self.__vacancies_by_city = {}
        self.__vacancies_amount = 0

    @property
    def data(self):
        return self.__data

    @data.setter
    def data(self, data):
        self.__data = data

    @property
    def profession_name(self):
        return self.__profession_name

    @profession_name.setter
    def profession_name(self, profession_name):
        self.__profession_name = profession_name

    def get_salary_by_year(self):
        salary_by_year = {}
        salary_amount = {}
        for vacancy in self.data:
            if (vacancy.published_at not in salary_by_year.keys()):
                salary_by_year[vacancy.published_at] = (vacancy.salary_from + vacancy.salary_to) / 2 * currency_to_rub[vacancy.salary_currency]
                salary_amount[vacancy.published_at] = 1
            else:
                salary_by_year[vacancy.published_at] += (vacancy.salary_from + vacancy.salary_to) / 2 * currency_to_rub[vacancy.salary_currency]
                salary_amount[vacancy.published_at] += 1
        for year in salary_by_year.keys():
            salary_by_year[year] = int(salary_by_year[year] / salary_amount[year])
        return dict(sorted(salary_by_year.items(), key=lambda item: item[0]))

    def get_vacancies_by_year(self):
        vacancies_by_year = {}
        for vacancy in self.data:
            if (vacancy.published_at not in vacancies_by_year.keys()):
                vacancies_by_year[vacancy.published_at] = 1
            else:
                vacancies_by_year[vacancy.published_at] += 1
            if (vacancy.area_name in self.__vacancies_by_city.keys()):
                self.__vacancies_by_city[vacancy.area_name] += 1
            else:
                self.__vacancies_by_city[vacancy.area_name] = 1
            self.__vacancies_amount += 1
        return dict(sorted(vacancies_by_year.items(), key=lambda item: item[0]))

    def get_salary_by_year_for_profession(self):
        salary_by_year_for_profession = {}
        salary_amount = {}
        for vacancy in self.data:
            if (self.profession_name in vacancy.name):
                if (vacancy.published_at not in salary_by_year_for_profession.keys()):
                    salary_by_year_for_profession[vacancy.published_at] = (vacancy.salary_from + vacancy.salary_to) / 2 * currency_to_rub[vacancy.salary_currency]
                    salary_amount[vacancy.published_at] = 1
                else:
                    salary_by_year_for_profession[vacancy.published_at] += (vacancy.salary_from + vacancy.salary_to) / 2 * currency_to_rub[vacancy.salary_currency]
                    salary_amount[vacancy.published_at] += 1
        for year in salary_by_year_for_profession.keys():
            salary_by_year_for_profession[year] = int(salary_by_year_for_profession[year] / salary_amount[year])
        if (salary_by_year_for_profession == {}):
            salary_by_year_for_profession = {2022: 0}
        return dict(sorted(salary_by_year_for_profession.items(), key=lambda item: item[0]))

    def get_vacancies_by_year_for_profession(self):
        vacancy_by_year_for_profession = {}
        for vacancy in self.data:
            if (self.profession_name in vacancy.name):
                if (vacancy.published_at not in vacancy_by_year_for_profession.keys()):
                    vacancy_by_year_for_profession[vacancy.published_at] = 1
                else:
                    vacancy_by_year_for_profession[vacancy.published_at] += 1
        if (vacancy_by_year_for_profession == {}):
            vacancy_by_year_for_profession = {2022: 0}
        return dict(sorted(vacancy_by_year_for_profession.items(), key=lambda item: item[0]))

    def get_salary_by_city(self):
        salary_by_city = {}
        salary_amount = {}
        for vacancy in self.data:
            if (self.__vacancies_by_city[vacancy.area_name] >= floor(self.__vacancies_amount / 100)):
                if (vacancy.area_name in salary_by_city.keys()):
                    salary_by_city[vacancy.area_name] += (vacancy.salary_from + vacancy.salary_to) / 2 * currency_to_rub[vacancy.salary_currency]
                    salary_amount[vacancy.area_name] += 1
                else:
                    salary_by_city[vacancy.area_name] = (vacancy.salary_from + vacancy.salary_to) / 2 * currency_to_rub[vacancy.salary_currency]
                    salary_amount[vacancy.area_name] = 1
        for city in salary_by_city.keys():
            salary_by_city[city] = int(salary_by_city[city] / salary_amount[city])
        salary_by_city = dict(sorted(salary_by_city.items(), key=lambda item: item[1], reverse=True))
        return dict(zip(list(salary_by_city.keys())[:10], list(salary_by_city.values())[:10]))

    def get_vacancies_by_city(self):
        vacancies_by_city = {}
        for city in self.__vacancies_by_city.keys():
            if (self.__vacancies_by_city[city] >= floor(self.__vacancies_amount / 100)):
                vacancies_by_city[city] = round(self.__vacancies_by_city[city] / self.__vacancies_amount, 4)
        vacancies_by_city = dict(sorted(vacancies_by_city.items(), key=lambda item: item[1], reverse=True))
        return dict(zip(list(vacancies_by_city.keys())[:10], list(vacancies_by_city.values())[:10]))

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
