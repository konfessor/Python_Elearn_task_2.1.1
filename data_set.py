import csv

class DataSet:
    def __init__(self, file_name: str):
        self.__file_name = file_name
        self.__data = self.universal_csv_parser()
        self.__vacancies_objects = [Vacancy(item) for item in self.data]

    @property
    def file_name(self):
        return self.__file_name

    @file_name.setter
    def file_name(self, file_name):
        self.__file_name = file_name

    @property
    def data(self):
        return self.__data

    @data.setter
    def data(self, data):
        self.__data = data

    @property
    def vacancies_objects(self):
        return self.__vacancies_objects

    @vacancies_objects.setter
    def vacancies_objects(self, vacancies_objects):
        self.__vacancies_objects = vacancies_objects

    def universal_csv_parser(self) -> list:
        with open(self.file_name, "r", encoding="utf-8-sig") as file:
            reader = csv.reader(file)
            header = []
            info = []
            for line in reader:
                if (header == []):
                    header = line
                    lineLength = len(header)
                else:
                    if ("" not in line and len(line) == lineLength):
                        info.append(dict(zip(header, line)))
        if (header == []):
            print("Пустой файл")
            exit()
        return info
