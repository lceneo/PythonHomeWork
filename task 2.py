import csv
import datetime
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles.borders import Border, Side
import matplotlib.pyplot as plt
from builtins import input
import re


class Vacancy:
    currency_to_rub = {
        "KGS": 0.76,
        "AZN": 35.68,
        "BYR": 23.91,
        "EUR": 59.90,
        "GEL": 21.74,
        "KZT": 0.13,
        "RUR": 1,
        "UAH": 1.64,
        "USD": 60.66,
        "UZS": 0.0055,
    }

    def __init__(self, object_vacancy):
        self.name = object_vacancy['name']
        salary_from = (int)((float)("".join(object_vacancy['salary_from'].split())))
        salary_to = (int)((float)("".join(object_vacancy['salary_to'].split())))
        self.salary = (salary_from + salary_to) * self.currency_to_rub[object_vacancy['salary_currency']] // 2
        self.area_name = object_vacancy['area_name']
        self.published_at = datetime.datetime.strptime(object_vacancy['published_at'], '%Y-%m-%dT%H:%M:%S%z')


class DataSet:
    def __init__(self, file_name: str, vacancies_objects: list):
        self.file_name = file_name
        self.vacancies_objects = vacancies_objects

    def refactor_html(self, raw_html):
       return(re.compile('<.*?>'), '', raw_html)

    def csv_read(self):
        vacancies = []
        name_list = []
        with open(self.file_name, encoding='utf-8-sig') as r_file:
            file_reader = csv.reader(r_file, delimiter=",")
            count = 0
            for row in file_reader:
                if count == 0:
                    count += 1
                    name_list = row
                else:
                    if "" in row or len(row) != len(name_list):
                        continue
                    vacancies.append(row)
        if len(name_list) == 0:
            print('Пустой файл')
            exit()
        if len(vacancies) == 0:
            print('Нет данных')
            exit()
        return (vacancies, name_list)

    def fill_vacancies(self):
        (vacancies, list_naiming) = self.csv_read()
        self.vacancies_objects = self.csv_filer(vacancies, list_naiming)
        
    def csv_filer(self, reader, list_naming):
        vacancies = list()
        for row in reader:
            current = {}
            for i in range(len(row)):
                current[list_naming[i]] = row[i]
            vacancies.append(Vacancy(current))
        return vacancies



class Tuple:
    totalSalary = 0
    count = 0
    def __init__(self, totalSalary: int, count: int):
        self.totalSalary = totalSalary
        self.count = count


class InputData:
    years_stats = {
    }

    cities_stats = {
    }

    vacancy_stats = {
    }

    def start_input(self):
        self.file_name = input('Введите название файла: ')
        self.profession = input('Введите название профессии: ')
        self.city_count = 0

    def count_vacancies(self, vacancies: list):
        for vacancy in vacancies:
            self.city_count += 1
            year = int(vacancy.published_at.year)

            if vacancy.area_name not in self.cities_stats.keys():
                self.cities_stats[vacancy.area_name] = Tuple(vacancy.salary, 1)
            else:
                self.cities_stats[vacancy.area_name].totalSalary += vacancy.salary
                self.cities_stats[vacancy.area_name].count += 1
                
            if year not in self.years_stats.keys():
                self.years_stats[year] = Tuple(vacancy.salary, 1)
                self.vacancy_stats[year] = Tuple(0, 0)
            else:
                self.years_stats[year].totalSalary += vacancy.salary
                self.years_stats[year].count += 1

            if self.profession in vacancy.name:
                self.vacancy_stats[year].totalSalary += vacancy.salary
                self.vacancy_stats[year].count += 1

    def update_stats(self):
        for year in self.years_stats.keys():
            self.years_stats[year].totalSalary = int(self.years_stats[year].totalSalary // self.years_stats[year].count)

        remove_data = list()
        for city in self.cities_stats.keys():
            percentage = round(self.cities_stats[city].count / self.city_count, 4)
            if percentage < 0.01:
                remove_data.append(city)
            else:
                self.cities_stats[city].totalSalary = int(self.cities_stats[city].totalSalary // self.cities_stats[city].count)
                self.cities_stats[city].count = percentage
        for year in self.vacancy_stats.keys():
            if self.vacancy_stats[year].count != 0:
                self.vacancy_stats[year].totalSalary = int(self.vacancy_stats[year].totalSalary // self.vacancy_stats[year].count)
        for city in remove_data:
            del [self.cities_stats[city]]

    
    def print_info(self, str_info: str, dict: dict, value_name: str):
        marker = False
        print(str_info, end='')
        ind = 0
        for year in dict.keys():
            if ind == 0:
                print(' {', end='')
                ind += 1
            printEnd = ', '
            if year == max(dict.keys()):
                printEnd = ''
                marker = True
            print(f"{year}: {getattr(dict[year], value_name)}", end=printEnd)
        if marker:
            print('}')

    def get_city_print(self, str_data: str, dict: dict, names: list, value_name):
        flag = False
        print(str_data, end='')
        ind = 0
        for name in names:
            if ind == 0:
                print(' {', end='')
            printEnd = ', '
            if ind == len(names) - 1:
                printEnd = ''
                flag = True
            print(f"'{name}': {getattr(dict[name], value_name)}", end=printEnd)
            ind += 1
        if flag:
            print('}')

    def get_answer(self):

        cities_sorted = sorted(self.cities_stats, key=lambda x: self.cities_stats[x].totalSalary, reverse=True)
        self.print_info("Динамика уровня зарплат по годам:", self.years_stats, "totalSalary")
        self.print_info("Динамика уровня зарплат по годам для выбранной профессии:", self.vacancy_stats, "totalSalary")

        self.print_info("Динамика количества вакансий по годам:", self.years_stats, "count")

        self.print_info("Динамика количества вакансий по годам для выбранной профессии:", self.vacancy_stats, "count")
        del cities_sorted[10:]
        self.get_city_print("Уровень зарплат по городам (в порядке убывания):", self.cities_stats,
                              cities_sorted, "totalSalary")
        cities_sorted = sorted(self.cities_stats, key=lambda x: self.cities_stats[x].count, reverse=True)
        del cities_sorted[10:]
        self.get_city_print("Доля вакансий по городам (в порядке убывания):", self.cities_stats,
                              cities_sorted, "count")

    def get_sorted_cities(self, attr_name: str):
        current = {}
        sorted_names = sorted(self.cities_stats, key=lambda x: getattr(self.cities_stats[x], attr_name), reverse=True)
        del sorted_names[10:]
        for name in sorted_names:
            current[name] = self.cities_stats[name]
        return current


class Report:
    def generate_image(self, inputer: InputData, filename: str):
        plt.rcParams.update({'font.size': 8})
        fig, axis = plt.subplots(2, 2)
        self.define_graphs(inputer, filename, fig, axis)
        plt.savefig(filename)
    

    def vertical_graph(self, axis, value_x, value_y_1, value_y_2, name_first, name_second, name):
        axis.grid(axis="y")
        axis.set_title(name, fontsize=16)
        axis.bar([v - 0.2 for v in value_x], value_y_1, label=name_second, width=0.6)
        axis.bar([v + 0.2 for v in value_x], value_y_2, label=name_first, width=0.6)
        axis.tick_params(axis="x", labelrotation=90)
        axis.legend()

    
    def cirle_diagramm(self, axis, names: list, values: list, name):
        axis.set_title(name, fontsize=18)
        names.append("Другие")
        values.append(1 - sum(values))
        axis.pie(values, labels=names)
        plt.axis('equal')

    def zontal_graph(self, axis, value_x, y_val, name):
        axis.set_title(name, fontsize=18)
        axis.grid(axis="x")
        axis.barh(value_x, y_val)
        axis.invert_yaxis()
    
    def define_graphs(self, inputer: InputData, filename: str, fig, axis):
        self.vertical_graph(axis[0, 1], inputer.years_stats.keys(),
                              [inputer.years_stats[key].count for key in inputer.years_stats],
                              [inputer.vacancy_stats[key].count for key in inputer.vacancy_stats],
                        "Количество ваканский", f"Количество ваканский {inputer.profession}", "Количество вакансий по годам")
        self.vertical_graph(axis[0, 0], inputer.years_stats.keys(),
                              [inputer.years_stats[key].totalSalary for key in inputer.years_stats],
                              [inputer.vacancy_stats[key].totalSalary for key in inputer.vacancy_stats],
                        "Средняя з/п", f"з/п {inputer.profession}", "Уровень зарплат по годам")
        sorted_cities_by_salary =  inputer.get_sorted_cities("totalSalary")
        self.zontal_graph(axis[1, 0], [key for key in sorted_cities_by_salary],
                                  [sorted_cities_by_salary[key].totalSalary for key in sorted_cities_by_salary],
                            "Уровень зарплат по городам")
        sorted_cities_by_count = inputer.get_sorted_cities("count")
        self.cirle_diagramm(axis[1, 1], [key for key in sorted_cities_by_salary],
                                 [sorted_cities_by_salary[key].count for key in sorted_cities_by_salary],
                             "Доля ваканский по городам")
        fig.tight_layout(h_pad=1)
        fig.set_size_inches(17, 10)


inputer = InputData()
inputer.start_input()
dataset = DataSet(inputer.file_name, list())
dataset.fill_vacancies()
inputer.count_vacancies(dataset.vacancies_objects)
inputer.update_stats()
inputer.get_answer()

report = Report()
report.generate_image(inputer, "graph.png")
