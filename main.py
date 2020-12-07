# -*- coding: utf-8 -*-

import json
import time
import requests
import openpyxl


wb = openpyxl.Workbook()
sht = wb['Sheet']
names_job = ['Кассир', 'Грузчик', 'Мерчендайзер']
metro_lines = json.loads(requests.get('https://api.hh.ru/metro/1').content.decode())['lines']
global all_vacancies
all_vacancies = []
nums = [0, 1, 2, 3, 4]
chars = ['A', 'B', 'C', 'D', 'E']

def getPage(page, name_job, metro, emp_num):
    params = {
        'text': 'NAME:' + name_job,
        'area': 1,
        'page': page,
        'per_page': 100,
        'metro': metro,
        'employment': emp_num
    }
    req = requests.get('https://api.hh.ru/vacancies', params)
    data = req.content.decode()
    req.close()
    time.sleep(0.25)
    return data

def getVacancies(name_job, station, emp_num):
    vacancies_list = []
    for page in range(20):
        vacancies = json.loads(getPage(page, name_job, station['id'], emp_num))['items']
        if not vacancies:
            return vacancies_list
        for vacancie in vacancies:
                vacancies_list.append([vacancie['id'], name_job, vacancie['name'], emp_num, station['name']])

def getVacinciesFromMetroLine(name_job, metro_line):
    vacancies = []
    for station in metro_line['stations']:
        for emp_num in ['full', 'part']:
            vacancies += getVacancies(name_job, station, emp_num)
    global all_vacancies
    all_vacancies += vacancies
    print(len(all_vacancies))


if __name__ == '__main__':
    for name_job in names_job:
        for metro_line in metro_lines:
            getVacinciesFromMetroLine(name_job, metro_line)
    for counter_rows in range(1, len(all_vacancies) + 1):
        for num, char in zip(nums, chars):
            sht[char + str(counter_rows)] = all_vacancies[counter_rows - 1][num]
    wb.save('test.xlsx')
