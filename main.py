import datetime

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


# Функция проверки на пустоту ячейки
def isnan(x):
    try:
        if x == '-':
            return True
        return np.isnan(x)
    except:
        return False


def view_profit(df, start_date=(2021, 1, 1), finish_date=(2100, 1, 1)):
    # На вход подаем бд и необходимый период
    # Функция сохраняет изображение с графиком в текущую директорию
    index = 0
    start_date = datetime.datetime(start_date[0], start_date[1], start_date[2], 0, 0)
    finish_date = datetime.datetime(finish_date[0], finish_date[1], finish_date[2], 0, 0)
    while index < len(df):
        if isnan(df[index]["receiving_date"]) or df[index]["receiving_date"] == "-":
            df.remove(df[index])
        else:
            index += 1
    df = sorted(df, key=lambda x: x["receiving_date"])
    profit = []
    profit_value = []
    for line in df:
        if start_date <= line["receiving_date"] <= finish_date:
            profit.append(line["receiving_date"])
            if profit_value:
                profit_value.append(profit_value[-1] + line["sum"])
            else:
                profit_value.append(line["sum"])
    plt.plot(profit, profit_value, color="green")
    plt.savefig('Ans2.png')


# Парсим excel таблицу
df = pd.read_excel("data.xlsx")
data_frame = df.to_dict(orient="records")
# Удаляем пустой столбец
for line in data_frame:
    line.pop("Unnamed: 5")
# решение первого вопроса
ans1 = 0
for line in data_frame:
    if line["status"] == "ОПЛАЧЕНО" and not isnan(line["receiving_date"]):
        if datetime.datetime(2021, 7, 1, 0, 0) <= line["receiving_date"] <= datetime.datetime(2021, 7, 31, 0, 0):
            ans1 += line["sum"]
print(f"Ответ на первый вопрос: {round(ans1, 2)}")
# Решение второго вопроса
view_profit(data_frame, (2021, 6, 1), (2021, 6, 30))
print("Ответ на второй вопрос лежит в файле \"Ans2.png\"")
# Решение третьего вопроса
managers = {}
for line in data_frame:
    if line["status"] == "ОПЛАЧЕНО" and not isnan(line["receiving_date"]):
        if datetime.datetime(2021, 9, 1, 0, 0) <= line["receiving_date"] <= datetime.datetime(2021, 9, 30, 0, 0):
            if line['sale'] in managers.keys():
                managers[line['sale']] += line['sum']
            else:
                managers[line['sale']] = line['sum']
print(f"Ответ на третий вопрос: {max(managers, key=lambda x: managers[x])}")
# Решение четвертого вопроса
count_new = 0
count_current = 0
for line in data_frame:
    if line["status"] == "ОПЛАЧЕНО" and not isnan(line["receiving_date"]):
        if datetime.datetime(2021, 10, 1, 0, 0) <= line["receiving_date"] <= datetime.datetime(2021, 10, 31, 0, 0):
            if line['new/current'] == 'новая':
                count_new += 1
            if line['new/current'] == 'текущая':
                count_current += 1
    ans4 = 'new' if count_new > count_current else 'current'
print(f"Ответ на четвертый вопрос: {ans4}")
# Решение пятого вопроса
ans5 = 0
for line in data_frame:
    if not isnan(line["receiving_date"]):
        if datetime.datetime(2021, 5, 1, 0, 0) <= line["receiving_date"] <= datetime.datetime(2021, 5, 31, 0, 0):
            ans5 += 1
print(f"Ответ на пятый вопрос: {ans5}")
# Решение задачи
managers = {}
for line in data_frame:
    if not isnan(line["receiving_date"]):
        if line["receiving_date"] >= datetime.datetime(2021, 7, 1, 0, 0):
            if line["receiving_date"] == "ОПЛАЧЕНО":
                summ = line['sum'] * 0.07
            elif line["receiving_date"] != "ПРОСРОЧЕНО":
                if line['sum'] > 10000:
                    summ = line['sum'] * 0.05
                else:
                    summ = line['sum'] * 0.03
            if line['sale'] in managers.keys():
                managers[line['sale']] += summ
            else:
                managers[line['sale']] = summ
print('Ответ на задачу')
for manager in managers.keys():
    print(f"{manager} - {managers[manager]}")
