# использован Отчет по форме № 5-ТН за 2018 год, сводный в целом по Москве -
# https://www.nalog.ru/rn77/related_activities/statistics_and_analytics/forms/8994216/

import xlrd  # библиотека для работы с excel, установка - pip install xlrd
import pymongo  # библиотека для работы с mongodb(установить mongodb и создать базу данных python)
import plotly.express as px  # библиотека для визуализации данных и отображения их в браузере, использует также pandas
import json

keys = []  # создаем листы для строк и заголовков для далнейшего наполнения
rows = []
doc = xlrd.open_workbook('5TN_za_2018.xls')  # открываем документ
for sheet in doc.sheets():  # перебор всех листов в доке
    for row_num in range(sheet.nrows):  # перебор всех строк в листе
        row = sheet.row_values(row_num)
        if row[1] == '':  # должен быть не пустой второй столбец(так пропустим подзаголовки)
            continue
        if row[1].isdigit():  # если второй столбец числовой то это нужные данные
            rows.append(row)
        elif not keys:  # самый первый не числовой второй столбец это заголовки
            keys = row

dicts = []
for values in rows:
    dicts.append(dict(zip(keys, values)))  # из листов заголовков и данных создадим словарь для импорта в базу данных

# подключаемся к серверу монгодб предполагается что работает локально и без пароля, используем созданную базу "python"
client = pymongo.MongoClient('localhost', 27017)
db = client["python"]
col = db["python"]  # создаем коллекцию "python" если она еще не создана
col.drop()  # очищаем коллекцию перед добавлением
for my_dict in dicts:
    col.insert_one(my_dict)  # добавляем все наши данные

# получим данные из базы данных в json
data = []
cursor = col.find({}, {'_id': False})
for document in cursor:
    data.append(json.dumps(document))

indicators = []
indicator_values = []
for item in data:
    i = json.loads(item)
    i['Значение показателя'] = (round(i['Значение показателя']))  # округляем значение до целого
    if i['Значение показателя'] < 3000000:  # чтобы график не был нагруженный берем значения от 3 млн. и меньше 20 млн.
        continue
    if i['Значение показателя'] > 20000000:
        continue
    del i['Код строки']  # не нужные параметр не используем
    indicators.append(i['Показатели'])
    indicator_values.append(i['Значение показателя'])

fig = px.bar(data, x=indicators, y=indicator_values,
             hover_data=[indicators], color=indicators,
             labels={'x': 'Показатели', 'y': 'Значение показателя'})  # строим график и отображаем его
fig.show()
