import pandas as pd
import requests
import fake_useragent
from datetime import timedelta


def get_order(name, date):
    rest_id = {
        'Белгород-1': 21,
        'Белгород-2': 583,
        'Белгород-3': 879,
        'Кириши-1': 870,
        'Мурманск-2': 1276,
        'Петергоф-1': 162,
        'Петрозаводск-1': 160,
        'Петрозаводск-2': 1128,
        'Сосновый бор-1': 330,
        'Шушары-1': 622
    }
    long_id = {
        'Белгород-1': '000D3A240C719A8711E68ABA13F81FAA',
        'Белгород-2': '000D3A2155A180E411E79C94155E4FA5',
        'Белгород-3': '000D3A24D2B7A94311E8BD9D5101A4B5',
        'Кириши-1': '000D3A29FF6BA94411E8A9D1E5F6EA20',
        'Мурманск-2': '000D3AAC9DCABB2E11EBEFC4A7C197B4',
        'Петергоф-1': '000D3A240C719A8711E68ABA13F8FFD4',
        'Петрозаводск-1': '000D3A240C719A8711E68ABA13F8FDB7',
        'Петрозаводск-2': '000D3A22FA54A81511EA943D4BBA8970',
        'Сосновый бор-1': '000D3A240C719A8711E68ABA13F9FDB2',
        'Шушары-1': '000D3A26B5B080F311E7CDBFAB091E32'
    }
    login = {
        'Белгород-1': 'dev_bel',
        'Белгород-2': 'dev_bel',
        'Белгород-3': 'dev_bel',
        'Кириши-1': 'dev_krs',
        'Мурманск-2': 'dev_mmk',
        'Петергоф-1': 'dev_ptf',
        'Петрозаводск-1': 'dev_ptz',
        'Петрозаводск-2': 'dev_ptz',
        'Сосновый бор-1': 'dev_sbr',
        'Шушары-1': 'dev_ssr'
    }
    password = {
        'Белгород-1': 'A42Et0Ik',
        'Белгород-2': 'A42Et0Ik',
        'Белгород-3': 'A42Et0Ik',
        'Кириши-1': 'il0y91I5',
        'Мурманск-2': 'q4E22ou5',
        'Петергоф-1': 'ZD585QrH',
        'Петрозаводск-1': 'i2Yml0q7',
        'Петрозаводск-2': 'i2Yml0q7',
        'Сосновый бор-1': 's23C8Dd7',
        'Шушары-1': 'xnH65V53'
    }
    session = requests.Session()
    user = fake_useragent.UserAgent().random
    log_data = {
        'CountryCode': 'Ru',
        'login': login[name],
        'password': password[name]
    }
    header = {
        'user-agent': user
    }
    log_link = 'https://auth.dodopizza.ru/Authenticate/LogOn'
    session.post(log_link, data=log_data, headers=header)
    orders_data = {
        'handover': {
            'link': 'https://officemanager.dodopizza.ru/Reports/OrderHandoverTime/Export',
            'data': {
                "unitsIds": long_id[name],
                "beginDate": date,
                "endDate": date,
                "orderTypes": "Delivery",
                "Export": "Экспорт+в+Excel"
            }
        },
        'delivery': {
            'link': 'https://officemanager.dodopizza.ru/Reports/CourierTasks/Export',
            'data': {
                "unitId": rest_id[name],
                "beginDate": date,
                "endDate": date,
                "statuses": [
                    "Queued",
                    "Ordering",
                    "Paused"
                ]
            }
        }
    }

    for order in orders_data:
        response = session.post(orders_data[order]['link'], data=orders_data[order]['data'], headers=header)
        try:
            with open(f'./export/{order}_{name}.xlsx', 'wb') as file:
                file.write(response.content)
                file.close()
        except FileNotFoundError as fne:
            print(str(fne))
    session.close()


def deltatime(t):
    if t < timedelta(0):
        t = timedelta(days=1) - t
        t = str(t).split(' ')[-1]
        t = '-' + t
    else:
        t = str(t).split(' ')[-1]
    return t


def deltatimedelta(x):
    try:
        x = timedelta(hours=x.hour, minutes=x.minute, seconds=x.second)
    except AttributeError:
        x = timedelta(0)
    return x


def change(name):
    df_delivery = pd.read_excel(f'./export/delivery_{name}.xlsx', skiprows=4)
    df_handover = pd.read_excel(f'./export/handover_{name}.xlsx', skiprows=6)

    counter = 0
    for i in df_delivery['Начало']:
        df_type = df_delivery.loc[df_delivery['Начало'] == i]
        if df_type.iloc[0]['Тип'] == 'На заказе':
            df_time = df_delivery.loc[df_delivery['Окончание'] == i]
            try:
                time = df_time.iloc[0]['Начало']
            except IndexError:
                time = 0
            df_delivery.at[counter, 'Постановка в очередь'] = time
        else:
            df_delivery.at[counter, 'Постановка в очередь'] = 0
        counter += 1
    df_delivery = df_delivery.loc[df_delivery['Тип'] == 'На заказе']

    df_orders_rec = pd.DataFrame()
    counter = 0
    for i in df_delivery['№ заказа']:
        m = i.split(', ')
        for j in m:
            df_orders_rec = pd.concat([df_orders_rec, df_delivery[df_delivery['№ заказа'] == i]], ignore_index=True)
            df_orders_rec.at[counter, 'orders'] = j
            counter += 1
    df_orders_rec = df_orders_rec[
        ['Начало', 'Длительность', 'Прогнозное время', '№ заказа', 'Количество заказов', 'Постановка в очередь',
         'orders', 'Фамилия']]
    df_orders_rec['orders'] = pd.to_numeric(df_orders_rec['orders'], errors='coerce')
    df_orders_rec['Постановка в очередь'] = df_orders_rec['Постановка в очередь'].apply(lambda x: str(x).split(' ')[-1])
    df_orders_rec['Начало'] = df_orders_rec['Начало'].apply(lambda x: str(x).split(' ')[-1])

    df_orders_rec['Длительность'] = df_orders_rec['Длительность'].apply(deltatimedelta)
    df_orders_rec['Прогнозное время'] = df_orders_rec['Прогнозное время'].apply(deltatimedelta)

    df_orders_rec['Разница поездки'] = df_orders_rec['Прогнозное время'] - df_orders_rec['Длительность']
    df_orders_rec['Разница поездки'] = df_orders_rec['Разница поездки'].apply(deltatime)

    df_orders_rec = df_orders_rec[
        ['orders', 'Постановка в очередь', 'Начало', 'Разница поездки', 'Количество заказов', 'Фамилия']
    ]
    df_orders_rec.rename(columns={'orders': 'Номер заказа', 'Начало': 'Начало поездки',
                                  'Количество заказов': 'Заказов'},
                         inplace=True)

    df_handover = df_handover[['Номер заказа', 'Дата и время', 'Ожидание', 'Приготовление', 'Ожидание на полке']]
    df_handover['Номер заказа'] = df_handover['Номер заказа'].apply(lambda x: x.split('-')[0])
    df_handover['Дата и время'] = df_handover['Дата и время'].apply(lambda x: x.replace(microsecond=0))
    df_handover['Ожидание'] = df_handover['Ожидание'].apply(deltatimedelta)
    df_handover['Приготовление'] = df_handover['Приготовление'].apply(deltatimedelta)
    df_handover['Ожидание на полке'] = df_handover['Ожидание на полке'].apply(deltatimedelta)
    df_handover['Ожидание на полке'] = df_handover['Ожидание на полке'].apply(
        lambda x: str(x).split(' ')[-1])
    df_handover['Заказ приготовился'] = df_handover['Ожидание'] + df_handover['Приготовление'] + df_handover[
        'Дата и время']
    df_handover['Заказ приготовился'] = df_handover['Заказ приготовился'].apply(lambda x: str(x).split(' ')[-1])
    df_handover['Отправка на заказ'] = df_handover['Ожидание'] + df_handover['Приготовление'] + df_handover[
        'Ожидание на полке'] + df_handover['Дата и время']
    df_handover['Отправка на заказ'] = df_handover['Отправка на заказ'].apply(lambda x: str(x).split(' ')[-1])
    df_handover['Дата и время'] = df_handover['Дата и время'].apply(lambda x: str(x).split(' ')[-1])
    df_handover = df_handover[
        ['Номер заказа', 'Дата и время', 'Отправка на заказ', 'Ожидание на полке', 'Заказ приготовился']]
    df_handover['Номер заказа'] = df_handover['Номер заказа'].astype('int64')

    df = df_orders_rec.merge(df_handover, on='Номер заказа', how='left')
    df = df[['Номер заказа', 'Постановка в очередь', 'Заказ приготовился', 'Ожидание на полке', 'Отправка на заказ',
             'Начало поездки', 'Разница поездки', 'Заказов', 'Фамилия']]
    df = df.loc[df['Заказов'] != 1]

    df['Начало поездки'] = pd.to_timedelta(df['Начало поездки'])
    df['Отправка на заказ'] = pd.to_timedelta(df['Отправка на заказ'])
    df['compare'] = (df['Начало поездки'] - df['Отправка на заказ'] < timedelta(seconds=20)) & (
            df['Отправка на заказ'] - df['Начало поездки'] < timedelta(seconds=20))
    df = df[df['compare'] == False]
    df = df[['Номер заказа', 'Постановка в очередь', 'Заказ приготовился', 'Ожидание на полке', 'Отправка на заказ',
             'Начало поездки', 'Разница поездки', 'Заказов', 'Фамилия']]
    df['Отправка на заказ'] = df['Отправка на заказ'].apply(lambda x: str(x).split(' ')[-1])
    df['Начало поездки'] = df['Начало поездки'].apply(lambda x: str(x).split(' ')[-1])
    df.to_excel(f'./export/flip_{name}.xlsx')


def start():
    get_order('Белгород-2', '2022-06-23')
    change('Белгород-2')


if __name__ == '__main__':
    start()
