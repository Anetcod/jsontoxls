import xlwt


def main(data: dict):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('main')

    to_mark, r_mark, ar_mark, average_mark = 0, 0, 0, 0
    to_ammount, r_ammount, ar_ammount, amount_full = 0, 0, 0, 0
    car_count = 0

    for car in list(data.items()):
        name = car[0]
        car = car[1]
        car_to_mark = car["ТО"]['rating']
        car_r_mark = car["Ремонт"]['rating']
        car_ar_mark = car["Аварийный ремонт"]['rating']
        car_to_ammount = car["ТО"]['price']
        car_r_ammount = car["Ремонт"]['price']
        car_ar_ammount = car["Аварийный ремонт"]['price']
        to_mark += car_to_mark
        r_mark += car_r_mark
        ar_mark += car_ar_mark
        to_ammount += car_to_ammount
        r_ammount += car_r_ammount
        ar_ammount += car_ar_ammount
        car_average_mark = (car_to_mark + car_r_mark + car_ar_mark) / 3
        car_ammount_full = car_to_ammount + car_r_ammount + car_ar_ammount
        average_mark += car_average_mark
        amount_full += car_ammount_full

        ws.write(3 + car_count, 0, name)
        ws.write(3 + car_count, 1, car_to_mark)
        ws.write(3 + car_count, 2, car_to_ammount)
        ws.write(3 + car_count, 3, car_r_mark)
        ws.write(3 + car_count, 4, car_r_ammount)
        ws.write(3 + car_count, 5, car_ar_mark)
        ws.write(3 + car_count, 6, car_ar_ammount)
        ws.write(3 + car_count, 8, car_ammount_full)
        ws.write(3 + car_count, 7, car_average_mark)
        
        car_count += 1


    ws.write_merge(0, 1, 0, 0, "Техника")
    ws.write_merge(0, 1, 7, 7, "Средняя оценка")
    ws.write_merge(0, 1, 8, 8, "Общая сумма")

    ws.write(2, 0, 'Сводка')
    ws.write(2, 1, to_mark / car_count)
    ws.write(2, 2, to_ammount)
    ws.write(2, 3, r_mark / car_count)
    ws.write(2, 4, r_ammount)
    ws.write(2, 5, ar_mark / car_count)
    ws.write(2, 6, ar_ammount)
    ws.write(2, 8, amount_full)
    ws.write(2, 7, average_mark / car_count)

    ws.write_merge(0, 0, 1, 2, 'ТО')

    ws.write(1, 1, 'Оценка')
    ws.write(1, 2, 'Сумма')

    ws.write_merge(0, 0, 3, 4, 'Ремонт')
    ws.write(1, 3, 'Оценка')
    ws.write(1, 4, 'Сумма')

    ws.write_merge(0, 0, 5, 6, 'Аварийный ремонт')
    ws.write(1, 5, 'Оценка')
    ws.write(1, 6, 'Сумма')

    wb.save('example.xls')

    return True

if __name__ == '__main__':
    null = 0
    data = {"\u0420\u0438\u0447\u0442\u0440\u0430\u043a": {"\u0422\u041e": {"price": 3500, "rating": 3}, "\u0420\u0435\u043c\u043e\u043d\u0442": {"price": 400, "rating": 2}, "\u0410\u0432\u0430\u0440\u0438\u0439\u043d\u044b\u0439 \u0440\u0435\u043c\u043e\u043d\u0442": {"price": 600, "rating": 4}},
            "\u041f\u043e\u0433\u0440\u0443\u0437\u0447\u0438\u043a": {"\u0422\u041e": {"price": 1700, "rating": 5}, "\u0420\u0435\u043c\u043e\u043d\u0442": {"price": 1400, "rating": 3}, "\u0410\u0432\u0430\u0440\u0438\u0439\u043d\u044b\u0439 \u0440\u0435\u043c\u043e\u043d\u0442": {"price": 800, "rating": 2}}}
    main(data)
