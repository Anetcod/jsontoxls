import xlwt


def main(data):
    if not data:
        return None

    wb = xlwt.Workbook()
    ws = wb.add_sheet('main')
    car_count = 0
    car_values = list(data['machines'].values())
        
    for date in car_values[0]['dates']:
        ws.write(
            1, 1 + list(car_values[0]['dates'].keys()).index(date), date)

    for car in car_values:
        _ = 0 if car_count == 0 else 1
        car_dates = list(car['dates'].keys())

        ws.write(2 + (car_count * 7 + 2 * _), 0, car['name'])
        ws.write(3 + (car_count * 7 + 2 * _), 0, 'Ударов')
        ws.write(4 + (car_count * 7 + 2 * _), 0, 'Выключение')
        ws.write(5 + (car_count * 7 + 2 * _), 0, 'Начальный заряд АКБ')
        ws.write(6 + (car_count * 7 + 2 * _), 0, "Выключение")
        ws.write(7 + (car_count * 7 + 2 * _), 0, "Быстрый заряд")
        ws.write(8 + (car_count * 7 + 2 * _), 0, "Температура 1")
        ws.write(9 + (car_count * 7 + 2 * _), 0, "Температура 2")

        for date in (car['dates'].items()):
            ws.write(3 + (car_count * 7 + 2 * _), 1 +
                     car_dates.index(date[0]), date[1]['0'])
            ws.write(4 + (car_count * 7 + 2 * _), 1 +
                     car_dates.index(date[0]), date[1]['1'])
            ws.write(5 + (car_count * 7 + 2 * _), 1 +
                     car_dates.index(date[0]), date[1]['2'])
            ws.write(6 + (car_count * 7 + 2 * _), 1 +
                     car_dates.index(date[0]), date[1]['3'])
            ws.write(7 + (car_count * 7 + 2 * _), 1 +
                     car_dates.index(date[0]), date[1]['4'])
            ws.write(8 + (car_count * 7 + 2 * _), 1 +
                     car_dates.index(date[0]), date[1]['5'])
            ws.write(9 + (car_count * 7 + 2 * _), 1 +
                     car_dates.index(date[0]), date[1]['6'])

        car_count += 1

    wb.save('example.xls')

    return True


if __name__ == '__main__':
    data = {"machines": {"1": {
        "dates": {"24.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "25.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "26.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "27.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "28.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "29.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "30.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}}, "name": "\u0420\u0438\u0447\u0442\u0440\u0430\u043a 305"},
        "2": {"dates": {"24.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "25.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "26.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "27.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "28.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "29.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}, "30.06.2019": {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0}}, "name": "\u0420\u0438\u0447\u0442\u0440\u0430\u043a 306"}}}
    main(data)
