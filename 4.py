import xlwt


def main(data: dict):
    if not data:
        return None

    wb = xlwt.Workbook()
    ws = wb.add_sheet('main')

    dates = list(list(data.values())[0]['data'].keys())
    car_count = 0

    false = xlwt.easyxf('pattern: pattern solid, fore_colour orange;'
                        'font: colour white, bold True;')
    true = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue;'
                       'font: colour white, bold True;')

    ws.write(0, 0, 'Борт')

    for date in dates:
        ws.write(
            0, 1 + dates.index(date), date)

    for car in list(data.values()):
        ws.write(1 + car_count, 0, car['name'])
        column = 1

        for _data in list(car['data'].values()):
            if _data == 0:
                ws.write(1 + car_count, column, 'п', false)

            else:
                ws.write(1 + car_count, column, 'р', true)

            column += 1

        car_count += 1

    wb.save('example.xls')

    return True

if __name__ == '__main__':
    data = {"1": {"data": {"26.08.2019": 0, "27.08.2019": 0, "28.08.2019": 0, "29.08.2019": 0, "30.08.2019": 0, "31.08.2019": 0, "01.09.2019": 0}, "name": "\u041f\u043e\u0433\u0440\u0443\u0437\u0447\u0438\u043a 305"}, "2": {
        "data": {"26.08.2019": 0, "27.08.2019": 1, "28.08.2019": 0, "29.08.2019": 0, "30.08.2019": 0, "31.08.2019": 0, "01.09.2019": 0}, "name": "\u0420\u0438\u0447\u0442\u0440\u0430\u043a 306"}}
    main(data)
