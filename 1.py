import xlwt


def main(data: dict):
    if not data:
        return None

    def toFixed(numObj, digits=2):
        return f"{numObj:.{digits}f}"

    wb = xlwt.Workbook()
    ws = wb.add_sheet('main')

    ws.write_merge(0, 0, 1, 3, "Общее по всем операторам")

    ws.write(1, 1, 'Время общее')
    ws.write(1, 2, 'Время активное')
    ws.write(1, 3, 'Эффективность')

    worker_count = 0

    for date in list(data.values()):
        ws.write(2 + worker_count, 0, date['name'])
        ws.write(2 + worker_count, 1, toFixed(date['common']))
        ws.write(2 + worker_count, 2, toFixed(date['active']))
        ws.write(2 + worker_count, 3, ('0' if date['common'] == 0 else (
            toFixed(date['active'] / date['common'] * 100))) + '%')
        worker_count += 1

    wb.save('example.xls')


if __name__ == '__main__':
    data = {"1": {"common": 4.920756944444448, "active": 4.0, "name": "\u0425\u043e\u043b\u043e\u043f \u0425\u043e\u043b\u043e\u043f\u043e\u0432"},
        "2": {"common": 3.5, "active": 1.2, "name": "\u0420\u0430\u0431\u043e\u0447\u0438\u0439 \u0420\u0430\u0431\u043e\u0442\u0430\u0435\u0442"}}
    main(data)
