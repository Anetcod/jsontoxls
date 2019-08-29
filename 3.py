import xlwt


def main(data: dict):
    def toFixed(numObj, digits=2):
        return f"{numObj:.{digits}f}"

    wb = xlwt.Workbook()
    ws = wb.add_sheet('main')

    ws.write_merge(0, 0, 1, 3, "Общее по всем машинам")

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

    return True

if __name__ == '__main__':
    data = {"1": {"common": 0, "active": 0, "name": "\u041f\u043e\u0433\u0440\u0443\u0437\u0447\u0438\u043a 305"},
            "2": {"common": 1.2170694444444454, "active": 0.0, "name": "\u0420\u0438\u0447\u0442\u0440\u0430\u043a 306"}}
    main(data)
