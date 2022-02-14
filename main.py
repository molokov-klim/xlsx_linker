import xlsxwriter
# from xlsxwriter.utility import xl_rowcol_to_cell
import openpyxl

Alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
            'V', 'W', 'X', 'Y', 'Z']
req_list = {}


def read_write_xlsx():
    # открытие книги
    book = "2.xlsx"
    wb = openpyxl.reader.excel.load_workbook(filename=book, data_only=True)

    # назначение страницы для чтения (2 лист)
    wb.active = 1
    sheet = wb.active

    # Ищем столбик 'ID Req' на втором листе, потом записываем в дикт значения столбика (ключ-ячейка, значение-содержимое ячейки)
    for column in range(0, 99):
        # print(column)
        # print("sheet[Alphabet[column] + str('1')]: ", sheet[Alphabet[column] + str('1')])
        if sheet[Alphabet[column] + str('1')].value == 'Req ID':
            for row in range(1, 999):
                if sheet[Alphabet[column] + str(row)].value == None:
                    break
                # print("sheet[Alphabet[column] + str(row)].value: ", sheet[Alphabet[column] + str(row)].value)
                key = Alphabet[column] + str(row)
                # print("key: ", key)
                value = sheet[Alphabet[column] + str(row)].value
                # print("value: ", value)
                req_list[str(key)] = str(value)
                # print("req_list: ", req_list)
        if sheet[Alphabet[column] + str('1')].value == None:
            break

    # назначение страницы для чтения (1 лист)
    wb.active = 0
    sheet = wb.active

    # Ищем столбик 'ID Req' на первом листе, потом записываем из дикта значения в столбик и добавляем ссылку на источник
    for column in range(0, 99):
        if sheet[Alphabet[column] + str('1')].value == 'Req ID':
            for row in range(2, len(req_list)+1):

                key = Alphabet[column] + str(row)
                value = sheet[Alphabet[column] + str(row)].value


                for key_dict in req_list:
                    print("key_dict", key_dict)
                    print("req_list[key_dict]: ", req_list[key_dict])
                    print("key: ", key)
                    print("value: ", value)
                    if req_list[key_dict] == value:
                        link = str(book)+"#Requirements!"+key_dict
                        sheet[key].value = value
                        sheet[key].hyperlink = str(link)
                        sheet[key].style = "Hyperlink"
                        print("link: ", link)
                        print("sheet[Alphabet[column] + str(row)]: ", sheet[Alphabet[column] + str(row)])
                        print("sheet[Alphabet[column] + str(row)].value: ", sheet[Alphabet[column] + str(row)].value)

                if sheet[Alphabet[column] + str(row)].value == None:
                    break

        if sheet[Alphabet[column] + str('1')].value == None:
            break

    print("req_list: ", req_list)

    wb.save(book)
    wb.close()



if __name__ == '__main__':
    read_write_xlsx()


