import xlsxwriter


def write_dictionary_as_excel(file_name, data: dict):
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet('Dictionary')
    row, col = 0, 0
    for key, values in data.items():
        col = 0
        if type(key) == tuple:
            column_values = key + values
        else:
            column_values = (key,) + values
        for value in column_values:
            worksheet.write(row, col, value)
            col += 1
        row += 1
    workbook.close()
