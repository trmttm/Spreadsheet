from ...Entities import Accounts
from ...Entities import Worksheets


def create_spreadsheet_data(accounts: Accounts, worksheets: Worksheets) -> tuple:
    spreadsheet_data_list = []
    for worksheet in worksheets.values():
        name, account_ids, rows, columns, fields, field_to_format = worksheet.get_worksheet_data
        formats = accounts.get_formats(account_ids, field_to_format, fields)
        group_of_cell_values = accounts.get_group_of_cell_values(account_ids, fields)

        worksheet_data = name, group_of_cell_values, rows, columns, formats
        spreadsheet_data_list.append(worksheet_data)
    spreadsheet_datas = tuple(spreadsheet_data_list)
    return spreadsheet_datas


def create_spreadsheet_model(gateway_model: tuple) -> tuple:
    instructions_list = []
    file_name, spreadsheet_datas = gateway_model
    for worksheet_data in spreadsheet_datas:
        worksheet_name, group_of_values, rows, cols, group_of_formats = worksheet_data

        for values, row, col_initial, formats in zip(group_of_values, rows, cols, group_of_formats):
            for n, (cell_value, cell_format) in enumerate(zip(values, formats)):
                row_data = (cell_value, worksheet_name, (row, col_initial + n), cell_format)
                instructions_list.append(row_data)
    instructions = tuple(instructions_list)
    spreadsheet_model = file_name, instructions
    return spreadsheet_model
