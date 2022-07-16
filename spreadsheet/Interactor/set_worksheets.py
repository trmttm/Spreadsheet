from spreadsheet.Entities import Entities


def execute(workheets: Entities.worksheets, number_of_periods: int) -> dict:
    columns_widths = {}
    narrow = 2
    wide = 30
    uom_width = 10
    value_width = 15

    for sheet_name in workheets.worksheet_names:
        columns_widths[sheet_name] = {}
        worksheet = workheets.get_worksheet(sheet_name)
        number_of_vertical_columns = sum(tuple(1 if 'vertical' in clm_name else 0 for clm_name in worksheet.fields))

        up_to_name = worksheet.initial_column + 1
        next_to_name = up_to_name + 1
        uom_col = next_to_name + 1
        next_to_uom = uom_col + 1
        start_of_vertical_accounts = next_to_uom + 1
        start_of_values = next_to_uom + number_of_vertical_columns + 1

        for c in range(up_to_name + 1):
            columns_widths[sheet_name][c] = narrow
        columns_widths[sheet_name][next_to_name] = wide
        columns_widths[sheet_name][uom_col] = uom_width
        columns_widths[sheet_name][next_to_uom] = narrow

        for c in range(start_of_vertical_accounts, start_of_vertical_accounts + number_of_vertical_columns + 1):
            columns_widths[sheet_name][c] = narrow
        for c in range(start_of_values, start_of_values + number_of_periods + 1):
            columns_widths[sheet_name][c] = value_width

    return columns_widths
