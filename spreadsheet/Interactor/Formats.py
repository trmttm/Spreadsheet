from typing import Callable

from ..Entities import CellFormat
from ..Entities import Worksheet
from ..Entities import Worksheets


def create_input_format(create_new_format: Callable[[], CellFormat]) -> CellFormat:
    new_format = create_new_format()
    new_format.set_text_color('blue')
    return new_format


def create_domestic_input_format(create_new_format: Callable[[], CellFormat]) -> CellFormat:
    new_format = create_new_format()
    new_format.set_text_color('orange')
    return new_format


def create_heading_format(create_new_format: Callable[[], CellFormat]) -> CellFormat:
    new_format = create_new_format()
    new_format.highlight('4F81BD')
    new_format.set_text_color('white')
    return new_format


def create_subtotal_format(create_new_format: Callable[[], CellFormat]) -> CellFormat:
    new_format = create_new_format()
    new_format.set_top_border(color='4F81BD')
    new_format.set_bottom_border('double', '4F81BD')
    return new_format


def create_number_format(create_new_format: Callable[[], CellFormat], number_format: str) -> CellFormat:
    new_format = create_new_format()
    new_format.set_number_format(number_format)
    return new_format


def set_number_formats(number_format: str, target_accounts: tuple, worksheets: Worksheets):
    for account_id in target_accounts:
        worksheet = worksheets.identify_worksheet(account_id)
        if worksheet is not None:
            values_format = worksheet.get_values_format(account_id)
            values_format.update({'number_format': number_format})
            worksheet.set_values_format(account_id, values_format)


def format_domestic_inputs_(direct_links: tuple, format_: dict, input_accounts: tuple, input_sheet_name: str,
                            worksheets: Worksheets):
    account_id = tuple(to_ for (from_, to_, shift) in direct_links if from_ in input_accounts)
    for worksheet_name, worksheet in worksheets.items():
        if worksheet_name == input_sheet_name:
            continue
        for account in worksheet.account_ids:
            if account in account_id:
                worksheet.set_values_format(account, format_)


def synchronize_format(account_from, sheet_from: Worksheet, account_to, sheet_to: Worksheet):
    values_format_from = sheet_from.get_values_format(account_from)
    values_format = sheet_to.get_values_format(account_to)
    values_format.update({'number_format': values_format_from['number_format']})
    sheet_to.set_values_format(account_to, values_format)


def set_input_formats(format_input, input_sheet: Worksheet, sorted_input_ids):
    for account_id in sorted_input_ids:
        input_sheet.set_values_format(account_id, format_input)


def queue_input_to_clients_synchronize_commands(clients: tuple, format_commands_args_kwargs: list, input_id,
                                                worksheet: Worksheet, input_sheet: Worksheet):
    for client in clients:
        args = (input_id, input_sheet, client, worksheet)
        format_commands_args_kwargs.append((synchronize_format, args))


def queue_input_to_domestic_synchronize_commands(domestic_input_id, format_commands_args_kwargs: list, input_id,
                                                 input_sheet, worksheet: Worksheet):
    args = (input_id, input_sheet, domestic_input_id, worksheet)
    format_commands_args_kwargs.append((synchronize_format, args))


def set_format_to_cells(cell_format, target_accounts, worksheets):
    for account_id in target_accounts:
        worksheet = worksheets.identify_worksheet(account_id)
        if worksheet is not None:
            worksheet.set_format_to_all_fields(account_id, cell_format)


def queue_format_synchronizer_commands(format_commands_args_kwargs, queue):
    for format_command, *args_kwargs in format_commands_args_kwargs:
        queue(synchronize_format, *args_kwargs)
