from . import DirectLinks
from . import Formats
from . import implementation as impl
from . import util
from ..Entities import Account
from ..Entities import Accounts
from ..Entities import Worksheet


def insert_worksheet_name_to_input_sheet(accounts: Accounts, input_sheet: Worksheet, nop: int, sheet_names) -> tuple:
    n_th_worksheet = 0
    new_account_ids_list = []
    for row in range(max(input_sheet.rows)):
        if row not in input_sheet.rows:
            new_account_id = f'input_title_{row - input_sheet.initial_row}'
            args = accounts, input_sheet, n_th_worksheet, new_account_id, nop, row, sheet_names
            _insert_sheet_name(*args)
            new_account_ids_list.append(new_account_id)
            n_th_worksheet += 1
    return tuple(new_account_ids_list)


def _insert_sheet_name(accounts: Accounts, input_worksheet: Worksheet, n_th_worksheet, new_account_id,
                       number_of_periods: int, row: int, sheet_names):
    worksheet_name = _identify_worksheet_name(n_th_worksheet, sheet_names)
    new_account = _create_account_that_show_worksheet_name(accounts, new_account_id, worksheet_name)
    _set_empty_values_for_the_account_showing_worksheet_name(new_account, number_of_periods)
    input_worksheet.add_account(row - input_worksheet.initial_row, new_account_id)
    input_worksheet.set_default_formats(new_account_id)


def _identify_worksheet_name(n_th_worksheet: int, worksheet_names) -> str:
    return worksheet_names[n_th_worksheet]


def _create_account_that_show_worksheet_name(accounts: Accounts, new_account_id, worksheet_name: str) -> Account:
    accounts.create_new_account(new_account_id, worksheet_name)
    new_account = accounts.get_account(new_account_id)
    return new_account


def _set_empty_values_for_the_account_showing_worksheet_name(new_account: Account, number_of_periods: int):
    new_account.set_values(tuple('' for _ in range(number_of_periods)))


def link_account_ids_sheet_name_to_period(period_id, account_ids: tuple, accounts: Accounts, input_sheet: Worksheet,
                                          nop: int):
    if period_id is not None:
        to_period = input_sheet.get_value_addresses_without_sheet_name_row_locked_with_eq_sign(period_id, nop)
        for account_id in account_ids:
            account = accounts.get_account(account_id)
            account.set_values(to_period)


def create_input_sheet_creation_args(id_to_text: dict, sorted_input_accounts_with_none: tuple) -> tuple:
    sorted_input_accounts = tuple(i for i in sorted_input_accounts_with_none if i is not None)
    sorted_input_texts = tuple(id_to_text[input_id] for input_id in sorted_input_accounts)
    input_rows = ()
    row = 0
    for input_account in sorted_input_accounts_with_none:
        if input_account is None:
            pass
        else:
            input_rows += (row,)
        row += 1
    return input_rows, sorted_input_accounts, sorted_input_texts


def _create_domestic_inputs(accounts: Accounts, input_ac_id, input_account, row: int, worksheet: Worksheet):
    domestic_input_id = util.get_domestic_input_id(worksheet.name, input_ac_id)
    impl.add_account_to_worksheet(row, domestic_input_id, input_account.name, accounts, worksheet)
    domestic_input_account = accounts[domestic_input_id]
    domestic_input_account.indent()
    return domestic_input_id


def extract_inputs_(accounts: Accounts, format_commands_args_kwargs: list, input_to_clients_dictionary: dict,
                    input_sheet: Worksheet, inputs: tuple, new_links: list, old_links: list, worksheet: Worksheet):
    fcak = format_commands_args_kwargs
    icd = input_to_clients_dictionary
    for input_id in worksheet.account_ids:
        if input_id in inputs:
            input_ac, r, clients = accounts.get_account(input_id), worksheet.get_row(input_id), icd[input_id]
            domestic_input_id = dii = _create_domestic_inputs(accounts, input_id, input_ac, r, worksheet)
            DirectLinks.connect_domestic_input(clients, domestic_input_id, input_id, new_links, old_links)
            Formats.queue_input_to_clients_synchronize_commands(clients, fcak, input_id, worksheet, input_sheet)
            Formats.queue_input_to_domestic_synchronize_commands(dii, fcak, input_id, input_sheet, worksheet)


def create_input_to_clients_dictionary(direct_links: tuple, inputs_accounts: tuple):
    input_to_clients_dictionary = icd = dict(zip(inputs_accounts, tuple(() for _ in inputs_accounts)))
    for from_, to_, shift in direct_links:
        if from_ in inputs_accounts:
            input_id = from_
            input_to_clients_dictionary[input_id] += (to_,)
    return icd


def indent_input_accounts_if_necessary(inputs: tuple, needs_indentation: bool, no_indent_accounts: tuple) -> tuple:
    if not needs_indentation:
        no_indent_accounts += inputs
    return no_indent_accounts


def get_input_accounts_segregated_by_worksheets(input_accounts: tuple, worksheet_data: dict) -> tuple:
    input_accounts_segregated_by_worksheets = ()
    for sheet_name, contents in worksheet_data.items():
        input_accounts_in_the_worksheet = tuple(c for c in contents if c in input_accounts)
        input_accounts_segregated_by_worksheets += input_accounts_in_the_worksheet
        if len(input_accounts_in_the_worksheet) > 0:
            input_accounts_segregated_by_worksheets += (None,)
    return input_accounts_segregated_by_worksheets
