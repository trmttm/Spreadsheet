from ..Interactor import DirectLinks
from ..Interactor import util


def set_user_defined_function_(account, account_id, arguments_ids, function_name, input_sheet_name, nop,
                               worksheets):
    worksheet_account = worksheets.identify_worksheet(account_id)
    argument_addresses = []
    for argument_id in arguments_ids:
        args = argument_id, account_id, input_sheet_name, worksheets
        if DirectLinks.is_direct_link_from_domestic_input(*args):
            domestic_input_id = util.get_domestic_input_id(worksheet_account.name, argument_id)
            argument_id = domestic_input_id
        worksheet_argument = worksheets.identify_worksheet(argument_id)

        with_sheet_name = worksheet_account != worksheet_argument
        if with_sheet_name:
            addresses = worksheet_argument.get_value_addresses_with_sheet_name_without_lock(argument_id, nop)
        else:
            addresses = worksheet_argument.get_value_addresses_without_sheet_name_without_lock(argument_id, nop)
        argument_addresses.append(addresses)
    user_defined_functions = []
    for addresses in zip(*argument_addresses):
        user_defined_functions.append(f'={function_name}{addresses}'.replace("'", ''))
    account.set_values(tuple(user_defined_functions))
