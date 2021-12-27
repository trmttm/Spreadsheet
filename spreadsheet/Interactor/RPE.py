import rpe_to_normal

from . import VAc
from . import rpe_converter as rpec
from ..Entities import Accounts
from ..Entities import Worksheets


def convert_rpe_to_cell_address_(accounts: Accounts, constant_ids: tuple, id_to_text: dict, input_sheet_name: str,
                                 nop: int, operator_ids: tuple, rpes, vertical_accounts: tuple, worksheets: Worksheets):
    vertical_accounts = {} if vertical_accounts is None else vertical_accounts
    vcac = VAc.get_vertical_calculation_account_ids(nop, vertical_accounts)
    rpes_converted_list = []
    for owner_id, rpes_formula in rpes:
        if owner_id in vertical_accounts:
            # No values of original vertical accounts.
            continue

        rpe_list = []
        vertical_dependencies = VAc.get_vertical_dependencies(owner_id, vcac, vertical_accounts)
        owner_sheet = worksheets.identify_worksheet(owner_id)

        for element in rpes_formula:
            if element in operator_ids:
                rpec.set_operator(rpe_to_normal.operator[id_to_text[element]], rpe_list)
            elif (owner_id in vcac) and (element in vertical_dependencies):
                field = rpec.get_field_from_element(element, vertical_dependencies, VAc.create_vertical_field_name)
                rpec.set_vertical_calculation_account(owner_id, field, nop, owner_sheet, rpe_list)
            elif element in accounts.all_account_ids:
                element_sheet = rpec.get_element_sheet(element, input_sheet_name, owner_sheet, worksheets)
                rpec.set_account(element, element_sheet, nop, owner_id, owner_sheet, rpe_list, vcac)
            elif element in constant_ids:
                rpec.set_element_as_is(id_to_text[element].replace(',', ''), rpe_list, nop)
            else:
                rpec.set_element_as_is(element, rpe_list, nop)
        rpe_converted = tuple(rpe_list)
        rpes_converted_list.append((accounts[owner_id], rpe_converted))
    rpes_converted = tuple(rpes_converted_list)
    return rpes_converted
