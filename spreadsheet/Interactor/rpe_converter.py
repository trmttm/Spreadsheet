from typing import Callable

from ..Entities import Worksheet
from ..Entities import Worksheets


def set_element_as_is(element, rpe_converted_list: list, number_of_periods: int):
    elements = tuple(element for _ in range(number_of_periods))
    rpe_converted_list.append(elements)


def set_operator(operator, rpe_converted_list: list):
    rpe_converted_list.append(operator)


def get_field_from_element(element, vertical_dependencies: tuple, field_name_factory: Callable):
    n = vertical_dependencies.index(element)
    field = field_name_factory(n)
    return field


def set_vertical_calculation_account(owner_id, field, nop: int, owner_sheet, rpe_converted_list: list):
    addresses = owner_sheet.get_addresses_without_sheet_name_col_locked(owner_id, field, nop)
    first_address = tuple(addresses[0] for _ in addresses)
    rpe_converted_list.append(first_address)


def set_account(element, element_sheet: Worksheet, nop: int, owner_id, owner_sheet: Worksheet, rpe_converted_list: list,
                v_calc_ac_ids: tuple):
    with_sheet_name = owner_sheet.name != element_sheet.name
    if with_sheet_name:
        if owner_id in v_calc_ac_ids:
            addresses = element_sheet.get_value_addresses_with_sheet_name_row_locked(element, nop)
        else:
            addresses = element_sheet.get_value_addresses_with_sheet_name_without_lock(element, nop)
    else:
        if owner_id in v_calc_ac_ids:
            addresses = element_sheet.get_value_addresses_without_sheet_name_row_locked(element, nop)
        else:
            addresses = element_sheet.get_value_addresses_without_sheet_name_without_lock(element, nop)
    rpe_converted_list.append(addresses)


def get_element_sheet(element, input_sheet_name: str, owner_sheet: Worksheet, worksheets: Worksheets) -> Worksheet:
    priorities = (owner_sheet.name, input_sheet_name,)
    element_sheet = worksheets.identify_worksheet(element, priorities)
    return element_sheet
