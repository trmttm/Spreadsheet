from . import util
from ..Entities import Accounts
from ..Entities import Worksheets


def get_how_many_vertical_reference_columns_are_needed(vertical_accounts: dict) -> int:
    max_vertical_accounts = 0
    for vertical_account, vertical_dependencies in vertical_accounts.items():
        max_vertical_accounts = max(max_vertical_accounts, len(vertical_dependencies))
    return max_vertical_accounts


def get_vertical_calculation_account_id(original_account_id: str, number: int) -> str:
    return f'vertical_{original_account_id}_{number}'


def get_original_vertical_account(vertical_calculation_account_id: str) -> int:
    return int(vertical_calculation_account_id.split('_')[1].split('_')[0])


def get_vertical_total_account_id(original_account_id: str) -> str:
    return f'vertical_{original_account_id}_total'


def create_vertical_field_name(n: int) -> str:
    new_field_name = f'vertical_{n}'
    return new_field_name


def get_vertical_calculation_account_ids(nop: int, vertical_accounts: dict) -> tuple:
    v_calc_ac_ids = tuple()
    for vertical_account_id in vertical_accounts.keys():
        v_calc_ac_ids += tuple(get_vertical_calculation_account_id(vertical_account_id, i) for i in range(nop))
    return v_calc_ac_ids


def get_vertical_dependencies(owner_id, v_calc_ac_ids: tuple, vertical_accounts: dict):
    if owner_id in v_calc_ac_ids:
        vertical_dependencies: tuple = vertical_accounts[get_original_vertical_account(owner_id)]
    else:
        vertical_dependencies = ()
    return vertical_dependencies


def insert_new_vertical_columns(max_vertical_accounts: int, position, worksheets: Worksheets):
    for sheet_name, worksheet in worksheets.items():
        for n in range(max_vertical_accounts):
            # Add field to all of worksheets
            new_field_name = create_vertical_field_name(n)
            worksheet.add_field(new_field_name, position)


def add_vertical_fields_to_accounts_and_worksheets(accounts: Accounts, max_vertical_accounts: int):
    for account_id in accounts.all_account_ids:
        account = accounts.get_account(account_id)
        for n in range(max_vertical_accounts):
            new_field_name = create_vertical_field_name(n)
            account.add_new_attribute(new_field_name, '')


def link_vertical_accounts(accounts: Accounts, input_ids: tuple, nop: int, vertical_accounts: dict,
                           worksheets: Worksheets):
    for original_ac_id, vertical_dependencies in vertical_accounts.items():
        vertical_ac_ids = tuple(get_vertical_calculation_account_id(original_ac_id, n) for n in range(nop))
        worksheet = w = worksheets.identify_worksheet(vertical_ac_ids[0])
        for i, vertical_ac_id in enumerate(vertical_ac_ids):
            vertical_account = accounts.get_account(vertical_ac_id)
            for n, v_dependency_id in enumerate(vertical_dependencies):
                vertical_field = create_vertical_field_name(n)
                if v_dependency_id in input_ids:
                    v_dependency_id = util.get_domestic_input_id(worksheet.name, v_dependency_id)
                addresses = w.get_value_addresses_without_sheet_name_row_locked_with_eq_sign(v_dependency_id, nop)
                vertical_account.add_new_attribute(vertical_field, addresses[i])


def remove_redundant_vertical_calculation_formulas(accounts: Accounts, nop: int, vertical_accounts: dict):
    # Purpose is to prevent circular reference
    for vertical_account in vertical_accounts:
        for period in range(nop):
            account_id = get_vertical_calculation_account_id(vertical_account, period)
            account = accounts.get_account(account_id)
            existing_formula = account.values
            new_formula = tuple(v if period <= p else '' for (p, v) in enumerate(existing_formula))
            account.set_values(new_formula)


def add_vertical_calculation_account(j, new_sheet_contents: list, shape_id_to_text: dict, vertical_ac_id):
    vertical_calc_ac_id = get_vertical_calculation_account_id(vertical_ac_id, j)
    new_sheet_contents.append(vertical_calc_ac_id)
    shape_id_to_text[vertical_calc_ac_id] = f'{shape_id_to_text[vertical_ac_id]}_{j}'


def add_vertical_calculator_formula(rpes_dict: dict, rpes_mutable: list, vertical_ac_id, j):
    vertical_calc_ac_id = get_vertical_calculation_account_id(vertical_ac_id, j)
    rpe_formula = rpes_dict[vertical_ac_id]
    rpes_mutable.append((vertical_calc_ac_id, rpe_formula), )
    rpes_dict[vertical_calc_ac_id] = rpe_formula


def increment_total_formula(j, rpe_total_list: list, vertical_ac_id):
    vertical_calc_ac_id = get_vertical_calculation_account_id(vertical_ac_id, j)
    rpe_total_list.append(vertical_calc_ac_id)
    if j > 0:
        rpe_total_list.append('+')


def add_total_account(new_sheet_contents: list, shape_id_to_text: dict, vertical_ac_id):
    vertical_ac_total_id = get_vertical_total_account_id(vertical_ac_id)
    new_sheet_contents.append(vertical_ac_total_id)
    shape_id_to_text[vertical_ac_total_id] = f'Total {shape_id_to_text[vertical_ac_id]}'


def add_total_formula(rpe_total_list, rpes_dict, rpes_mutable, vertical_ac_id, sub_total_accounts_mutable):
    vertical_ac_total_id = get_vertical_total_account_id(vertical_ac_id)
    rpe_total_formula = tuple(rpe_total_list)
    rpes_mutable.append((vertical_ac_total_id, rpe_total_formula), )
    rpes_dict[vertical_ac_total_id] = rpe_total_formula
    sub_total_accounts_mutable.append(vertical_ac_total_id, )


def add_direct_links_from_total_to_vertical_account_(vertical_accounts: dict) -> tuple:
    new_direct_links = []
    for vertical_account in vertical_accounts:
        total_id = get_vertical_total_account_id(vertical_account)
        new_direct_link = (total_id, vertical_account, 0)
        new_direct_links.append(new_direct_link)
    return tuple(new_direct_links)


def insert_vertical_account_columns_(accounts: Accounts, vertical_accounts: dict, worksheets: Worksheets):
    max_vertical_accounts = get_how_many_vertical_reference_columns_are_needed(vertical_accounts)
    insert_new_vertical_columns(max_vertical_accounts, 3, worksheets)
    add_vertical_fields_to_accounts_and_worksheets(accounts, max_vertical_accounts)


def create_vertical_account_rpe_and_sub_totals(number_of_periods: int, rpes: tuple, shape_id_to_text: dict,
                                               vertical_accounts: dict, sub_total_accounts: tuple, sheets_data: dict):
    rpes_mutable = list(rpes)
    sub_total_acs_mutable = list(sub_total_accounts)
    rpes_dict = util.create_rpe_dictionary(rpes_mutable)
    for each_sheet_name, each_sheet_contents in sheets_data.items():
        new_sheet_contents = []
        for sheet_content in each_sheet_contents:
            new_sheet_contents.append(sheet_content)
            if sheet_content in vertical_accounts:
                vertical_ac_id = sheet_content
                rpe_total = []
                for j in range(number_of_periods):
                    add_vertical_calculation_account(j, new_sheet_contents, shape_id_to_text, vertical_ac_id)
                    add_vertical_calculator_formula(rpes_dict, rpes_mutable, vertical_ac_id, j)
                    increment_total_formula(j, rpe_total, vertical_ac_id)
                add_total_account(new_sheet_contents, shape_id_to_text, vertical_ac_id)
                add_total_formula(rpe_total, rpes_dict, rpes_mutable, vertical_ac_id, sub_total_acs_mutable)
        sheets_data[each_sheet_name] = tuple(new_sheet_contents)
    return tuple(rpes_mutable), tuple(sub_total_acs_mutable)
