from interface_spreadsheet import SpreadsheetABC

from .Entities import Entities
from .EntityGateways import Gateways
from .ExternalAPI.XlsxWriter import save_as_excel
from .ExternalAPI.XlsxWriter import write_dictionary_as_excel
from .Interactor import Interactor
from .Interactor import VAc
from .Presenters import Presenters


class Spreadsheet(SpreadsheetABC):
    def __init__(self):
        entities = Entities()
        presenters = Presenters()
        gateways = Gateways(entities)
        gateways.attach_to_spreadsheet(save_as_excel)
        self._interactor = Interactor(entities, presenters, gateways)

    def export(self, **kwargs) -> tuple:
        rpes, sub_total_accounts = alter_kwargs_to_set_up_vertical_accounts(kwargs)

        # Unpack kwargs
        input_accounts = get_kwargs('inputs', kwargs)
        input_values = get_kwargs('input_values', kwargs)
        insert_sheet_name = get_kwargs('insert_sheet_name', kwargs)
        workbook_name = get_kwargs('workbook_name', kwargs)
        shape_id_to_text = get_kwargs('shape_id_to_address', kwargs)
        operator_ids = get_kwargs('operators', kwargs)
        constant_ids = get_kwargs('constants', kwargs)
        worksheets_data = get_kwargs('sheets_data', kwargs)
        direct_links = get_kwargs('direct_links_mutable', kwargs)
        user_defined_function = get_kwargs('user_defined_function', kwargs)
        format_data = get_kwargs('format_data', kwargs)
        number_format_data = get_kwargs('number_format_data', kwargs)
        number_of_periods = get_kwargs('nop', kwargs)
        vertical_accounts = get_kwargs('vertical_acs', kwargs) or {}
        heading_accounts = get_formatted_accounts(format_data, 'heading')
        whole_number_accounts = get_formatted_accounts(number_format_data, 'whole number')
        one_digit_accounts = get_formatted_accounts(number_format_data, '1-digit')
        two_digit_accounts = get_formatted_accounts(number_format_data, '2-digit')
        percent_accounts = get_formatted_accounts(number_format_data, '%')

        # A few decision making
        no_indent = heading_accounts + sub_total_accounts + tuple(vertical_accounts.keys())

        # Control Interactor
        """
        In principle, 
            Any operation that includes changing rows or columns of accounts must be done BEFORE:
                1) converting RPE 
                
            Formatting:
                1) should be at the end,
                2) all accounts at the same time,
                3) by careful order between accounts.
                ...because of overwriting nature of updating account formats.
        """
        interactor = self._interactor
        interactor.set_number_of_periods(number_of_periods)
        interactor.set_direct_links(direct_links)
        interactor.set_vertical_accounts(vertical_accounts)
        interactor.handle_inputs(input_accounts, input_values, insert_sheet_name, shape_id_to_text, worksheets_data)
        interactor.create_new_worksheets(shape_id_to_text, worksheets_data, no_indent)
        interactor.insert_vertical_account_columns(vertical_accounts)

        # ===========ROW / COLUMN fixes here. Change them before here!=================================================
        add_sensitivity_sheet = get_kwargs('add_sensitivity_sheet', kwargs)
        if add_sensitivity_sheet:
            target_accounts = get_kwargs('target_accounts', kwargs)
            variables = get_kwargs('selected_variables', kwargs)
            shape_id_to_delta = get_kwargs('shape_id_to_delta', kwargs)
            command_file_name = get_kwargs('command_file_name', kwargs)

            rpes = interactor.modify_input_sheet_for_prior_to_creating_sensitivity_sheet(rpes)
            sensitivity_sheet = interactor.create_sensitivity_sheet(target_accounts, variables, shape_id_to_delta)
            interactor.set_formula_to_input_sensitivities(sensitivity_sheet)
            interactor.add_tornado_chart(sensitivity_sheet)
            interactor.add_spider_chart(sensitivity_sheet)
            vba_file = kwargs.get('vba_file', None)
            interactor.delegate_commands_to_vba(sensitivity_sheet, workbook_name, command_file_name, vba_file)

        args_to_rpe_conversion = rpes, constant_ids, shape_id_to_text, operator_ids, vertical_accounts
        interactor.set_formulas_from_rpe(*args_to_rpe_conversion)
        interactor.extract_inputs_and_create_domestic_inputs()
        interactor.link_vertical_accounts(vertical_accounts)
        interactor.add_direct_links_from_total_to_vertical_account(vertical_accounts)
        interactor.remove_redundant_vertical_calculation_formulas(vertical_accounts)
        interactor.connect_direct_links()

        if user_defined_function is not None:
            function_name = user_defined_function['name']
            account_id = user_defined_function['account_ids']
            arguments_ids = user_defined_function['arguments']
            interactor.set_user_defined_function(function_name, account_id, arguments_ids)

        # ===========Formatting must be after modifying DirectLinks / Domestic Inputs!=================================
        interactor.format_subtotal(sub_total_accounts)
        interactor.format_subtotal(tuple(vertical_accounts.keys()))
        interactor.format_whole_number(whole_number_accounts)
        interactor.format_one_digit(one_digit_accounts)
        interactor.format_two_digit(two_digit_accounts)
        interactor.format_percent(percent_accounts)
        interactor.format_domestic_inputs()
        interactor.format_heading(heading_accounts)
        interactor.execute_delayed_format_commands()

        feedback = interactor.export_spreadsheet(workbook_name)
        return feedback

    @staticmethod
    def save_dictionary_as_spreadsheet(file_name, dictionary: dict):
        write_dictionary_as_excel(file_name, dictionary)


def get_kwargs(key: str, kwargs: dict):
    return kwargs[key] if key in kwargs else None


def get_formatted_accounts(format_data: dict, key: str) -> tuple:
    accounts = ()
    if format_data is not None:
        if key in format_data:
            accounts = format_data[key]
    return accounts


def alter_kwargs_to_set_up_vertical_accounts(kwargs) -> tuple:
    shape_id_to_text = get_kwargs('shape_id_to_address', kwargs)
    number_of_periods = get_kwargs('nop', kwargs)
    format_data = get_kwargs('format_data', kwargs)
    worksheets_data = get_kwargs('sheets_data', kwargs)
    vertical_accounts = get_kwargs('vertical_acs', kwargs) or {}
    sub_total_accounts = get_formatted_accounts(format_data, 'total')
    rpes = get_kwargs('rpes', kwargs)

    args = number_of_periods, rpes, shape_id_to_text, vertical_accounts, sub_total_accounts, worksheets_data
    rpes, sub_total_accounts = VAc.create_vertical_account_rpe_and_sub_totals(*args)
    return rpes, sub_total_accounts
