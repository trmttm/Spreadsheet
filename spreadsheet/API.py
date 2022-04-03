from interface_spreadsheet import SpreadsheetABC

from .Entities import Entities
from .EntityGateways import Gateways
from .ExternalAPI.XlsxWriter import save_as_excel
from .ExternalAPI.XlsxWriter import write_dictionary_as_excel
from .Interactor import Interactor
from .Interactor import VAc
from .Interactor import breakdown_account
from .Presenters import Presenters


class Spreadsheet(SpreadsheetABC):
    _field_uom = 'UOM'

    def __init__(self):
        entities = Entities()
        presenters = Presenters()
        gateways = Gateways(entities)
        gateways.attach_to_spreadsheet(save_as_excel)
        self._interactor = Interactor(entities, presenters, gateways)

    def export(self, **kwargs) -> tuple:
        rpes, sub_total_accounts = alter_kwargs_to_set_up_vertical_accounts(kwargs)
        rpes = breakdown_account.handle_breakdown_accounts(rpes, kwargs)
        # Unpack kwargs
        input_accounts = kwargs.get('inputs', None)
        input_values = kwargs.get('input_values', None)
        insert_sheet_name = kwargs.get('insert_sheet_name', None)
        workbook_name = kwargs.get('workbook_name', None)
        shape_id_to_text = kwargs.get('shape_id_to_text', None)
        operator_ids = kwargs.get('operators', None)
        constant_ids = kwargs.get('constants', None)
        worksheets_data = kwargs.get('sheets_data', None)
        modules_data = kwargs.get('modules_data', worksheets_data)
        direct_links = kwargs.get('direct_links', None)
        user_defined_function = kwargs.get('user_defined_function', None)
        format_data = kwargs.get('format_data', None)
        number_format_data = kwargs.get('number_format_data', None)
        number_of_periods = kwargs.get('nop', None)
        vertical_accounts = kwargs.get('vertical_acs', {})
        shape_id_to_uom = kwargs.get('shape_id_to_uom', {})
        breakdown_accounts = kwargs.get('breakdown_accounts', ())
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
        interactor.handle_inputs(input_accounts, input_values, insert_sheet_name, shape_id_to_text, modules_data)
        interactor.create_new_worksheets(shape_id_to_text, worksheets_data, no_indent)
        interactor.insert_vertical_account_columns(vertical_accounts)
        interactor.insert_uom(self._field_uom)

        # ===========ROW / COLUMN fixes here. Change them before here!=================================================
        add_sensitivity_sheet = kwargs.get('add_sensitivity_sheet', None)
        if add_sensitivity_sheet:
            target_accounts = kwargs.get('target_accounts', None)
            variables = kwargs.get('selected_variables', None)
            shape_id_to_delta = kwargs.get('shape_id_to_delta', None)
            command_file_name = kwargs.get('command_file_name', None)

            f = interactor.modify_input_sheet_for_prior_to_creating_sensitivity_sheet
            rpes = f(rpes, self._field_uom, shape_id_to_uom)
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
        interactor.set_uom_to_accounts(shape_id_to_uom, self._field_uom)

        if user_defined_function is not None:
            function_name = user_defined_function['name']
            account_id = user_defined_function['account_ids']
            arguments_ids = user_defined_function['arguments']
            interactor.set_user_defined_function(function_name, account_id, arguments_ids)

        # ===========Formatting must be after modifying DirectLinks / Domestic Inputs!=================================
        interactor.format_subtotal(sub_total_accounts)
        interactor.format_subtotal(tuple(vertical_accounts.keys()))
        interactor.format_breakdown(breakdown_accounts)
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


def get_formatted_accounts(format_data: dict, key: str) -> tuple:
    accounts = ()
    if format_data is not None:
        if key in format_data:
            accounts = format_data[key]
    return accounts


def alter_kwargs_to_set_up_vertical_accounts(kwargs) -> tuple:
    shape_id_to_text = kwargs.get('shape_id_to_text', None)
    number_of_periods = kwargs.get('nop', None)
    format_data = kwargs.get('format_data', None)
    worksheets_data = kwargs.get('sheets_data', None)
    vertical_accounts = kwargs.get('vertical_acs', None) or {}
    sub_total_accounts = get_formatted_accounts(format_data, 'total')
    rpes = kwargs.get('rpes', None)

    args = number_of_periods, rpes, shape_id_to_text, vertical_accounts, sub_total_accounts, worksheets_data
    rpes, sub_total_accounts = VAc.create_vertical_account_rpe_and_sub_totals(*args)
    return rpes, sub_total_accounts
