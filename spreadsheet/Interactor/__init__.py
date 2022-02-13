from . import DirectLinks
from . import Formats
from . import Formulas
from . import Inputs
from . import RPE
from . import Sensitivity
from . import UDF
from . import UOM
from . import VAc
from . import implementation as impl
from . import rpe_converter as rpec
from . import util
from ..BoundaryIn import BoundaryInABC
from ..Entities import Entities
from ..Entities import SensitivitySheet
from ..Entities import VBAcommands
from ..EntityGatewaysABS import GatewaysABC
from ..Presenters import PresentersABC


class Interactor(BoundaryInABC):
    input_sheet_name = 'Inputs'
    number_of_periods = 10
    default_number_format = "#,##0.0_);[Red]▲#,##0.0"
    whole_number_format = "#,##0;[Red]▲#,##0"
    one_digit_number_format = "#,##0.0_);[Red]▲#,##0.0"
    two_digit_number_format = "#,##0.00_);[Red]▲#,##0.00"
    percent_number_format = "#,##0 %;[Red]▲#,##0 %"
    _sh_sens = 'Sensitivity'

    def __init__(self, entities: Entities, presenters: PresentersABC, gate_way: GatewaysABC):
        self._accounts = entities.accounts
        self._worksheets = entities.worksheets
        self._additional_sheet_datas = entities.additional_sheet_datas
        self._cls_sensitivity_sheet = entities.cls_sensitivity_sheet
        self._cls_vba_command_sheet = entities.cls_vba_command_sheet
        self._temporary_data = entities.temporary_data
        self._cls_cell_format = entities.cls_cell_format
        self._presenters = presenters
        self._gate_way = gate_way
        self._charts = entities.charts

        self._define_formats()

    def _create_new_format(self):
        new_format = self._cls_cell_format()
        new_format.set_number_format(self.default_number_format)
        return new_format

    def _define_formats(self):
        create_new_format = self._create_new_format
        self._format_default = create_new_format()
        self._format_input = Formats.create_input_format(create_new_format)
        self._format_domestic_input = Formats.create_domestic_input_format(create_new_format)
        self._format_subtotal = Formats.create_subtotal_format(create_new_format)
        self._format_heading = Formats.create_heading_format(create_new_format)
        self._format_whole_number = Formats.create_number_format(create_new_format, self.whole_number_format)
        self._format_one_digits = Formats.create_number_format(create_new_format, self.one_digit_number_format)
        self._format_two_digits = Formats.create_number_format(create_new_format, self.two_digit_number_format)
        self._format_percent = Formats.create_number_format(create_new_format, self.percent_number_format)

    def handle_inputs(self, input_accounts: tuple, input_values: tuple, insert_sheet_name: bool, shape_id_to_text: dict,
                      worksheets_data: dict):
        """
        if you pass worksheets_data, then inputs will be segregated by worksheets.
        if you pass modules_data, then inputs will be segregated by modules.
        """
        self.set_input_accounts(input_accounts)
        inputs_segregated = Inputs.get_input_accounts_segregated_by_worksheets(input_accounts, worksheets_data)
        self.create_input_sheet(inputs_segregated, shape_id_to_text)
        if insert_sheet_name:
            worksheets_that_have_input_accounts = []
            for sheet_name, accounts in worksheets_data.items():
                a = set(accounts).intersection(set(input_accounts))
                if a != set():
                    worksheets_that_have_input_accounts.append(sheet_name)
            self.insert_worksheet_name_to_input_sheet(tuple(worksheets_that_have_input_accounts))
        self.set_inputs_values(input_accounts, input_values)
        self.set_period_account_if_any()

    @property
    def input_worksheet(self):
        return self._worksheets.get_worksheet(self.input_sheet_name)

    def set_number_of_periods(self, number_of_periods: int):
        self.number_of_periods = number_of_periods

    def set_inputs_values(self, input_account_ids: tuple, group_of_values: tuple):
        for account_id, values in zip(input_account_ids, group_of_values):
            self.set_input_values(account_id, values)

    def set_input_values(self, input_account_id, values: tuple):
        self._accounts[input_account_id].set_values(values)

    def create_input_sheet(self, sorted_input_accounts_with_none: tuple, id_to_text: dict):
        arguments = Inputs.create_input_sheet_creation_args(id_to_text, sorted_input_accounts_with_none)
        input_worksheet = self._worksheets.create_new_worksheet(self.input_sheet_name)
        args = *arguments, self._accounts, input_worksheet
        impl.add_accounts_to_worksheet(*args)
        Formats.set_input_formats(self._format_input, input_worksheet, arguments[1])

    def insert_worksheet_name_to_input_sheet(self, worksheet_names: tuple):
        args = self._accounts, self.input_worksheet, self.number_of_periods, worksheet_names
        new_accounts = Inputs.insert_worksheet_name_to_input_sheet(*args)
        for new_account_id in new_accounts:
            self._temporary_data.add_account_ids_inserted_in_input_sheet(new_account_id)

    def extract_inputs_and_create_domestic_inputs(self):
        inputs_accounts = ia = self._temporary_data.input_accounts
        icd = Inputs.create_input_to_clients_dictionary(self.direct_links, inputs_accounts)

        for ws in self._worksheets.all_worksheets:
            if ws.name == self.input_sheet_name:
                continue

            new_links, old_links, fcak = [], [], []
            Inputs.extract_inputs_(self._accounts, fcak, icd, self.input_worksheet, ia, new_links, old_links, ws)
            DirectLinks.update_direct_links(new_links, old_links, self.add_direct_link, self.remove_direct_link)
            Formats.queue_format_synchronizer_commands(fcak, self.add_delayed_format_command)

    def create_new_worksheets(self, id_to_text: dict, worksheets_data: dict, no_indent_accounts: tuple = ()):
        args1 = self.input_accounts, self.sheet_names_were_inserted_in_input_sheet, no_indent_accounts
        no_indents = Inputs.indent_input_accounts_if_necessary(*args1)
        args2 = self._accounts, no_indents, self._format_default, id_to_text, self._worksheets, worksheets_data
        impl.create_new_worksheets(*args2)

    def set_formulas_from_rpe(self, rpes: tuple, constants: tuple, id_to_text: dict, operators: tuple,
                              vertical_acs: dict = None):
        rpe_converted = self.convert_rpe_to_cell_address(rpes, constants, id_to_text, operators, vertical_acs)
        self.set_formulas(rpe_converted)

    def convert_rpe_to_cell_address(self, rpes: tuple, constants: tuple, id_to_text: dict, operators: tuple,
                                    vertical_acs: dict = None) -> tuple:
        nop, input_sheet_name, worksheets = self.number_of_periods, self.input_sheet_name, self._worksheets
        args = self._accounts, constants, id_to_text, input_sheet_name, nop, operators, rpes, vertical_acs, worksheets
        return RPE.convert_rpe_to_cell_address_(*args)

    def set_formulas(self, rpes_converted: tuple):
        Formulas.set_formulas(self.number_of_periods, rpes_converted)

    # Vertical account methods
    def insert_vertical_account_columns(self, vertical_accounts: dict):
        VAc.insert_vertical_account_columns_(self._accounts, vertical_accounts, self._worksheets)

    def link_vertical_accounts(self, vertical_accounts: dict):
        args = self._accounts, self.input_accounts, self.number_of_periods, vertical_accounts, self._worksheets
        VAc.link_vertical_accounts(*args)

    def add_direct_links_from_total_to_vertical_account(self, vertical_accounts):
        new_direct_links = VAc.add_direct_links_from_total_to_vertical_account_(vertical_accounts)
        for new_direct_link in new_direct_links:
            self.add_direct_link(new_direct_link)

    def remove_redundant_vertical_calculation_formulas(self, vertical_accounts: dict):
        VAc.remove_redundant_vertical_calculation_formulas(self._accounts, self.number_of_periods, vertical_accounts)

    def connect_direct_links(self):
        args = self._accounts, self.direct_links, self.input_sheet_name, self.number_of_periods, self._worksheets
        DirectLinks.connect_direct_links(*args)

    def set_user_defined_function(self, name: str, account_id, arguments: tuple):
        account = self._accounts.get_account(account_id)
        args = account, account_id, arguments, name, self.input_sheet_name, self.number_of_periods, self._worksheets
        UDF.set_user_defined_function_(*args)

    # UOM
    def insert_uom(self, field_name: str):
        worksheets = self._worksheets
        position = 3
        for sheet_name, worksheet in worksheets.items():
            worksheet.add_field(field_name, position)

    def set_uom_to_accounts(self, shape_id_to_uom: dict, field_name: str):
        UOM.set_uom_to_accounts(self._accounts, self.direct_links, field_name, shape_id_to_uom)

    # Formatting
    def format_subtotal(self, subtotal_account_ids: tuple):
        self._set_format_to_cells(self._format_subtotal, subtotal_account_ids)

    def format_heading(self, heading_account_ids: tuple):
        self._set_format_to_cells(self._format_heading, heading_account_ids)

    def _set_format_to_cells(self, cell_format, target_accounts: tuple):
        Formats.set_format_to_cells(cell_format, target_accounts, self._worksheets)

    def format_whole_number(self, whole_number_accounts: tuple):
        self._set_number_format(self.whole_number_format, whole_number_accounts)

    def format_one_digit(self, one_digit_accounts: tuple):
        self._set_number_format(self.one_digit_number_format, one_digit_accounts)

    def format_two_digit(self, two_digit_accounts: tuple):
        self._set_number_format(self.two_digit_number_format, two_digit_accounts)

    def format_percent(self, percent_accounts: tuple):
        self._set_number_format(self.percent_number_format, percent_accounts)

    def format_domestic_inputs(self):
        args = self.direct_links, self._format_domestic_input, self.input_accounts, self.input_sheet_name, self._worksheets
        Formats.format_domestic_inputs_(*args)

    def _set_number_format(self, number_format, target_accounts):
        Formats.set_number_formats(number_format, target_accounts, self._worksheets)

    # Sensitivity Sheet
    def modify_input_sheet_for_prior_to_creating_sensitivity_sheet(self, rpes, field_uom: str, uom_dict: dict) -> tuple:
        mutable_rpes = list(rpes)

        input_accounts = self.input_accounts
        input_sheet = self._worksheets.get_worksheet(self.input_sheet_name)
        # Insert 2 rows
        for input_account in input_accounts:
            input_sheet.insert_row_above(input_account, 2)

        for input_id in self.input_accounts:
            row = input_sheet.account_id_to_row[input_id]
            input_account = self._accounts.get_account(input_id)
            # 1) Add new Input Row
            new_input_id = f'new_input_{input_id}'
            impl.add_account_to_worksheet(row - 2, new_input_id, input_account.name, self._accounts, input_sheet)
            new_input_account = self._accounts.get_account(new_input_id)
            # 2) Add Sensitivity Row
            sensitivity_account_id = Sensitivity.get_input_sensitivity_account_id(input_id)
            impl.add_account_to_worksheet(row - 1, sensitivity_account_id, self._sh_sens, self._accounts, input_sheet)
            sensitivity_account = self._accounts.get_account(sensitivity_account_id)
            # 3) Set Formula
            mutable_rpes.append((input_id, [new_input_id, sensitivity_account_id, '*']))
            # 4) Set default Input Values
            new_input_account.set_values(input_account.values)
            # 5) Link Sensitivity formula (delayed command necessary?)
            sensitivity_account.set_values(tuple(1 for _ in range(self.number_of_periods)))
            # 6) Formatting
            input_sheet.set_values_font_color(input_id, 'black')
            self.format_percent((sensitivity_account_id,))
            self._add_delayed_command_number_format_synchronize(input_id, input_sheet, new_input_id, input_sheet)
            # 7) Group Rows
            input_sheet.group_row(input_id)
            input_sheet.group_row(sensitivity_account_id)
            # 8) Indent if necessary
            if self.sheet_names_were_inserted_in_input_sheet:
                new_input_account.indent()
                sensitivity_account.indent()
            # 9) Add UOM to [new_input]s
            uom = uom_dict.get(input_id, '')
            new_input_account.add_new_attribute(field_uom, uom)

        return tuple(mutable_rpes)

    def create_sensitivity_sheet(self, target_accounts: tuple, selected_variables: tuple,
                                 shape_id_to_delta: dict) -> SensitivitySheet:
        input_accounts = self.input_accounts

        shape_id_to_address = self._create_shape_id_to_address_name(input_accounts + target_accounts)
        formats = self._format_input, self._format_domestic_input, self._format_whole_number, self._format_two_digits, self._format_percent
        sensitivity_sheet = self._cls_sensitivity_sheet(self._sh_sens, input_accounts, shape_id_to_address, *formats)

        args = target_accounts, self.number_of_periods, self._worksheets
        target_to_address = Sensitivity.create_target_account_text_to_address(*args)
        sheet_data = sensitivity_sheet.get_worksheet_data(target_to_address, selected_variables, shape_id_to_delta)

        self._additional_sheet_datas.add(sheet_data)
        return sensitivity_sheet

    def delegate_commands_to_vba(self, sensitivity_sheet: SensitivitySheet, workbook_name: str, command_file_name: str,
                                 vba_file=None):
        sheet_vba_command = self.create_vba_command_sheet(workbook_name, command_file_name)
        new_format = self._create_new_format()
        new_format.set_text_color('white')
        sheet_vba_command.set_format(new_format)
        sheet_vba_command.add_command('add_data_table', *sensitivity_sheet.data_table_tornado_arguments)
        sheet_vba_command.add_command('add_data_table', *sensitivity_sheet.data_table_spider_arguments)
        sheet_vba_command.add_command('select_sheet', self.input_sheet_name)
        self.save_vba_command_sheet(sheet_vba_command.path, sheet_vba_command, vba_file)

    def create_vba_command_sheet(self, file_path, my_book_name: str):
        sheet = self._cls_vba_command_sheet(file_path, my_book_name)
        return sheet

    def save_vba_command_sheet(self, file_name: str, sheet: VBAcommands, vba_file=None):
        sheet_data = sheet.get_worksheet_data()

        kwargs = {'vba_file': vba_file} if vba_file is not None else {}
        self.export_separate_workbook(file_name, (sheet_data,), **kwargs)

    def set_formula_to_input_sensitivities(self, sheet: SensitivitySheet):
        for input_account in self.input_accounts:
            sensitivity_id = Sensitivity.get_input_sensitivity_account_id(input_account)
            sensitivity_account = self._accounts.get_account(sensitivity_id)
            address = sheet.get_address_sensitivity_applied(input_account)
            sensitivity_account.set_values(tuple(address for _ in range(self.number_of_periods)))

    def _create_shape_id_to_address_name(self, account_ids: tuple) -> dict:
        shifts = tuple(1 if self._accounts.get_account(account).is_indented else 0 for account in account_ids)
        return impl.create_shape_id_to_address_name(account_ids, self._worksheets, shifts)

    # Graph
    def add_tornado_chart(self, sheet: SensitivitySheet):
        graph_inputs = sheet.graph_inputs_tornado
        chart_title = graph_inputs[0]
        border_color = 'black'
        show_label = False
        position = graph_inputs[1]
        address_values_list = graph_inputs[2]
        names = graph_inputs[3]
        category = graph_inputs[4]
        fill_colors = ('red', 'blue')

        charts = self._charts
        chart_id = charts.add_new_stacked_bar_chart()
        charts.set_chart_sheet_name(chart_id, 'Tornado')
        charts.set_worksheet_name(chart_id, self._sh_sens)
        charts.set_chart_title(chart_id, chart_title)
        charts.set_x_axis_label(chart_id, 'Low vs Hi')
        charts.set_y_axis_label(chart_id, 'Variables')
        charts.set_chart_position(chart_id, position)

        for n, values in enumerate(address_values_list):
            charts.add_data_series(chart_id, values, names[n], category, fill_colors[n], border_color, show_label)

    def add_spider_chart(self, sheet: SensitivitySheet):
        graph_inputs = sheet.graph_inputs_spider
        chart_title = graph_inputs[0]
        border_color = 'black'
        show_label = False
        position = graph_inputs[1]
        address_values_list = graph_inputs[2]
        names = graph_inputs[3]
        category = graph_inputs[4]
        fill_colors = ('red', 'blue')

        charts = self._charts
        chart_id = charts.add_new_line_chart()
        charts.set_chart_sheet_name(chart_id, 'Spider')
        charts.set_worksheet_name(chart_id, self._sh_sens)
        charts.set_chart_title(chart_id, chart_title)
        charts.set_x_axis_label(chart_id, 'Cases')
        charts.set_y_axis_label(chart_id, 'Variables')
        charts.set_chart_position(chart_id, position)

        for n, values in enumerate(address_values_list):
            charts.add_data_series(chart_id, values, names[n], category, None, None, show_label)

    # Temporary Data
    def set_input_accounts(self, input_accounts: tuple):
        self._temporary_data.set_input_accounts(input_accounts)

    @property
    def input_accounts(self) -> tuple:
        return self._temporary_data.input_accounts

    def set_period_account_if_any(self):
        for input_account in self.input_accounts:
            account = self._accounts.get_account(input_account)
            if account.name == 'Period':
                self._temporary_data.set_period_account(input_account)
                break

    def set_direct_links(self, direct_links: tuple):
        self._temporary_data.set_direct_links(direct_links)

    def add_direct_link(self, direct_link: tuple):
        self._temporary_data.add_direct_link(direct_link)

    def remove_direct_link(self, direct_link: tuple):
        self._temporary_data.remove_direct_link(direct_link)

    def set_vertical_accounts(self, vertical_accounts: dict):
        self._temporary_data.set_vertical_accounts(vertical_accounts)

    @property
    def direct_links(self) -> tuple:
        return self._temporary_data.direct_links

    @property
    def sheet_names_were_inserted_in_input_sheet(self) -> bool:
        return self._temporary_data.account_ids_of_worksheet_names_inserted_in_input_sheet != ()

    def add_delayed_format_command(self, method, args: tuple = (), kwargs: dict = None):
        self._temporary_data.add_delayed_format_command(method, args, kwargs)

    def _add_delayed_command_number_format_synchronize(self, account_from, sheet_from, account_to, sheet_to):
        args = (account_from, sheet_from, account_to, sheet_to)
        method = Formats.synchronize_format
        kwargs = {}
        self.add_delayed_format_command(method, args, kwargs)
        sheet_from.set_format(account_to, sheet_from.field_values, self._format_input)

    def execute_delayed_format_commands(self):
        self._temporary_data.execute_delayed_format_commands()

    # Export
    def export_spreadsheet(self, file_name: str, **options) -> tuple:
        self._handle_temporary_data()

        w = self._worksheets
        sheet_names = tuple(name for name in w.worksheet_names if w.get_worksheet(name).groupings_exist)
        dictionaries = tuple(w.get_worksheet(name).row_to_levels for name in sheet_names)
        sheet_to_rows_to_levels_dictionary = dict(zip(sheet_names, dictionaries))
        options.update({
            'additional_sheets': self._additional_sheet_datas.data,
            'levels_dict': sheet_to_rows_to_levels_dictionary,
            'charts': self._charts.chart_model,
        })
        feed_back = self._gate_way.export_spreadsheet(file_name, **options)

        return feed_back

    def export_separate_workbook(self, file_name: str, spreadsheet_data, **kwargs) -> tuple:
        gateway_model = file_name, spreadsheet_data
        feed_back = self._gate_way.save_spreadsheet_file(gateway_model, **kwargs)
        return feed_back

    def _handle_temporary_data(self):
        account_ids_sheet_name = self._temporary_data.account_ids_of_worksheet_names_inserted_in_input_sheet
        self.format_heading(account_ids_sheet_name)
        self.format_whole_number(account_ids_sheet_name)
        self.link_period_to_input_worksheet_names(account_ids_sheet_name)
        self._temporary_data.clear()  # remove potential side effects

    def link_period_to_input_worksheet_names(self, ac_ids):
        args = self._temporary_data.period_account, ac_ids, self._accounts, self.input_worksheet, self.number_of_periods
        Inputs.link_account_ids_sheet_name_to_period(*args)
