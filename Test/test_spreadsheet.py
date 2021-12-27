import unittest

from . import test_cases as tc


class TestInteractor(unittest.TestCase):
    def setUp(self) -> None:
        from ..spreadsheet.Interactor import Interactor
        from ..spreadsheet.Entities import Entities
        from ..spreadsheet.Presenters import Presenters
        from ..Test import Catcher
        from ..spreadsheet.EntityGateways import Gateways
        from ..spreadsheet.ExternalAPI.XlsxWriter import save_as_excel

        entities = Entities()
        presenters = Presenters()
        gateways = Gateways(entities)
        catcher = Catcher()
        presenters.attach_to_response_model_receiver(catcher.add_response_model)
        gateways.attach_to_gateway_model_receiver(catcher.add_gateway_model)
        gateways.attach_to_spreadsheet(catcher.add_gateway_last_model)
        gateways.attach_to_spreadsheet(save_as_excel)
        interactor = Interactor(entities, presenters, gateways)

        self._interactor = interactor
        self._catcher = catcher

    def test_export_input_sheets_as_excel(self):
        interactor = self._interactor
        catcher = self._catcher
        t = tc.test_case_export_input_sheet_as_excel()

        for request_model, expected_gateway_model, expected_spreadsheet_model in zip(t[0], t[1], t[2]):
            workbook_name = request_model[0]
            account_ids = request_model[1]
            input_account_ids = request_model[2]
            input_texts = request_model[3]
            values_volume = request_model[4]
            values_price = request_model[5]

            interactor.set_input_accounts(input_account_ids)
            interactor.create_input_sheet(input_account_ids, input_texts)
            interactor.set_input_values(account_ids[0], values_volume)
            interactor.set_input_values(account_ids[1], values_price)
            interactor.export_spreadsheet(workbook_name)

            self.assertEqual(catcher.gateway_models[-1], expected_gateway_model)
            self.assertEqual(catcher.gateway_last_models[-1], expected_spreadsheet_model)

    def test_export_calculation_sheet_as_excel(self):
        interactor = self._interactor
        catcher = self._catcher

        expected_gateway_models, request_models = tc.test_case_export_calculation_sheet_as_excel()

        for request_model, expected_gateway_model in zip(request_models, expected_gateway_models):
            workbook_name = request_model[0]
            id_to_text = request_model[1]
            input_accounts = request_model[2]
            input_values = request_model[3]
            operator_ids = request_model[4]
            constant_ids = ()
            worksheets_data = request_model[5]
            rpes = request_model[6]

            interactor.set_number_of_periods(5)
            interactor.set_input_accounts(input_accounts)
            interactor.create_input_sheet(input_accounts, id_to_text)
            interactor.set_inputs_values(input_accounts, input_values)
            interactor.create_new_worksheets(id_to_text, worksheets_data)
            rpe_converted = interactor.convert_rpe_to_cell_address(rpes, constant_ids, id_to_text, operator_ids)
            interactor.set_formulas(rpe_converted)
            interactor.extract_inputs_and_create_domestic_inputs()
            interactor.connect_direct_links()
            interactor.export_spreadsheet(workbook_name)

            self.assertEqual(catcher.gateway_models[-1], expected_gateway_model)

    def test_external_api(self):
        from ..spreadsheet.API import Spreadsheet
        expected_gateway_models, request_models = tc.test_cases_external_api()

        for n, (request_model, expected_gateway_model) in enumerate(zip(request_models, expected_gateway_models)):
            spreadsheet = Spreadsheet()
            spreadsheet._interactor = self._interactor  # replace interactor with test_interactor with catcher attached.
            new_signature = {'workbook_name': request_model[0],
                             'shape_id_to_address': request_model[1],
                             'inputs': request_model[2],
                             'input_values': request_model[3],
                             'operators': request_model[4],
                             'constants': request_model[5],
                             'sheets_data': request_model[6],
                             'rpes': request_model[7],
                             'nop': request_model[8],
                             'direct_links_mutable': request_model[9], }
            spreadsheet.export(**new_signature)
            gateway_model_caught = self._catcher.gateway_models[-1]

            self.assertEqual(gateway_model_caught[0], expected_gateway_model[0])
            sheet_data_caught = gateway_model_caught[1]
            sheet_data_expected = expected_gateway_model[1]
            same_number_of_sheets = len(sheet_data_caught) == len(sheet_data_expected)
            self.assertEqual(same_number_of_sheets, True)
            if same_number_of_sheets:
                for sh_n, (sheet_data, sheet_data_expected) in enumerate(zip(sheet_data_caught, sheet_data_expected)):
                    self.assertEqual(sheet_data[0], sheet_data_expected[0])
                    self.assertEqual(sheet_data[1], sheet_data_expected[1])
                    self.assertEqual(sheet_data[2], sheet_data_expected[2])
                    self.assertEqual(sheet_data[3], sheet_data_expected[3])

                    format_data = sheet_data[4]
                    format_data_expected = sheet_data_expected[4]
                    self.assertEqual(len(format_data), len(format_data_expected))
                    for format_n, (each_format_data, expected) in enumerate(zip(format_data, format_data_expected)):
                        self.assertEqual(each_format_data, expected)
            self.setUp()

    def test_sensitivity_sheet(self):
        input_accounts = (0, 1, 2, 3, 4, 5)
        shape_id_to_text = {
            0: 'Input 0',
            1: 'Input 1',
            2: 'Input 2',
            3: 'Input 3',
            4: 'Input 4',
            5: 'Input 5',
        }
        selected_variables = (0, 1, 3, 4)
        shape_id_to_delta = {
            0: 10,
            1: 20,
            2: 15,
            3: 10,
            4: 10,
            5: 20,
        }

        from ..spreadsheet.Entities import SensitivitySheet
        from ..spreadsheet.ExternalAPI.XlsxWriter import save_as_excel
        from ..spreadsheet.EntityGateways.Spreadsheet.implementation import create_spreadsheet_model
        file_name = f'src/Test/Test Sensitivity Sheet.xlsx'
        sensitivity_sheet = SensitivitySheet('Tests', input_accounts, shape_id_to_text)
        sheet_data = sensitivity_sheet.get_worksheet_data({}, selected_variables, shape_id_to_delta)
        gateway_model = file_name, (sheet_data,)
        spreadsheet_model = create_spreadsheet_model(gateway_model)
        save_as_excel(spreadsheet_model)


if __name__ == '__main__':
    unittest.main()
