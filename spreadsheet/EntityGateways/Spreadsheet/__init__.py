from . import implementation as impl
from ...Entities import Accounts
from ...Entities import Worksheets
from ...EntityGateways.GatewayABC import GatewayABC


class SpreadsheetExporter(GatewayABC):
    @staticmethod
    def create_gateway_model_for_spreadsheet(file_name: str, accounts: Accounts, worksheets: Worksheets,
                                             additional_sheet_datas: tuple):
        spreadsheet_data = impl.create_spreadsheet_data(accounts, worksheets)
        gateway_model = file_name, spreadsheet_data + additional_sheet_datas
        return gateway_model

    def create_model(self, gateway_model, **kwargs):
        spreadsheet_model = impl.create_spreadsheet_model(gateway_model)
        return spreadsheet_model
