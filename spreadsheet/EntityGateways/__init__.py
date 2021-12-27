from .Spreadsheet import SpreadsheetExporter
from ..Entities import Entities
from ..EntityGatewaysABS import GatewaysABC


class Gateways(GatewaysABC):
    def __init__(self, entities: Entities):
        self._entities = entities
        self._observers = []

        self._spreadsheet = SpreadsheetExporter()

    def _notify(self, gateway_model):
        for observer in self._observers:
            observer(gateway_model)

    def attach_to_gateway_model_receiver(self, observer):
        self._observers.append(observer)

    def attach_to_spreadsheet(self, observer):
        self._spreadsheet.attach(observer)

    def export_spreadsheet(self, file_name: str, **options) -> tuple:
        additional_sheets = options['additional_sheets'] if 'additional_sheets' in options else ()
        args = file_name, self._entities.accounts, self._entities.worksheets, additional_sheets
        gateway_model = self._spreadsheet.create_gateway_model_for_spreadsheet(*args)

        self._notify(gateway_model)
        return self.save_spreadsheet_file(gateway_model, **options)

    def save_spreadsheet_file(self, gateway_model, **kwargs):
        feedback = self._spreadsheet.export(gateway_model, **kwargs)
        return feedback
