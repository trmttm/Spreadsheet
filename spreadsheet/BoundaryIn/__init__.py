from abc import ABC
from abc import abstractmethod


class BoundaryInABC(ABC):
    @abstractmethod
    def set_input_accounts(self, input_accounts: tuple):
        pass

    @abstractmethod
    def insert_worksheet_name_to_input_sheet(self, worksheet_names: tuple):
        pass

    @abstractmethod
    def set_period_account_if_any(self):
        pass

    @abstractmethod
    def set_inputs_values(self, input_account_ids: tuple, group_of_values: tuple):
        pass

    @abstractmethod
    def create_input_sheet(self, input_account_ids: tuple, id_to_text: dict):
        pass

    @abstractmethod
    def export_spreadsheet(self, file_name: str, **options) -> tuple:
        pass

    @abstractmethod
    def set_input_values(self, input_account_id, values: tuple):
        pass

    @abstractmethod
    def connect_direct_links(self):
        pass
