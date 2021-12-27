from abc import ABC
from abc import abstractmethod


class GatewaysABC(ABC):
    @abstractmethod
    def attach_to_gateway_model_receiver(self, observer):
        pass

    @abstractmethod
    def attach_to_spreadsheet(self, observer):
        pass

    @abstractmethod
    def export_spreadsheet(self, file_name: str, **options) -> tuple:
        pass

    @abstractmethod
    def save_spreadsheet_file(self, gateway_model, **kwargs):
        pass
