from typing import Type

from .Account import Account
from .Accounts import Accounts
from .AdditionalSheetDatas import AdditionalSheetDatas
from .CellFormat import CellFormat
from .Charts import Charts
from .SpecialSheets import SensitivitySheet
from .SpecialSheets import VBAcommands
from .TemporaryData import TemporaryData
from .Worksheet import Worksheet
from .Worksheets import Worksheets


class Entities:
    def __init__(self):
        self._accounts = Accounts()
        self._cell_format = CellFormat
        self._worksheets = Worksheets()
        self._additional_sheet_datas = AdditionalSheetDatas()
        self._temporary_data = TemporaryData()
        self._cls_sensitivity_sheet = SensitivitySheet
        self._cls_vba_command_sheet = VBAcommands
        self._charts = Charts()

    @property
    def accounts(self) -> Accounts:
        return self._accounts

    @property
    def cls_cell_format(self) -> Type[CellFormat]:
        return self._cell_format

    @property
    def worksheets(self) -> Worksheets:
        return self._worksheets

    @property
    def additional_sheet_datas(self) -> AdditionalSheetDatas:
        return self._additional_sheet_datas

    @property
    def cls_sensitivity_sheet(self) -> Type[SensitivitySheet]:
        return self._cls_sensitivity_sheet

    @property
    def cls_vba_command_sheet(self) -> Type[VBAcommands]:
        return self._cls_vba_command_sheet

    @property
    def temporary_data(self) -> TemporaryData:
        return self._temporary_data

    @property
    def charts(self) -> Charts:
        return self._charts
