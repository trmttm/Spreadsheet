class AdditionalSheetDatas:
    def __init__(self):
        self._data = ()

    def add(self, sheet_data: tuple):
        self._data += (sheet_data,)

    def clear(self):
        self._data = ()

    @property
    def data(self) -> tuple:
        return self._data
