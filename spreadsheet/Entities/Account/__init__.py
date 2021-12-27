from . import implementation as impl


class Account:
    def __init__(self, account_id, text: str):
        self._account_id = account_id
        self._name = text
        self._values = ()
        self._indented = False

    @property
    def name(self) -> str:
        return self._name

    @property
    def account_id(self):
        return self._account_id

    @property
    def values(self) -> tuple:
        return self._values

    @property
    def blank(self) -> str:
        return ''

    def get_cell_values(self, fields: dict) -> tuple:
        values = impl.get_cell_values(self, fields)
        if self.is_indented:
            values = ('',) + (values[0],) + values[2:]
        return values

    def set_values(self, values: tuple):
        self._values = values

    def add_new_attribute(self, attribute: str, value):
        # UOM, Comments, as users may want to add.
        setattr(self, attribute, value)

    def __repr__(self):
        return self._name

    def indent(self):
        self._indented = True

    @property
    def is_indented(self) -> bool:
        return self._indented
