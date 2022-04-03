import copy

from . import implementation as impl


def wrapper_equal_sign(addresses: tuple) -> tuple:
    return tuple(f'={a}' for a in addresses)


class Worksheet:
    field_name = 'name'
    field_values = 'values'
    field_blank = 'blank'

    def __init__(self, name: str, initial_row=1, initial_column=1):
        self._name = name
        self._row_to_account_id = {}
        self._initial_row = initial_row
        self._initial_column = initial_column
        self._fields: tuple = (self.field_name, self.field_blank, self.field_blank, self.field_values)

        self._field_to_format = {}
        self._account_to_levels = {}
        self._filed_to_width = {}

    @property
    def name(self) -> str:
        return self._name

    @property
    def account_ids(self) -> tuple:
        return tuple(self._row_to_account_id[key] for key in sorted(self._row_to_account_id.keys()))

    @property
    def get_worksheet_data(self) -> tuple:
        return self._name, self.account_ids, self.rows, self._columns, self._fields, self._field_to_format

    @property
    def account_id_to_row(self) -> dict:
        return dict(zip(self._row_to_account_id.values(), self._row_to_account_id.keys()))

    @property
    def rows(self) -> tuple:
        rows = tuple(self._initial_row + r for r in self._row_to_account_id.keys())
        rows_sorted = sorted(rows)
        return tuple(rows_sorted)

    @property
    def initial_row(self) -> int:
        return self._initial_row

    @property
    def _columns(self) -> tuple:
        return tuple(self._initial_column for _ in self._row_to_account_id)

    def has_account(self, account_id) -> bool:
        return account_id in self.account_ids

    def get_row(self, account_id) -> int:
        if account_id in self.account_id_to_row:
            return self.account_id_to_row[account_id]

    def get_account(self, row):
        if row in self._row_to_account_id:
            return self._row_to_account_id[row]

    def set_initial_row(self, initial_row: int):
        self._initial_row = initial_row

    def set_initial_column(self, initial_column: int):
        self._initial_column = initial_column

    def add_account(self, row: int, account_id):
        self._row_to_account_id[row] = account_id

    def set_default_formats(self, account_id):
        self._field_to_format[account_id] = dict(zip(self._fields, tuple(dict() for _ in self._fields)))

    def set_format_to_all_fields_with_negative_index(self, account_id, format_dict: dict, negative_index: tuple = ()):
        for n, field in enumerate(self._fields):
            if n not in negative_index:
                self.set_format(account_id, field, format_dict)

    def set_format_to_all_fields(self, account_id, format_dict: dict):
        for field in self._fields:
            self.set_format(account_id, field, format_dict)

    def set_values_format(self, account_id, format_dict: dict):
        self.set_format(account_id, self.field_values, format_dict)

    def set_format(self, account_id, field: str, format_dict: dict):
        if account_id in self._field_to_format:
            if field in self._field_to_format[account_id]:
                self._field_to_format[account_id][field].update(format_dict)
            else:
                self._field_to_format[account_id][field] = format_dict
        else:
            self._field_to_format[account_id] = {field: format_dict}

    def get_values_format(self, account_id) -> dict:
        if account_id in self._field_to_format:
            return copy.deepcopy(self._field_to_format[account_id][self.field_values])
        else:
            return {}

    def set_values_font_color(self, account_id, color: str):
        updated_format = self.get_values_format(account_id)
        updated_format.update({'color': color})
        self.set_values_format(account_id, updated_format)

    def add_field(self, name: str, position: int):
        new_fields_list = []
        for n, field_name in enumerate(self._fields):
            if n == position:
                new_fields_list.append(name)
            new_fields_list.append(field_name)
        self._fields = tuple(new_fields_list)

        for account_id, format_dict in self._field_to_format.items():
            self._field_to_format[account_id][name] = {}

    def clear_data(self):
        self._row_to_account_id = {}

    def insert_row_above(self, account_id, insert: int = 1):
        current_account_row = self.get_row(account_id)
        old_data = copy.deepcopy(self._row_to_account_id)
        self.clear_data()

        for row, ac_id in old_data.items():
            if row < current_account_row:
                self.add_account(row, ac_id)
            else:
                self.add_account(row + insert, ac_id)

    def group_row(self, account_id, level=1):
        self._account_to_levels[account_id] = level

    @property
    def groupings_exist(self):
        return self._account_to_levels != {}

    @property
    def row_to_levels(self) -> dict:
        r = self._initial_row
        rows = tuple(r + self.get_row(account_id) for account_id in self._account_to_levels.keys())
        levels = tuple(self._account_to_levels.values())
        row_to_levels = dict(zip(rows, levels))
        return row_to_levels

    def set_column_width(self, field, width: int):
        self._filed_to_width[field] = width

    @property
    def column_width(self) -> dict:
        column_width_dictionary = dict(zip(self._fields, tuple(None for _ in range(len(self._fields)))))
        column_width_dictionary.update(self._filed_to_width)
        return column_width_dictionary

    def __repr__(self):
        return self._name

    # Get Addresses Values
    def _get_addresses(self, account_id, field, number_of_periods, sheet_name, rl, cl, shift) -> tuple:

        initial_row = self._initial_row
        initial_column = self._initial_column
        fields = self._fields
        ac_id_to_row = self.account_id_to_row

        row = initial_row + ac_id_to_row[account_id]
        nop = number_of_periods
        initial_col = initial_column + fields.index(field) + shift
        addresses = tuple(impl.get_address_str(row, c + initial_col, sheet_name, rl, cl) for c in range(nop))

        return addresses

    # Get Address Name
    def get_address_name_without_lock(self, account_id, shift: int = 0) -> str:
        sheet_name = ''
        return self._get_addresses_without_lock(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_locked(self, account_id, shift: int = 0) -> str:
        sheet_name = ''
        return self._get_addresses_locked(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_row_locked(self, account_id, shift: int = 0) -> str:
        sheet_name = ''
        return self._get_addresses_row_locked(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_col_locked(self, account_id, shift: int = 0) -> str:
        sheet_name = ''
        return self._get_addresses_col_locked(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_with_sheet_name_without_lock(self, account_id, shift: int = 0) -> str:
        sheet_name = self._name
        return self._get_addresses_without_lock(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_with_sheet_name_locked(self, account_id, shift: int = 0) -> str:
        sheet_name = self._name
        return self._get_addresses_locked(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_with_sheet_name_row_locked(self, account_id, shift: int = 0) -> str:
        sheet_name = self._name
        return self._get_addresses_row_locked(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_with_sheet_name_col_locked(self, account_id, shift: int = 0) -> str:
        sheet_name = self._name
        return self._get_addresses_col_locked(account_id, self.field_name, 1, sheet_name, shift)[0]

    def get_address_name_without_lock_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_without_sheet_name_without_lock_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    def get_address_name_locked_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_without_sheet_name_locked_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    def get_address_name_row_locked_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_without_sheet_name_row_locked_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    def get_address_name_col_locked_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_without_sheet_name_col_locked_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    def get_address_name_with_sheet_name_without_lock_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_with_sheet_name_without_lock_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    def get_address_name_with_sheet_name_locked_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_with_sheet_name_locked_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    def get_address_name_with_sheet_name_row_locked_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_with_sheet_name_row_locked_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    def get_address_name_with_sheet_name_col_locked_with_eq_sign(self, account_id, shift: int = 0) -> str:
        return self.get_addresses_with_sheet_name_col_locked_with_eq_sign(account_id, self.field_name, 1, shift)[0]

    # Get Addresses Values Row / Col lock
    def _get_addresses_without_lock(self, account_id, field, nop, sheet_name, shift) -> tuple:
        rl = False
        cl = False
        return self._get_addresses(account_id, field, nop, sheet_name, rl, cl, shift)

    def _get_addresses_locked(self, account_id, field, nop, sheet_name, shift) -> tuple:
        rl = True
        cl = True
        return self._get_addresses(account_id, field, nop, sheet_name, rl, cl, shift)

    def _get_addresses_row_locked(self, account_id, field, nop, sheet_name, shift) -> tuple:
        rl = True
        cl = False
        return self._get_addresses(account_id, field, nop, sheet_name, rl, cl, shift)

    def _get_addresses_col_locked(self, account_id, field, nop, sheet_name, shift) -> tuple:
        rl = False
        cl = True
        return self._get_addresses(account_id, field, nop, sheet_name, rl, cl, shift)

    # Get Addresses Values with Sheet name
    def get_addresses_with_sheet_name_without_lock(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_without_lock(account_id, field, nop, self.name, shift)

    def get_addresses_with_sheet_name_row_locked(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_row_locked(account_id, field, nop, self.name, shift)

    def get_addresses_with_sheet_name_col_locked(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_col_locked(account_id, field, nop, self.name, shift)

    def get_addresses_with_sheet_name_locked(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_locked(account_id, field, nop, self.name, shift)

    # Get Addresses Values without Sheet name
    def get_addresses_without_sheet_name_without_lock(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_without_lock(account_id, field, nop, '', shift)

    def get_addresses_without_sheet_name_row_locked(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_row_locked(account_id, field, nop, '', shift)

    def get_addresses_without_sheet_name_col_locked(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_col_locked(account_id, field, nop, '', shift)

    def get_addresses_without_sheet_name_locked(self, account_id, field, nop, shift=0) -> tuple:
        return self._get_addresses_locked(account_id, field, nop, '', shift)

    # Get Addresses Values with Equal sign
    def get_addresses_with_sheet_name_without_lock_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_with_sheet_name_without_lock(account_id, field, nop, shift))

    def get_addresses_with_sheet_name_row_locked_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_with_sheet_name_row_locked(account_id, field, nop, shift))

    def get_addresses_with_sheet_name_col_locked_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_with_sheet_name_col_locked(account_id, field, nop, shift))

    def get_addresses_with_sheet_name_locked_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_with_sheet_name_locked(account_id, field, nop, shift))

    def get_addresses_without_sheet_name_without_lock_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_without_sheet_name_without_lock(account_id, field, nop, shift))

    def get_addresses_without_sheet_name_row_locked_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_without_sheet_name_row_locked(account_id, field, nop, shift))

    def get_addresses_without_sheet_name_col_locked_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_without_sheet_name_col_locked(account_id, field, nop, shift))

    def get_addresses_without_sheet_name_locked_with_eq_sign(self, account_id, field, nop, shift=0) -> tuple:
        return wrapper_equal_sign(self.get_addresses_without_sheet_name_locked(account_id, field, nop, shift))

    # Get Value Addresses
    def get_value_addresses_with_sheet_name_without_lock(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_without_lock(account_id, self.field_values, nop, self.name, shift)

    def get_value_addresses_with_sheet_name_row_locked(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_row_locked(account_id, self.field_values, nop, self.name, shift)

    def get_value_addresses_with_sheet_name_col_locked(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_col_locked(account_id, self.field_values, nop, self.name, shift)

    def get_value_addresses_with_sheet_name_locked(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_locked(account_id, self.field_values, nop, self.name, shift)

    def get_value_addresses_without_sheet_name_without_lock(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_without_lock(account_id, self.field_values, nop, '', shift)

    def get_value_addresses_without_sheet_name_row_locked(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_row_locked(account_id, self.field_values, nop, '', shift)

    def get_value_addresses_without_sheet_name_col_locked(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_col_locked(account_id, self.field_values, nop, '', shift)

    def get_value_addresses_without_sheet_name_locked(self, account_id, nop, shift=0) -> tuple:
        return self._get_addresses_locked(account_id, self.field_values, nop, '', shift)

    def get_value_addresses_with_sheet_name_without_lock_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_with_sheet_name_without_lock_with_eq_sign(account_id, self.field_values, nop, shift)

    def get_value_addresses_with_sheet_name_row_locked_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_with_sheet_name_row_locked_with_eq_sign(account_id, self.field_values, nop, shift)

    def get_value_addresses_with_sheet_name_col_locked_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_with_sheet_name_col_locked_with_eq_sign(account_id, self.field_values, nop, shift)

    def get_value_addresses_with_sheet_name_locked_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_with_sheet_name_locked_with_eq_sign(account_id, self.field_values, nop, shift)

    def get_value_addresses_without_sheet_name_without_lock_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_without_sheet_name_without_lock_with_eq_sign(account_id, self.field_values, nop,
                                                                               shift)

    def get_value_addresses_without_sheet_name_row_locked_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_without_sheet_name_row_locked_with_eq_sign(account_id, self.field_values, nop, shift)

    def get_value_addresses_without_sheet_name_col_locked_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_without_sheet_name_col_locked_with_eq_sign(account_id, self.field_values, nop, shift)

    def get_value_addresses_without_sheet_name_locked_with_eq_sign(self, account_id, nop, shift=0) -> tuple:
        return self.get_addresses_without_sheet_name_locked_with_eq_sign(account_id, self.field_values, nop, shift)

    # Get Value Addresses Application
    def get_target_address(self, account_id, nop: int) -> str:
        return self.get_value_addresses_with_sheet_name_without_lock(account_id, nop)[0]
