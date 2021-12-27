from typing import Any
from typing import Dict

from . import implementation as impl
from ..Account import Account


class Accounts(Dict[Any, Account]):
    def __init__(self):
        dict.__init__(self)

    @property
    def all_account_ids(self) -> tuple:
        return tuple(self.keys())

    def get_account(self, account_id) -> Account:
        return self[account_id]

    def get_group_of_cell_values(self, account_ids: tuple, fields: dict) -> tuple:
        return impl.get_group_of_cell_values(self, account_ids, fields)

    def get_formats(self, account_ids: tuple, field_to_format: dict, fields: tuple) -> tuple:
        formats_list = []
        for account_id in account_ids:
            account = self.get_account(account_id)
            formats = []
            for attribute in fields:
                if attribute in ['values']:
                    values = getattr(account, attribute)
                    formats += [field_to_format[account_id][attribute] for _ in values]
                else:
                    formats.append(field_to_format[account_id][attribute])
            formats_list.append(tuple(formats))
        formats = tuple(formats_list)
        return formats

    def create_new_account(self, account_id, text):
        new_account = Account(account_id, text)
        self[account_id] = new_account

