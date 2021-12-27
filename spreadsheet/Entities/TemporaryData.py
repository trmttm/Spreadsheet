from typing import Any
from typing import Callable
from typing import Dict


class TemporaryData:
    # Temporary states, data to prevent from having to pass arguments around within the controller (API) layer
    _account_ids_of_worksheet_names_inserted_in_input_sheet = '_account_ids_of_worksheet_names_inserted_in_input_sheet'
    _input_accounts = '_input_accounts'
    _period_account = '_period_account'
    _direct_links = '_direct_links'
    _vertical_accounts = '_vertical_accounts'
    _delayed_format_command = '_delayed_format_command'

    def __init__(self):
        self._data = {
            self._account_ids_of_worksheet_names_inserted_in_input_sheet: (),
            self._input_accounts: (),
            self._period_account: None,
            self._direct_links: (),
            self._vertical_accounts: {},
            self._delayed_format_command: (),
        }

    def add_account_ids_inserted_in_input_sheet(self, account_id):
        self._data[self._account_ids_of_worksheet_names_inserted_in_input_sheet] += (account_id,)

    @property
    def account_ids_of_worksheet_names_inserted_in_input_sheet(self) -> tuple:
        return self._data[self._account_ids_of_worksheet_names_inserted_in_input_sheet]

    def set_input_accounts(self, input_accounts: tuple):
        self._data[self._input_accounts] = input_accounts

    @property
    def input_accounts(self) -> tuple:
        return self._data[self._input_accounts]

    def set_period_account(self, period_account):
        self._data[self._period_account] = period_account

    @property
    def period_account(self):
        return self._data[self._period_account]

    # Limiting all modifications to direct_links_mutable to the following methods.
    def set_direct_links(self, direct_links: tuple):
        self._data[self._direct_links] = direct_links

    def add_direct_link(self, direct_link: tuple):
        self._data[self._direct_links] += (direct_link,)

    def remove_direct_link(self, direct_link: tuple):
        direct_links_mutable = list(self._data[self._direct_links])
        index_ = direct_links_mutable.index(direct_link)
        del direct_links_mutable[index_]
        self.set_direct_links(tuple(direct_links_mutable))

    @property
    def direct_links(self) -> tuple:
        return self._data[self._direct_links]

    def set_vertical_accounts(self, vertical_accounts: Dict[Any, tuple]):
        self._data[self._vertical_accounts] = vertical_accounts

    def add_delayed_format_command(self, method: Callable, args: tuple = (), kwargs: dict = None):
        kwargs = {} if kwargs is None else kwargs
        self._data[self._delayed_format_command] += ((method, args, kwargs),)

    def execute_delayed_format_commands(self):
        for method, args, kwargs in self._data[self._delayed_format_command]:
            method(*args, **kwargs)

    @property
    def vertical_accounts(self) -> Dict[Any, tuple]:
        return self._data[self._vertical_accounts]

    def clear(self):
        self.__init__()
