from typing import Callable

from . import util
from ..Entities import Account
from ..Entities import Accounts
from ..Entities import Worksheet
from ..Entities import Worksheets


def connect_direct_links(accounts: Accounts, direct_links: tuple, input_sheet_name: str, nop: int,
                         worksheets: Worksheets):
    for ac_from, account_to_id, shift in direct_links:
        sheet_from = worksheets.identify_worksheet(ac_from)
        sheet_to = worksheets.identify_worksheet(account_to_id)
        ac_to = accounts.get_account(account_to_id)
        _connect_direct_links(ac_from, ac_to, input_sheet_name, nop, shift, sheet_from, sheet_to)


def _connect_direct_links(ac_from: Account, ac_to: Account, input_sheet_name: str, nop: int, shift: int,
                          sheet_from: Worksheet, sheet_to: Worksheet):
    args = ac_from, nop, shift
    is_link_from_input_sheet = (sheet_from.name == input_sheet_name)
    with_worksheet = sheet_to.name != sheet_from.name
    if is_link_from_input_sheet:
        addresses = sheet_from.get_value_addresses_with_sheet_name_row_locked_with_eq_sign(*args)
    elif with_worksheet:
        addresses = sheet_from.get_value_addresses_with_sheet_name_without_lock_with_eq_sign(*args)
    else:
        addresses = sheet_from.get_value_addresses_without_sheet_name_without_lock_with_eq_sign(*args)
    ac_to.set_values(addresses)


def is_direct_link_from_domestic_input(account_from, account_to, input_sheet_name, worksheets: Worksheets) -> bool:
    worksheet_from = worksheets.identify_worksheet(account_from)
    worksheet_to = worksheets.identify_worksheet(account_to)
    if worksheet_from.name == input_sheet_name:
        domestic_input_id = util.get_domestic_input_id(worksheet_to.name, account_from)
        if domestic_input_id in worksheet_to.account_ids:
            return True
    return False


def connect_domestic_input(clients: tuple, domestic_input_id, input_id, new_links: list, old_links: list):
    new_links.append((input_id, domestic_input_id, 0))
    for client in clients:
        new_links.append((domestic_input_id, client, 0))
        old_links.append((input_id, client, 0))


def update_direct_links(new_links: list, old_links: list, add: Callable, remove: Callable):
    for new_direct_link in new_links:
        add(new_direct_link)
    for old_direct_link in old_links:
        remove(old_direct_link)
