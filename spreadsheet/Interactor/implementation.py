from ..Entities import Accounts
from ..Entities import Worksheet
from ..Entities import Worksheets


def add_account_to_worksheet(row: int, account_id, account_text: str, accounts: Accounts, worksheet: Worksheet):
    add_accounts_to_worksheet((row,), (account_id,), (account_text,), accounts, worksheet)


def add_accounts_to_worksheet(rows: tuple, account_ids: tuple, account_texts: tuple, accounts: Accounts,
                              worksheet: Worksheet):
    for row, account_id, text in zip(rows, account_ids, account_texts):
        if account_id not in accounts:
            accounts.create_new_account(account_id, text)
        worksheet.add_account(row, account_id)
        worksheet.set_default_formats(account_id)


def create_shape_id_to_address_name(account_ids: tuple, worksheets: Worksheets, shifts: tuple) -> dict:
    addresses = []
    for account_id, shift in zip(account_ids, shifts):
        sheet = worksheets.identify_worksheet(account_id)
        address = sheet.get_address_name_with_sheet_name_locked_with_eq_sign(account_id, shift)
        addresses.append(address)
    shape_id_to_address_text = dict(zip(account_ids, addresses))
    return shape_id_to_address_text


def create_new_worksheets(accounts: Accounts, accounts_not_to_indent: tuple, default_format: dict, id_to_text: dict,
                          worksheets: Worksheets, worksheets_data: dict):
    for sheet_name, worksheet_contents in worksheets_data.items():
        new_worksheet = worksheets.create_new_worksheet(sheet_name)
        rows = tuple(n for (n, c) in enumerate(worksheet_contents) if str(c) != 'blank')
        account_ids = tuple(c for c in worksheet_contents if str(c) != 'blank')
        account_texts = tuple(id_to_text[ac] for ac in account_ids)

        add_accounts_to_worksheet(rows, account_ids, account_texts, accounts, new_worksheet)
        for account_id in account_ids:
            new_worksheet.set_values_format(account_id, default_format)

            if account_id not in accounts_not_to_indent:
                account = accounts.get_account(account_id)
                account.indent()
