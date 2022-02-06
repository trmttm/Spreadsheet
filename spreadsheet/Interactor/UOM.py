from ..Entities import Accounts


def set_uom_to_accounts(accounts: Accounts, direct_links: tuple, field_name: str, shape_id_to_uom: dict):
    for shape_id, uom in shape_id_to_uom.items():
        shape_id_int = convert_shape_id_to_integer(shape_id)
        account = accounts.get_account(shape_id_int)
        account.add_new_attribute(field_name, uom)
    for account_from, account_to, _ in direct_links:
        shape_id_int = convert_shape_id_to_integer(account_to)
        account = accounts.get_account(shape_id_int)

        # If account_from's UOM is '' and account_to's is not, then take account_to's UOM
        uom = shape_id_to_uom.get(account_from, '') or shape_id_to_uom.get(account_to, '')

        account.add_new_attribute(field_name, uom)


def convert_shape_id_to_integer(shape_id):
    try:
        shape_id_int = int(shape_id)
    except ValueError:
        shape_id_int = shape_id
    return shape_id_int
