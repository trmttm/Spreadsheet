def get_group_of_cell_values(accounts, account_ids: tuple, fields: dict) -> tuple:
    cell_values_list = []
    for account_id in account_ids:
        account = accounts.get_account(account_id)
        cell_values = account.get_cell_values(fields)
        cell_values_list.append(cell_values)
    group_of_cell_values = tuple(cell_values_list)
    return group_of_cell_values
