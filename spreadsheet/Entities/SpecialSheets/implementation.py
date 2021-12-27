def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def get_address_str(row: int, column: int, sheet_name: str = '', row_lock=False, col_lock=False) -> str:
    column_str = column_string(column + 1)
    sheet_name = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
    sheet = f'{sheet_name}!' if sheet_name != '' else ''
    row_lock_str = '$' if row_lock else ''
    col_lock_str = '$' if col_lock else ''
    return f'{sheet}{col_lock_str}{column_str}{row_lock_str}{row + 1}'
