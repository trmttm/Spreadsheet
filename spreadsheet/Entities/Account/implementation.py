def get_cell_values(account, fields: dict) -> tuple:
    cell_values = []
    for attribute in fields:
        if attribute == 'values':
            cell_values += [v for v in getattr(account, attribute)]
        else:
            try:
                cell_values.append(getattr(account, attribute))
            except AttributeError:
                cell_values.append('')
    cell_values = tuple(cell_values)
    return cell_values
