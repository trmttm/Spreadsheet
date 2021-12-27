from typing import Iterable


def get_domestic_input_id(worksheet_name: str, account_id) -> str:
    return f'{"domestic_input_"}{worksheet_name}_{account_id}'


def create_rpe_dictionary(rpes: Iterable) -> dict:
    keys = tuple(item[0] for item in rpes)
    formulas = tuple(item[1] for item in rpes)
    rpes_dict = dict(zip(keys, formulas))
    return rpes_dict
