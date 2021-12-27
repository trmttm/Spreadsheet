from ..Entities import Worksheets


def get_input_sensitivity_account_id(input_id) -> str:
    return f'input_sensitivity_{input_id}'


def create_target_account_text_to_address(target_account_ids: tuple, nop: int, worksheets: Worksheets) -> dict:
    def get_address(account_id_) -> str:
        return worksheets.identify_worksheet(account_id_).get_target_address(account_id_, nop)

    target_addresses = tuple(get_address(target_account_id) for target_account_id in target_account_ids)
    target_name_to_address = dict(zip(target_account_ids, target_addresses))
    return target_name_to_address
