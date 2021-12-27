from typing import Any
from typing import Dict
from typing import Tuple
from typing import Union

from .Worksheet import Worksheet


class Worksheets(Dict[Any, Worksheet]):
    def __init__(self):
        dict.__init__(self)

    def create_new_worksheet(self, name: str, initial_row=1, initial_column=1) -> Worksheet:
        if name not in self:
            new_worksheet = Worksheet(name, initial_row, initial_column)
            self[name] = new_worksheet
            return new_worksheet
        else:
            return self[name]

    @property
    def all_worksheets(self) -> Tuple[Worksheet]:
        return tuple(self.values())

    @property
    def worksheet_names(self) -> tuple:
        return tuple(self.keys())

    def get_worksheet(self, sheet_name) -> Worksheet:
        return self[sheet_name]

    def identify_worksheet(self, account_id, priorities: tuple = (), ) -> Union[None, Worksheet]:
        worksheets_list = []
        for worksheet in self.values():
            if worksheet.has_account(account_id):
                worksheets_list.append(worksheet)

        worksheet_dict = dict(zip(tuple(w.name for w in worksheets_list), worksheets_list))
        n = len(worksheet_dict)
        if n == 0:
            return None
        elif n == 1:
            return worksheets_list[0]
        else:
            for priority in priorities:
                if priority in worksheet_dict:
                    return worksheet_dict[priority]
        return worksheets_list[0]
