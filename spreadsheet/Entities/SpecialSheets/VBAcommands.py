class VBAcommands:
    _command_sheet_name = 'commands'

    def __init__(self, file_path: str, my_book_name: str = 'Commands.xlsm'):
        self._path = '/'.join(file_path.split('/')[:-1]) + '/'
        self._my_book_name = my_book_name
        self._file_name = target_file_name = file_path.split('/')[-1]
        self._f = {}

        self._setup_commands = (
            ('open_workbook', (file_path,))
            ,
        )
        self._commands = ()

        self._end_commands = (
            ('save_workbook', (target_file_name,)),
            ('close_workbook', (my_book_name,)),
        )

    @property
    def path(self) -> str:
        return self._path + self._my_book_name

    @property
    def commands(self) -> tuple:
        return self._setup_commands + self._commands + self._end_commands

    def add_command(self, method_name: str, *args):
        new_args = (self._file_name,) + args
        self._commands += ((method_name, new_args),)

    def set_format(self, format_dict):
        self._f = format_dict

    def get_worksheet_data(self) -> tuple:
        name = self._command_sheet_name
        group_of_cell_values_list = []
        max_number_of_columns = 0
        for n, command in enumerate(self.commands):
            row_values = (n, command[0]) + command[1:][0]
            group_of_cell_values_list.append(row_values)
            max_number_of_columns = max(max_number_of_columns, len(row_values))
        group_of_cell_values = tuple(group_of_cell_values_list)
        rows = tuple(range(len(group_of_cell_values)))
        columns = tuple(0 for _ in range(len(group_of_cell_values)))
        formats = tuple(tuple(self._f for _ in range(max_number_of_columns)) for _ in range(len(group_of_cell_values)))
        worksheet_data = name, group_of_cell_values, rows, columns, formats
        return worksheet_data
