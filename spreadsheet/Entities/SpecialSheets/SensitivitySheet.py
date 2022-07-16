class SensitivitySheet:

    def __init__(self, name: str, input_accounts: tuple, shape_id_to_address: dict, format_input: dict = None,
                 format_link: dict = None, format_whole: dict = None, format_2_digit: dict = None,
                 format_percent: dict = None):
        self._name = name
        self._input_accounts = input_accounts
        self._shape_id_to_address = shape_id_to_address
        self._initial_row = 0
        self._initial_column = 0
        self._format_input = format_input or {}
        self._format_link = format_link or {}
        self._format_whole = format_whole or {}
        self._format_2_digit = format_2_digit or {}
        self._format_percent = format_percent or {}

        self._start_row = 0

        self._number_of_selected_variables = 0
        self._graph_title = ''
        self._row_start_variables_xl = 0

    @property
    def data_table_tornado_arguments(self) -> tuple:
        start = self._row_start_variables_xl
        end = start + self._number_of_selected_variables - 1
        args = self._name, f'AB{start - 1}:AE{end}', 'C1', 'C2'
        return args

    @property
    def data_table_spider_arguments(self) -> tuple:
        start = self._row_start_variables_xl
        end = start + self._number_of_selected_variables - 1
        args = self._name, f'AS{start - 1}:BD{end}', 'C1', 'C2'
        return args

    def get_worksheet_data(self, target_to_address: dict, selected_variables: tuple, shape_id_to_delta: dict) -> tuple:
        self._number_of_selected_variables = len(selected_variables)
        name = self._name
        input_accounts = self._input_accounts
        number_of_variables = len(input_accounts)
        shape_id_to_address = self._shape_id_to_address

        # Selectors [Case, Variable, Period, Target]
        cell_values_dict = {
            0: {1: 'Case Selector',
                2: 1,
                3: '=C1=1',
                },
            1: {1: 'Variable selector',
                2: 1,
                3: '="variable number: "&C2',
                },
            3: {
                1: 'Period',
                2: 0,
            },
            4: {
                1: 'Target Values',
                2: 1,
                3: f'=INDEX($B$6:$B${6 + len(target_to_address) - 1},$C$5)',
                4: f'=INDEX($C$6:$C${6 + len(target_to_address) - 1},$C$5)',
                5: '="0,0.0"',
            },
        }

        # Target Values
        self._start_row = max(cell_values_dict.keys()) + 1
        for n, (target_id, target_address) in enumerate(target_to_address.items()):
            cell_values_dict[self._start_row + n] = {
                0: n + 1,
                1: self._shape_id_to_address[target_id],
                2: f'=OFFSET({target_address},0,$C$4)',
            }

        # A few items
        self._start_row = max(cell_values_dict.keys()) + 1
        start_row = self._start_row + 1
        start_row_xl = start_row + 1
        row_first_input_xl = start_row_xl + 2
        cell_values_dict[start_row] = {
            19: f'=COUNTIF(T{row_first_input_xl}:T{row_first_input_xl + (number_of_variables - 1)},TRUE)',
            24: f'=ROW(Y{start_row_xl + 1})',
            27: 'Tornado - input variables',
            37: '''="Sensitivity by inputs : "&$D$5&" (current case ="&TEXT($E$5,$F$5)&" period="&C4&")"''',
            41: f'=OFFSET(AN{start_row_xl + 1},T{start_row_xl},0)*20%',
            44: 'Spider - input variables',
            59: '''="Spider : "&$B$3&" (current case ="&TEXT($D$3,"0,0")&" "&$C$3&")"''',
        }

        self._graph_title = f"='{self._name}'!$AL${start_row_xl}"

        # Titles
        self._start_row = max(cell_values_dict.keys())
        self._row_start_variables_xl = self._start_row + 3
        cell_values_dict[self._start_row + 1] = {
            0: 'No',
            1: "Sensitivity Analysis on Green inputs, RED border is irregular formula",
            2: 'delta: high case [%]',
            # 3: 'Bigger = better?',
            4: 'Base',
            5: 'Low',
            6: 'High',
            7: 4,
            8: 5,
            9: 6,
            10: 7,
            11: 8,
            12: 9,
            13: 10,
            14: 11,
            15: 12,
            16: 13,
            17: 14,
            18: '',
            19: 'Variable Selected',
            20: 'Base',
            21: 'Test',
            22: 'Applied in model',
            23: '',
            24: 'Selected?',
            25: 'Sort',
            26: '',
            27: "=$E$5",
            28: 1,
            29: 2,
            30: 3,
            31: '',
            32: 'Low vs Base',
            33: 'High vs Base',
            34: 'abs delta',
            35: 'Large',
            36: 'Match',
            37: 'Name',
            38: 'Low vs Base',
            39: 'High vs Base',
            40: 'label',
            41: 'label',
            42: 'Elimination',
            43: '',
            44: "=E5",
            45: f'=H{self._row_start_variables_xl - 1}',
            46: f'=I{self._row_start_variables_xl - 1}',
            47: f'=J{self._row_start_variables_xl - 1}',
            48: f'=K{self._row_start_variables_xl - 1}',
            49: f'=L{self._row_start_variables_xl - 1}',
            50: f'=M{self._row_start_variables_xl - 1}',
            51: f'=N{self._row_start_variables_xl - 1}',
            52: f'=O{self._row_start_variables_xl - 1}',
            53: f'=P{self._row_start_variables_xl - 1}',
            54: f'=Q{self._row_start_variables_xl - 1}',
            55: f'=R{self._row_start_variables_xl - 1}',
            56: '',
            57: '',
            58: '',
            59: f'=H{self._row_start_variables_xl}',
            60: f'=I{self._row_start_variables_xl}',
            61: f'=J{self._row_start_variables_xl}',
            62: f'=K{self._row_start_variables_xl}',
            63: f'=L{self._row_start_variables_xl}',
            64: f'=M{self._row_start_variables_xl}',
            65: f'=N{self._row_start_variables_xl}',
            66: f'=O{self._row_start_variables_xl}',
            67: f'=P{self._row_start_variables_xl}',
            68: f'=Q{self._row_start_variables_xl}',
            69: f'=R{self._row_start_variables_xl}',
        }

        # Calculations
        self._start_row = max(cell_values_dict.keys()) + 1
        start_row = self._start_row
        start_row_xl = start_row + 1
        last_row = start_row + number_of_variables
        last_row_xl = last_row
        for n, input_account in enumerate(input_accounts):
            row = start_row + n + 1
            cell_values_dict[n + start_row] = {
                0: n + 1,
                # 1: shape_id_to_address[input_account],
                1: f'''={shape_id_to_address[input_account]}&" [+/-"&TEXT(ABS(C{row}),"0.0")&"%]"''',
                2: shape_id_to_delta[input_account] + n * 0.00001,
                4: 1,
                5: f'=$E{row}-$C{row}/100',
                6: f'=$E{row}+$C{row}/100',
                7: .5,
                8: .6,
                9: .7,
                10: .8,
                11: .9,
                12: 1,
                13: 1.1,
                14: 1.2,
                15: 1.3,
                16: 1.4,
                17: 1.5,
                19: input_account in selected_variables,
                20: f'=INDEX(E{row}:R{row},$C$1)',
                21: f'=$A{row}=$C$2',
                22: f'=IF(V{row},U{row},E{row})',
                24: f'=MAX(0,ROW(T{row})*T{row}-$Y${start_row_xl - 2})',
                25: f'=COUNTIF($Y${start_row_xl}:$Y${last_row_xl},">0")' if n == 0 else f'=+MAX(0,Z{row - 1}-1)',
                27: f'=IFERROR(LARGE($Y${start_row_xl}:$Y${last_row_xl},Z{row}),"")',
                32: f'=AD{row}-AC{row}',
                33: f'=AE{row}-AC{row}',
                34: f'=ABS(AG{row})+ABS(AH{row})',
                35: f'=LARGE(OFFSET($AI${start_row_xl},0,0,$Z${start_row_xl},1),Z{row})',
                36: f'=INDEX($AB${start_row_xl}:$AB${last_row_xl},MATCH(AJ{row},$AI${start_row_xl}:$AI${last_row_xl},0))',
                37: f'=IF(AND(AM{row}=0,AN{row}=0),"",INDEX($B${start_row_xl}:$B${last_row_xl},MATCH(AK{row},$A${start_row_xl}:$A${last_row_xl},0)))',
                38: f'=INDEX($AG${start_row_xl}:$AG${last_row_xl},MATCH(AJ{row},$AI${start_row_xl}:$AI${last_row_xl},0))',
                39: f'=INDEX($AH${start_row_xl}:$AH${last_row_xl},MATCH(AJ{row},$AI${start_row_xl}:$AI${last_row_xl},0))',
                40: f'=-$AP$2' if n == 0 else f'=AO{row - 1}',
                41: f'=$AP$2' if n == 0 else f'=AP{row - 1}',
                42: f'=(AM{row}=0)*(AN{row}=0)' if n == 0 else f'=(AM{row}=0)*(AN{row}=0)+AQ{row - 1}',
                44: f'=AB{row}',
                57: f'=INDEX($AL${start_row_xl}:$AL${last_row_xl},MATCH(AS{row},$AK${start_row_xl}:$AK${last_row_xl},0))',
                58: f'=BF{row}',
                59: f'=AT{row}',
                60: f'=AU{row}',
                61: f'=AV{row}',
                62: f'=AW{row}',
                63: f'=AX{row}',
                64: f'=AY{row}',
                65: f'=AZ{row}',
                66: f'=BA{row}',
                67: f'=BB{row}',
                68: f'=BC{row}',
                69: f'=BD{row}',
            }

        max_rows = len(cell_values_dict.keys())
        max_column = 0
        for row, data in cell_values_dict.items():
            max_column = max(max_column, max(data.keys()))
        max_column += 1

        group_of_cell_values = []
        rows_to_loop_in_range = max_rows + 2
        for row in range(rows_to_loop_in_range):
            if row in cell_values_dict:
                values = []
                for column in range(max_column):
                    if column in cell_values_dict[row]:
                        values.append(cell_values_dict[row][column])
                    else:
                        values.append('')
                group_of_cell_values.append(values)
        group_of_cell_values = tuple(group_of_cell_values)

        rows = tuple(self._initial_row + r for r in cell_values_dict.keys())
        rows_sorted = sorted(rows)
        rows = tuple(rows_sorted)
        columns = tuple(self._initial_column for _ in cell_values_dict.keys())
        formats = self._get_sheet_format(max_column, rows, len(target_to_address))

        return name, group_of_cell_values, rows, columns, formats

    def _get_sheet_format(self, max_column, rows, number_of_targets: int) -> tuple:
        formats_list = []
        target_row_start = 5
        target_row_end = target_row_start + number_of_targets - 1
        variables_title_row = target_row_end + 3
        for row in rows:
            if row in (0, 1, 3, 4):
                formats_row_list = []
                for column in range(max_column):
                    f = {}
                    if column in (2, 5):
                        f.update(self._format_input)
                    f.update(self._format_whole)
                    formats_row_list.append(f)
            elif target_row_start <= row <= target_row_end:
                formats_row_list = []
                for column in range(max_column):
                    f = {}
                    if column in (1,):
                        f.update(self._format_link)
                    f.update(self._format_whole)
                    formats_row_list.append(f)
            elif variables_title_row == row:
                formats_row_list = []
                percent_columns = tuple(range(59, 70))
                for column in range(max_column):
                    f = {}
                    if column in percent_columns:
                        f.update(self._format_percent)
                    else:
                        f.update(self._format_whole)
                    formats_row_list.append(f)
            elif variables_title_row + 1 <= row:
                formats_row_list = []
                input_columns = (2, 4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19)
                two_digit_columns = ()
                percent_columns = (4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 22)
                link_columns = (1,)
                for column in range(max_column):
                    f = {}
                    if column in input_columns:
                        f.update(self._format_input)
                    elif column in link_columns:
                        f.update(self._format_link)

                    if column in two_digit_columns:
                        f.update(self._format_2_digit)
                    elif column in percent_columns:
                        f.update(self._format_percent)
                    else:
                        f.update(self._format_whole)

                    formats_row_list.append(f)
            else:
                formats_row_list = [{} for _ in range(max_column)]
            formats_list.append(tuple(formats_row_list))
        formats = tuple(formats_list)
        return formats

    def get_row(self, input_account_id) -> int:
        return self._start_row + self._input_accounts.index(input_account_id)

    def get_address_sensitivity_applied(self, input_account_id) -> str:
        row = self.get_row(input_account_id) + 1
        return f"='{self._name}'!$W${row}"

    @property
    def graph_inputs_tornado(self) -> tuple:
        title = self._graph_title
        start = self._row_start_variables_xl
        end = start + self._number_of_selected_variables - 1

        position = f"$D$1"
        address_values_list = (f"'{self._name}'!$AM${start}:$AM${end}", f"'{self._name}'!$AN${start}:$AN${end}")
        address_names = f"'{self._name}'!$AM${start - 1}", f"'{self._name}'!$AN${start - 1}"
        address_category = f"'{self._name}'!$AL${start}:$AL${end}"
        return title, position, address_values_list, address_names, address_category

    @property
    def graph_inputs_spider(self) -> tuple:
        title = self._graph_title
        start = self._row_start_variables_xl
        end = start + self._number_of_selected_variables - 1

        position = f"$D$1"

        address_values_list = tuple(f"'{self._name}'!$BH${row}:$BR${row}" for row in range(start, end + 1))
        address_names = tuple(f"'{self._name}'!$BG${row}" for row in range(start, end + 1))
        address_category = f"'{self._name}'!$BH${start - 1}:$BR${start - 1}"
        return title, position, address_values_list, address_names, address_category
