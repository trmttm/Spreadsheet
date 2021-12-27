import os

import os_identifier


def test_case_export_input_sheet_as_excel():
    if os_identifier.is_mac:
        file_path_common = f'{os.getcwd().replace(".", "")}Tests/pickles/'
    else:
        file_path_common = 'C:\\Users\\Yamaka\\Documents\\FMDesigner/Tests/pickles\\'
    workbook_name1 = wb1 = f'{file_path_common}export_input_sheet1.xlsx'
    request_model1 = wb1, (0, 1), (0, 1), ('Volume', 'Price'), tuple(range(10)), tuple(range(10, 20))

    expected_gateway_model1 = (workbook_name1,
                               (('Inputs',
                                 (('Volume', '', '', 0, 1, 2, 3, 4, 5, 6, 7, 8, 9),
                                  ('Price', '', '', 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)),
                                 (1, 2),
                                 (1, 1),
                                 (({},
                                   {},
                                   {},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                  ({},
                                   {},
                                   {},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                   {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}))),))
    expected_spreadsheet_model1 = (workbook_name1,
                                   (('Volume', 'Inputs', (1, 1), {}),
                                    ('', 'Inputs', (1, 2), {}),
                                    ('', 'Inputs', (1, 3), {}),
                                    (0,
                                     'Inputs',
                                     (1, 4),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (1,
                                     'Inputs',
                                     (1, 5),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (2,
                                     'Inputs',
                                     (1, 6),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (3,
                                     'Inputs',
                                     (1, 7),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (4,
                                     'Inputs',
                                     (1, 8),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (5,
                                     'Inputs',
                                     (1, 9),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (6,
                                     'Inputs',
                                     (1, 10),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (7,
                                     'Inputs',
                                     (1, 11),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (8,
                                     'Inputs',
                                     (1, 12),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (9,
                                     'Inputs',
                                     (1, 13),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    ('Price', 'Inputs', (2, 1), {}),
                                    ('', 'Inputs', (2, 2), {}),
                                    ('', 'Inputs', (2, 3), {}),
                                    (10,
                                     'Inputs',
                                     (2, 4),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (11,
                                     'Inputs',
                                     (2, 5),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (12,
                                     'Inputs',
                                     (2, 6),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (13,
                                     'Inputs',
                                     (2, 7),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (14,
                                     'Inputs',
                                     (2, 8),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (15,
                                     'Inputs',
                                     (2, 9),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (16,
                                     'Inputs',
                                     (2, 10),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (17,
                                     'Inputs',
                                     (2, 11),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (18,
                                     'Inputs',
                                     (2, 12),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                    (19,
                                     'Inputs',
                                     (2, 13),
                                     {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'})))

    request_models = (request_model1,)
    expected_gateway_models = (expected_gateway_model1,)
    expected_spreadsheet_models = (expected_spreadsheet_model1,)
    return request_models, expected_gateway_models, expected_spreadsheet_models


def test_case_export_calculation_sheet_as_excel():
    from .blank_dummy import Blank
    if os_identifier.is_mac:
        file_path_common = f'{os.getcwd().replace(".", "")}Tests/pickles/'
    else:
        file_path_common = 'C:\\Users\\Yamaka\\Documents\\FMDesigner/Tests/pickles\\'
    request_model1 = (
        f'{file_path_common}export_calculation1.xlsx',
        {0: 'Volume', 1: 'Price', 2: 'Revenue', 3: 'x'},
        (0, 1),
        (tuple(range(5)), tuple(range(10, 15))),
        (3,),
        {'Revenue Sheet': (0, 1, Blank(), 2)},
        ((2, (0, 1, 3)),),
    )
    expected_gateway_model1 = (
        (f'{file_path_common}export_calculation1.xlsx',
         (('Inputs',
           (('Volume', '', '', 0, 1, 2, 3, 4), ('Price', '', '', 10, 11, 12, 13, 14)),
           (1, 2),
           (1, 1),
           (({},
             {},
             {},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
            ({},
             {},
             {},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'color': 'blue', 'number_format': '#,##0.0_);[Red]▲#,##0.0'}))),
          ('Revenue Sheet',
           (('',
             'Volume',
             '',
             '=Inputs!E$2',
             '=Inputs!F$2',
             '=Inputs!G$2',
             '=Inputs!H$2',
             '=Inputs!I$2'),
            ('',
             'Price',
             '',
             '=Inputs!E$3',
             '=Inputs!F$3',
             '=Inputs!G$3',
             '=Inputs!H$3',
             '=Inputs!I$3'),
            ('', 'Revenue', '', '=E2*E3', '=F2*F3', '=G2*G3', '=H2*H3', '=I2*I3')),
           (1, 2, 4),
           (1, 1, 1),
           (({}, {}, {}, {}, {}, {}, {}, {}),
            ({}, {}, {}, {}, {}, {}, {}, {}),
            ({},
             {},
             {},
             {'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'number_format': '#,##0.0_);[Red]▲#,##0.0'},
             {'number_format': '#,##0.0_);[Red]▲#,##0.0'}))))))
    request_models = (request_model1,)
    expected_gateway_models = (expected_gateway_model1,)
    return expected_gateway_models, request_models


def test_cases_external_api():
    from .blank_dummy import Blank
    blank = Blank()
    if os_identifier.is_mac:
        file_path_common = f'{os.getcwd().replace(".", "")}Tests/pickles/'
    else:
        file_path_common = 'C:\\Users\\Yamaka\\Documents\\FMDesigner/Tests/pickles\\'

    request_model1 = (
        f'{file_path_common}Excel.xlsx',
        {19: 'Excess Cash', 20: 'Excess Deficit', 21: 'Plug', 22: 'Cash', 24: 'Non cash assets',
         25: 'Liabilities NC', 26: 'Total Equity', 27: 'Non Cash CA', 28: 'Non Current Assets',
         2: '+',
         3: '+', 4: '-', 5: 'Net Cash', 0: 'max', 1: '0', 6: '0', 7: 'min', 8: '-1', 9: 'x',
         10: 'Non Plug CL', 11: 'Total Assets', 12: 'Non Current Assets', 13: 'Current Assets',
         14: 'Cash', 15: 'Non Cash', 16: '+', 17: '+', 18: 'Total Liabilities',
         23: 'Non Current Liabilities', 29: 'Current Liabilities', 30: 'Non Plug CL',
         31: 'Plug - Revolver', 32: '+', 33: '+', 34: 'Total Equity', 35: 'Other Equity',
         36: 'Retained Earnings', 37: 'Paid in Capital', 38: '+', 39: 'Non Cash',
         40: 'Non Current Assets', 41: 'Total Equity', 42: 'Non Plug CL',
         43: 'Non Current Liabilities', 44: 'Cash', 45: 'Plug', 46: 'Total Assets',
         47: 'Total Liabilities', 48: 'Total Equity', 49: 'Dr Cr match', 50: '+', 51: '-',
         52: 'Total Liabilities', 53: 'Total Equity', 54: 'Total Assets'},
        (12, 15, 23, 30, 35, 36, 37),
        ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
         (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
         (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
         (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
         (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
         (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
         (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
        (2, 3, 4, 0, 7, 9, 16, 17, 32, 33, 38, 50, 51),
        (1, 6, 8),
        {
            'Sheet1': (
                14, 15, 13, 12, 11, blank, 30, 31, 29, 23, 18, blank, 37, 36, 35, 34, blank, 26, 10,
                25,
                blank, 27, 28, 24, blank, 5, blank, 19, 22, 20, 21, blank, 46, 47, 48, 49, blank)},
        ((24, (27, 28, 2)), (5, (26, 10, 3, 25, 3, 24, 4)), (19, (1, 5, 0)), (20, (5, 6, 7)),
         (21, (8, 20, 9)), (13, (14, 15, 16)), (11, (12, 13, 17)), (29, (30, 31, 32)),
         (18, (23, 29, 33)), (34, (35, 36, 38, 37, 38)), (49, (46, 47, 48, 50, 51))),
        10,
        ((30, 10, 0), (19, 22, 0), (34, 48, 0), (23, 25, 0), (21, 31, 0), (34, 26, 0), (11, 46, 0),
         (12, 28, 0), (22, 14, 0), (15, 27, 0), (18, 47, 0)),)

    request_model2 = (f'{file_path_common}bb_output.xlsx',
                      {0: 'AA', 1: 'AA BB', 2: 'increase', 3: 'decrease', 4: '+', 5: '-', 6: 'BB'},
                      (2, 3),
                      ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0), (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
                      (4, 5),
                      (),
                      {'Sheet1': (1, 2, 3, 0, blank)},
                      ((0, (1, 2, 4, 3, 5)),),
                      10,
                      ((0, 1, -1),))
    # from FMDesigner unit tests (gateway models created from test pickles)
    request_models_from_pickles = (
        (f'{file_path_common}operation_ave.xlsx',
         {0: 'A', 1: 'B', 2: 'C', 3: 'ave', 4: 'D', 5: 'E'},
         (0, 1, 4, 5),
         ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
         (3,),
         (),
         {'Sheet1': (0, 1, 4, 5, 2, blank)},
         ((2, (0, 1, 3, 4, 3, 5, 3)),),
         10,
         ()),

        (f'{file_path_common}operation_eq.xlsx',
         {0: 'A', 1: 'B', 2: 'C', 3: '='},
         (0, 1),
         ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0), (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
         (3,),
         (),
         {'Sheet1': (0, 1, 2)},
         ((2, (0, 1, 3)),),
         10,
         ()),

        (f'{file_path_common}operation_ge.xlsx',
         {0: 'A', 1: 'B', 2: 'C', 3: '>='},
         (0, 1),
         ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0), (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
         (3,),
         (),
         {'Sheet1': (0, 1, 2)},
         ((2, (0, 1, 3)),),
         10,
         ()),

        (f'{file_path_common}operation_min.xlsx',
         {0: 'A', 1: 'B', 2: 'C', 3: 'min', 4: 'D', 5: 'E'},
         (0, 1, 4, 5),
         ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
         (3,),
         (),
         {'Sheet1': (0, 1, 4, 5, 2, blank)},
         ((2, (0, 1, 3, 4, 3, 5, 3)),),
         10,
         ()),

        (f'{file_path_common}operation_abs.xlsx',
         {0: 'Cash from BS',
          1: 'Cash from CFWF',
          3: 'Difference - absolute',
          4: '-',
          6: 'abs'},
         (0, 1),
         ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0), (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
         (4, 6),
         (),
         {'Sheet1': (0, 1, 3, blank)},
         ((3, (0, 1, 4, 6)),),
         10,
         ()),

        (f'{file_path_common}test_socket_saved_as_a_module.xlsx',
         {0: 'D', 1: 'X', 2: 'A', 3: 'B', 4: 'C', 5: '+', 'Socket_to_D': 'x'},
         (0, 2, 3),
         ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
          (0, 0, 0, 0, 0, 0, 0, 0, 0, 0)),
         ('Socket_to_D', 5),
         (),
         {'Sheet From': (2, 3, 4), 'Sheet1': (0, 1)},
         ((1, (0, 4, 'Socket_to_D')), (4, (2, 3, 5))),
         10,
         ()),

        (f'{file_path_common}test_auto_relay_across_sheets.xlsx',
         {0: 'D', 1: 'X', 2: 'D', 'Socket_to_D': 'x'},
         (0,),
         ((0, 0, 0, 0, 0, 0, 0, 0, 0, 0),),
         ('Socket_to_D',),
         (),
         {'Sheet1': (0,), 'Sheet2': (1,)},
         ((1, (0, 'Socket_to_D')),),
         10,
         ())
    )

    expected_gateway_model1 = (f'{file_path_common}Excel.xlsx', (('Inputs', (
        ('Non Cash', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        ('Non Current Assets', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        ('Non Plug CL', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        ('Non Current Liabilities', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        ('Paid in Capital', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        ('Retained Earnings', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        ('Other Equity', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2, 3, 4, 5, 6, 7), (1, 1, 1, 1, 1, 1, 1),
                                                                  (({}, {}, {},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}))),
                                                                 ('Sheet1', (('',
                                                                              'Cash',
                                                                              '',
                                                                              '=E30',
                                                                              '=F30',
                                                                              '=G30',
                                                                              '=H30',
                                                                              '=I30',
                                                                              '=J30',
                                                                              '=K30',
                                                                              '=L30',
                                                                              '=M30',
                                                                              '=N30'),
                                                                             ('',
                                                                              'Non Cash',
                                                                              '',
                                                                              '=Inputs!E$2',
                                                                              '=Inputs!F$2',
                                                                              '=Inputs!G$2',
                                                                              '=Inputs!H$2',
                                                                              '=Inputs!I$2',
                                                                              '=Inputs!J$2',
                                                                              '=Inputs!K$2',
                                                                              '=Inputs!L$2',
                                                                              '=Inputs!M$2',
                                                                              '=Inputs!N$2'),
                                                                             ('',
                                                                              'Current Assets',
                                                                              '',
                                                                              '=E2+E3',
                                                                              '=F2+F3',
                                                                              '=G2+G3',
                                                                              '=H2+H3',
                                                                              '=I2+I3',
                                                                              '=J2+J3',
                                                                              '=K2+K3',
                                                                              '=L2+L3',
                                                                              '=M2+M3',
                                                                              '=N2+N3'),
                                                                             ('',
                                                                              'Non Current Assets',
                                                                              '',
                                                                              '=Inputs!E$3',
                                                                              '=Inputs!F$3',
                                                                              '=Inputs!G$3',
                                                                              '=Inputs!H$3',
                                                                              '=Inputs!I$3',
                                                                              '=Inputs!J$3',
                                                                              '=Inputs!K$3',
                                                                              '=Inputs!L$3',
                                                                              '=Inputs!M$3',
                                                                              '=Inputs!N$3'),
                                                                             ('',
                                                                              'Total Assets',
                                                                              '',
                                                                              '=E5+E4',
                                                                              '=F5+F4',
                                                                              '=G5+G4',
                                                                              '=H5+H4',
                                                                              '=I5+I4',
                                                                              '=J5+J4',
                                                                              '=K5+K4',
                                                                              '=L5+L4',
                                                                              '=M5+M4',
                                                                              '=N5+N4'),
                                                                             ('',
                                                                              'Non Plug CL',
                                                                              '',
                                                                              '=Inputs!E$4',
                                                                              '=Inputs!F$4',
                                                                              '=Inputs!G$4',
                                                                              '=Inputs!H$4',
                                                                              '=Inputs!I$4',
                                                                              '=Inputs!J$4',
                                                                              '=Inputs!K$4',
                                                                              '=Inputs!L$4',
                                                                              '=Inputs!M$4',
                                                                              '=Inputs!N$4'),
                                                                             ('',
                                                                              'Plug - Revolver',
                                                                              '',
                                                                              '=E32',
                                                                              '=F32',
                                                                              '=G32',
                                                                              '=H32',
                                                                              '=I32',
                                                                              '=J32',
                                                                              '=K32',
                                                                              '=L32',
                                                                              '=M32',
                                                                              '=N32'),
                                                                             ('',
                                                                              'Current Liabilities',
                                                                              '',
                                                                              '=E8+E9',
                                                                              '=F8+F9',
                                                                              '=G8+G9',
                                                                              '=H8+H9',
                                                                              '=I8+I9',
                                                                              '=J8+J9',
                                                                              '=K8+K9',
                                                                              '=L8+L9',
                                                                              '=M8+M9',
                                                                              '=N8+N9'),
                                                                             ('',
                                                                              'Non Current Liabilities',
                                                                              '',
                                                                              '=Inputs!E$5',
                                                                              '=Inputs!F$5',
                                                                              '=Inputs!G$5',
                                                                              '=Inputs!H$5',
                                                                              '=Inputs!I$5',
                                                                              '=Inputs!J$5',
                                                                              '=Inputs!K$5',
                                                                              '=Inputs!L$5',
                                                                              '=Inputs!M$5',
                                                                              '=Inputs!N$5'),
                                                                             ('',
                                                                              'Total Liabilities',
                                                                              '',
                                                                              '=E11+E10',
                                                                              '=F11+F10',
                                                                              '=G11+G10',
                                                                              '=H11+H10',
                                                                              '=I11+I10',
                                                                              '=J11+J10',
                                                                              '=K11+K10',
                                                                              '=L11+L10',
                                                                              '=M11+M10',
                                                                              '=N11+N10'),
                                                                             ('',
                                                                              'Paid in Capital',
                                                                              '',
                                                                              '=Inputs!E$6',
                                                                              '=Inputs!F$6',
                                                                              '=Inputs!G$6',
                                                                              '=Inputs!H$6',
                                                                              '=Inputs!I$6',
                                                                              '=Inputs!J$6',
                                                                              '=Inputs!K$6',
                                                                              '=Inputs!L$6',
                                                                              '=Inputs!M$6',
                                                                              '=Inputs!N$6'),
                                                                             ('',
                                                                              'Retained Earnings',
                                                                              '',
                                                                              '=Inputs!E$7',
                                                                              '=Inputs!F$7',
                                                                              '=Inputs!G$7',
                                                                              '=Inputs!H$7',
                                                                              '=Inputs!I$7',
                                                                              '=Inputs!J$7',
                                                                              '=Inputs!K$7',
                                                                              '=Inputs!L$7',
                                                                              '=Inputs!M$7',
                                                                              '=Inputs!N$7'),
                                                                             ('',
                                                                              'Other Equity',
                                                                              '',
                                                                              '=Inputs!E$8',
                                                                              '=Inputs!F$8',
                                                                              '=Inputs!G$8',
                                                                              '=Inputs!H$8',
                                                                              '=Inputs!I$8',
                                                                              '=Inputs!J$8',
                                                                              '=Inputs!K$8',
                                                                              '=Inputs!L$8',
                                                                              '=Inputs!M$8',
                                                                              '=Inputs!N$8'),
                                                                             ('',
                                                                              'Total Equity',
                                                                              '',
                                                                              '=E16+E15+E14',
                                                                              '=F16+F15+F14',
                                                                              '=G16+G15+G14',
                                                                              '=H16+H15+H14',
                                                                              '=I16+I15+I14',
                                                                              '=J16+J15+J14',
                                                                              '=K16+K15+K14',
                                                                              '=L16+L15+L14',
                                                                              '=M16+M15+M14',
                                                                              '=N16+N15+N14'),
                                                                             ('',
                                                                              'Total Equity',
                                                                              '',
                                                                              '=E17',
                                                                              '=F17',
                                                                              '=G17',
                                                                              '=H17',
                                                                              '=I17',
                                                                              '=J17',
                                                                              '=K17',
                                                                              '=L17',
                                                                              '=M17',
                                                                              '=N17'),
                                                                             ('',
                                                                              'Non Plug CL',
                                                                              '',
                                                                              '=E8',
                                                                              '=F8',
                                                                              '=G8',
                                                                              '=H8',
                                                                              '=I8',
                                                                              '=J8',
                                                                              '=K8',
                                                                              '=L8',
                                                                              '=M8',
                                                                              '=N8'),
                                                                             ('',
                                                                              'Liabilities NC',
                                                                              '',
                                                                              '=E11',
                                                                              '=F11',
                                                                              '=G11',
                                                                              '=H11',
                                                                              '=I11',
                                                                              '=J11',
                                                                              '=K11',
                                                                              '=L11',
                                                                              '=M11',
                                                                              '=N11'),
                                                                             ('',
                                                                              'Non Cash CA',
                                                                              '',
                                                                              '=E3',
                                                                              '=F3',
                                                                              '=G3',
                                                                              '=H3',
                                                                              '=I3',
                                                                              '=J3',
                                                                              '=K3',
                                                                              '=L3',
                                                                              '=M3',
                                                                              '=N3'),
                                                                             ('',
                                                                              'Non Current Assets',
                                                                              '',
                                                                              '=E5',
                                                                              '=F5',
                                                                              '=G5',
                                                                              '=H5',
                                                                              '=I5',
                                                                              '=J5',
                                                                              '=K5',
                                                                              '=L5',
                                                                              '=M5',
                                                                              '=N5'),
                                                                             ('',
                                                                              'Non cash assets',
                                                                              '',
                                                                              '=E23+E24',
                                                                              '=F23+F24',
                                                                              '=G23+G24',
                                                                              '=H23+H24',
                                                                              '=I23+I24',
                                                                              '=J23+J24',
                                                                              '=K23+K24',
                                                                              '=L23+L24',
                                                                              '=M23+M24',
                                                                              '=N23+N24'),
                                                                             ('',
                                                                              'Net Cash',
                                                                              '',
                                                                              '=E19+E20+E21-E25',
                                                                              '=F19+F20+F21-F25',
                                                                              '=G19+G20+G21-G25',
                                                                              '=H19+H20+H21-H25',
                                                                              '=I19+I20+I21-I25',
                                                                              '=J19+J20+J21-J25',
                                                                              '=K19+K20+K21-K25',
                                                                              '=L19+L20+L21-L25',
                                                                              '=M19+M20+M21-M25',
                                                                              '=N19+N20+N21-N25'),
                                                                             ('',
                                                                              'Excess Cash',
                                                                              '',
                                                                              '=max(0, E27)',
                                                                              '=max(0, F27)',
                                                                              '=max(0, G27)',
                                                                              '=max(0, H27)',
                                                                              '=max(0, I27)',
                                                                              '=max(0, J27)',
                                                                              '=max(0, K27)',
                                                                              '=max(0, L27)',
                                                                              '=max(0, M27)',
                                                                              '=max(0, N27)'),
                                                                             ('',
                                                                              'Cash',
                                                                              '',
                                                                              '=E29',
                                                                              '=F29',
                                                                              '=G29',
                                                                              '=H29',
                                                                              '=I29',
                                                                              '=J29',
                                                                              '=K29',
                                                                              '=L29',
                                                                              '=M29',
                                                                              '=N29'),
                                                                             ('',
                                                                              'Excess Deficit',
                                                                              '',
                                                                              '=min(E27, 0)',
                                                                              '=min(F27, 0)',
                                                                              '=min(G27, 0)',
                                                                              '=min(H27, 0)',
                                                                              '=min(I27, 0)',
                                                                              '=min(J27, 0)',
                                                                              '=min(K27, 0)',
                                                                              '=min(L27, 0)',
                                                                              '=min(M27, 0)',
                                                                              '=min(N27, 0)'),
                                                                             ('',
                                                                              'Plug',
                                                                              '',
                                                                              '=-1*E31',
                                                                              '=-1*F31',
                                                                              '=-1*G31',
                                                                              '=-1*H31',
                                                                              '=-1*I31',
                                                                              '=-1*J31',
                                                                              '=-1*K31',
                                                                              '=-1*L31',
                                                                              '=-1*M31',
                                                                              '=-1*N31'),
                                                                             ('',
                                                                              'Total Assets',
                                                                              '',
                                                                              '=E6',
                                                                              '=F6',
                                                                              '=G6',
                                                                              '=H6',
                                                                              '=I6',
                                                                              '=J6',
                                                                              '=K6',
                                                                              '=L6',
                                                                              '=M6',
                                                                              '=N6'),
                                                                             ('',
                                                                              'Total Liabilities',
                                                                              '',
                                                                              '=E12',
                                                                              '=F12',
                                                                              '=G12',
                                                                              '=H12',
                                                                              '=I12',
                                                                              '=J12',
                                                                              '=K12',
                                                                              '=L12',
                                                                              '=M12',
                                                                              '=N12'),
                                                                             ('',
                                                                              'Total Equity',
                                                                              '',
                                                                              '=E17',
                                                                              '=F17',
                                                                              '=G17',
                                                                              '=H17',
                                                                              '=I17',
                                                                              '=J17',
                                                                              '=K17',
                                                                              '=L17',
                                                                              '=M17',
                                                                              '=N17'),
                                                                             ('',
                                                                              'Dr Cr match',
                                                                              '',
                                                                              '=E34-(E35+E36)',
                                                                              '=F34-(F35+F36)',
                                                                              '=G34-(G35+G36)',
                                                                              '=H34-(H35+H36)',
                                                                              '=I34-(I35+I36)',
                                                                              '=J34-(J35+J36)',
                                                                              '=K34-(K35+K36)',
                                                                              '=L34-(L35+L36)',
                                                                              '=M34-(M35+M36)',
                                                                              '=N34-(N35+N36)')),
                                                                  (1, 2, 3, 4, 5, 7,
                                                                   8, 9, 10, 11, 13,
                                                                   14, 15, 16, 18,
                                                                   19, 20, 22, 23,
                                                                   24, 26, 28, 29,
                                                                   30, 31, 33, 34,
                                                                   35, 36), (
                                                                      1, 1, 1, 1, 1, 1,
                                                                      1, 1, 1, 1, 1, 1,
                                                                      1, 1, 1, 1, 1, 1,
                                                                      1, 1, 1, 1, 1, 1,
                                                                      1, 1, 1, 1, 1), ((
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                               'color': 'orange'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                                       (
                                                                                           {},
                                                                                           {},
                                                                                           {},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                           {
                                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'})))))

    expected_gateway_model2 = (f'{file_path_common}bb_output.xlsx', (('Inputs', (
        ('increase', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('decrease', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2),
                                                                      (1, 1), (({},
                                                                                {},
                                                                                {},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'}),
                                                                               ({},
                                                                                {},
                                                                                {},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'},
                                                                                {
                                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                    'color': 'blue'}))),
                                                                     ('Sheet1', ((
                                                                                     '',
                                                                                     'AA BB',
                                                                                     '',
                                                                                     '=D5',
                                                                                     '=E5',
                                                                                     '=F5',
                                                                                     '=G5',
                                                                                     '=H5',
                                                                                     '=I5',
                                                                                     '=J5',
                                                                                     '=K5',
                                                                                     '=L5',
                                                                                     '=M5'),
                                                                                 (
                                                                                     '',
                                                                                     'increase',
                                                                                     '',
                                                                                     '=Inputs!E$2',
                                                                                     '=Inputs!F$2',
                                                                                     '=Inputs!G$2',
                                                                                     '=Inputs!H$2',
                                                                                     '=Inputs!I$2',
                                                                                     '=Inputs!J$2',
                                                                                     '=Inputs!K$2',
                                                                                     '=Inputs!L$2',
                                                                                     '=Inputs!M$2',
                                                                                     '=Inputs!N$2'),
                                                                                 (
                                                                                     '',
                                                                                     'decrease',
                                                                                     '',
                                                                                     '=Inputs!E$3',
                                                                                     '=Inputs!F$3',
                                                                                     '=Inputs!G$3',
                                                                                     '=Inputs!H$3',
                                                                                     '=Inputs!I$3',
                                                                                     '=Inputs!J$3',
                                                                                     '=Inputs!K$3',
                                                                                     '=Inputs!L$3',
                                                                                     '=Inputs!M$3',
                                                                                     '=Inputs!N$3'),
                                                                                 (
                                                                                     '',
                                                                                     'AA',
                                                                                     '',
                                                                                     '=E2+E3-E4',
                                                                                     '=F2+F3-F4',
                                                                                     '=G2+G3-G4',
                                                                                     '=H2+H3-H4',
                                                                                     '=I2+I3-I4',
                                                                                     '=J2+J3-J4',
                                                                                     '=K2+K3-K4',
                                                                                     '=L2+L3-L4',
                                                                                     '=M2+M3-M4',
                                                                                     '=N2+N3-N4')),
                                                                      (1, 2, 3, 4),
                                                                      (1, 1, 1, 1),
                                                                      (({}, {}, {},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'}),
                                                                       ({}, {}, {},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'}),
                                                                       ({}, {}, {},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                            'color': 'orange'}),
                                                                       ({}, {}, {},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                        {
                                                                            'number_format': '#,##0.0_);[Red]▲#,##0.0'})))))
    expected_gateway_models_from_pickles = (
        (f'{file_path_common}operation_ave.xlsx', (('Inputs', (
            ('A', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('B', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
            ('D', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('E', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2, 3, 4),
                                                    (1, 1, 1, 1), (({}, {}, {}, {
            'number_format': '#,##0.0_);[Red]▲#,##0.0', 'color': 'blue'}, {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'blue'}, {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {}, {
                                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                       'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {}, {
                                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                       'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {}, {
                                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                       'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}))),
                                                   ('Sheet1', (('', 'A', '',
                                                                '=Inputs!E$2',
                                                                '=Inputs!F$2',
                                                                '=Inputs!G$2',
                                                                '=Inputs!H$2',
                                                                '=Inputs!I$2',
                                                                '=Inputs!J$2',
                                                                '=Inputs!K$2',
                                                                '=Inputs!L$2',
                                                                '=Inputs!M$2',
                                                                '=Inputs!N$2'), (
                                                                   '', 'B', '',
                                                                   '=Inputs!E$3',
                                                                   '=Inputs!F$3',
                                                                   '=Inputs!G$3',
                                                                   '=Inputs!H$3',
                                                                   '=Inputs!I$3',
                                                                   '=Inputs!J$3',
                                                                   '=Inputs!K$3',
                                                                   '=Inputs!L$3',
                                                                   '=Inputs!M$3',
                                                                   '=Inputs!N$3'), (
                                                                   '', 'D', '',
                                                                   '=Inputs!E$4',
                                                                   '=Inputs!F$4',
                                                                   '=Inputs!G$4',
                                                                   '=Inputs!H$4',
                                                                   '=Inputs!I$4',
                                                                   '=Inputs!J$4',
                                                                   '=Inputs!K$4',
                                                                   '=Inputs!L$4',
                                                                   '=Inputs!M$4',
                                                                   '=Inputs!N$4'), (
                                                                   '', 'E', '',
                                                                   '=Inputs!E$5',
                                                                   '=Inputs!F$5',
                                                                   '=Inputs!G$5',
                                                                   '=Inputs!H$5',
                                                                   '=Inputs!I$5',
                                                                   '=Inputs!J$5',
                                                                   '=Inputs!K$5',
                                                                   '=Inputs!L$5',
                                                                   '=Inputs!M$5',
                                                                   '=Inputs!N$5'), (
                                                                   '', 'C', '',
                                                                   '=average(E2, E3, E4, E5)',
                                                                   '=average(F2, F3, F4, F5)',
                                                                   '=average(G2, G3, G4, G5)',
                                                                   '=average(H2, H3, H4, H5)',
                                                                   '=average(I2, I3, I4, I5)',
                                                                   '=average(J2, J3, J4, J5)',
                                                                   '=average(K2, K3, K4, K5)',
                                                                   '=average(L2, L3, L4, L5)',
                                                                   '=average(M2, M3, M4, M5)',
                                                                   '=average(N2, N3, N4, N5)')),
                                                    (1, 2, 3, 4, 5),
                                                    (1, 1, 1, 1, 1), (({}, {}, {}, {
                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                       'color': 'orange'}, {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'}))))),
        (f'{file_path_common}operation_eq.xlsx', (('Inputs', (
            ('A', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('B', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2), (1, 1),
                                                   (({},
                                                     {},
                                                     {},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'}),
                                                    ({},
                                                     {},
                                                     {},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'}))),
                                                  ('Sheet1', (('', 'A', '',
                                                               '=Inputs!E$2',
                                                               '=Inputs!F$2',
                                                               '=Inputs!G$2',
                                                               '=Inputs!H$2',
                                                               '=Inputs!I$2',
                                                               '=Inputs!J$2',
                                                               '=Inputs!K$2',
                                                               '=Inputs!L$2',
                                                               '=Inputs!M$2',
                                                               '=Inputs!N$2'), (
                                                                  '', 'B', '',
                                                                  '=Inputs!E$3',
                                                                  '=Inputs!F$3',
                                                                  '=Inputs!G$3',
                                                                  '=Inputs!H$3',
                                                                  '=Inputs!I$3',
                                                                  '=Inputs!J$3',
                                                                  '=Inputs!K$3',
                                                                  '=Inputs!L$3',
                                                                  '=Inputs!M$3',
                                                                  '=Inputs!N$3'), (
                                                                  '', 'C', '', '=E2=E3',
                                                                  '=F2=F3', '=G2=G3',
                                                                  '=H2=H3', '=I2=I3',
                                                                  '=J2=J3', '=K2=K3',
                                                                  '=L2=L3', '=M2=M3',
                                                                  '=N2=N3')), (1, 2, 3),
                                                   (1, 1, 1), (({}, {}, {}, {
                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                      'color': 'orange'}, {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'}),
                                                               ({}, {}, {}, {
                                                                   'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                   'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                    'color': 'orange'}),
                                                               ({}, {}, {}, {
                                                                   'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                {
                                                                    'number_format': '#,##0.0_);[Red]▲#,##0.0'}))))),
        (f'{file_path_common}operation_ge.xlsx', (('Inputs', (
            ('A', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('B', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2), (1, 1),
                                                   (({},
                                                     {},
                                                     {},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'}),
                                                    ({},
                                                     {},
                                                     {},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'},
                                                     {
                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                         'color': 'blue'}))),
                                                  ('Sheet1', (('', 'A', '',
                                                               '=Inputs!E$2',
                                                               '=Inputs!F$2',
                                                               '=Inputs!G$2',
                                                               '=Inputs!H$2',
                                                               '=Inputs!I$2',
                                                               '=Inputs!J$2',
                                                               '=Inputs!K$2',
                                                               '=Inputs!L$2',
                                                               '=Inputs!M$2',
                                                               '=Inputs!N$2'), (
                                                                  '', 'B', '',
                                                                  '=Inputs!E$3',
                                                                  '=Inputs!F$3',
                                                                  '=Inputs!G$3',
                                                                  '=Inputs!H$3',
                                                                  '=Inputs!I$3',
                                                                  '=Inputs!J$3',
                                                                  '=Inputs!K$3',
                                                                  '=Inputs!L$3',
                                                                  '=Inputs!M$3',
                                                                  '=Inputs!N$3'), (
                                                                  '', 'C', '',
                                                                  '=E2>=E3', '=F2>=F3',
                                                                  '=G2>=G3', '=H2>=H3',
                                                                  '=I2>=I3', '=J2>=J3',
                                                                  '=K2>=K3', '=L2>=L3',
                                                                  '=M2>=M3',
                                                                  '=N2>=N3')),
                                                   (1, 2, 3), (1, 1, 1), (({}, {},
                                                                           {}, {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'}),
                                                                          ({}, {},
                                                                           {}, {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                               'color': 'orange'}),
                                                                          ({}, {},
                                                                           {}, {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                           {
                                                                               'number_format': '#,##0.0_);[Red]▲#,##0.0'}))))),
        (f'{file_path_common}operation_min.xlsx', (('Inputs', (
            ('A', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('B', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
            ('D', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('E', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2, 3, 4),
                                                    (1, 1, 1, 1), (({}, {}, {}, {
            'number_format': '#,##0.0_);[Red]▲#,##0.0', 'color': 'blue'}, {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'blue'}, {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {}, {
                                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                       'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {}, {
                                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                       'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}),
                                                                   ({}, {}, {}, {
                                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                       'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'},
                                                                    {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'blue'}))),
                                                   ('Sheet1', (('', 'A', '',
                                                                '=Inputs!E$2',
                                                                '=Inputs!F$2',
                                                                '=Inputs!G$2',
                                                                '=Inputs!H$2',
                                                                '=Inputs!I$2',
                                                                '=Inputs!J$2',
                                                                '=Inputs!K$2',
                                                                '=Inputs!L$2',
                                                                '=Inputs!M$2',
                                                                '=Inputs!N$2'), (
                                                                   '', 'B', '',
                                                                   '=Inputs!E$3',
                                                                   '=Inputs!F$3',
                                                                   '=Inputs!G$3',
                                                                   '=Inputs!H$3',
                                                                   '=Inputs!I$3',
                                                                   '=Inputs!J$3',
                                                                   '=Inputs!K$3',
                                                                   '=Inputs!L$3',
                                                                   '=Inputs!M$3',
                                                                   '=Inputs!N$3'), (
                                                                   '', 'D', '',
                                                                   '=Inputs!E$4',
                                                                   '=Inputs!F$4',
                                                                   '=Inputs!G$4',
                                                                   '=Inputs!H$4',
                                                                   '=Inputs!I$4',
                                                                   '=Inputs!J$4',
                                                                   '=Inputs!K$4',
                                                                   '=Inputs!L$4',
                                                                   '=Inputs!M$4',
                                                                   '=Inputs!N$4'), (
                                                                   '', 'E', '',
                                                                   '=Inputs!E$5',
                                                                   '=Inputs!F$5',
                                                                   '=Inputs!G$5',
                                                                   '=Inputs!H$5',
                                                                   '=Inputs!I$5',
                                                                   '=Inputs!J$5',
                                                                   '=Inputs!K$5',
                                                                   '=Inputs!L$5',
                                                                   '=Inputs!M$5',
                                                                   '=Inputs!N$5'), (
                                                                   '', 'C', '',
                                                                   '=min(E2, E3, E4, E5)',
                                                                   '=min(F2, F3, F4, F5)',
                                                                   '=min(G2, G3, G4, G5)',
                                                                   '=min(H2, H3, H4, H5)',
                                                                   '=min(I2, I3, I4, I5)',
                                                                   '=min(J2, J3, J4, J5)',
                                                                   '=min(K2, K3, K4, K5)',
                                                                   '=min(L2, L3, L4, L5)',
                                                                   '=min(M2, M3, M4, M5)',
                                                                   '=min(N2, N3, N4, N5)')),
                                                    (1, 2, 3, 4, 5),
                                                    (1, 1, 1, 1, 1), (({}, {}, {}, {
                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                       'color': 'orange'}, {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'orange'}),
                                                                      ({}, {}, {}, {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                       {
                                                                           'number_format': '#,##0.0_);[Red]▲#,##0.0'}))))),
        (f'{file_path_common}operation_abs.xlsx', (('Inputs', (
            ('Cash from BS', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
            ('Cash from CFWF', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2), (1, 1), (({}, {}, {}, {
            'number_format': '#,##0.0_);[Red]▲#,##0.0', 'color': 'blue'}, {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}, {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}),
                                                                                        ({}, {}, {},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'},
                                                                                         {
                                                                                             'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                             'color': 'blue'}))),
                                                   ('Sheet1', (('', 'Cash from BS',
                                                                '', '=Inputs!E$2',
                                                                '=Inputs!F$2',
                                                                '=Inputs!G$2',
                                                                '=Inputs!H$2',
                                                                '=Inputs!I$2',
                                                                '=Inputs!J$2',
                                                                '=Inputs!K$2',
                                                                '=Inputs!L$2',
                                                                '=Inputs!M$2',
                                                                '=Inputs!N$2'), (
                                                                   '', 'Cash from CFWF',
                                                                   '', '=Inputs!E$3',
                                                                   '=Inputs!F$3',
                                                                   '=Inputs!G$3',
                                                                   '=Inputs!H$3',
                                                                   '=Inputs!I$3',
                                                                   '=Inputs!J$3',
                                                                   '=Inputs!K$3',
                                                                   '=Inputs!L$3',
                                                                   '=Inputs!M$3',
                                                                   '=Inputs!N$3'), ('',
                                                                                    'Difference - absolute',
                                                                                    '',
                                                                                    '=abs(E2-E3)',
                                                                                    '=abs(F2-F3)',
                                                                                    '=abs(G2-G3)',
                                                                                    '=abs(H2-H3)',
                                                                                    '=abs(I2-I3)',
                                                                                    '=abs(J2-J3)',
                                                                                    '=abs(K2-K3)',
                                                                                    '=abs(L2-L3)',
                                                                                    '=abs(M2-M3)',
                                                                                    '=abs(N2-N3)')),
                                                    (1, 2, 3), (1, 1, 1), (({}, {},
                                                                            {}, {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'}),
                                                                           ({}, {},
                                                                            {}, {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                'color': 'orange'}),
                                                                           ({}, {},
                                                                            {}, {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                            {
                                                                                'number_format': '#,##0.0_);[Red]▲#,##0.0'}))))),
        (f'{file_path_common}test_socket_saved_as_a_module.xlsx', (('Inputs', (
            ('A', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), ('B', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
            ('D', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)), (1, 2, 4), (1, 1, 1), (({}, {}, {}, {
            'number_format': '#,##0.0_);[Red]▲#,##0.0', 'color': 'blue'}, {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}), ({}, {}, {}, {
            'number_format': '#,##0.0_);[Red]▲#,##0.0', 'color': 'blue'}, {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                           'color': 'blue'}, {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'},
                                                                                                          {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'},
                                                                                                          {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'},
                                                                                                          {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'},
                                                                                                          {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'},
                                                                                                          {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'},
                                                                                                          {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'},
                                                                                                          {
                                                                                                              'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                                              'color': 'blue'}),
                                                                                 ({}, {}, {}, {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'blue'}))),
                                                                   ('Sheet From', ((
                                                                                       '',
                                                                                       'A',
                                                                                       '',
                                                                                       '=Inputs!E$2',
                                                                                       '=Inputs!F$2',
                                                                                       '=Inputs!G$2',
                                                                                       '=Inputs!H$2',
                                                                                       '=Inputs!I$2',
                                                                                       '=Inputs!J$2',
                                                                                       '=Inputs!K$2',
                                                                                       '=Inputs!L$2',
                                                                                       '=Inputs!M$2',
                                                                                       '=Inputs!N$2'),
                                                                                   (
                                                                                       '',
                                                                                       'B',
                                                                                       '',
                                                                                       '=Inputs!E$3',
                                                                                       '=Inputs!F$3',
                                                                                       '=Inputs!G$3',
                                                                                       '=Inputs!H$3',
                                                                                       '=Inputs!I$3',
                                                                                       '=Inputs!J$3',
                                                                                       '=Inputs!K$3',
                                                                                       '=Inputs!L$3',
                                                                                       '=Inputs!M$3',
                                                                                       '=Inputs!N$3'),
                                                                                   (
                                                                                       '',
                                                                                       'C',
                                                                                       '',
                                                                                       '=E2+E3',
                                                                                       '=F2+F3',
                                                                                       '=G2+G3',
                                                                                       '=H2+H3',
                                                                                       '=I2+I3',
                                                                                       '=J2+J3',
                                                                                       '=K2+K3',
                                                                                       '=L2+L3',
                                                                                       '=M2+M3',
                                                                                       '=N2+N3')),
                                                                    (1, 2, 3),
                                                                    (1, 1, 1), (({},
                                                                                 {},
                                                                                 {},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'}),
                                                                                ({},
                                                                                 {},
                                                                                 {},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                     'color': 'orange'}),
                                                                                ({},
                                                                                 {},
                                                                                 {},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                 {
                                                                                     'number_format': '#,##0.0_);[Red]▲#,##0.0'}))),
                                                                   ('Sheet1', (('',
                                                                                'D',
                                                                                '',
                                                                                '=Inputs!E$5',
                                                                                '=Inputs!F$5',
                                                                                '=Inputs!G$5',
                                                                                '=Inputs!H$5',
                                                                                '=Inputs!I$5',
                                                                                '=Inputs!J$5',
                                                                                '=Inputs!K$5',
                                                                                '=Inputs!L$5',
                                                                                '=Inputs!M$5',
                                                                                '=Inputs!N$5'),
                                                                               ('',
                                                                                'X',
                                                                                '',
                                                                                "=E2*'Sheet From'!E4",
                                                                                "=F2*'Sheet From'!F4",
                                                                                "=G2*'Sheet From'!G4",
                                                                                "=H2*'Sheet From'!H4",
                                                                                "=I2*'Sheet From'!I4",
                                                                                "=J2*'Sheet From'!J4",
                                                                                "=K2*'Sheet From'!K4",
                                                                                "=L2*'Sheet From'!L4",
                                                                                "=M2*'Sheet From'!M4",
                                                                                "=N2*'Sheet From'!N4")),
                                                                    (1, 2), (1, 1),
                                                                    (({}, {}, {}, {
                                                                        'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                        'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'orange'}),
                                                                     ({}, {}, {}, {
                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                      {
                                                                          'number_format': '#,##0.0_);[Red]▲#,##0.0'}))))),
        (f'{file_path_common}test_auto_relay_across_sheets.xlsx', (('Inputs', (
            ('D', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),), (1,), (1,), (({}, {}, {},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'},
                                                                         {'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                          'color': 'blue'}),)),
                                                                   ('Sheet1', (('', 'D', '',
                                                                                '=Inputs!E$2',
                                                                                '=Inputs!F$2',
                                                                                '=Inputs!G$2',
                                                                                '=Inputs!H$2',
                                                                                '=Inputs!I$2',
                                                                                '=Inputs!J$2',
                                                                                '=Inputs!K$2',
                                                                                '=Inputs!L$2',
                                                                                '=Inputs!M$2',
                                                                                '=Inputs!N$2'),),
                                                                    (1,), (1,), (({}, {}, {}, {
                                                                       'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                       'color': 'orange'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'}, {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'},
                                                                                  {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'},
                                                                                  {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'},
                                                                                  {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'},
                                                                                  {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'},
                                                                                  {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'},
                                                                                  {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'},
                                                                                  {
                                                                                      'number_format': '#,##0.0_);[Red]▲#,##0.0',
                                                                                      'color': 'orange'}),)),
                                                                   ('Sheet2', (('',
                                                                                'X',
                                                                                '',
                                                                                '=Inputs!E2*1',
                                                                                '=Inputs!F2*1',
                                                                                '=Inputs!G2*1',
                                                                                '=Inputs!H2*1',
                                                                                '=Inputs!I2*1',
                                                                                '=Inputs!J2*1',
                                                                                '=Inputs!K2*1',
                                                                                '=Inputs!L2*1',
                                                                                '=Inputs!M2*1',
                                                                                '=Inputs!N2*1'),),
                                                                    (1,), (1,), ((
                                                                                     {},
                                                                                     {},
                                                                                     {},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'},
                                                                                     {
                                                                                         'number_format': '#,##0.0_);[Red]▲#,##0.0'}),)))),
    )

    request_models = (request_model1, request_model2,)
    request_models += request_models_from_pickles

    expected_gateway_models = (expected_gateway_model1, expected_gateway_model2,)
    expected_gateway_models += expected_gateway_models_from_pickles

    return expected_gateway_models, request_models
