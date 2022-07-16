from typing import Dict
from typing import Tuple

import xlsxwriter


def save_as_excel(spreadsheet_model: tuple, **options) -> str:
    feedback = ''
    worksheets = {}
    workbook_name, instructions = spreadsheet_model
    sheet_to_rows_to_levels = options.get('levels_dict', {})
    workbook = xlsxwriter.Workbook(workbook_name)
    for instruction in instructions:
        cell_value = instruction[0]
        sheet_name = instruction[1]
        rows_to_level = sheet_to_rows_to_levels.get(sheet_name, {})
        row, col = instruction[2]
        format_ = instruction[3]
        cell_format = workbook.add_format(format_interpreter(format_))
        if sheet_name not in worksheets:
            worksheet = workbook.add_worksheet(sheet_name)
            worksheets[sheet_name] = worksheet
        else:
            worksheet = worksheets[sheet_name]
        worksheet.write(row, col, cell_value, cell_format)
        if row in rows_to_level:
            worksheet.set_row(row, None, None, {'level': rows_to_level[row], 'hidden': True})

    for sheet_name, column_to_width in options.get('column_widths', {}).items():  # Set column width
        if sheet_name not in worksheets:
            continue
        worksheet = worksheets[sheet_name]
        for col, width in column_to_width.items():
            worksheet.set_column(col, col, width)

    for sheet in workbook.worksheets():
        add_charts(workbook, sheet.name, **options)

    if workbook_name.split('.')[1] == "xlsm":
        vba_binary = options.get('vba_file', None)
        if vba_binary is not None:
            workbook.add_vba_project(vba_binary, is_stream=True)  # PyOxidizer compatible
        else:
            workbook.add_vba_project('./src/vbaProject.bin')  # This is legacy file/system way.

    try:
        workbook.close()
        return 'success' if feedback == '' else feedback
    except xlsxwriter.exceptions.FileCreateError:
        return 'for Windows, workbook must be closed first!'
    except xlsxwriter.exceptions.EmptyChartSeries as e:
        return f'something wrong with graph. {e}'


def add_charts(workbook: xlsxwriter.Workbook, sheet_name: str, **kwargs):
    if 'charts' not in kwargs:
        return
    if sheet_name not in kwargs['charts']:
        return
    chart_model: Dict[str, Tuple[dict]] = kwargs['charts']
    if sheet_name not in chart_model:
        return

    worksheet = workbook.get_worksheet_by_name(sheet_name)

    for each_chart_model in chart_model[sheet_name]:
        chart_id = each_chart_model['id']
        chart_type = each_chart_model['chart_type']
        sub_type = each_chart_model['sub_type']
        address_values = each_chart_model['address_values']
        address_names = each_chart_model['address_names']
        address_categories = each_chart_model['address_categories']
        fill_colors = each_chart_model['fill_colors']
        border_colors = each_chart_model['border_colors']
        bool_data_labels = each_chart_model['bool_data_labels']
        chart_title = each_chart_model['chart_title']
        x_axis_label = each_chart_model['x_axis_label']
        y_axis_label = each_chart_model['y_axis_label']
        chart_position = each_chart_model['chart_position']
        chart_sheet_name = each_chart_model['chart_sheet_name']

        options = {
            'type': chart_type,
            'sub_type': sub_type,
        }
        chart = workbook.add_chart(options)
        chart.set_title({
            'name': 'Title',
            'name_formula': chart_title,
            'name_font': {'size': 12}
        })
        chart.set_x_axis({'name': x_axis_label})
        chart.set_y_axis({'name': y_axis_label})
        chart.set_legend({'position': 'bottom'})
        for a_names, a_categories, values, color_f, color_b, bool_label in zip(address_names, address_categories,
                                                                               address_values, fill_colors,
                                                                               border_colors, bool_data_labels):
            options = {
                'name': a_names,
                'categories': a_categories,
                'values': values,
                'data_labels': {'value': bool_label},
                'overlap': 100,
            }
            if color_f is not None:
                options.update({'fill': {'color': color_f}})
            if color_b is not None:
                options.update({'border': {'color': color_b}})
            chart.add_series(options)

        separate_sheet = True
        if separate_sheet:
            chart_sheet = workbook.add_chartsheet(chart_sheet_name)
            chart_sheet.set_chart(chart)
        else:
            worksheet.insert_chart(chart_position, chart)


key_converter = {
    'color': 'font_color',
    'border_top': 'top',
    'border_bottom': 'bottom',
    'border_left': 'left',
    'border_right': 'right',

    'border_top_color': 'top_color',
    'border_bottom_color': 'bottom_color',
    'border_left_color': 'left_color',
    'border_right_color': 'right_color',

    'number_format': 'num_format',

    'highlight': 'bg_color',
}


def convert_value(value) -> int:
    """
    <Border style>

    Index	Name	    Weight	Style
    0	    None	        0
    1	    Continuous	    1	-----------
    2	    Continuous	    2	-----------
    3	    Dash	        1	- - - - - -
    4	    Dot	            1	. . . . . .
    5	    Continuous	    3	-----------
    6	    Double	        3	===========
    7	    Continuous	    0	-----------
    8	    Dash	        2	- - - - - -
    9	    Dash Dot	    1	- . - . - .
    10	    Dash Dot	    2	- . - . - .
    11	    Dash Dot Dot	1	- . . - . .
    12	    Dash Dot Dot	2	- . . - . .
    13	    SlantDash Dot	2	/ - . / - .
    """
    converter = {
        'normal': 1,
        'thick': 2,
        'thicker': 5,
        'thin': 7,
        'double': 6
    }
    return converter[value] if value in converter else value


def format_interpreter(cell_format: dict) -> dict:
    new_cell_format = {}
    for key, value in cell_format.items():
        new_key = key_converter.get(key, key)
        if new_key in ['top', 'bottom', 'left', 'right', ]:
            value = convert_value(value)
        new_cell_format[new_key] = value
    return new_cell_format
