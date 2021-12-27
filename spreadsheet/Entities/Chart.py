class Chart:
    def __init__(self, chart_id, chart_type: str, sub_type: str = None):
        self._id = chart_id
        self._chart_type = chart_type
        self._sub_type = sub_type
        self._sheet_name = None
        self._address_values = ()
        self._address_names = ()
        self._address_categories = ()
        self._fill_colors = ()
        self._border_colors = ()
        self._bool_data_labels = ()
        self._chart_title = None
        self._x_axis_label = None
        self._y_axis_label = None
        self._chart_position = None
        self._chart_sheet_name = 'Chart'

    @property
    def chart_type(self) -> str:
        return self._chart_type

    def add_data_series(self, address_values: str, address_name: str = None, address_categories: str = None,
                        fill_color: str = None, border_color: str = None, add_data_lable: bool = None):
        self._address_values += (address_values,)
        self._address_names += (address_name,)
        self._address_categories += (address_categories,)
        self._fill_colors += (fill_color,)
        self._border_colors += (border_color,)
        self._bool_data_labels += (add_data_lable,)

    def set_chart_title(self, value):
        self._chart_title = value

    def set_worksheet_name(self, sheet_name: str):
        self._sheet_name = sheet_name

    @property
    def sheet_name(self) -> str:
        return self._sheet_name

    def set_x_axis_label(self, value):
        self._x_axis_label = value

    def set_y_axis_label(self, value):
        self._y_axis_label = value

    def set_chart_position(self, address_position: str):
        self._chart_position = address_position

    def set_chart_sheet_name(self, name: str):
        self._chart_sheet_name = name

    @property
    def chart_model(self) -> dict:
        return {
            'id': self._id,
            'chart_type': self._chart_type,
            'sub_type': self._sub_type,
            'sheet_name': self._sheet_name,
            'address_values': self._address_values,
            'address_names': self._address_names,
            'address_categories': self._address_categories,
            'fill_colors': self._fill_colors,
            'border_colors': self._border_colors,
            'bool_data_labels': self._bool_data_labels,
            'chart_title': self._chart_title,
            'x_axis_label': self._x_axis_label,
            'y_axis_label': self._y_axis_label,
            'chart_position': self._chart_position,
            'chart_sheet_name': self._chart_sheet_name,
        }
