from typing import Any
from typing import Dict
from typing import Tuple

from .Chart import Chart


class Charts:

    def __init__(self):
        self._chart_id = 0
        self._data: Dict[Any, Chart] = {}

    # Chart Creations
    def add_new_chart(self, chart_type: str, sub_type: str = None):
        new_chart_id = self._chart_id
        self._data[self._chart_id] = Chart(new_chart_id, chart_type, sub_type)
        self._chart_id += 1
        return new_chart_id

    def set_worksheet_name(self, chart_id, sheet_name: str):
        chart = self.get_chart(chart_id)
        chart.set_worksheet_name(sheet_name)

    def add_new_line_chart(self):
        return self.add_new_chart('line')

    def add_new_stacked_bar_chart(self):
        return self.add_new_chart('bar', 'stacked')

    def get_chart(self, chart_id) -> Chart:
        return self._data[chart_id]

    # Adding Data Series
    def set_chart_position(self, chart_id, address: str):
        chart = self.get_chart(chart_id)
        chart.set_chart_position(address)

    def add_data_series(self, chart_id, address_values: str, address_name: str = None, address_category: str = None,
                        fill_color: str = None, border_color: str = None, add_data_label: bool = False):
        chart = self.get_chart(chart_id)
        chart.add_data_series(address_values, address_name, address_category, fill_color, border_color, add_data_label)

    def set_chart_title(self, chart_id, chart_title: str):
        chart = self.get_chart(chart_id)
        chart.set_chart_title(chart_title)

    def set_x_axis_label(self, chart_id, label: str):
        chart = self.get_chart(chart_id)
        chart.set_x_axis_label(label)

    def set_y_axis_label(self, chart_id, label: str):
        chart = self.get_chart(chart_id)
        chart.set_y_axis_label(label)

    def set_chart_sheet_name(self, chart_id, name: str):
        chart = self.get_chart(chart_id)
        chart.set_chart_sheet_name(name)

    @property
    def all_chart_ids(self) -> tuple:
        return tuple(self._data.keys())

    @property
    def all_charts(self) -> Tuple[Chart]:
        return tuple(self._data.values())

    @property
    def chart_model(self) -> Dict[str, Tuple[dict]]:
        chart_model = {}
        for chart in self.all_charts:
            sheet_name = chart.sheet_name
            if sheet_name in chart_model:
                chart_model[sheet_name] += (chart.chart_model,)
            else:
                chart_model[sheet_name] = (chart.chart_model,)
        return chart_model
