class CellFormat(dict):
    def __init__(self, *args, **kwargs):
        dict.__init__(self, *args, **kwargs)

    def set_number_format(self, number_format: str):
        self['number_format'] = number_format

    def set_text_color(self, color: str):
        self['color'] = color

    def highlight(self, color: str):
        self['highlight'] = color

    def set_border(self, position: str, style: str = 'normal', color: str = 'black'):
        """
        <Styles>
        'normal'
        'thick'
        'thicker'
        'thin'
        'double
        """
        self[f'border_{position}'] = style
        self[f'border_{position}_color'] = color

    def set_top_border(self, style: str = 'normal', color: str = 'black'):
        self.set_border('top', style, color)

    def set_bottom_border(self, style: str = 'normal', color: str = 'black'):
        self.set_border('bottom', style, color)

    def set_left_border(self, style: str = 'normal', color: str = 'black'):
        self.set_border('left', style, color)

    def set_right_border(self, style: str = 'normal', color: str = 'black'):
        self.set_border('right', style, color)
