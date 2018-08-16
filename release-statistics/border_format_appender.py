# solution is taken from:
# https://stackoverflow.com/questions/21599809/python-xlsxwriter-set-border-around-multiple-cells/37907013
# see also https://gist.github.com/pankaj28843/d8c9c548a5a761be7ae6


class BorderFormatAppender:

    def __init__(self, workbook):
        self.workbook = workbook

    def append_borders_vertical(self, cell_format):
        return self.append_borders(cell_format, is_left=True, is_right=True)

    def append_border_left(self, cell_format):
        return self.append_borders(cell_format, is_left=True)

    def append_border_right(self, cell_format):
        return self.append_borders(cell_format, is_right=True)

    def append_borders(self, cell_format, is_left=False, is_right=False, is_top=False, is_bottom=False):
        extra_props = {}
        # ['top', 'bottom', 'left', 'right']
        if is_left:
            extra_props['left'] = 1
        if is_right:
            extra_props['right'] = 1
        if is_top:
            extra_props['top'] = 1
        if is_bottom:
            extra_props['bottom'] = 1

        return self.add_to_format(cell_format, extra_props)

    def add_to_format(self, existing_format, extra_properties):
        """Give a format you want to extend and a dict of the properties you want to
        extend it with, and you get them returned in a single format"""
        new_dict = {}
        for key, value in existing_format.__dict__.items():
            if (value != 0) and (value != {}) and (value != None):
                new_dict[key] = value
        del new_dict['escapes']

        uber_dict = new_dict.copy()
        uber_dict.update(extra_properties)
        uber_dict.pop('dxf_format_indices', None)
        return self.workbook.add_format(uber_dict)
