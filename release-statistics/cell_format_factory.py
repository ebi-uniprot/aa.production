class CellFormatFactory:
    """helper to work with cell formatting properties, allowing chaining in the code

    Do not forget to call ``make`` function at the end of creation. Example usage:
        cell_format_def = CellFormatDef(workbook, {})

        fmt_res = cell_format_def \
            .add_properties({}) \
            .add_properties({'font_name': 'Arial'}) \
            .add_properties({'font_size': 10}).make()
    """

    def __init__(self, workbook, base_properties):
        self.base_properties = base_properties
        self.workbook = workbook

    def add_properties(self, extra_properties):
        uber_dict = self.base_properties.copy()
        uber_dict.update(extra_properties)
        return CellFormatFactory(self.workbook, uber_dict)

    def make(self):
        return self.workbook.add_format(self.base_properties)
