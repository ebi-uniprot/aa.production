import unittest
import mock
from cell_format_factory import CellFormatFactory


class TestCellFormatFactory(unittest.TestCase):

    def test_init_preserves_workbook_and_base_properties(self):
        param_workbook = mock.Mock()
        cell_format_factory = CellFormatFactory(param_workbook, {'a': 1, 'b': 'c'})

        # Assert
        self.assertDictEqual(cell_format_factory.base_properties, {'a': 1, 'b': 'c'})
        self.assertIs(cell_format_factory.workbook, param_workbook)

    def test_add_properties(self):
        cell_format_factory1 = CellFormatFactory(None, {'a': 1})
        cell_format_factory2 = cell_format_factory1.add_properties({'b': 'cd'})
        #
        self.assertIsNot(cell_format_factory2, cell_format_factory1, 'add_properties should return a new object')
        self.assertDictEqual(cell_format_factory1.base_properties, {'a': 1},
                             'source format def properties should not be modified')
        self.assertDictEqual(cell_format_factory2.base_properties, {'a': 1, 'b': 'cd'},
                             'source format def properties should not be modified')

    def test_make_format(self):
        param_workbook = mock.Mock()
        cell_format_factory = CellFormatFactory(param_workbook,
                                            {'font_name': 'Arial', 'font_size': 8})
        # Act
        cell_format_factory.make()

        # Assert
        param_workbook.add_format.assert_called_once_with({'font_name': 'Arial', 'font_size': 8})

    def test_can_chain_expected_params(self):
        param_workbook = mock.Mock()
        param_workbook.add_format = mock.Mock(side_effect=stub_workbook_add_format)

        cell_format_factory = CellFormatFactory(param_workbook, {})
        fmt_res = cell_format_factory \
            .add_properties({}) \
            .add_properties({'font_name': 'Arial'}) \
            .add_properties({'font_size': 10}).make()

        # Assert
        self.assertIsNotNone(fmt_res)
        self.assertTrue(fmt_res.is_dict_ok)

    def test_can_chain_different_params(self):
        param_workbook = mock.Mock()
        param_workbook.add_format = mock.Mock(side_effect=stub_workbook_add_format)

        cell_format_factory = CellFormatFactory(param_workbook, {})
        fmt_res = cell_format_factory \
            .add_properties({'font_name': 'Arial'}) \
            .add_properties({'font_size': 8}).make()

        # Assert
        self.assertIsNone(fmt_res)


def stub_workbook_add_format(props):
    if {'font_name': 'Arial', 'font_size': 10} == props:
        # TODO: replace with mock object
        return Bunch(is_dict_ok=True)
    return None


if __name__ == '__main__':
    unittest.main()


# taken from http://code.activestate.com/recipes/52308-the-simple-but-handy-collector-of-a-bunch-of-named
class Bunch:
    def __init__(self, **kwds):
        self.__dict__.update(kwds)
