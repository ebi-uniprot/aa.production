import unittest
from xls_util import XlsUtil

# this is the placeholder for the XlsUtil test
# to make it run we need to set up packaging properly,
# so that we could split the whole project into logical units
class TestXlsUtilMethods(unittest.TestCase):
    def test_happy_path(self):
        xlsUtil = XlsUtil()
        self.assertEqual(xlsUtil.generate_deviations_sheet_name('report-2018_01.txt', 'report-2018_02.txt'), '2018_01-vs-2018_02')
        print(xlsUtil.generate_deviations_sheet_name('report-2018_01.txt', 'report-2018_02.txt'))
        # print(generateDifferencesSheetName('report-2018_01.txt', 'report-2018_02.txt'))
        # print(generateDifferencesSheetName('report-2018_01+++.txt', 'report-2018_02.txt'))
        # print(generateDifferencesSheetName('report-2018 _01.txt', 'report-2018_abc02.txt'))
        # print(generateDifferencesSheetName('report-2018 _01.txt++', 'report-2018_abc02.txt-'))
        # self.assertEqual(generateDifferencesSheetName('report-2018_01.txt', 'report-2018_02.txt'), 'FOO')
        pass

    def test_range_name_zero_change(self):
        xls_util = XlsUtil()
        self.assertEqual(xls_util.span_range('D15'), 'D15:D15')

    def test_range_col_increase(self):
        xls_util = XlsUtil()
        self.assertEqual(xls_util.span_range('M25', col_span=3), 'M25:O25')

    def test_range_row_increase(self):
        xls_util = XlsUtil()
        self.assertEqual(xls_util.span_range('B8', row_span=3), 'B8:B10')

    def test_column_exceeds_Z_row_stays(self):
        xls_util = XlsUtil()
        with self.assertRaises(Exception):
            xls_util.span_range('Z3', col_span=2)

    def test_column_exceeds_Z_row_increase(self):
        xls_util = XlsUtil()
        with self.assertRaises(Exception):
            xls_util.span_range('Z3', col_span=2, row_span=4)

if __name__ == '__main__':
    unittest.main()
