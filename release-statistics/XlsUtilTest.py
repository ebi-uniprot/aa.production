import unittest

# this is the placeholder for the XlsUtil test
# to make it run we need to set up packaging properly,
# so that we could split the whole project into logical units
class TestXlsUtilMethods(unittest.TestCase):
    def test_happy_path(self):
        # print(generateDifferencesSheetName('report-2018_01.txt', 'report-2018_02.txt'))
        # print(generateDifferencesSheetName('report-2018_01+++.txt', 'report-2018_02.txt'))
        # print(generateDifferencesSheetName('report-2018 _01.txt', 'report-2018_abc02.txt'))
        # print(generateDifferencesSheetName('report-2018 _01.txt++', 'report-2018_abc02.txt-'))
        # self.assertEqual(generateDifferencesSheetName('report-2018_01.txt', 'report-2018_02.txt'), 'FOO')
        pass

    # def test_split(self):
    #     s = 'hello world'
    #     self.assertEqual(s.split(), ['hello', 'world'])
    #     # check that s.split fails when the separator is not a string
    #     with self.assertRaises(TypeError):
    #         s.split(2)

if __name__ == '__main__':
    unittest.main()