"""
Run from top-level folder as module: 

$ python -m tests
"""

import unittest
from pathlib import Path
import datetime
from xlclass import Xlsx

tests_path = Path(__file__).resolve().parent
test_file = tests_path / "_test.xlsx"
test_CSV = tests_path / "_testCSV.csv"


class TestXlsx(unittest.TestCase):

    def setUp(self):
        # Add an Xlsx object as an attribute
        self.xl = Xlsx(test_file)

    # def test_path(self):  # TODO: Move this to next function
    #     self.assertTrue(str(self.xl.path).endswith("_test.xlsx"))

    # def test_init(self):
    #     # xls conversion
    #     # xlsx path
    #     pass

    def test_copy_sheet_data(self):
        self.xl_temp = Xlsx()
        self.xl_temp.copy_sheet_data(self.xl, {'B': 'A', 'D': 'B', 'C': 'C'})
        self.assertEqual(self.xl_temp.ws['A9'].value, 'Magenta')
        self.assertEqual(self.xl_temp.ws['C20'].value, 1900)

    def test_copy_csv_data(self):
        self.xl.ws.delete_rows(1, self.xl.ws.max_row)            
        self.xl.copy_csv_data(test_CSV)
        self.assertEqual(self.xl.ws['A1'].value, 'State')
        self.assertEqual(self.xl.ws['B1'].value, 'Number')
        self.assertEqual(self.xl.ws['A11'].value, 'Georgia')
        self.assertEqual(self.xl.ws['B11'].value, '7')

    def test_sort_and_replace(self):
        self.xl.sort_and_replace('B', startrow=2)
        self.assertEqual(self.xl.ws['B1'].value, 'Strings')
        self.assertEqual(self.xl.ws['B2'].value, 'Blue')
        self.assertEqual(self.xl.ws['E6'].value, 25.5)
        self.assertEqual(self.xl.ws['C20'].value, 500)

    def test_name_headers(self):
        self.xl.name_headers({'A': 'TestA', 'D': 'TestD'}, bold=True)
        self.assertEqual(self.xl.ws['A1'].value, 'TestA')
        self.assertEqual(self.xl.ws['B1'].value, 'Strings')
        self.assertEqual(self.xl.ws['D1'].value, 'TestD')
        self.assertEqual(self.xl.ws['B2'].value, 'Red')

    def test_get_matching_value(self):
        self.assertEqual(self.xl.get_matching_value('a', 'M', 'B'), 'SNES')

    def test_set_matching_value(self):
        self.xl.set_matching_value('a', 'O', 'E', 'TEST', startrow=4)
        self.assertEqual(self.xl.ws['E16'].value, 'TEST')

    def test_find_remove_row(self):
        self.xl.find_remove_row('b', 'Switch', startrow=1)
        self.assertFalse(self.xl.ws['E19'].value == 29.5)
        self.assertEqual(self.xl.ws['E19'].value, 433.0498)

    def test_find_replace(self):
        self.xl.find_replace('B', {'NES': 'TEST'}, ('AB', 'CD'), startrow=2)
        self.assertEqual(self.xl.ws['B13'].value, 'TEST')

    def test_move_values(self):
        self.xl.move_values('B', 'C', ('Gameboy', 'Red'))
        self.assertEqual(self.xl.ws['C15'].value, 'Gameboy')

    def test_verify_length(self):
        self.xl.verify_length('B', 4, 'red', startrow=2)  #TODO: Fix int/stop

    def test_highlight_rows(self):
        self.xl.highlight_rows('B', 'Cyan', 'yellow', startrow=2)

    def test_number_type_fix(self):
        self.xl.number_type_fix('C', 'i', startrow=2)

    def test_format_date(self):
        self.assertIsInstance(self.xl.ws['D7'].value, datetime.datetime)
        self.xl.format_date('d', startrow=2)
        self.assertEqual(self.xl.ws['D7'].value, '04/25/2019')
        self.assertIsInstance(self.xl.ws['D7'].value, str)

    def test_format_currency(self):
        self.xl.format_currency('E', startrow=2) 
        self.assertIsInstance(self.xl.ws['E12'].value, float)

    def test_set_cell_size(self):
        self.xl.set_cell_size({'A': 48, 2: 20})

    def test_save(self):
        self.assertFalse(Path(f'{tests_path / "outfile.xlsx"}').exists())
        self.xl.save(f'{tests_path / "outfile.xlsx"}')
        self.assertTrue(Path(f'{tests_path / "outfile.xlsx"}').exists())
        Path(f'{tests_path / "outfile.xlsx"}').unlink()
        self.assertFalse(Path(f'{tests_path / "outfile.xlsx"}').exists())

    def test_generate_dictionary(self):
        _dict = self.xl.generate_dictionary(
            ('A', 'b', 'd', 'C', 'e'), keycol='b', hdrrow=1)
        self.assertEqual(_dict['Cyan']['Currency'], 48398.58)

    def test_generate_list(self):
        _list = self.xl.generate_list(startrow=4, stoprow=14)
        self.assertEqual(_list[0][1], 'Purple')
        self.assertEqual(_list[6][2], 900)
        with self.assertRaises(IndexError):
            _list[11][0]


if __name__ == '__main__':
    unittest.main()
