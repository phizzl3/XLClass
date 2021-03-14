"""
Run from top-level folder as module: 

$ python -m tests
"""

import unittest
from pathlib import Path

from xlclass import Xlsx

test_file = Path(__file__).resolve().parent / "_test.xlsx"


class TestXlsx(unittest.TestCase):

    def setUp(self):
        # Add an Xlsx object as an attribute
        self.xl = Xlsx(test_file)
        # Generate a list of original contents to reset when needed
        self.data = self.xl.generate_list()

    def restore_original_data(self):
        self.xl.ws.delete_rows(1, self.xl.ws.max_row)
        for rowvalues in self.data:
            self.xl.ws.append(rowvalues)

    def test_path(self):  #TODO: Move this to next function
        self.assertTrue(str(self.xl.path).endswith("_test.xlsx"))
        
    def test_init(self):
        # xls conversion
        # xlsx path
        pass
    
    def test_copy_sheet_data(self):
        # check cell data match
        pass
    
    def test_copy_csv_data(self):
        # check cell data match
        pass
    
    def test_sort_and_replace(self):
        # add random entry and sort and check cell values
        pass
    
    def test_name_headers(self):
        self.xl.name_headers({'A': 'TestA', 'D': 'TestD'}, bold=True)
        self.assertEqual(self.xl.ws['A1'].value, 'TestA')
        self.assertEqual(self.xl.ws['B1'].value, 'Strings')
        self.assertEqual(self.xl.ws['D1'].value, 'TestD')
    
    def test_get_matching_value(self):
        self.assertEqual(self.xl.get_matching_value('a', 'M', 'B'), 'SNES')
    
    def test_set_matching_value(self):
        self.xl.set_matching_value('a', 'O', 'E', 'TEST', startrow=4)
        self.assertEqual(self.xl.ws['E16'].value, 'TEST')
        self.restore_original_data()
    
    def test_find_remove_row(self):
        self.xl.find_remove_row('b', 'Switch', startrow=1)
        self.assertFalse(self.xl.ws['E19'].value == 29.5)
        self.assertEqual(self.xl.ws['E19'].value, 433.0498)
    
    def test_find_replace(self):
        self.xl.find_replace('B', {'NES': 'TEST'}, ('AB', 'CD'), startrow=2)
        self.assertEqual(self.xl.ws['B13'].value, 'TEST')
        self.restore_original_data()
    
    def test_move_values(self):
        self.xl.move_values('B', 'C', ('Gameboy', 'Red'))
        self.assertEqual(self.xl.ws['C15'].value, 'Gameboy')
        self.restore_original_data()
    
    def test_verify_length(self):
        # check lengths
        pass
    
    def test_highlight_rows(self):
        # not sure unless i save out and check
        pass
    
    def test_number_type_fix(self):
        # fix and read in and verify value
        pass
    
    def test_format_date(self):
        # fix and read in and verify value
        pass
    
    def test_format_currency(self):
        # fix and read in and verify value
        pass
    
    def test_set_cell_size(self):
        # set and verify on save out
        pass
    
    def test_save(self):
        # save out and verify path exists
        pass
    
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