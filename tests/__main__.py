"""
Run from top-level folder as module: 

$ python3 -m tests
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
        # name and check values
        pass
    
    def test_get_matching_value(self):
        # get and check value
        self.assertTrue(self.xl.get_matching_value('a', 'M', 'D') == 13000)
    
    def test_set_matching_value(self):
        self.xl.set_matching_value(
            'a', 'O', 'E', 'TEST', startrow=4)
        self.assertTrue(self.xl.get_matching_value(
            'a', 'O', 'E', startrow=2) == 'TEST')
        self.restore_original_data()
    
    def test_find_remove_row(self):
        # remove and then search for value
        pass
    
    def test_find_replace(self):
        self.xl.find_replace(
            'B', {'NES': 'TEST'}, ('AB', 'CD'), startrow=2)
        self.assertTrue(self.xl.get_matching_value(
            'a', 'L', 'B', startrow=2) == 'TEST')
        self.restore_original_data()
    
    def test_move_values(self):
        # move and check value
        pass
    
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
        # generate and check values
        pass
    
    def test_generate_list(self):
        # generate and check values
        pass
    
    




if __name__ == '__main__':
    unittest.main()