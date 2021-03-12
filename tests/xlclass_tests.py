import unittest
from pathlib import Path

from xlclass import Xlsx

test_file = Path(__file__).resolve().parent / "_test.xlsx"


class TestXlsx(unittest.TestCase):

    def setUp(self):
        self.xl = Xlsx(test_file)

    def test_print(self):

        print('running')
        self.assertTrue(str(self.xl.path).endswith("_test.xlsx"))


if __name__ == '__main__':
    unittest.main()


def run_tests():
    print('done')
