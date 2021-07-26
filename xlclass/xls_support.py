import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows


def conver_xls()

# Convert xls to xlsx data using Pandas/Xlrd
            if str(filepath).endswith(".xls"):
                try:
                    # Read data from xls and create xlsx object
                    df = pd.read_excel(filepath, sheet_name=sheetname)
                    self.path = filepath
                    self.wb = openpyxl.Workbook()
                    self.ws = self.wb.active
                    self.ws.title = sheetname
                    # Copy row data from xls to new xlsx object
                    for row in dataframe_to_rows(df):
                        self.ws.append(row)
                    # Remove index row/colum created by Pandas
                    self.ws.delete_cols(1, 1)
                    self.ws.delete_rows(1, 1)

                except Exception as e:
                    print(f"\n Error: {e}\n Error converting from xls be sure "
                          "to include sheetname argument when passing file.")
                    input(" \"sheetname='Invoice'\", etc\n ENTER to close...")
                    exit(" Exiting...")