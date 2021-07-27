import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl


def conver_xls(obj, filepath, sheetname):

    # Convert xls to xlsx data using Pandas/Xlrd
    if not filepath:
        input("Error converting from xls.\nBe sure to include the sheetname "
              "when passing your filepath.")
        exit("Exiting...")
        
    # Read data from xls and create xlsx object
    df = pd.read_excel(filepath, sheet_name=sheetname)
    obj.path = filepath
    obj.wb = openpyxl.Workbook()
    obj.ws = obj.wb.active
    obj.ws.title = sheetname
    
    # Copy row data from xls to new xlsx object
    for row in dataframe_to_rows(df):
        obj.ws.append(row)
                
    # Remove index row/colum created by Pandas
    obj.ws.delete_cols(1, 1)
    obj.ws.delete_rows(1, 1)

