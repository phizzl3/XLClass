"""
Module containing class for working with Excel *.xlsx files using Openpyxl.
Generates an Xlsx object with Openpyxl Workbook/Worksheet objects as 
attributes for use with the enclosed methods.

Requirements:
* openpyxl==3.0.6

xls support Requirements:
* pandas==1.2.2
* xlrd==2.0.1
"""

from .xlsx_class import Xlsx

__version__ = '07.28.2021'

__all__ = ['Xlsx']
