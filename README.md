# XLClass

 A Class I put together for working with Excel files using openpyxl for
 my common uses.  

Generates an Xlsx object with Openpyxl Workbook/Worksheet objects as
attributes for use with the enclosed methods.  

## Requirements for .xlsx support only

* openpyxl==3.0.6

**Delete** the *xls_support.py* file from the folder if you do not need
to convert *.xls* files.  

```bash
$ pip install -r requirements-xlsx_only.txt
```

## Requirements to also add .xls support

* openpyxl==3.0.6
* pandas==1.2.2
* xlrd==2.0.1


``bash
$ pip install -r requirements.txt
```
