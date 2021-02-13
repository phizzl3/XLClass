# Xlclass

 A Class I put together for working with Excel files using openpyxl for my common uses.  

Generates an Xlsx object with Openpyxl Workbook/Worksheet objects as attributes for use with the enclosed methods.  

## Requirements

* openpyxl

```bash
$ pip install -r requirements.txt
```

## Methods

### __init__(filepath=None, sheetname=None)

Initialize main attributes for Xlsx objects if Path points to 
an existing Excel file. Creates a blank Workbook/Worksheet object 
if no filepath is passed. If multiple sheets are present in passed 
Excel file, the name of the sheet you want to work with can be passed 
as a string to 'sheetname' or you can select needed sheet from a menu.

    Attrs:
        *.path (str/pathlib.Path, optional): Filepath information.
        *.wb (openpyxl.Workbook): Workbook object for Excel file.
        *.ws (openpyxl.Workbook.worksheet): Active sheet for Excel file. 

    Args:
        filepath (str/pathlib.Path, optional): str/Path object representing 
                *.xlsx input file.
        sheet (str, optional): Name representing which sheet you want to 
                work with. ex: 'Invoice'

### copy_sheet_data(other, columns)

Copy cell values from source Excel Worksheet to target Excel Worksheet using a passed dictionary of column letters.

    Args:
        other (Xlsx): Input Xlsx Excel file object to copy values from.
        columns (dict{str: str}): Dictionary of column letters representing the source and 
                target column letters to copy the values. ex: {'A': 'C', 'D': 'B'}

    Returns:
        self

### copy_csv_data(incsv)

Copy all values from csv file to target Excel Worksheet.

    Args:
        incsv (pathlib.Path): Path object representing a csv file. 

    Returns:
        self

### sort_and_replace(sortcol, startrow=1)

Sort and replace cell values based on values of a specific column. Use this BEFORE any cell formatting, etc as it DELETES the values and then replaces them after sorting.

    Args:
        sortcol (str): Column letter containing the values to use as "keys" to 
                sort row data by. ex: 'A'
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### name_headers(headers, hdrrow=1, bold=False)

Cycle through header row and fill cells with values.

    Args:
        headers (dict{str:str}): Str pairs of columns and header name values. 
                ex: {'A': 'Name'}
        hdrrow (int, optional): Row to be used for headers. Defaults to 1.
        bold (bool, optional): Option to bold values in headers. Defaults to False.

    Returns:
        self

### get_matching_value(srchcol, srchval, retcol, startrow=1)

Search column for a value and return the corresponding value from another column in the same row.

    Args:
        srchcol (str): Column letter to search for a value. ex: 'A'
        srchval (str): Value to search column for. ex: 'Total'
        retcol (str): Column letter containing the corresponding value 
                to be returned. ex: 'B'
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        str: Value from corresponding cell in the same row as search value. 
                Returns False if value search value is not found. 

### set_matching_value(srchcol, srchval, trgtcol, setval, startrow=1)

Search column for a value and set a corresponding value in another column in the same row.  

    Args:
        srchcol (str): Column letter to search for a value. ex: 'A'
        srchval (str): Value to search column for. ex: 'Total'
        trgtcol (str): Column letter to set with corresponding value. ex: 'B'
        setval (str): Value to insert into target cell.
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### find_remove_row(col, srch, startrow=1)

Remove row based on a specific value found in a column.

    Args:
        col (str): Column letter to search for the needed value.
        srch (str): Value to search for.
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### find_replace(col, fndrplc, skip=None, startrow=1)

Search column for a string value and replace it the value is not listed in 'skip'.

    Args:
        col (str): Column letter to search for the needed values.
        fndrplc (dict{str: str}): Dictionary of pairs to find and replace. 
                ex: {'find': 'replace'}.
        skip (list(str), optional): List of values to ignore when replacing.
                Defaults to None.
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### move_values(scol, tcol, vals, startrow=1)

Search source column for passed list of values and move them to target column.  

    Args:
        scol (str): Source column letter to search for values. ex: 'A'
        tcol (str): Target column letter to move located values to. ex: 'B'
        vals (list(str)): List of str values to move. ex: ('name', '20')
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### verify_length(col, length, fillcolor, skip=None, startrow=1)

 Cycle through values in a column to verify their length marking cells of an incorrect length with a background fill color.  

    Args:
        col (str): Column to search for values. ex: 'B'
        length (int): Total character length for correct values.
        fillcolor (str): Background fill color selection from COLORS dict.
        skip (list(str), optional): List of string values to skip when evaluating.
                Defaults to None.
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### highlight_rows(col, srch, fillcolor, startrow)

 Search row for specified str value and fill entire row with specified background fill color when found.  

    Args:
        col (str): Column to search for value. ex: 'B'
        srch (str): Str value to search cells for.
        fillcolor (str): Background fill color selection from COLORS dict.
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### number_type_fix(col, numtype, startrow=1)

Quick fix for cells that contain numbers formatted as text/str data. Cycle through cells replacing str formatted values with int/float values.

    Args:
        col (str): Column containing data to convert
        numtype (str): 'i' or 'f' indicating which type of number 
                values the column contains (int/float)
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### format_date(col, startrow=1)

 Format str date value to (MM/DD/YYYY).

    Args:
        col (str): Column containing date values.
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### format_currency(col, startrow=1)

Format str currency value to ($0,000.00).

    Args:
        col (str): Column containing currency values.
        startrow (int, optional): Starting row number where values begin. 
                    Defaults to 1.

    Returns:
        self

### set_cell_size(pairs)

Selects rows and columns and adjusts their sizes using a dictionary of pairs of rows or columns along with corresponding height or width to adjust the size of cells from each pair. If dict key is type: str, adjusts column width. If dict key is type: int, adjusts row height.  

    Args:
        pairs (dict): Dictionary of column/row keys with target size  
                as their values.

    Returns:
        self

### save(savepath)

Duplicates openpyxl's save function so it can be called on the object without needing the .wb attribute. Saves the Excel file to the specified filepath or Path location.  
ex: 'C:\files\MyFile.xlxs'

    Args:
        savepath (str or pathlib.Path): Output file location (including filename) 
                    for your output file.

### generate_dictionary(keycol, datacols, hdrrow=1, datastartrow=2)

Read the headers and cells from the spreadsheet and use them to generate
a dictionary of the data. Data listed in *keycol* on spreadsheet will need to be a series of unique values to be used as keys or the information assigned will be overwritten each time a duplicate key is found. 

    Args:
        keycol (str): Column letter where the data that will be used as the dictionary 
                keys is located.
        datacols (list): List of string column letters where needed data is located.
        hdrrow (int, optional) Row number containing the headers in the spreadsheet. 
                Defaults to 1.
        datastartrow (int, optional) Row number where the needed data starts. 
                     Defaults to 2.

    Returns:
        dict: Dictionary generated from the data in the spreadsheet.  

### generate_list(startrow=1, stoprow=None)

Generates a list of lists containing all cell values from startrow to stoprow (inclusive).
(Use _list.pop(0) on returned list to get a separate headers list if present/needed.)

    Args:
        startrow (int, optional): First row to pull data from to generate the list. Defaults to 1.
        stoprow (int, optional): Last row to pull data from to generate the list. 
                If no value is passed, it will pull data from all rows after startrow. Defaults to None.

    Returns:
        list: List of lists containing the values read from the cells.
