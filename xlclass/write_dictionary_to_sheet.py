"""

Writes nested dictionary data to an Xlsx object.

"""


def _write_dictionary_to_sheet(
    xlsx, data_dict: dict, header_row: int = None, start_row: int = None
) -> None:
    """Receives an Xlsx object and a nested dictionary:
    {"Row1": {"Key/Header": Value, "Key/Header": Value,},}
    Generates a list of keys to use as headers. Writes the headers and
    row data to the Xlsx object. If no header_row is passed, the default
    will be row 1 and the default start_row will be the header_row+1.

        Args:
            xlsx (Xlsx): Xlsx object to write data to.
            data_dict (dict): Nested dictionary with data to write to Xlsx.
            header_row (int, optional): Row number to write headers. Defaults to None.
            start_row (int, optional): Row number to start writing data. Defaults to None.
    """

    # Sets header and start rows if either is not passed, and
    # corrects start row if it is the same as the header row.
    if not header_row:
        header_row = 1
    if not start_row or (start_row == header_row):
        start_row = header_row + 1

    # Adds all keys to a single list to be used as headers.
    # (looks at all dictionaries for keys in case some keys
    # aren't present in every nested dictionary.)
    keys = []
    for row_data in data_dict.values():
        for key in row_data:
            if key not in keys:
                keys.append(key)

    # Writes keys/headers to xlsx sheet object in the specified row.
    for column_number, header in enumerate(keys, 1):
        xlsx.ws.cell(row=header_row, column=column_number, value=header)

    # Writes matching row data to each cell in the xlsx sheet object
    # by using the column headers to search for the value's key.
    for row_number, row_data in enumerate(data_dict.values(), start=start_row):
        for column_number, _key in enumerate(row_data, 1):
            xlsx.ws.cell(
                row=row_number,
                column=column_number,
                value=row_data.get(
                    xlsx.ws.cell(row=header_row, column=column_number).value
                ),
            )
