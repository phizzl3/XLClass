def search_matching_value(xl: Xlsx, 
                        header_srch_value: str, row_srch_value: str) -> str:
    # header to look for: "Total"
    # match row to look for: "Dollars"
    # Set variables to 0
    search_column, search_row = 0, 0
    for row in xl.ws.iter_rows():
        for cell_number, cell_data in enumerate(row, 1):
            if cell_data.value == header_srch_value:
                search_column += cell_number
            if cell_data.value == row_srch_value:
                search_row = True
            if search_row:
                if cell_number == search_column:
                    return cell_data.value