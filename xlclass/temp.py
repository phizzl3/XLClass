

header_row = 1

# create an empty dictionary to store the data
data = {}

# iterate over the rows of the sheet
for row, row_data in enumerate(test.ws.iter_rows(), 1):
    if row < header_row:
        continue
    # if the row is the first one, assign the values as the keys
    if row == header_row:
        keys = [cell.value for cell in row_data]
    # otherwise, assign the values as a nested dictionary with the keys
    else:
        values = [cell.value for cell in row_data]
        data[f"{row:0>4}"] = dict(zip(keys, values))
