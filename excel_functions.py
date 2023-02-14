def copy_range(start_col, start_row, end_col, end_row, sheet):
    """
    Copy range of cells as a nested list
    :param start_col:
    :param start_row:
    :param end_col:
    :param end_row:
    :param sheet:
    :return:
    """
    range_selected = []
    # Loops through selected Rows
    for i in range(start_row, end_row + 1, 1):
        # Appends the row to a RowSelected list
        row_selected = []
        for j in range(start_col, end_col + 1, 1):
            row_selected.append(sheet.cell(row=i, column=j).value)
        # Adds the RowSelected List and nests inside the rangeSelected
        range_selected.append(row_selected)

    return range_selected


def paste_range(start_col, start_row, end_col, end_row, sheet_receiving, copied_data):
    """
    Paste data from copyRange into template sheet
    :param start_col:
    :param start_row:
    :param end_col:
    :param end_row:
    :param sheet_receiving:
    :param copied_data:
    :return:
    """
    count_row = 0
    for i in range(start_row, end_row + 1, 1):
        count_col = 0
        for j in range(start_col, end_col + 1, 1):
            sheet_receiving.cell(row=i, column=j).value = copied_data[count_row][count_col]
            count_col += 1
        count_row += 1
