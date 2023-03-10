try:
    import xlrd

    def XLSDictReader(f, sheet_index=0):
        data = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
        book = xlrd.open_workbook(file_contents=data)
        sheet = book.sheet_by_index(sheet_index)

        def item(i, j):
            return (sheet.cell_value(0, j), sheet.cell_value(i, j))

        return (dict(item(i, j) for j in range(sheet.ncols)) \
                for i in range(1, sheet.nrows))

except ImportError:
    XLSDictReader = None
