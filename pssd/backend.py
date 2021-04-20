import xlrd

def import_doc(data): 
    """ 
    imports the excel file and returns a sheet format
    
    data = data file
    """
    loc = (data)
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    return sheet

def print_data(sheet):
    """
    prints out the data from the sheet

    sheet = the sheet that is to be printed
    """
    for i in range(sheet.nrows):
        print(
            sheet.cell_value(i, 0), 
            sheet.cell_value(i, 1), 
            sheet.cell_value(i, 2), 
            sheet.cell_value(i, 3),
            sheet.cell_value(i, 4), 
            sheet.cell_value(i, 5)
            )


def main():
    data = ('test.xlsx')
    doc = import_doc(data)
    print_data(doc)

if __name__ == "__main__":
    main()
