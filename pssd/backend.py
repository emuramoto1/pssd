import xlrd


def import_doc(data):
    """
    imports the excel file and returns a sheet format

    data = data file
    """
    loc = data
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
            # sheet.cell_value(i, 0),
            sheet.cell_value(i, 1),
            # sheet.cell_value(i, 2),
            # sheet.cell_value(i, 3),
            # sheet.cell_value(i, 4),
            # sheet.cell_value(i, 5),
        )

def class_list():
    '''creates a list of the first 3 letters in a course code'''
    return [
        "ACC",
        "LAW",
        "SME",
        "TAX",
        "AHS",
        "ARB",
        "ART",
        "CHN",
        "CVA",
        "ENG",
        "FRN",
        "HUM",
        "JPN",
        "LIT",
        "LVA",
        "MUS",
        "PHL",
        "PHO",
        "RHT",
        "SPN",
        "ECN",
        "EPS",
        "FIN",
        "AMS",
        "ANT",
        "HIS",
        "HSS",
        "MDS",
        "POL",
        "SOC",
        "ASM",
        "FME",
        "MOB",
        "COM",
        "MKT",
        "NST",
        "QTM",
        "SCN",
        "FYS",
        "IMH",
        "IND",
        "INH",
        "SEN",
        "WRT",
        "MIS",
        "OIM",
    ]

def is_course(value):
    """identifies whether or not the value is a course"""
    lst = value.split()
    return lst[0] in class_list()

def return_data(sheet):
    '''return data in column 2 as a list'''
    lst = []
    for i in range(sheet.nrows):
        lst.append(sheet.cell_value(i, 1))
    return lst

def course_code_seperator(row, sheet):
    ''' seperates the course code from the title and puts both into a dictionary'''
    d = {}
    lst = []
    string = sheet.cell_value(row, 1)
    lst = string.split()
    d['course_code'] = lst[0] + ' ' + lst[1]
    d['course_title'] = ' '.join(lst[3:])
    return d

def course_info(row, sheet):
    """outputs a dictionary with the credit count, course code, and course title for a row of information"""
    d = course_code_seperator(row, sheet)
    d["credits"] = sheet.cell_value(row, 4)
    return d 


        



def main():
    data = "test.xlsx"
    # classes = "data/class.txt"
    sheet = import_doc(data)
    # print_data(sheet)
    # print(is_course('HSS'))
    # print(return_data(sheet)[0])
    # print(cell_course(sheet))
    print(course_info(33, sheet))


if __name__ == "__main__":
    main()