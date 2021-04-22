import xlrd
import pprint

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
        "MIS",
        "MFE",
        "LIB",
        "CGE",
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

def transfer_classes():
    return [
        'ECN 2000',
        'RHT 1000',
        'XXX 1200',
        'HSS 2050',
        'CVA 2050',
        'HSS 2001', 
        'CHN 1200',
        'LVA 2050',
        'FRN 1200',
        'ITL 1200',
        'JPN 1200',
        'SPN 1200',
        'QTM 1000' 
        ]

def micro():
    return ['SME 2031']

def is_course(value):
    """identifies whether or not the value is a course"""
    value = str(value)
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

def study_abroad_course_info(row, sheet):
    """outputs a dictionary with credit count, course code, and course title for a row of information for study abroad classes"""
    d = course_code_seperator(row, sheet)
    d["credits"] = sheet.cell_value(row, 2)
    return d 

def excel_scanner(sheet):
    """Scans an excel file and outputs a dictionary with a student's course information"""
    d = {}
    count = 0
    no_credit = ['W', 'NCP', 'F', 'NCF']
    for i in range(sheet.nrows):
        if sheet.cell_value(i, 2) not in no_credit:
            if sheet.cell_value(i,1) != '':
                if is_course(sheet.cell_value(i, 1)):
                    if sheet.cell_value(i, 2) in [1, 2, 3, 4]:
                        d["class_" + str(count)] = study_abroad_course_info(i, sheet)  
                        count +=1
                    else:
                        d["class_" + str(count)] = course_info(i, sheet)
                        count += 1
    return d

    # for i in d: 
    #     if d[i]['course_code'] in transfer_classes():
    #         d[i]['credits'] = 4.0
    #     elif d[i]['course_code'] in micro():
    #         d[i]['credits'] = 3.0
    # return d

def count_credits(sheet): 
    """Counts the number of credits a student has"""
    d = excel_scanner(sheet)
    count = 0 
    for i in d:
        count += d[i]['credits']
    return count

def main():
    data = "test0.xlsx"
    data1 = "test1.xlsx"
    data_sara = "test_sara.xlsx"
    # classes = "data/class.txt"
    sheet = import_doc(data)
    sheet1 = import_doc(data1)
    sheet_sara = import_doc(data_sara)
    # print_data(sheet)
    # print(is_course('HSS'))
    # print(return_data(sheet)[0])
    # print(cell_course(sheet))
    pprint.pprint(excel_scanner(sheet))
    # excel_scanner(sheet)
    print(count_credits(sheet))
    pprint.pprint(excel_scanner(sheet1))
    print(count_credits(sheet1))
    pprint.pprint(excel_scanner(sheet_sara))
    print(count_credits(sheet_sara))


if __name__ == "__main__":
    main()