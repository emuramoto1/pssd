import xlrd
import pprint
import pdfrw

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
    """creates a list of the first 3 letters in a course code"""
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
        "ECN 2000",
        "RHT 1000",
        "XXX 1200",
        "HSS 2050",
        "CVA 2050",
        "HSS 2001",
        "CHN 1200",
        "LVA 2050",
        "FRN 1200",
        "ITL 1200",
        "JPN 1200",
        "SPN 1200",
        "QTM 1000",
    ]


def micro():
    return ["SME 2031"]


def is_course(value):
    """identifies whether or not the value is a course"""
    value = str(value)
    lst = value.split()
    return lst[0] in class_list()


def return_data(sheet):
    """return data in column 2 as a list"""
    lst = []
    for i in range(sheet.nrows):
        lst.append(sheet.cell_value(i, 1))
    return lst


def course_code_seperator(row, sheet):
    """ seperates the course code from the title and puts both into a dictionary"""
    d = {}
    lst = []
    string = sheet.cell_value(row, 1)
    lst = string.split()
    d["course_code"] = lst[0] + " " + lst[1]
    d["course_title"] = " ".join(lst[3:])
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
    no_credit = ["W", "NCP", "F", "NCF"]
    for i in range(sheet.nrows):
        if sheet.cell_value(i, 2) not in no_credit:
            if sheet.cell_value(i, 1) != "":
                if is_course(sheet.cell_value(i, 1)):
                    if sheet.cell_value(i, 2) in [1, 2, 3, 4]:
                        d["class_" + str(count)] = study_abroad_course_info(i, sheet)
                        count += 1
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
        count += d[i]["credits"]
    return count


def temp_dictionary(ce, rm, ce_value):
    d = {}
    d[ce] = ce_value
    d[rm] = True
    return d

def count_credits_ce(d):
    count = 0 
    for i in d: 
        for key in d[i]:
            if key[:2] == "CE":
                count += d[i][key]
    return count


def change_keys_discover(sheet):
    d = excel_scanner(sheet)
    new_d = {}
    for key in d:
        # for k in d[key].items():
        if d[key]["course_code"] == "FME 1000":
            new_d[key] = temp_dictionary("CE_FME", "RM_FME", 3)
        elif d[key]["course_code"] == "FME 1001":
            new_d[key] = temp_dictionary("CE_FME2", "RM_FME2", 4)
        elif d[key]["course_code"] == "ACC 1000":
            new_d[key] = temp_dictionary("CE_ACC", "RM_ACC", 4)
        elif d[key]["course_code"] == "LAW 1000":
            new_d[key] = temp_dictionary("CE_LAW", "RM_LAW", 4)
        elif d[key]["course_code"] == "FYS 1000":
            new_d[key] = temp_dictionary("CE_FYS", "RM_FYS", 1)
        elif d[key]["course_code"] == "QTM 1000":
            new_d[key] = temp_dictionary("CE_QTM1", "RM_QTM1", 4)
        elif d[key]["course_code"] == "QTM 1010":
            new_d[key] = temp_dictionary("CE_QTM2", "RM_QTM2", 4)
        elif d[key]["course_code"] == "RHT 1000":
            new_d[key] = temp_dictionary("CE_RHT1", "RM_RHT1", 4)
        elif d[key]["course_code"] == "RHT 1001":
            new_d[key] = temp_dictionary("CE_RHT2", "RM_RHT2", 4)
        elif d[key]["course_code"] == "AHS 1000":
            new_d[key] = temp_dictionary("CE_AHS", "RM_AHS", 4)
        elif d[key]["course_code"][:-2] == "NST 10":
            new_d[key] = temp_dictionary("CE_NST1", "RM_NST1", 4)
    a = {"CE_DiscoverTotal":count_credits_ce(new_d)}
    new_d["discover_total"] = a
    return new_d

def change_keys_explore(sheet):
    d = excel_scanner(sheet)
    new_d = {}
    for key in d:
        for i in d[key].items():
            if d[key]["course_code"] == "SME 2001":
                new_d[key] = temp_dictionary("CE_SME2001", "RM_SME2001", 3)
            elif d[key]["course_code"] == "SME 2002":
                new_d[key] = temp_dictionary("CE_SME2002", "RM_SME2002", 3)
            elif d[key]["course_code"] == "SME 2011":
                new_d[key] = temp_dictionary("CE_SME2011", "RM_SME2011", 3)
            elif d[key]["course_code"] == "SME 2012":
                new_d[key] = temp_dictionary("CE_SME2012", "RM_SME2012", 3)
            elif d[key]["course_code"] == "SME 2021":
                new_d[key] = temp_dictionary("CE_SME2021", "RM_SME2021", 3)
            elif d[key]["course_code"] == "SME 2031":
                new_d[key] = temp_dictionary("CE_SME2031", "RM_SME2031", 3)
            elif d[key]["course_code"] == "ECN 2000":
                new_d[key] = temp_dictionary("CE_ECN2000", "RM_ECN2000", 4)
            elif d[key]["course_code"][:6] == "HSS 20":
                new_d[key] = temp_dictionary("CE_HSS", "RM_HSS", 4)
            elif d[key]["course_code"][:6] == "LVA 20":
                new_d[key] = temp_dictionary("CE_LVA", "RM_LVA", 4)
            elif d[key]["course_code"][:6] == "CVA 20":
                new_d[key] = temp_dictionary("CE_CVA", "RM_CVA", 4)
    #         elif d[key]["course_code"][:-2] == "NST 10":
    #             new_d[key] = temp_dictionary("CE_NST1", "RM_NST1", 4)
    a = {"CE_ExploreTotal":count_credits_ce(new_d)}
    new_d["explore_total"] = a
    return new_d

ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'

def fill_pdf(input_pdf_path, output_pdf_path, data_dict):
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    for page in template_pdf.pages:
        if page[ANNOT_KEY] != None:
            annotations = page[ANNOT_KEY]
            for annotation in annotations:
                if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                    if annotation[ANNOT_FIELD_KEY]:
                        key = annotation[ANNOT_FIELD_KEY][1:-1]
                        for i in data_dict:
                            if key in data_dict[i].keys():
                                if type(data_dict[i][key]) == bool:
                                    if data_dict[i][key] == True:
                                        annotation.update(pdfrw.PdfDict(
                                            AS=pdfrw.PdfName('Yes')))
                                else:
                                    annotation.update(
                                        pdfrw.PdfDict(V='{}'.format(data_dict[i][key]))
                                    )
                                    annotation.update(pdfrw.PdfDict(AP=''))
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

def main():
    data = "test0.xlsx"
    sheet = import_doc(data)
    data1 = "test1.xlsx"
    sheet1 = import_doc(data1)
    data_sara = "test_sara.xlsx"
    sheet_sara = import_doc(data_sara)
    input_pdf_path = "template.pdf"
    output_pdf_path = "output_template.pdf"
    # print_data(sheet)
    # print(is_course('HSS'))
    # print(return_data(sheet)[0])
    # print(cell_course(sheet))
    # pprint.pprint(excel_scanner(sheet))
    # excel_scanner(sheet)
    # print(count_credits(sheet))
    # pprint.pprint(excel_scanner(sheet1))
    # print(count_credits(sheet1))
    # pprint.pprint(excel_scanner(sheet_sara))
    # print(count_credits(sheet_sara))
    data_dict = change_keys_discover(sheet) | change_keys_explore(sheet)
    print(data_dict)
    fill_pdf(input_pdf_path, output_pdf_path, data_dict)

if __name__ == "__main__":
    main()