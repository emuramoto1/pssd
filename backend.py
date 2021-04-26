import xlrd
import pprint
import pdfrw

#!/usr/bin/env python
# Class Variables


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


# Exploring the Data


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


def count_credits(sheet):
    """Counts the number of credits a student has"""
    d = excel_scanner(sheet)
    count = 0
    for i in d:
        count += d[i]["credits"]
    return count


# Helper Functions


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


def temp_dictionary_ce_rm(ce, rm, ce_value):
    """creates dict for course credits and the checkmark"""
    d = {}
    d[ce] = ce_value
    d[rm] = True
    return d


def count_credits_ce(d):
    """counts the number of credits in a dictionary with 'ce'"""
    count = 0
    for i in d:
        for key in d[i]:
            if key[:2] == "CE":
                count += d[i][key]
    return count


def clear_used_entries(original, used):
    """
    clears the used entries from the original dictionary

    original: original dictionary
    used: dictionary to be removed
    returns new dictionary that has all of the original keys that have not been used yet
    """
    for key in used:
        if key in original:
            original.pop(key)
    return original


def temp_dictionary_ce_rm_cc_ct(ce, rm, cc, ct, ce_value, cc_value, ct_value):
    """creates a dict for ce, rm, cc, and ct"""
    d = {}
    d[ce] = ce_value
    d[rm] = True
    d[cc] = cc_value
    d[ct] = ct_value
    return d


### DISCOVER


def dict_discover(sheet):
    """creates the dictionary for discover section"""
    d = excel_scanner(sheet)
    new_d = {}
    for key in d:
        if d[key]["course_code"] == "FME 1000":
            new_d[key] = temp_dictionary_ce_rm("CE_FME", "RM_FME", 3)
        elif d[key]["course_code"] == "FME 1001":
            new_d[key] = temp_dictionary_ce_rm("CE_FME2", "RM_FME2", 4)
        elif d[key]["course_code"] == "ACC 1000":
            new_d[key] = temp_dictionary_ce_rm("CE_ACC", "RM_ACC", 4)
        elif d[key]["course_code"] == "LAW 1000":
            new_d[key] = temp_dictionary_ce_rm("CE_LAW", "RM_LAW", 4)
        elif d[key]["course_code"] == "FYS 1000":
            new_d[key] = temp_dictionary_ce_rm("CE_FYS", "RM_FYS", 1)
        elif d[key]["course_code"] == "QTM 1000":
            new_d[key] = temp_dictionary_ce_rm("CE_QTM1", "RM_QTM1", 4)
        elif d[key]["course_code"] == "QTM 1010":
            new_d[key] = temp_dictionary_ce_rm("CE_QTM2", "RM_QTM2", 4)
        elif d[key]["course_code"] == "RHT 1000":
            new_d[key] = temp_dictionary_ce_rm("CE_RHT1", "RM_RHT1", 4)
        elif d[key]["course_code"] == "RHT 1001":
            new_d[key] = temp_dictionary_ce_rm("CE_RHT2", "RM_RHT2", 4)
        elif d[key]["course_code"] == "AHS 1000":
            new_d[key] = temp_dictionary_ce_rm("CE_AHS", "RM_AHS", 4)
        elif d[key]["course_code"][:-2] == "NST 10":
            new_d[key] = temp_dictionary_ce_rm("CE_NST1", "RM_NST1", 4)
    return new_d


def dict_after_discover(sheet):
    """remaining classes after removing the discover section"""
    original = excel_scanner(sheet)
    used = dict_discover(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_discovertotal(sheet):
    """create dict for credits in discover"""
    d = dict_discover(sheet)
    nest_d = {}
    count = 0
    for key in d:
        for i in d[key]:
            if i[:2] == "CE":
                count += d[key][i]
    nest_d["CE_DiscoverTotal"] = count
    d["discover_total"] = nest_d
    return d


### EXPLORE


def change_keys_explore(sheet):
    """create dict for the explore section"""
    d = dict_after_discover(sheet)
    new_d = {}
    count_cva = 0
    count_lva = 0
    count_hss = 0
    for key in d:
        if d[key]["course_code"] == "SME 2001":
            new_d[key] = temp_dictionary_ce_rm("CE_SME2001", "RM_SME2001", 3)
        elif d[key]["course_code"] == "SME 2002":
            new_d[key] = temp_dictionary_ce_rm("CE_SME2002", "RM_SME2002", 3)
        elif d[key]["course_code"] == "SME 2011":
            new_d[key] = temp_dictionary_ce_rm("CE_SME2011", "RM_SME2011", 3)
        elif d[key]["course_code"] == "SME 2012":
            new_d[key] = temp_dictionary_ce_rm("CE_SME2012", "RM_SME2012", 3)
        elif d[key]["course_code"] == "SME 2021":
            new_d[key] = temp_dictionary_ce_rm("CE_SME2021", "RM_SME2021", 3)
        elif d[key]["course_code"] == "SME 2031":
            new_d[key] = temp_dictionary_ce_rm("CE_SME2031", "RM_SME2031", 3)
        elif d[key]["course_code"] == "ECN 2000":
            new_d[key] = temp_dictionary_ce_rm("CE_ECN2000", "RM_ECN2000", 4)
        # CVA/HSS/LVA does not match because it does not match correctly in the PDF file
        elif d[key]["course_code"][:6] == "CVA 20" and count_cva == 0:
            new_d[key] = temp_dictionary_ce_rm("CE_HSS", "RM_HSS", 4)
            count_cva = 1
        elif d[key]["course_code"][:6] == "HSS 20" and count_hss == 0:
            new_d[key] = temp_dictionary_ce_rm("CE_LVA", "RM_LVA", 4)
            count_hss = 1
        elif d[key]["course_code"][:6] == "LVA 20" and count_lva == 0:
            new_d[key] = temp_dictionary_ce_rm("CE_CVA", "RM_CVA", 4)
            count_lva = 1
    # a = {"CE_ExploreTotal":count_credits_ce(new_d)}
    # new_d["explore_total"] = a
    return new_d


def dict_before_ILA4(sheet):
    """removes the used classes"""
    original = dict_after_discover(sheet)
    used = change_keys_explore(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_ILA4(sheet):
    """creates a dict for hss/css/lva section of pdf"""
    d = dict_before_ILA4(sheet)
    new_d = {}
    ILA4 = ["HSS", "CVA", "LVA"]
    count = 0
    for key in d:
        if d[key]["course_code"][:3] in ILA4 and count == 0:
            new_d[key] = {
                "CE_ILA4": 4,
                "RM_ILA4": True,
                "Explore_ILA4Description": d[key]["course_title"],
            }
            count = 1
    return new_d


def dict_after_ILA4(sheet):
    """removes lva/cva/hss class"""
    original = dict_before_ILA4(sheet)
    used = change_keys_ILA4(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_qtmnst(sheet):
    """creates dict for qtm3/nst2"""
    d = dict_after_ILA4(sheet)
    new_d = {}
    count = 0
    for key in d:
        if d[key]["course_code"][:6] == "QTM 20" and count == 0:
            new_d[key] = {"CE_NSTQTM": 4, "RM_NSTQTM": True, "Explore_CheckQTM": True}
            count = 1
        elif d[key]["course_code"][:6] == "NST 20" and count == 0:
            new_d[key] = {"CE_NSTQTM": 4, "RM_NSTQTM": True, "Explore_CheckNST": True}
            count = 1
    return new_d


def dict_after_qtmnst(sheet):
    """removes the qtm3/nst2 class from the main dict"""
    original = dict_after_ILA4(sheet)
    used = change_keys_qtmnst(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_exploretotal(sheet):
    """creates dictionary for explore total credits"""
    d = change_keys_qtmnst(sheet) | change_keys_ILA4(sheet) | change_keys_explore(sheet)
    nest_d = {}
    count = 0
    for key in d:
        for i in d[key]:
            if i[:2] == "CE":
                count += d[key][i]
    nest_d["CE_ExploreTotal"] = count
    d["explore_total"] = nest_d
    return d


def dict_explore(sheet):
    """combines and creates the dictionary for the explore section"""
    d = (
        change_keys_qtmnst(sheet)
        | change_keys_ILA4(sheet)
        | change_keys_explore(sheet)
        | change_keys_exploretotal(sheet)
    )
    return d


### FOCUS


def change_keys_asm(sheet):
    """creates dictionary for asm"""
    d = dict_after_qtmnst(sheet)
    new_d = {}
    for key in d:
        if d[key]["course_code"] == "ASM 3300":
            new_d[key] = temp_dictionary_ce_rm("CE_ASM", "RM_ASM", 4)
    return new_d


def dict_after_asm(sheet):
    """removes asm from main dict"""
    original = dict_after_qtmnst(sheet)
    used = change_keys_asm(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_46XX(sheet):
    """creates dictionary for 46XX class"""
    d = dict_after_asm(sheet)
    new_d = {}
    for key in d:
        if d[key]["course_code"][4:6] == "46":
            new_d[key] = temp_dictionary_ce_rm("CE_ASM", "RM_ASM", 4)
    return new_d


def dict_after_46XX(sheet):
    """removes 46XX from main dict"""
    original = dict_after_asm(sheet)
    used = change_keys_46XX(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_ALA(sheet):
    """creates dict for advanced liberal arts"""
    # ALAE = Advanced Liberal Art
    d = dict_after_46XX(sheet)
    new_d = {}
    count = 1
    for key in d:
        if d[key]["course_code"][5] == "6" and count < 4:
            count += 1
            new_d[key] = temp_dictionary_ce_rm_cc_ct(
                "CE_ALA" + str(count),
                "RM_ALA" + str(count),
                "CC_ALA" + str(count),
                "CT_ALA" + str(count),
                d[key]["credits"],
                d[key]["course_code"],
                d[key]["course_title"],
            )
    return new_d


def dict_after_ALA(sheet):
    """removes advanced liberal arts from main dict"""
    original = dict_after_46XX(sheet)
    used = change_keys_ALA(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_AE(sheet):
    """creates dict for advanced electives"""
    # AE stands for Advanced Elective
    d = dict_after_ALA(sheet)
    new_d = {}
    count = 0
    for key in d:
        if d[key]["course_code"][5] in ["5", "6"] and count < 4:
            count += 1
            new_d[key] = temp_dictionary_ce_rm_cc_ct(
                "CE_AE" + str(count),
                "RM_AE" + str(count),
                "CC_AE" + str(count),
                "CT_AE" + str(count),
                d[key]["credits"],
                d[key]["course_code"],
                d[key]["course_title"],
            )
    return new_d


def dict_after_AE(sheet):
    """removes advanced electives from main dict"""
    original = dict_after_ALA(sheet)
    used = change_keys_AE(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_FE(sheet):
    """creates dict for free electives"""
    # FE stands for Free Elective
    d = dict_after_AE(sheet)
    new_d = {}
    count = 0
    for key in d:
        if d[key]["course_code"][5] in ["1", "2", "5", "6"] and count < 3:
            count += 1
            new_d[key] = temp_dictionary_ce_rm_cc_ct(
                "CE_FE" + str(count),
                "RM_FE" + str(count),
                "CC_FE" + str(count),
                "CT_FE" + str(count),
                d[key]["credits"],
                d[key]["course_code"],
                d[key]["course_title"],
            )
    return new_d


def dict_after_FE(sheet):
    """removes free electives from main dict"""
    original = dict_after_AE(sheet)
    used = change_keys_FE(sheet)
    d = clear_used_entries(original, used)
    return d


def change_keys_focustotal(sheet):
    """creates dict for total credits in focus section"""
    d = (
        change_keys_asm(sheet)
        | change_keys_46XX(sheet)
        | change_keys_ALA(sheet)
        | change_keys_AE(sheet)
        | change_keys_FE(sheet)
    )
    nest_d = {}
    count = 0
    for key in d:
        for i in d[key]:
            if i[:2] == "CE":
                count += d[key][i]
    nest_d["CE_FocusTotal"] = count
    d["focus_total"] = nest_d
    return d


# Extra Courses


def change_keys_extra(sheet):
    """creates dict for extra courses"""
    d = dict_after_FE(sheet)
    new_d = {}
    count = 0
    for key in d:
        if count < 3:
            count += 1
            new_d[key] = temp_dictionary_ce_rm_cc_ct(
                "CE_Extra" + str(count),
                "RM_Extra" + str(count),
                "CC_Extra" + str(count),
                "CT_Extra" + str(count),
                d[key]["credits"],
                d[key]["course_code"],
                d[key]["course_title"],
            )
    return new_d


def dict_after_extra(sheet):
    """removes extra courses from main dict"""
    original = dict_after_FE(sheet)
    used = change_keys_extra(sheet)
    d = clear_used_entries(original, used)
    return d


def dict_for_all_total(sheet):
    """combines all the dicts needed to calculate total credits"""
    new_d = {}
    d = [
        change_keys_46XX(sheet),
        change_keys_AE(sheet),
        change_keys_asm(sheet),
        change_keys_explore(sheet),
        change_keys_discovertotal(sheet),
        change_keys_exploretotal(sheet),
        change_keys_extra(sheet),
        change_keys_FE(sheet),
        change_keys_focustotal(sheet),
        change_keys_ILA4(sheet),
        change_keys_qtmnst(sheet),
    ]
    for i in d:
        if i is not None:
            new_d = new_d | i
    return new_d


def change_keys_all_total(sheet):
    """sums the credits for all classes and creates a dict with that value"""
    d = dict_for_all_total(sheet)
    nest_d = {}
    count = 0
    for key in d:
        for i in d[key]:
            if i[:2] == "CE":
                if i[-5:] != "Total":
                    count += d[key][i]
    nest_d["CE_AllTotal"] = count
    d["all_total"] = nest_d
    return d


# Read and write PDF

ANNOT_KEY = "/Annots"
ANNOT_FIELD_KEY = "/T"
ANNOT_VAL_KEY = "/V"
ANNOT_RECT_KEY = "/Rect"
SUBTYPE_KEY = "/Subtype"
WIDGET_SUBTYPE_KEY = "/Widget"


def fill_pdf(input_pdf_path, output_pdf_path, data_dict):
    """reads and writes into the pdf template"""
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
                                        annotation.update(
                                            pdfrw.PdfDict(AS=pdfrw.PdfName("Yes"))
                                        )
                                else:
                                    annotation.update(
                                        pdfrw.PdfDict(V="{}".format(data_dict[i][key]))
                                    )
                                    annotation.update(pdfrw.PdfDict(AP=""))
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)


def write(sheet):
    """This function writes the classes to the pdf"""
    data = sheet
    sheet = import_doc(data)
    input_pdf_path0 = "template.pdf"
    output_pdf_path0 = "output_template.pdf"

    data_dict = (
        dict_discover(sheet)
        | dict_explore(sheet)
        | change_keys_asm(sheet)
        | change_keys_46XX(sheet)
        | change_keys_ALA(sheet)
    )

    fill_pdf(input_pdf_path0, output_pdf_path0, data_dict)
    input_pdf_path1 = "output_template.pdf"
    output_pdf_path1 = "output_template1.pdf"

    data_dict1 = change_keys_AE(sheet)
    fill_pdf(input_pdf_path1, output_pdf_path1, data_dict1)
    input_pdf_path2 = "output_template1.pdf"
    output_pdf_path2 = "output_template2.pdf"

    data_dict2 = change_keys_FE(sheet) | change_keys_focustotal(sheet)
    fill_pdf(input_pdf_path2, output_pdf_path2, data_dict2)
    input_pdf_path2 = "output_template2.pdf"
    output_pdf_path2 = "static/output_template3.pdf"

    data_dict2 = change_keys_extra(sheet) | change_keys_all_total(sheet)
    return fill_pdf(input_pdf_path2, output_pdf_path2, data_dict2)

def main(): 
    data = "test1.xlsx"
    # sheet = import_doc(data)
    # input_pdf_path0 = "template.pdf"
    # output_pdf_path0 = "output_template.pdf"

    # data_dict = (
    #     dict_discover(sheet)
    #     | dict_explore(sheet)
    #     | change_keys_asm(sheet)
    #     | change_keys_46XX(sheet)
    #     | change_keys_ALA(sheet)
    # )

    # fill_pdf(input_pdf_path0, output_pdf_path0, data_dict)
    # input_pdf_path1 = "output_template.pdf"
    # output_pdf_path1 = "output_template1.pdf"

    # data_dict1 = change_keys_AE(sheet)
    # fill_pdf(input_pdf_path1, output_pdf_path1, data_dict1)
    # input_pdf_path2 = "output_template1.pdf"
    # output_pdf_path2 = "output_template2.pdf"

    # data_dict2 = change_keys_FE(sheet) | change_keys_focustotal(sheet)
    # fill_pdf(input_pdf_path2, output_pdf_path2, data_dict2)
    # input_pdf_path2 = "output_template2.pdf"
    # output_pdf_path2 = "output_template3.pdf"

    # data_dict2 = change_keys_extra(sheet) | change_keys_all_total(sheet)
    # fill_pdf(input_pdf_path2, output_pdf_path2, data_dict2)
    write(data)

if __name__ == "__main__":
    main()