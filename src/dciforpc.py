import sys
import os
import logging
import Backlog_Handler as BH
import ExcelInterface as EI
import InputConfigParser as ICF


def getDciInfo(dciBook, requirement):
    logging.info("dciBook, requirement---------->", dciBook, requirement)
    requirement = requirement.strip()
    dciInfo = {
        "dciSignal": "",
        "arch": "",
        # "network": "",
        # "pc": "",
        "thm": "",
        # "framename": "",
        "dciReq": "",
        # "proj_param": "",
        # "dciThematic": ""
    }
    # maxrow = dciBook.sheets['MUX'].range('A' + str(dciBook.sheets['MUX'].cells.last_cell.row)).end('up').row
    # logging.info(maxrow)
    for sheet in dciBook.sheets:
        logging.info("Searching in Sheet...")
        logging.info(sheet)
        # logging.info("sheet name =*" + sheet.name + "*")
        try:
            if sheet.name.strip() == "MUX":
                maxrow = (dciBook.sheets['MUX'].range('A' + str(dciBook.sheets['MUX'].cells.last_cell.row)).end(
                    'up').row)
                sheet_value = sheet.used_range.value
                logging.info(maxrow)
                logging.info(
                    "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                searchResult = EI.searchDataInExcelCache(sheet_value, requirement)
                logging.info("searchResult-------searchResult--------->", searchResult)
                if searchResult["count"] > 0:
                    logging.info("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Success!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    for cellPosition in searchResult["cellPositions"]:
                        logging.info("getSignal")
                        x, y = cellPosition
                        logging.info(f"123yyyyyyy {y}")
                        dciInfo["dciSignal"] = (str(EI.getDataFromCell(sheet, (x, y + 2))))
                        dciInfo["arch"] = (str(EI.getDataFromCell(sheet, (x, y + 3))))
                        # dciInfo["network"] = (str(EI.getDataFromCell(sheet, (x, y + 9))))
                        # dciInfo["pc"] = (str(EI.getDataFromCell(sheet, (x, y + 10))))
                        dciInfo["thm"] = (str(EI.getDataFromCell(sheet, (x, y + 15))).encode('utf-8').strip())
                        # dciInfo["framename"] = (str(EI.getDataFromCell(sheet, (x, y + 9)))) + "/" + (
                        #     str(EI.getDataFromCell(sheet, (x, y + 8)))) + "/" + (str(EI.getDataFromCell(sheet, (x, y + 7))))
                        dciInfo["dciReq"] = (str(EI.getDataFromCell(sheet, (x, 1))))
                        # dciInfo["proj_param"] = (str(EI.getDataFromCell(sheet, (x, y + 3))))
                        # dciInfo["dciThematic"] = (str(EI.getDataFromCell(sheet, (x, 17))).encode('utf-8').strip())
                        logging.info(dciInfo)
                        break
        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            logging.info(
            f"Req is not present in the DCI files present in the input folder.{ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
    dciBook.close()
    return dciInfo


# NEAR1.2
def dciArch(dciInfo):
    dciInfo = dciInfo["arch"].replace("_", " ")
    # Split the modified string by spaces
    dciList = dciInfo.split()
    # Get the last element of the list
    lastElement = dciList[-1]
    # logging.info the modified list and the last element
    logging.info("Modified List:", dciList)
    logging.info("Last Element:", lastElement)
    return lastElement


def getThematiccode_for_DCI(thematic_string):
    them_code = ''
    # Split the string by newline characters to get individual clauses
    clauses = thematic_string.split("\n")

    # Remove empty clauses
    clauses = [clause for clause in clauses if clause.strip()]

    # Add parenthesis to each clause along with "OR"
    formatted_clauses = ["(" + clause + ")" if "OR" not in clause else clause for clause in clauses]

    # Join the formatted clauses with "\n" to get the final result
    modified_effective = "\n".join(formatted_clauses)

    logging.info(modified_effective)
    thematics_code = BH.grepThematicsCode(modified_effective)
    logging.info("thematics_code--------->", thematics_code)
    modified_thematics_code = thematics_code.replace(") (", ") OR (")

    logging.info(modified_thematics_code)

    them_line = BH.createCombination(modified_thematics_code)
    logging.info("them_line---------->", them_line)
    # Split the input string using the pipe character
    them_code = them_line.split('|')
    print("them_codethem_code-------->",them_code)
    them_code = list(set([code for line in them_line.split('\n') for code in line.split('|')]))
    print("them_code list:", them_code)
    return them_code, them_line


if __name__ == '__main__':
    ICF.loadConfig()
    # effective = 'BFC_01 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nBFC_02 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nDXD_00 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_03 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_05 AND IOP_07 AND LVM_03 AND LYQ_02'
    effective = 'BFC_01 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nBFC_02 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nDXD_00 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_03 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_05 AND IOP_07 AND LVM_03 AND LYQ_02'
    them_code = getThematiccode_for_DCI(effective)
    # inconnues_list = Thématiques_inconnue.split('|')
    # non_applicables_list = Thématiques_non_applicable.split('|')
    # them_line_list = them_code
    # code = ''
    # # Check if any code in them_line is present in Thématiques_inconnues or Thématiques_non_applicables
    # for code in them_line_list:
    #     if code in inconnues_list:
    #         inconnues_list_flag = 1
    #         logging.info(f"Code {code} is present in Thématiques_inconnues.")
    #         Thématiques_inconnues_codes.append(code)
    #         Thématiques_inconnues_flags.append(inconnues_list_flag)
    #
    #     else:
    #         inconnues_list_flag = 2
    #         logging.info(f"Code {code} is not present in Thématiques_inconnues.")
    #         Thématiques_inconnues_flags.append(inconnues_list_flag)
    #
    #     if code in non_applicables_list:
    #         non_applicables_list_flag = 1
    #         logging.info(f"Code {code} is present in Thématiques_non_applicables.")
    #         Thématiques_non_applicable_codes.append(code)
    #         Thématiques_non_applicables_flags.append(non_applicables_list_flag)
    #
    #     else:
    #         non_applicables_list_flag = 2
    #         logging.info(f"Code {code} is not present in Thématiques_non_applicables.")
    #         Thématiques_non_applicables_flags.append(non_applicables_list_flag)
    # Thématiques_inconnues_codes_result = ','.join(Thématiques_inconnues_codes)
    # logging.info("Thématiques_inconnues_codes_result--->", Thématiques_inconnues_codes_result)
    # Thématiques_non_applicable_codes_result = ','.join(Thématiques_non_applicable_codes)
    # logging.info("Thématiques_non_applicable_codes_result--->", Thématiques_non_applicable_codes_result)
    # result = findlines(them_code, Thématiques_non_applicable_codes_result)
    # logging.info("Thematics lines applicable to project or Arch from the requirement------>", result)
    #
    #
    #
    #


