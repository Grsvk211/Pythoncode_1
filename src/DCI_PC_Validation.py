import sys
import os
import logging
import Backlog_Handler as BH
import ExcelInterface as EI
import InputConfigParser as ICF
import time
import re
import DocumentSearch as DS


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
            print(
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
    logging.info("them_codethem_code-------->",them_code)
    them_code = list(set([code for line in them_line.split('\n') for code in line.split('|')]))
    logging.info("them_code list:", them_code)
    return them_code, them_line


def getthematics(Config_sheet):
    path = ICF.getInputFolder() + "\\" + EI.findInputFiles()[19]
    logging.info("path ---->", path)
    Campagnec_Book = EI.openExcel(path)
    Campagnec_Book.activate()
    Campagne_sheet = Campagnec_Book.sheets[Config_sheet]
    sheet_value = Campagne_sheet.used_range.value
    Fun_name4 = EI.searchDataInColCache(sheet_value, 1, 'Function(s) or type of test')
    if not Fun_name4:
        Fun_name4 = EI.searchDataInColCache(sheet_value, 2, 'Function(s) or type of test')
    row, col = Fun_name4['cellPositions'][0]
    time.sleep(2)
    logging.info(row, col)
    Fun_name4_1 = EI.searchDataInColCache(sheet_value, 5, 'Thématiques inconnues')
    if not Fun_name4_1:
        Fun_name4_1 = EI.searchDataInColCache(sheet_value, 6, 'Thématiques inconnues')
    row, col1 = Fun_name4_1['cellPositions'][0]
    time.sleep(2)
    logging.info("row, col0000------------------>", row, col1)

    Thématiques_inconnues = EI.getDataFromCell(Campagne_sheet, (row, col1 + 1))
    logging.info("Thématiques_inconnues---->", Thématiques_inconnues)

    Fun_name4_2 = EI.searchDataInColCache(sheet_value, 5, 'Thématiques non applicables')
    if not Fun_name4_2:
        Fun_name4_2 = EI.searchDataInColCache(sheet_value, 6, 'Thématiques non applicables')
    row, col2 = Fun_name4_2['cellPositions'][0]
    time.sleep(2)
    logging.info("row, col0000------------------>", row, col2)
    Thématiques_non_applicables = EI.getDataFromCell(Campagne_sheet, (row, col2 + 1))
    logging.info("Thématiques_non_applicables---->", Thématiques_non_applicables)

    NT_values = EI.getDataFromCell(Campagne_sheet, (row, col2 + 3))
    logging.info("NT_values------>",NT_values)
    NA_values = EI.getDataFromCell(Campagne_sheet, (row+1, col2 + 3))
    logging.info("NA_values------>", NA_values)
    NT_value = int(re.search(r':\s*(\d+)', NT_values).group(1)) if re.search(r':\s*(\d+)', NT_values) else None
    NA_value = int(re.search(r':\s*(\d+)', NA_values).group(1)) if re.search(r':\s*(\d+)', NA_values) else None
    logging.info("NT_value------->", NT_value, NA_value)
    return Thématiques_inconnues, Thématiques_non_applicables, NT_value, NA_value


def getthematics_in_list(Thématiques_non_applicables):
    # Thématiques_non_applicables = "| ABY_02 | AEM_01 | AFC_04 | AGG_01 | AMD_02 | AMG_01 | APT_04 | AWB_01 | AXG_01 | AXW_02 | AZX_02 | BPR_01 | CAE_04 | CAE_05 | CMP_02 | CRA_01 | CWF_03 | D12_01 | D41_01 | D41_02 | D59_02 | D7A_02 | D7K_01 | D7U_02 | D7U_04 | D7W_00 | D7W_03 | DAO_01 | DE7_03 | DI6_04 | DO8_00 | DRG_40 | DUE_07 | DUE_08 | DVQ_64 | DVQ_67 | DXD_00 | DXD_03 | DXD_05 | ENK_00 | EPD_00 | ICS_07 | IPD_00 | IWA_01 | IWH_00 | IWW_00 | LVM_02 | LXA_01"
    Thématiques_non_applicables = Thématiques_non_applicables.strip('|')
    # Split the string into a list and remove leading/trailing whitespaces from each element
    thematiques_list = [theme.strip() for theme in Thématiques_non_applicables.split('|')]
    # Remove any empty strings from the list
    thematiques_lists = list(filter(None, thematiques_list))
    # logging.info the list
    logging.info(thematiques_lists)
    return thematiques_lists


def searchDataInExcelCache(value, keyword):
    # value = sheet.used_range.value
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    # x is the index of column
    # i is the value of column
    # y is the index of row
    # j is the value of cell
    if keyword=="":
        return searchResult

    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if j is not None:
                if keyword in str(j):
                    searchResult["count"] = searchResult["count"] + 1
                    searchResult["cellPositions"].append((x + 1, y + 1))
                    searchResult["cellValue"].append(j)

    return searchResult


def getReqVer(req):
    if req.find('(') != -1:
        new_reqName = req.split("(")[0].split()[0] if len(req.split("(")) > 0 else ""
        new_reqVer = req.split("(")[1].split(")")[0] if len(req.split("(")) > 1 else ""
    else:
        new_reqName = req.split()[0] if len(req.split()) > 0 else ""
        new_reqVer = req.split()[1] if len(req.split()) > 1 else ""
    return new_reqName.strip(), new_reqVer.strip()


def search_requirement(requirement_id, file_path):
    searchResult = ''
    table = ''
    reqName, reqVer = getReqVer(requirement_id)
    variations = [
        reqName + "(" + reqVer + ")",
        reqName + " (" + reqVer + ")",
        reqName + "  (" + reqVer + ")",
        reqName + " " + reqVer,
        reqName + "  " + reqVer,
    ]

    searchResult = None  # Initialize con to None

    for variation in variations:
        logging.info("variation--------->", variation)
        searchResult, table, file_path = DS.find_requirement_content(file_path, variation)
        # logging.info(f"con $----> {variation}: {con}")
        logging.info(f"con $----> {variation}: {searchResult}")
        if searchResult and searchResult[0]:  # Check if the first element of con is not empty
            break  # Break the loop if content is found
    return searchResult, table, file_path


def not_thematic(project, Thématiques_non_applicables, Thématiques_inconnues, desired_architecture):
    logging.info("project---------->",project)
    Thématiques_inconnues_after_check = ''
    Thématiques_non_applicable = Thématiques_non_applicables
    logging.info("ko")
    result_tuples = []
    if Thématiques_non_applicable is not None:
        thematiques_lists = getthematics_in_list(Thématiques_non_applicable)
        silhouette_files = [file for file in os.listdir(ICF.getInputFolder()) if file.lower().endswith((".xlsm", ".xlsx")) and "Silhouette" in file]
        matching_files = []
        for file in silhouette_files:
            # Check if any substring in the project matches the file name
            if any(substring in project for substring in file.split('_')):
                matching_files.append(file)
        print("matching_files--------------->", matching_files)
        path = ICF.getInputFolder() + "\\" + matching_files[0]
        logging.info("path ---->", path)
        Silhouette_Book = EI.openExcel(path)
        Silhouette_Book.activate()
        Silhouettes_sheet = Silhouette_Book.sheets['Silhouettes']
        sheet_value = Silhouettes_sheet.used_range.value
        # for project in project:
        print("project---------->", project)
        try:
            for result_string in thematiques_lists:
                logging.info("thematiques_lists result_string------>", result_string)
                cleaned_result_string = result_string.strip()
                Fun_name9 = EI.searchDataInColCache(sheet_value, 1, cleaned_result_string)
                if not Fun_name9:
                    Fun_name9 = EI.searchDataInColCache(sheet_value, 2, cleaned_result_string)
                # logging.info("Fun_name9----->",Fun_name9)
                # Check if 'cellPositions' is not empty before accessing its elements
                if Fun_name9['cellPositions']:
                    row, col = Fun_name9['cellPositions'][0]
                    time.sleep(2)
                    # logging.info(row, col)
                    Project_cell = searchDataInExcelCache(sheet_value, project)
                    # logging.info("project cells----------->", Project_cell)
                    row18, col18 = Project_cell['cellPositions'][0]
                    project1 = EI.getDataFromCell(Silhouettes_sheet, (row18, col18))
                    if project == project1:
                        # logging.info("hiiiiiiiii")
                        Silhouettes_sheet_Valeur = EI.getDataFromCell(Silhouettes_sheet, (row, col18))
                        logging.info("ValeurValeurValeurValeur--->", Silhouettes_sheet_Valeur)
                        if Silhouettes_sheet_Valeur == '--' or Silhouettes_sheet_Valeur == 'X' or Silhouettes_sheet_Valeur == 'opt':
                            logging.info("v--------------->", result_string)
                            result_tuples.append((cleaned_result_string, Silhouettes_sheet_Valeur))
            logging.info("result_tuples---------------->", result_tuples)

        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(f"{ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
        #
        Thématiques_sheet = Silhouette_Book.sheets['Thématiques']
        sheet_value1 = Thématiques_sheet.used_range.value
        # for project in project:
        logging.info("project---------->", project)
        result__strings = []
        result_Xstrings = []
        try:
            NEA_R1_cell_Xdatas = []
            NEA_R1_cell__datas = []
            if Thématiques_inconnues is not None:
                Thématiques_inconnuess = getthematics_in_list(Thématiques_inconnues)
            else:
                project_text_bold = "\033[1m" + "In Campagne Config Thématiques inconnues cell is empty." + "\033[0m"
                print(project_text_bold)
                Thématiques_inconnuess = []
            for result_string in Thématiques_inconnuess:
                logging.info("Thématiques_inconnuess result_string------>", result_string)
                cleaned_result_string = result_string.strip()
                Fun_name9 = EI.searchDataInColCache(sheet_value1, 1, cleaned_result_string)
                logging.info("Fun_name9----->", Fun_name9)
                # Check if 'cellPositions' is not empty before accessing its elements
                if Fun_name9['cellPositions']:
                    row23, col23 = Fun_name9['cellPositions'][0]
                    time.sleep(2)
                    logging.info("row23, col23--------->", row23, col23)
                    logging.info("desired_architecture----------->", desired_architecture)
                    # Extracting the version part
                    version = desired_architecture.split()[1]

                    if version.startswith("R1."):
                        desired_architecture = "NEA R1.x"
                        logging.info(desired_architecture)
                    else:
                        desired_architecture = "NEA R1"
                        logging.info(desired_architecture)
                    logging.info("sheet_value1, desired_architecture-------->", desired_architecture)
                    NEA_R1_cell = searchDataInExcelCache(sheet_value1, desired_architecture)
                    logging.info("NEA_R1_cell--------->", NEA_R1_cell)
                    if NEA_R1_cell['cellPositions']:
                        row15, col15 = NEA_R1_cell['cellPositions'][0]
                        logging.info("row, col15----------->", row15, col15)
                        # NEA_R1_cell_data = EI.getDataFromCell(sheet_value1, (row23, col15))
                        cell_data = EI.getDataFromCell(Thématiques_sheet, (row23, col15))
                        logging.info(f'{result_string} thematics Arch {cell_data}')
                        if cell_data == 'X':
                            NEA_R1_cell_Xdatas.append(cell_data)
                            result_Xstrings.append(result_string)
                        if cell_data == '--':
                            NEA_R1_cell__datas.append(cell_data)
                            result__strings.append(result_string)
                else:
                    logging.info(f'Thematics {result_string} not present in the Silhouette file Thematcs sheet. Check in the EC file.')

            logging.info("NEA_R1_cell__datas---------->",NEA_R1_cell__datas, result__strings)
            Thématiques_inconnues_after_check = '|' + '|'.join(result__strings)
            logging.info("Thématiques_inconnues_after_check--------->", Thématiques_inconnues_after_check)

            if result__strings:
                Thématiques_inconnues_after_check = '|' + '|'.join(result__strings)
                logging.info("Thématiques_inconnues_after_check--------->", Thématiques_inconnues_after_check)

            if result_Xstrings:
                logging.info("NEA_R1_cell_Xdatas---------->", NEA_R1_cell_Xdatas, result_Xstrings)
                logging.info(f'check the thematcis {result_Xstrings} are applicable for the {desired_architecture}. when searching for Thématiques_inconnues.')

        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(f"{ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
        Silhouette_Book.close()
    else:
        project_text_bold = "\033[1m" + "In Campagne Config Thématiques non applicables cell is empty." + "\033[0m"
        print(project_text_bold)
    return result_tuples, Thématiques_inconnues_after_check


if __name__ == '__main__':
    ICF.loadConfig()
    # effective = 'BFC_01 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nBFC_02 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nDXD_00 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_03 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_05 AND IOP_07 AND LVM_03 AND LYQ_02'
    effective = 'BFC_01 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nBFC_02 AND IOP_07 AND LVM_03 AND LYQ_01\nOR\nDXD_00 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_03 AND IOP_07 AND LVM_03 AND LYQ_02\nOR\nDXD_05 AND IOP_07 AND LVM_03 AND LYQ_02'
    them_code = getThematiccode_for_DCI(effective)
    # print("them_code------------>", them_code)
