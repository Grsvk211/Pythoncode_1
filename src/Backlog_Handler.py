import sys

import ExcelInterface as EI
import os
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import time
from lexer import Lexer
from thmParser import Parser
import re
import collections
# from openpyxl import load_workbook
import json
import xlwings as xw
import InputConfigParser as ICF
import logging

UpdateHMIInfoCb=None
#from BusinessLogic import displayInformation

# Open JSON file
f_backlog = open('../user_input/Backlog.json', "r")

# Convert JSON to DICT
backlog = json.load(f_backlog)


def displayInfoInGUI(func):
    global UpdateHMIInfoCb
    UpdateHMIInfoCb = func


def getKPIpath():
    return backlog["DocPath"]["KPI"]


def getReferentialpath():
    return backlog["DocPath"]["Referential"]


def getArchitecture():
    return backlog["DocPath"]["ARCH"]


# def getArchitecture():
#     return backlog["DocPath"]["ARCH"]


def getRequirementList():
    return backlog["DocPath"]["Requirement"]


def myRange(start, end, step):
    i = start
    while i < end:
        yield i
        i += step
    yield end


def mergeDict(dict1, dict2):
    dict3 = dict1.copy()
    for key in dict3:
        if key == "count":
            dict3[key] = dict3[key] + dict2[key]
        if key == "cellPositions":
            dict3[key] = dict3[key] + dict2[key]
        if key == "cellValue":
            dict3[key] = dict3[key] + dict2[key]
    return dict3


def searchDataInExcel_(sheet, cellRange, keyword):
    start_time = time.time()
    logging.info(f'searchDataInExcel start time: {start_time}')
    value = sheet.used_range.value
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    # x is the index of column
    # i is the value of column
    # y is the index of row
    # j is the value of cell
    if keyword == "":
        return searchResult

    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if j is not None:
                if keyword in str(j):
                    searchResult["count"] = searchResult["count"] + 1
                    searchResult["cellPositions"].append((x + 1, y + 1))
                    searchResult["cellValue"].append(j)

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'searchDataInExcel end execution time: {execution_time}')

    return searchResult


def searchDataInExcelCache(value, cellRange, keyword):
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


def threadFun(path_to_sheet, sheetName, cellRange, keyword):
    time.sleep(1)
    # logging.info("sheetname", sheetName)
    x, y = cellRange
    # logging.info("searching data in col", x, y)
    c = 0
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    try:
        sheet = EI.openExcel(path_to_sheet).sheets[sheetName]
        for row in range(1, int(y + 1)):
            cellValue = str(sheet.range(row, x).value)
            if keyword in cellValue:
                searchResult["cellPositions"].append(tuple((row, x)))
                searchResult["cellValue"].append(cellValue)
                c = c + 1
                searchResult["count"] = c
    except:
        pass
    return searchResult


def threadFunCol(path_to_sheet, sheetName, cellRange, keyword):
    time.sleep(1)
    # logging.info("\nIn threadFunCol function sheetname", sheetName)
    start, end, y = cellRange
    # logging.info("searching data in col", start, end, y)
    c = 0
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    try:
        sheet = EI.openExcel(path_to_sheet).sheets[sheetName]
        for row in range(start, end + 1):
            cellValue = str(sheet.range(row, y).value)
            if keyword in cellValue:
                searchResult["cellPositions"].append(tuple((row, y)))
                searchResult["cellValue"].append(cellValue)
                c = c + 1
                searchResult["count"] = c
            # logging.info("row = ", row)
    except:
        pass
    # logging.info("threadFunCol funtion returning - ", searchResult)
    return searchResult


def grepThematicsCode(rawThematics):
    logging.info("grepThematicsCode(rawThematics--------->",rawThematics)
    start_time = time.time()
    logging.info(f'grepThematicsCode start time: {start_time}')

    try:
        rawThematics = rawThematics[((rawThematics.index(']')) + 1):]
    except:
        rawThematics = rawThematics
    logging.info("rawThematics before - ", rawThematics)
    rawThematics = rawThematics.replace("(", " ( ")
    rawThematics = rawThematics.replace(")", " ) ")
    logging.info("rawThematics after - ", rawThematics)
    thematics_code = ['AND']
    for a in rawThematics.split(" "):
        # logging.info("a - ", a)
        if re.search("[a-zA-Z0-9]{3}[(][0-9]{2}[)]", a) is not None:
            a = a.replace("(", "_")
            a = a.replace(")", "")
            a = a.strip()
            thematics_code.append(a)
            logging.info("thematics_code = ", thematics_code)
        else:
            if a.find("AND") == 0:
                if (thematics_code[-1] != "AND"):
                    thematics_code.append(a)
            if a.find("OR") == 0:
                if (thematics_code[-1] != "OR"):
                    thematics_code.append(a)
            if (re.search("{", a)) is not None:
                thematics_code.append("(")
            if (re.search("}", a)) is not None:
                thematics_code.append(")")
            if (re.search("[(][(][(]", a)) is not None:
                thematics_code.append(re.findall("[(((]", a)[0])
            if (re.search("[)][)][)]", a)) is not None:
                thematics_code.append(re.findall("[)))]", a)[0])
            if (re.search("[(][(]", a)) is not None:
                thematics_code.append(re.findall("[((]", a)[0])
            if (re.search("[)][)]", a)) is not None:
                thematics_code.append(re.findall("[))]", a)[0])
            if (re.search("[(]", a)) is not None:
                thematics_code.append(re.findall("[(]", a)[0])
            if (re.search("[)]", a)) is not None:
                thematics_code.append(re.findall("[)]", a)[0])
            if re.search("[a-zA-Z0-9]{3}_[0-9]{2}", a) is not None:
                thematics_code.append(" ( " + (re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", a)[0]) + " ) ")
    if len(re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", thematics_code[0])) == 0:
        if thematics_code[0] != "(":
            # logging.info("removing first element", thematics_code[0], re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", thematics_code[0]))
            thematics_code.remove(thematics_code[0])
    # logging.info("Thematic = ", thematics_code)
    # logging.info("Thematic code final(1) = ", ''.join(thematics_code))
    openBracket = []
    closeBracket = []
    for i in range(len(thematics_code)):
        if thematics_code[i] == "(":
            openBracket.append(i)
        if thematics_code[i] == ")":
            closeBracket.append(i)
    # logging.info("Indices = ", openBracket, closeBracket)
    # logging.info("Thematic code = ", thematics_code)
    for n, i in enumerate(thematics_code):
        # logging.info("N & i", n, i)
        if i.find('_') != -1:
            # logging.info("thm code",thematics_code[n+1])
            if n < (len(thematics_code) - 1):
                if thematics_code[n + 1].find('_') != -1:
                    thematics_code[n + 1] = ',' + thematics_code[n + 1]
    reducedThm = ' '.join(thematics_code)
    logging.info("Thematic code final(2)1 = ", reducedThm)
    reducedThm = reducedThm.replace("( )", "")
    logging.info("Thematic code final(2)2 = ", reducedThm)
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'grepThematicsCode end execution time: {execution_time}')

    return reducedThm


def createCombination(data):
    start_time = time.time()
    logging.info(f'createCombination start time: {start_time}')

    lexer = Lexer().get_lexer()
    tokens = lexer.lex(data)
    '''
    for token in tokens:
        logging.info(token)'''

    pg = Parser()
    pg.parse()
    parser = pg.get_parser()
    combinations = parser.parse(tokens).eval()

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'createCombination end execution time: {execution_time}')

    return combinations


def filterThemForArch(thematicLine, refEC, ARCH):
    # refBook = xw.Book(r"C:/Users/6451/Desktop/bsi_auto/09-11-2021/Modified/Aptest/Input/Referentiel_EC.xlsm")

    ListOfThematics = thematicLine.split("|")
    time.sleep(1)
    sheet = refEC.sheets['Liste EC']
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    sheet_Value = sheet.used_range.value
    # final_VSMr2_thm = []
    tempflagR1 = 0
    tempflagR2 = 0
    ListOfThematicsCopy = ListOfThematics.copy()
    for i in ListOfThematicsCopy:
        flagR1 = 0
        flagR2 = 0
        # try:
        # searchResults = searchDataInExcel(sheet, (maxrow, 7), i)
        searchResults = searchDataInExcelCache(sheet_Value, (maxrow, 7), i)
        # except:
        #     searchResults = searchDataInCol_1(sheet, (maxrow, 7), i)
        if searchResults["count"] != 0:
            x, y = searchResults["cellPositions"][0]
            applicableBSI = sheet.range(x, y + 38).value
            applicableR1 = sheet.range(x, y + 39).value
            applicableR2 = sheet.range(x, y + 40).value
            logging.info("Thematique = ", i, "Aplicable to = ", applicableBSI, applicableR1, applicableR2)
            # for BSi Arch
            if ARCH == "BSI":
                if applicableBSI == "Y":
                    pass
                else:
                    # ListOfThematics.remove(i)
                    logging.info("not applicable for BSI but its present in req\n")
                    return -1
            # for VSM Arch
            elif ARCH == "VSM":
                if (applicableR1 == "Y") and (applicableR2 == "Y"):
                    logging.info("Aplicable to R1 & R2")
                    pass
                elif (applicableR1 == "Y") or (applicableR2 == "Y"):
                    if (applicableR1 == "Y"):
                        flagR1 = 1
                        logging.info("NEA R1 applicable")
                        tempflagR1 = flagR1
                    elif (applicableR2 == "Y"):
                        flagR2 = 1
                        logging.info("NEA R2 applicable")
                        tempflagR2 = flagR2
                else:
                    # ListOfThematics.remove(i)
                    logging.info("not applicable for VSM but its present in req\n")
                    return -1
            else:
                logging.info("arch not found\n")
                pass
        else:
            logging.info("Thematique not found in referntial EC")
            return -1
    if (tempflagR1 == 1) and (tempflagR2 == 1):
        logging.info("CONFLICT")
        return -1
    else:
        i = 1
        thematicLine = ""
        for l in ListOfThematics:
            thematicLine = thematicLine + l
            if i < len(ListOfThematics):
                thematicLine = thematicLine + "|"
                i = i + 1
        return thematicLine


# Takes KPI folder name as input
# Return: A KPI folder can contain many KPI excel files, This method merges the KPI document name and
# its path and returns the list

def getKPIDocPath(path):
    docList = []
    documents = os.listdir(path)
    for d in documents:
        a = (path + "\\" + d)
        docList.append(a)
    return docList


# Takes List of KPI paths and requirement list ["REQ-1|REQ-2|REQ-3"] as input
# Returns list of thematic lines of each req (REQ-1|REQ-2|REQ-3)
#
def searchDataInKPI(docList, requirement):
    # logging.info(f"In searchDataInKPI - {docList}")
    #logging.info("\n------------------------------")
    start_time = time.time()
    #logging.info(f'searchDataInKPI start time: {start_time}')

    status = True
    thematiqueList = []
    reqNotFound = []
    foundReq = []
    reqVerNotFound=[]
    requirement = requirement.split("|")
    logging.info(f"requirement>> {requirement}")
    try:
        for n, doc in enumerate(docList):
            with XwApp(visible=True) as app:
                d = app.books.open(doc)
                # sheetflag = 0
                for sheet in d.sheets:
                    if sheet.name.find("REQ") != -1:
                        # maxCol = sheet.range('A10').end('right').last_cell.column
                        # #logging.info("maxCol - ", maxCol)
                        maxrow1 = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                        maxrow2 = sheet.range('B' + str(sheet.cells.last_cell.row)).end('up').row
                        sheet_value = sheet.used_range.value
                        if maxrow1 > maxrow2:
                            maxrow = maxrow1
                        else:
                            maxrow = maxrow2
                        for r in requirement:
                            if r not in reqNotFound:
                                with open('../Thematics_Report.txt', 'a') as f:
                                    f.writelines(
                                        "\n\n--------------------------------" + r + "--------------------------------\n")
                                flag = 0

                                if r.find("REQ-") != -1:
                                    # logging.info("r = ", r)
                                    req = r.split("(")[0]
                                    ver = r.split("(")[1].split(")")[0]
                                else:
                                    req = r
                                    ver = ""
                                # logging.info("req , ver = ", req, ver)
                                if req.find("REQ-") != -1:
                                    # result = searchDataInExcel(sheet, (maxrow, 2), req)
                                    result = searchDataInExcelCache(sheet_value, (maxrow, 2), req)
                                    # logging.info("Result =", result, result["cellPositions"])
                                    if result["count"] >= 1:
                                        # cellPositions = result["cellPositions"][0]
                                        foundVersion = False
                                        for count, cellPositions in enumerate(result["cellPositions"]):
                                            row, col = cellPositions
                                            version = EI.getDataFromCell(sheet, (row, col+1))
                                            if version is None:
                                                version = EI.getDataFromCell(sheet, (row, col + 2))
                                            # version = EI.getDataFromCellCache(sheet_value, (row, col + 1))
                                            logging.info("\n\nVersion in document - ", version, type(version))
                                            logging.info("Version in ts - ", ver,"\n\n")
                                            if type(version) is float:
                                                version = str(int(version))
                                                logging.info("Version in docc - ", version, "\n\n")
                                            if ver == version:
                                                # logging.info("fffff >> ")
                                                foundVersion = True
                                                with open('../Thematics_Report.txt', 'a') as f:
                                                    f.writelines(
                                                        "\n\nRequirement - " + r + " found in document " + str(
                                                            d) + "\n")
                                                # Logic to find effective expression
                                                # effective_col = searchDataInExcel(sheet, (100, 100),"Effectivity Expression")
                                                effective_col = searchDataInExcelCache(sheet_value, (100, 100),
                                                                                       "Effectivity Expression")
                                                effective_col["cellPositions"].sort()
                                                cells = effective_col["cellPositions"][0]
                                                effective_r, effective_c = cells
                                                # logging.info("effective_c - ", effective_c)
                                                effective = EI.getDataFromCell(sheet, (row, effective_c))
                                                # effective = EI.getDataFromCellCache(sheet_value, (row, effective_c))

                                                if effective is None:
                                                    effectiveExpression = "None"
                                                    with open('../Thematics_Report.txt', 'a') as f:
                                                        f.writelines("\nEffective expression -\n" + str(effective))
                                                else:
                                                    effective = str(effective.encode('utf-8').strip())
                                                    # logging.info("Result effective= ", effective)
                                                    # effectiveExpression = grepEffectiveExpression(effective)
                                                    data = grepThematicsCode(effective)

                                                    if data.strip().endswith("OR'") or data.strip().endswith("AND'"):
                                                        data=remove_trailing_and_or(data)
                                                    # logging.info(f"data1 {req} <--> {data}")
                                                    effectiveExpression=''
                                                    if len(data.strip())!=0:
                                                        effectiveExpression = createCombination(data)


                                                    # logging.info("Final Combination - ", effectiveExpression)
                                                    # thematiqueList.append(them_line)
                                                    with open('../Thematics_Report.txt', 'a') as f:
                                                        f.writelines("\nEffective expression -\n" + str(
                                                            effectiveExpression.encode(
                                                                'utf-8').strip()) + "\nFinal Combination from effective- " + effectiveExpression)

                                                # logic to get LCDV
                                                # lcdv_col = searchDataInExcel(sheet, (100, 100), "Lcdv")
                                                lcdv_col = searchDataInExcelCache(sheet_value, (100, 100), "Lcdv")
                                                lcdv_col["cellPositions"].sort()
                                                cells = lcdv_col["cellPositions"][0]
                                                lcdv_r, lcdv_c = cells
                                                # logging.info("lcdv_c - ", lcdv_c)
                                                lcdv = EI.getDataFromCell(sheet, (row, lcdv_c))
                                                # lcdv = EI.getDataFromCellCache(sheet_value, (row, lcdv_c))
                                                if lcdv is None:
                                                    if count == (result["count"] - 1):
                                                        with open('../Thematics_Report.txt', 'a') as f:
                                                            f.writelines("\n\nLCDV expression\n" + str(lcdv))
                                                        flag = 1
                                                        if effectiveExpression is not None and effectiveExpression != 'None' and effectiveExpression != '-1':
                                                            thematiqueList.append(effectiveExpression)
                                                            # sheetflag = 1
                                                            # compareLCDVandEffective(lcdv, effective)
                                                            foundReq.append(r)
                                                            break
                                                else:
                                                    try:
                                                        lcdv = str(lcdv.encode('utf-8').strip())
                                                        # logging.info("Result lcdv= ", lcdv)
                                                        data = grepThematicsCode(lcdv)
                                                        # logging.info(f"data2 {data}")
                                                        them_line = createCombination(data)
                                                        # logging.info("Simplified LCDV - ", data)
                                                        # logging.info("Final Combination - ", them_line)
                                                        # logging.info("effective expression",effectiveExpression)
                                                        if len(them_line) == 0:

                                                            them_line = effectiveExpression

                                                        with open('../Thematics_Report.txt', 'a') as f:
                                                            f.writelines("\nLCDV expression -\n" + str(lcdv.encode(
                                                                'utf-8').strip()) + "\nSimplified LCDV - " + data + "\nFinal Combination from LCDV- " + them_line)
                                                        flag = 1
                                                        # sheetflag = 1
                                                        if them_line is not None and them_line != 'None' and them_line != '-1':
                                                            thematiqueList.append(them_line)
                                                            compareLCDVandEffective(them_line, effectiveExpression)
                                                    except Exception as e:
                                                        # logging.info("In except - ", e)
                                                        them_line = str(e)
                                                        with open('../Thematics_Report.txt', 'a') as f:
                                                            f.writelines("\nError LCDV- " + them_line)
                                                        flag = 1
                                                        if effectiveExpression is not None and effectiveExpression != 'None' and effectiveExpression != '-1':
                                                            thematiqueList.append(effectiveExpression)
                                                        # sheetflag = 1
                                                    foundReq.append(r)
                                                    break
                                            else:
                                                if count == (result["count"] - 1):
                                                    # logging.info(f"No thematic for this requirement {req}")
                                                    effectiveExpression = "-1"
                                                    # un comment if need
                                                    # thematiqueList.append(effectiveExpression)

                                        if foundVersion == False:
                                            status = False
                                            if n == (len(docList) - 1):
                                               reqVerNotFound.append(r)
                                    else:
                                        # d.close()
                                        # app.kill()
                                        # time.sleep(2)
                                        if n == (len(docList) - 1):
                                            # logging.info("Req not found in any sheet please check manually")
                                            status = False
                                            with open('../Thematics_Report.txt', 'a') as f:
                                                f.writelines(
                                                    "\n\nRequirement" + r + " not found in any document. Please proceed manually")
                                            reqNotFound.append(r)
                                            # logging.info("reqNotFound - ", reqNotFound)
                                            effectiveExpression = "-1"
                                            # un comment if need
                                            # thematiqueList.append(effectiveExpression)
                                else:
                                    # result = searchDataInExcel(sheet, (maxrow, 1), req)
                                    result = searchDataInExcelCache(sheet_value, (maxrow, 1), req)
                                    # logging.info("Result =", result, result["cellPositions"])

                                    if result["count"] == 0:
                                        if (req.find('.') != -1):
                                            req = req.replace('.', '-')
                                            # logging.info("req = ", req)
                                            # result = searchDataInExcel(sheet, (maxrow, 1), req)
                                            result = searchDataInExcelCache(sheet_value, (maxrow, 1), req)
                                    if result["count"] >= 1:
                                        # cellPositions = result["cellPositions"][0]
                                        for count, cellPositions in enumerate(result["cellPositions"]):
                                            row, col = cellPositions
                                            version = EI.getDataFromCell(sheet, (row, col + 2))
                                            # #logging.info("Version in document - ", version)
                                            with open('../Thematics_Report.txt', 'a') as f:
                                                f.writelines(
                                                    "\n\nRequirement - " + r + " found in document " + str(d) + "\n")
                                            # Logic to find effective expression
                                            # effective_col = searchDataInExcel(sheet, (100, 100), "Effectivity Expression")
                                            effective_col = searchDataInExcelCache(sheet_value, (100, 100),
                                                                                   "Effectivity Expression")
                                            effective_col["cellPositions"].sort()
                                            cells = effective_col["cellPositions"][0]
                                            effective_r, effective_c = cells
                                            # logging.info("effective_c - ", effective_c)
                                            effective = EI.getDataFromCell(sheet, (row, effective_c))
                                            # effective = EI.getDataFromCellCache(sheet_value, (row, effective_c))

                                            if effective is None:
                                                effectiveExpression = "None"
                                                with open('../Thematics_Report.txt', 'a') as f:
                                                    f.writelines("\nEffective expression -\n" + str(effective))
                                            else:
                                                effective = str(effective.encode('utf-8').strip())
                                                # logging.info("Result effective= ", effective)
                                                # effectiveExpression = grepEffectiveExpression(effective)
                                                data = grepThematicsCode(effective)
                                                # logging.info(f"data3 {data}")
                                                if data.strip().endswith('OR') or data.strip().endswith('AND'):
                                                    data = remove_trailing_and_or(data)
                                                effectiveExpression = ''
                                                if len(data.strip()) != 0:
                                                    effectiveExpression = createCombination(data)

                                                # logging.info("Final Combination - ", effectiveExpression)
                                                # thematiqueList.append(them_line)
                                                with open('../Thematics_Report.txt', 'a') as f:
                                                    f.writelines("\nEffective expression -\n" + str(
                                                        effectiveExpression.encode(
                                                            'utf-8').strip()) + "\nFinal Combination from effective- " + effectiveExpression)

                                            # logic to get LCDV
                                            # lcdv_col = searchDataInExcel(sheet, (100, 100), "Lcdv")
                                            lcdv_col = searchDataInExcelCache(sheet_value, (100, 100), "Lcdv")
                                            lcdv_col["cellPositions"].sort()
                                            cells = lcdv_col["cellPositions"][0]
                                            lcdv_r, lcdv_c = cells
                                            # logging.info("lcdv_c - ", lcdv_c)
                                            lcdv = EI.getDataFromCell(sheet, (row, lcdv_c))
                                            # lcdv = EI.getDataFromCellCache(sheet_value, (row, lcdv_c))
                                            if lcdv is None:
                                                if count == (result["count"] - 1):
                                                    with open('../Thematics_Report.txt', 'a') as f:
                                                        f.writelines("\nLCDV expression -\n" + str(lcdv))
                                                    flag = 1
                                                    if effectiveExpression is not None and effectiveExpression != 'None' and effectiveExpression != '-1':
                                                        thematiqueList.append(effectiveExpression)
                                                        # sheetflag = 1
                                                        # compareLCDVandEffective(lcdv, effective)
                                                        foundReq.append(r)
                                                        break
                                            else:
                                                try:
                                                    lcdv = str(lcdv.encode('utf-8').strip())
                                                    # logging.info("Result lcdv= ", lcdv)
                                                    data = grepThematicsCode(lcdv)
                                                    # logging.info(f"data5 {data}")
                                                    them_line = createCombination(data)
                                                    # logging.info("Simplified LCDV - ", data)
                                                    # logging.info("Final Combination - ", them_line)
                                                    if len(them_line) == 0:
                                                        them_line = effectiveExpression
                                                    # thematiqueList.append(effectiveExpression)
                                                    with open('../Thematics_Report.txt', 'a') as f:
                                                        f.writelines("\nLCDV expression -\n" + str(lcdv.encode(
                                                            'utf-8').strip()) + "\nSimplified LCDV - " + data + "\nFinal Combination from LCDV- " + them_line)
                                                    flag = 1
                                                    # sheetflag = 1
                                                    if effectiveExpression is not None and effectiveExpression != 'None' and effectiveExpression != '-1':
                                                        thematiqueList.append(effectiveExpression)
                                                        compareLCDVandEffective(them_line, effectiveExpression)
                                                except Exception as e:
                                                    # logging.info("In except - ", e)
                                                    them_line = str(e)
                                                    with open('../Thematics_Report.txt', 'a') as f:
                                                        f.writelines("\nError in LCDV- " + them_line)
                                                    flag = 1
                                                    thematiqueList.append(effectiveExpression)
                                                    # sheetflag = 1
                                                foundReq.append(r)
                                                break
                                    else:
                                        # d.close()
                                        # app.kill()
                                        # time.sleep(2)
                                        if n == (len(docList) - 1):
                                            # logging.info("Req not found in any sheet please check manually")
                                            status = False
                                            with open('../Thematics_Report.txt', 'a') as f:
                                                f.writelines(
                                                    "\n\nRequirement" + r + " not found in any document. Please proceed manually")
                                            req = r.split("(")[0]
                                           # ver = r.split("(")[1].split(")")[0]
                                            result = searchDataInExcelCache(sheet_value, (maxrow, 1), req)

                                            #version = EI.getDataFromCell(sheet, (row, 3))
                                            # logging.info("Result =", result, result["cellPositions"])
                                            if result["count"] >= 1:
                                                reqVerNotFound.append(r)
                                            else:
                                                reqNotFound.append(r)
                                            # logging.info("reqNotFound - ", reqNotFound)
                                            effectiveExpression = "-1"
                                            # un comment if need
                                            # thematiqueList.append(effectiveExpression)
                                # if sheetflag == 1:
                                #     break
                                if flag == 1:
                                    pass
                                    # d.close()
                                    # app.kill()
                                    # time.sleep(2)
                                    # logging.info("List of thematiques from all requirement -\n", thematiqueList)
                                    # break
        with open('../Thematics_Report.txt', 'a') as f:
            f.writelines("\n-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*--*-*-*-*-*-\n")
        with open('../Thematics_Report.txt', 'a') as f:
            f.writelines("\n\nRequirements not found in the document -\n" + ', '.join(reqNotFound))
        with open('../Thematics_Report.txt', 'a') as f:
            f.writelines(
                "\n\nList of thematiques from all requirement found in document-\n" + '\n'.join(thematiqueList))

    except Exception as e:
        logging.info(f"Error .... {e}")
    # end_time = time.time()
    # execution_time = end_time - start_time

    # logging.info(f'searchDataInKPI end execution time: ')
    # logging.info("\n------------------------------")
    if reqNotFound:
        UpdateHMIInfoCb(f"Requirements not found in KPI:\n{','.join(reqNotFound)}.")
        # logging.info(f"Requirements not found in KPI:\n{','.join(reqNotFound)}.")
    if reqVerNotFound:
        UpdateHMIInfoCb(f"Requirements version not matched in KPI:\n{','.join(reqVerNotFound)}.")
        # logging.info(f"Requirements version not matched in KPI:\n{','.join(reqVerNotFound)}.")
    return status, thematiqueList


# fucntion to create thematique lines for all requirement in sheet
def createApplicableCombination(thematiqueList, refEC, ARCH):
    themLines = []
    applicableThemLines = []
    for them in thematiqueList:
        them = them.split("\n")
        themLines.append(them)
    logging.info("themLines - ", themLines)
    # themLines.sort(key=len)
    for thematic in themLines:
        logging.info("t = ", thematic)
        for t in thematic:
            applicableThemLine = filterThemForArch(''.join(t), refEC, ARCH)
            if applicableThemLine != -1:
                applicableThemLines.append(t)
    applicableThemLines.sort(key=len)
    applicableThemLines.reverse()
    logging.info("applicableThemLines - ", applicableThemLines)


# function to grep Effective expression
def grepEffectiveExpression(rawThematics):
    try:
        rawThematics = rawThematics[((rawThematics.index(']')) + 1):]
    except:
        rawThematics = rawThematics
    logging.info("rawThematics after - ", rawThematics)
    thematics_code = ['AND']
    for a in rawThematics.split(" "):
        if re.search("[a-zA-Z0-9]{3}[(][0-9]{2}[)]", a) is not None:
            # logging.info("IN IF")
            a = a.replace("(", "_")
            a = a.replace(")", "")
            a = a.strip()
            thematics_code.append(a)
            # logging.info("thematics_code = ", thematics_code)
        else:
            if a.find("AND") == 0:
                if (thematics_code[-1] != "AND"):
                    thematics_code.append("|")
            if a.find("OR") == 0:
                if (thematics_code[-1] != "OR"):
                    thematics_code.append("\n")
            if re.search("[a-zA-Z0-9]{3}_[0-9]{2}", a) is not None:
                thematics_code.append(re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", a)[0])
    if len(re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", thematics_code[0])) == 0:
        thematics_code.remove(thematics_code[0])
    for n, i in enumerate(thematics_code):
        # logging.info("N & i", n, i)
        if i.find('_') != -1:
            # logging.info("thm code",thematics_code[n+1])
            if n < (len(thematics_code) - 1):
                if thematics_code[n + 1].find('_') != -1:
                    thematics_code[n + 1] = ',' + thematics_code[n + 1]
    reducedThm = ' '.join(thematics_code)
    reducedThm = reducedThm.replace(" ", "")
    logging.info("Thematic code final(2) =\n", reducedThm)
    return reducedThm


def compareLCDVandEffective(lcdv, effective):
    lcdv = lcdv.split("\n")
    lcdv.sort(key=len)
    lcdv.reverse()
    effective = effective.split("\n")
    effective.sort(key=len)
    effective.reverse()
    logging.info("lcdv - ", lcdv)
    logging.info("effective - ", effective)
    for n, l in enumerate(lcdv):
        l = l.split("|")
        l.sort()
        l = '|'.join(l)
        lcdv[n] = l
    logging.info("lcdv(1) - ", lcdv)
    for n, i in enumerate(effective):
        i = i.split("|")
        i.sort()
        i = '|'.join(i)
        effective[n] = i
    logging.info("effective(1) - ", effective)
    if collections.Counter(effective) == collections.Counter(lcdv):
        logging.info("SAME")
        with open('../Thematics_Report.txt', 'a') as f:
            f.writelines("\n\nThematique combinations formed from LCDV & Effective expression are SAME\n")
    else:
        logging.info("DIFFERENT")
        with open('../Thematics_Report.txt', 'a') as f:
            f.writelines("\n\nThematique combinations formed from LCDV & Effective expression are DIFFERENT\n")


def sortFirst(val):
    return val[0]


def sortThematics(data):
    start_time = time.time()
    logging.info(f'sortThematics start time: {start_time}')

    countN = []
    res = []
    for n, i in enumerate(data):
        countN.append((i.count("\n"), n))
    countN.sort(key=sortFirst, reverse=True)
    for index in countN:
        res.append(data[index[1]])

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'sortThematics end execution time: {execution_time}')
    return res


def sortThmLine(thmLine):
    start_time = time.time()
    logging.info(f'\nsortThmLine start time: {start_time}')

    line = thmLine.split("\n")
    temp = []
    for l in line:
        thm = l.split("|")
        thm.sort()
        temp.append("|".join(thm))
    temp.sort()
    sortedThms = ("\n".join(temp))
    # logging.info("sorted ",sortedThms,"  ",thmLine)
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'sortThmLine end execution time: {execution_time}')

    return sortedThms


def compareThematics(currData, data):
    logging.info("compare ", currData, data)
    start_time = time.time()
    logging.info(f'\n compareThematics start time: {start_time}')
    currData = sortThmLine(currData)
    for d in data:
        sortedThms = sortThmLine(d)
        logging.info("sorted 1", sortedThms, "  ", currData)
        if currData == sortedThms:
            logging.info("True")
            end_time = time.time()
            execution_time = end_time - start_time
            logging.info(f'\n compareThematics end execution time: {execution_time}')
            return True
        else:
            pass
    return False


def sortComb(data):
    start_time = time.time()
    logging.info(f'\n sortComb start time: {start_time}')

    result = []
    for i in data:
        result.append(sortThmLine(i))

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n sortComb end execution time: {execution_time}')
    return result


def findContrary(name, thms):
    start_time = time.time()
    logging.info(f'\n findContrary start time: {start_time}')
    for t in thms:
        logging.info("names ", t.split("_")[0], name)
        # if t.split("_")[0] == name:
        #     return True

        if t.split("_")[0] == name:
            end_time = time.time()
            execution_time = end_time - start_time
            logging.info(f'\n findContrary end execution time: {execution_time}')
            return True

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n findContrary end execution time: {execution_time}')
    return False


def createLineCombinations(lines, data):
    start_time = time.time()
    logging.info(f'\n createLineCombinations start time: {start_time}')

    logging.info("l", lines)
    logging.info("d ", data)
    result = []
    for first in lines:
        tempo = []
        if data:
            for i in data:
                remLines = i.split("\n")
                for j in remLines:
                    intComb = first + '|' + j
                    tempo.append(intComb)
        else:
            if lines not in result:
                result.append(lines)
        result.append(tempo)

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n createLineCombinations end execution time: {execution_time}')
    return result


def removeContrary(thmLines):
    start_time = time.time()
    logging.info(f'\n removeContrary start time: {start_time}')
    result = []
    logging.info("thmLines Contrary:: ", thmLines)
    data = thmLines.copy()
    for req in data:
        logging.info("\nreq ", req)
        for line in req:
            temp = []
            data = line.split("|")
            [temp.append(x) for x in data if x not in temp]  # removing duplicates
            logging.info("temp >>****** ", temp)
            flag = 0
            for n, thm in enumerate(temp.copy()):
                thmName = thm.split("_")[0]
                if findContrary(thmName, temp[:n] + temp[n + 1:]):
                    flag = 1
                    break
                # if findContrary(thm, data[n+1]):
                #     flag = 1
                #     break
        logging.info("flag>>> ", flag)
        if flag == 1:
            logging.info("flagContraryReq>>> ", req)
            if req in thmLines:
                thmLines.remove(req)
        logging.info("$$$thmLines after reoved : ", thmLines)
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n removeContrary end execution time: {execution_time}')
    return thmLines


def checkCompatibility(lines):
    final = []
    for thmLine in lines.copy():
        result = removeContrary(thmLine)
        logging.info("thm ", thmLine)
        logging.info("res ", result)
        if result:
            if result not in final:
                final.append(result)
    logging.info("Check ", final)
    return final


def checkSubset(lines, datas):
    logging.info("data ", datas, lines)
    result = []
    for line in lines:
        for data in datas:
            # result=None
            for i in data:
                # logging.info("chk ",line,i)
                # if (all(x in i.split('|') for x in line.split('|'))):
                # if set(line.split('|')).issubset(set(i.split('|'))):
                if i == line:
                    if i not in result:
                        result.append(i)
    logging.info("Check result ", result)
    if result:
        return result
    else:
        return None


def flatten(data):
    result = []
    for i in data:
        for j in i:
            result.append(j)
    return result


def createFinalCombinations(lines):
    start_time = time.time()
    logging.info(f'\n createFinalCombinations start time: {start_time}')
    result = []
    for line in lines:
        # r=line.copy()
        logging.info("l ", line)
        res = []
        [res.append(x) for x in line if x not in res]
        # logging.info(res)
        temp = res.copy()
        for n, l in enumerate(temp):
            if n < len(res):
                if (checkSubset(l, temp[n + 1:])):
                    logging.info("in")
                    res.remove(l)
            else:
                # if (checkSubset(l,temp[n:])):
                #     logging.info("in")
                #     res.remove(l)
                pass
        if res:
            logging.info("ress ", res)
            result.append(res)
    temp = flatten(result)
    final = []
    [final.append(x) for x in temp if x not in final]
    # tempRes=result.copy()
    # for nn,j in enumerate(tempRes):
    #     if n<len(tempRes):
    #         if checkSubset(j, tempRes[n+1:]):
    #             result.remove(j)

    logging.info("result Final Combination ", final)
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n createFinalCombinations end execution time: {execution_time}')
    return final


def eliminateContrary(final):
    result = []
    for i in final:
        data = i.split("|")

        for n, d in enumerate(data.copy()):
            thmName = d.split("_")[0]
            flag = 0
            for j in data[n + 1:]:
                if j.split("_")[0] == thmName:
                    flag = 1
                    data.remove(j)
            if flag == 1:
                data.remove(d)
        result.append("|".join(data))

    return result


def checkContrary(data):
    intLines = []
    final = []
    firstLines = data[0].split("\n")
    logging.info("first", firstLines, data[1:], len(data))
    if len(data) > 1:
        intLines = createLineCombinations(firstLines, data[1:])
        logging.info("int", intLines)
        compatibleLines = checkCompatibility(intLines)
        logging.info("comp ", compatibleLines)
        final = createFinalCombinations(compatibleLines)
        # final=eliminateContrary(final)
    else:
        return firstLines

    # logging.info("resss ",result)
    logging.info("fff ", final)
    return final


def remove_trailing_and_or(input_string):
    if input_string.strip().endswith("'"):
        input_string=input_string.strip()[:-1]
    while input_string.strip().endswith('AND') or input_string.strip().endswith('OR'):
        if input_string.strip().endswith('AND'):
            input_string = input_string.strip()[:-3]
        elif input_string.strip().endswith('OR'):
            input_string = input_string.strip()[:-2]
    return input_string


def filterThemForArch(thematicLine, refEC, ARCH):
    # refBook = xw.Book(r"C:/Users/6451/Desktop/bsi_auto/09-11-2021/Modified/Aptest/Input/Referentiel_EC.xlsm")
    logging.info("In filterThemForArch function - ", thematicLine)
    ListOfThematics = thematicLine.split("|")
    time.sleep(1)
    sheet = refEC.sheets['Liste EC']
    logging.info("sheet = ", sheet)
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    logging.info(maxrow, ListOfThematics)
    sheet_value = sheet.used_range.value
    # final_VSMr2_thm = []
    tempflagR1 = 0
    tempflagR2 = 0
    ListOfThematicsCopy = ListOfThematics.copy()
    logging.info("ARCH -->> ", ARCH)
    logging.info("ListOfThematicsCopy -->> ", ListOfThematicsCopy)
    for i in ListOfThematicsCopy:
        flagR1 = 0
        flagR2 = 0
        logging.info("In filterThemForArch (maxrow, i, sheet)", maxrow, i, sheet)
        # try:
        logging.info("---------HI-3---------------")
        # searchResults = EI.searchDataInExcel(sheet, (maxrow, 7), i)
        # searchResults = EI.searchDataInCol(sheet, 7, i)
        searchResults = searchDataInExcelCache(sheet_value, 7, i)
        logging.info("---------HI-3---------------")
        # except:
        #     searchResults = searchDataInCol_1(sheet, (maxrow, 7), i)
        logging.info("searchresult------->a3:", searchResults)
        if searchResults["count"] != 0:
            x, y = searchResults["cellPositions"][0]
            applicableBSI = sheet.range(x, y + 38).value
            applicableR1 = sheet.range(x, y + 39).value
            applicableR2 = sheet.range(x, y + 40).value
            logging.info("Thematique = ", i, "\napplicableBSI = ", applicableBSI, "\napplicableR1 = ", applicableR1,
                  "\napplicableR2 = ", applicableR2)
            # for BSi Arch
            if ARCH == "BSI":
                if applicableBSI == "Y":
                    pass
                else:
                    # ListOfThematics.remove(i)
                    logging.info("not applicable for BSI but its present in req\n")
                    return -1
            # for VSM Arch
            elif ARCH == "VSM":
                if (applicableR1 == "Y") and (applicableR2 == "Y"):
                    logging.info("Aplicable to R1 & R2")
                    pass
                elif (applicableR1 == "Y") or (applicableR2 == "Y"):
                    if (applicableR1 == "Y"):
                        flagR1 = 1
                        logging.info("NEA R1 applicable")
                        tempflagR1 = flagR1
                    elif (applicableR2 == "Y"):
                        flagR2 = 1
                        logging.info("NEA R2 applicable")
                        tempflagR2 = flagR2
                else:
                    # ListOfThematics.remove(i)
                    logging.info("not applicable for VSM but its present in req\n")
                    return -1
            else:
                logging.info("arch not found\n")
                pass
        else:
            logging.info("Thematique not found in referntial EC")
            return -1
    logging.info("TempFlag = ", tempflagR1, tempflagR2)
    if (tempflagR1 == 1) and (tempflagR2 == 1):
        logging.info("CONFLICT")
        return -1
    else:
        logging.info("Returning thematics ", ListOfThematics)
        i = 1
        thematicLine = ""
        for l in ListOfThematics:
            thematicLine = thematicLine + l
            if i < len(ListOfThematics):
                thematicLine = thematicLine + "|"
                i = i + 1
        logging.info("Returning thematic line", thematicLine)
        return thematicLine


def is_identical(list_a, list_b):
    logging.info(f"len(list_a) {len(list_a)}")
    logging.info(f"len(list_b) {len(list_b)}")

    logging.info(f"list_a {list_a}")
    logging.info(f"list_b {list_b}")
    if len(list_a) != len(list_b):
        return False
    for i in list_a:
        if i not in list_b:
            logging.info("++++++++++**********++++++++++++")
            return False
    return True


def checkSubset(lines, datas):
    logging.info("data >>>", datas)
    logging.info("lines >>>", lines)
    result = []
    for line in lines:
        for data in datas:
            # result=None
            for i in data:
                logging.info("line, i ", line, ",", i)
                # if (all(x in i.split('|') for x in line.split('|'))):
                # if set(line.split('|')).issubset(set(i.split('|'))):
                # logging.info("check ",line)
                # logging.info("i ",i)
                logging.info(f'line.split("|") {line.split("|")}')
                if is_identical(i.split("|"), line.split("|")):
                    logging.info("same ", i, line)
                    if line not in result:
                        result.append(line)
    logging.info("Check result ", result)
    if result:
        return result
    else:
        return None


def checkClubbed(data, result):
    lastID = -1
    final = []
    for i in data:
        lines = i.split("\n")

        for line in lines:

            thms = line.split("|")

            for thm in thms:
                flag = 0
                # logging.info(flag,thm)
                for r in result:
                    # logging.info("clubbed ",thm,r.split("|"))
                    if thm in r.split("|"):

                        flag = 1

                    else:

                        pass
                if flag == 0:
                    # logging.info("Appending ",thm)
                    try:
                        final.append(data.index(i, lastID + 1))
                        lastID = data.index(i, lastID + 1)
                    except:
                        pass
    logging.info("IDs ", final)
    return result


def splitThmsByLine(data):
    result = []
    for i in data:
        temp = i.split('\n')
        result.append(temp)
    return result


def createThmLines(data, refEC, ARCH):
    # logging.info(args)
    # data=list(args)
    logging.info(data)
    result = []

    sortedThms = sortThematics(data)
    logging.info("sorted thms ", sortedThms)
    sortedCombs = sortComb(sortedThms)
    logging.info("sorted  ", sortedCombs)

    [result.append(x) for x in sortedCombs if x not in result]
    logging.info("final ", result, len(result))
    result = checkContrary(result)
    # f=["LYQ_01|ANF_02|AMO_01|ALW_02|BUP_00|AZC_00",
    #    "LYQ_01|ANF_02|AMO_02|ALW_02|BUP_00|AZC_00",
    #    "LYQ_01|ANF_02|AMO_03|ALW_02|BUP_00|EIV_01|AZC_00"]
    # result=f
    logging.info("after contrary check ", result)
    for i in result.copy():
        logging.info(filterThemForArch(i, refEC, ARCH))
        if filterThemForArch(i, refEC, ARCH) == -1:
            result.remove(i)
    IDs = checkClubbed(data, result)
    logging.info("IDDD ", IDs)
    return result


def flattenLines(data):
    result = []
    for i in data:
        temp = i.split('\n')
        for t in temp:
            result.append(t)
    return result


def findRejectedThms(line, data):
    start_time = time.time()
    logging.info(f'\n findRejectedThms start time: {start_time}')
    result = []
    logging.info("\nlineRThem: ", line)
    logging.info("dataRThem: ", data)
    logging.info("find ", line, data)
    for l in line:
        logging.info("yyyyyyyyyy", l)

        cnt = 0
        for ll in l:
            temp = []
            last_index = -1
            for i in data:
                logging.info("xxxxxxxxxxxxxxxxxxx", ll, i)

                for j in i:
                    logging.info(l, j)
                    logging.info("jj>> ", j)

                    if is_identical(ll.split('|'), j.split('|')):
                        cnt = cnt + 1
                        if last_index != len(data) - 1:
                            last_index = data.index(i, last_index + 1)
                        else:
                            last_index = len(data) - 1

                        logging.info("last index ", last_index)
                        if last_index not in temp:
                            temp.append(last_index)
                if temp not in result:
                    result.append(temp)
    logging.info("Rejected index ", result)
    final = []
    [final.append(x) for x in result if x not in final]
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n findRejectedThms end execution time: {execution_time}')
    return final


def findApplicatbility(thematic, refEC, arch):
    start_time = time.time()
    logging.info(f'\n findApplicatbility start time: {start_time}')
    ListOfThematics = thematic.split("|")
    time.sleep(1)
    sheet = refEC.sheets['Liste EC']
    logging.info("sheet............ ", sheet)
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    logging.info("Maxrow,ListOfThematics", maxrow, ", ", ListOfThematics)
    sheet_value = sheet.used_range.value
    listR1 = []
    listR2 = []
    listR1R2 = []
    flag_nea = 0
    ListOfThematicsCopy = ListOfThematics.copy()
    logging.info(f"ListOfThematicsCopy {ListOfThematicsCopy}")
    for i in ListOfThematicsCopy:
        logging.info("iThem.... ", i)
        # searchResults = EI.searchDataInCol(sheet, 7, i)
        searchResults = EI.searchDataInColCache(sheet_value, 7, i)
        logging.info("searchresultttt:", searchResults)
        if searchResults["count"] != 0:
            x, y = searchResults["cellPositions"][0]
            applicableBSI = sheet.range(x, y + 38).value
            applicableR1 = sheet.range(x, y + 39).value
            applicableR2 = sheet.range(x, y + 40).value
            logging.info("Thematique123 = ", i, "\nAplicable to = ", applicableR1, applicableR2)
            # for BSi Arch
            if arch == "BSI":
                if applicableBSI == "Y":
                    pass
                else:
                    # ListOfThematics.remove(i)
                    logging.info("not applicable for BSI but its present in req\n")
                    return -1, []
            elif arch == 'VSM':
                if (applicableR1 == "Y") and (applicableR2 == "Y"):
                    logging.info("Aplicable to R1 & R2............")
                    listR1R2.append(i)
                elif (applicableR1 == "Y") or (applicableR2 == "Y"):
                    if (applicableR1 == "Y"):
                        logging.info("NEA R1 applicable............")
                        listR1.append(i)
                    elif (applicableR2 == "Y"):
                        logging.info("NEA R2 applicable............")
                        listR2.append(i)
                else:
                    if applicableR1 == "N" and applicableR2 == "N" and applicableBSI == 'Y':
                        flag_nea = 1

    logging.info(f"listR2listR2 {listR2}")
    logging.info(f"listR2listR1 {listR1}")
    logging.info(f"listR1R2 {listR1R2}")

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n findApplicatbility end execution time: {execution_time}')
    if len(listR1) > 0 and len(listR2) > 0 or flag_nea == 1:
        return -1, []
    else:
        if len(listR1) > 0 and len(listR2) == 0 and len(listR1R2) >= 0:
            return 1, listR1
        if len(listR1) == 0 and len(listR2) > 0 and len(listR1R2) >= 0:
            return 2, listR2
        if len(listR1R2) > 0:
            return 3, listR1R2
        if len(listR1) == 0 and len(listR2) == 0 and len(listR1R2) == 0:
            return 0, []


def findContrary_BW_ThematicLine(thematicLine1, thematicLine2):
    start_time = time.time()
    logging.info(f'\n findContrary_BW_ThematicLine start time: {start_time}')

    logging.info("\nContname1 :: ", thematicLine1)
    logging.info("Contthms2 :: ", thematicLine2)

    splittedT1 = thematicLine1.split('|')
    for them1 in splittedT1:
        if them1.split("_")[0] in thematicLine2:
            findNum = re.findall(them1.split("_")[0] + "_[0-9]{2}", thematicLine2)
            if findNum[0].split("_")[1] != them1.split('_')[1]:
                end_time = time.time()
                execution_time = end_time - start_time
                logging.info(f'\n findContrary_BW_ThematicLine end execution time: {execution_time}')
                return True
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n findContrary_BW_ThematicLine end execution time: {execution_time}')
    return False


def createThematicCombination(them1, them2, refEC, arch):
    start_time = time.time()
    logging.info(f'\n createThematicCombination start time: {start_time}')
    logging.info("\n\nthem1>> ", them1)
    logging.info("Them2>> ", them2)

    finalThematic = []
    finalThematicList = []

    for thematicLine1 in them1:
        # thematicCodes = []
        joinedThems = []
        splitThematic1 = thematicLine1.split("|")
        for thematicLine2 in them2:
            thematicCodes = []
            findContraryy = findContrary_BW_ThematicLine(thematicLine1, thematicLine2)
            logging.info("findContraryy ||| ", findContraryy)
            # exit()
            if not findContraryy:
                splitThematic2 = thematicLine2.split("|")
                for thematicCode1 in splitThematic1:
                    if thematicCode1 not in thematicCodes:
                        thematicCodes.append(thematicCode1)
                for thematicCode2 in splitThematic2:
                    if thematicCode2 not in thematicCodes:
                        thematicCodes.append(thematicCode2)
                logging.info(f"splitThematic1 {splitThematic1}")
                logging.info(f"splitThematic2 {splitThematic2}")

                joinedThems.append('|'.join(thematicCodes))
                logging.info(f"\njoinedThems {joinedThems} \n")

        # logging.info(f"\nthemCode ============  {'|'.join(thematicCodes)} \n")
        if joinedThems:
            for jthem in joinedThems:
                # finalThematic.append('|'.join(thematicCodes))
                finalThematic.append(jthem)

    logging.info("finalThematic >>> ", finalThematic)
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'\n createThematicCombination end execution time: {execution_time}')

    return finalThematic


def createNewThmLines(data, refEC, ARCH):
    logging.info("111111111111111")
    start_time = time.time()
    logging.info(f'createNewThmLines start time: {start_time}')

    result = []
    accept = []
    reject = []
    acceptList = []
    logging.info("Input ", data)
    sortedThms = sortThematics(data)
    logging.info("sorted thms ", sortedThms)
    sortedCombs = sortComb(sortedThms)
    logging.info("sorted  ", sortedCombs)
    [result.append(x) for x in sortedCombs if x not in result]
    logging.info("final ", result, len(result))
    thmReqLines = splitThmsByLine(sortedCombs)
    logging.info("line ", thmReqLines)
    thmReqLines = removeContrary(thmReqLines)
    logging.info("After contary check ", thmReqLines)
    logging.info(f"len(thmReqLines) -> {len(thmReqLines)}")
    thematicCombination = []
    finalThematicCombination = []
    finalThematicList = []
    try:
        if len(thmReqLines) != 1:
            temp = thmReqLines.copy()
            # temp = data.copy()
            logging.info("temp >> ", temp)
            accept = []
            for n, line in enumerate(temp):
                # if thmLines[n+1:]:
                # logging.info(f"line1www {line}")
                # logging.info(f"thmReqLines[n + 1:] {thmReqLines[n + 1:]}")
                # logging.info(f"thmReqLines[n + 1:] {thmReqLines[n + 1]}")
                if finalThematicCombination:
                    line = finalThematicCombination
                logging.info(f"n , len(temp)-1 {n}  {len(temp) - 1}")
                if n != len(temp) - 1:
                    logging.info(f"\n============line {line}")
                    logging.info(f"\n============thmReqLines[n + 1] {thmReqLines[n + 1]}\n\n")
                    thematicCombination = createThematicCombination(line, thmReqLines[n + 1], refEC, ARCH)
                    logging.info(f"thematicCombination11 {thematicCombination}")
                    if thematicCombination:
                        finalThematicCombination = thematicCombination
                    else:
                        finalThematicCombination = line
                else:
                    # thematicCombination = createThematicCombination(line, thmReqLines[n], refEC, ARCH)
                    logging.info(f"thematicCombination12 {thematicCombination}")
                    logging.info(f"finalThematicCombination {finalThematicCombination}")
                    if finalThematicCombination:
                        # for t1 in thematicCombination:
                        for t1 in finalThematicCombination:
                            logging.info("\n\nt1======= ", t1)
                            thematicApplicability, thematicLines_AP = findApplicatbility(t1, refEC, ARCH)
                            logging.info("thematicApplicable1 === ", thematicApplicability)
                            if thematicApplicability == -1:
                                logging.info(
                                    f"\n______________________Thematic code in thematic line {t1} having R1 and R2 architecture")

                            else:
                                if thematicApplicability == 1 or thematicApplicability == 2 or thematicApplicability == 3:
                                    finalThematicList.append(t1)
                logging.info("createComb============ ", thematicCombination)

            logging.info(f"finalThematicList {finalThematicList}")
            if finalThematicList:
                if finalThematicList not in accept:
                    logging.info(f"&&&&&&&&&&&&checkdLine1212 >> {finalThematicList}")
                    # accept.append(finalThematicCombination)
                    accept.append(finalThematicList)
                    # break
        else:
            accept = thmReqLines

            for n, i in enumerate(accept.copy()):
                temp = i.copy()
                for a in i:
                    if filterThemForArch(a, refEC, ARCH) == -1:
                        temp.remove(a)
                accept[n] = temp
            acceptList = findRejectedThms(accept, splitThmsByLine(data))
            logging.info(f'acceptList {accept}')

        end_time = time.time()
        execution_time = end_time - start_time
        logging.info(f'createNewThmLines end execution time: {execution_time}')
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        logging.info(f"\nError: {ex} \nError line no. {exc_tb.tb_lineno} file name: {exp_fname}")

    return accept, acceptList


def Diff(li1, li2):
    li_dif = [i for i in li1 + li2 if i not in li1 or i not in li2]
    return li_dif


class XwApp(xw.App):
    def __enter__(self, *args, **kwargs):
        return super(*args, **kwargs)

    def __exit__(self, *args):
        for book in self.books:
            try:
                book.close()
            except:
                pass
        self.kill()


########################################################################
# Input:
#
# KPIDocList => list of  "KPI document path + name", call
# reqs => eg ["REQ-0743467(A)|REQ-0743468(B)|REQ-0743473(A)"]
# refEC => Path of refEC path+name
# arch => either BSI or VSM
#
# returns:
#
# combined thematic lines as a list
########################################################################


def getCombinedThematicLines(kpiDocList, reqs, refEC, arch):
    logging.info("\n------------------------------")
    start_time = time.time()
    logging.info(f'GetCombinedThematics start time: {start_time}')
    # ARCH = "VSM"
    logging.info("getCombinedThematicLines + ")
    combination = []
    with open('../Thematics_Report.txt', 'a') as f:
        f.writelines("Thematics - \n \n")
    logging.info("reqsThem : ", reqs)
    for r in reqs:
        logging.info("r -- ", r)
        r.replace("||", "|")
        logging.info("rrr---->", r)
        with open('../Thematics_Report.txt', 'a') as f:
            f.writelines("\n\n\nCreating Thematics for one sheet--> \n")
        status, thematiqueList = searchDataInKPI(kpiDocList, r)
        logging.info("thematiqueList---------->",thematiqueList)
        if len(thematiqueList) != 0:
            # Combining all thematic lines
            (accept, acceptList) = createNewThmLines(thematiqueList, refEC, arch)
            logging.info("accept, acceptList --", accept, acceptList)
            for combination in accept:
                logging.info("comb ", combination)
                with open('../Thematics_Report.txt', 'a') as f:
                    f.writelines("\n\n Requirements clubbed together - ")
                with open('../Thematics_Report.txt', 'a') as f:
                    f.writelines("\n" + str(reqs))
                with open('../Thematics_Report.txt', 'a') as f:
                    f.writelines("\n\n Thematiques combinations formed" + str(combination))
            with open('../Thematics_Report.txt', 'a') as f:
                f.writelines("\n" + str(reqs))
        else:
            with open('../Thematics_Report.txt', 'a') as f:
                f.writelines("\n\nThematiques not present for any of the requirment")
        if combination is None:
            combination = -1

        end_time = time.time()
        execution_time = end_time - start_time
        logging.info(f'GetCombinedThematics end execution time: {execution_time}')
        logging.info("\n+++++++++++++++++++++")
        return status, combination


def removeInterfaceReq(ts_reqs):
    reqList = []
    finalReqsList = []
    finalReqs = ""
    logging.info("\nts_reqs: ", ts_reqs)
    for reqIDs in ts_reqs:
        splitReqs = reqIDs.split("|")
        logging.info(f"splitReqs {splitReqs}")
        for reqID in splitReqs:
            if reqID.find("DCI") == -1:
                reqList.append(reqID)
    logging.info(f"\nreqList {reqList}")
    if reqList:
        finalReqs = '|'.join(reqList)
        finalReqsList.append(finalReqs)
    logging.info(f"\nfinalReqs {finalReqs}")
    logging.info(f"\nfinalReqsList {finalReqsList}")
    return finalReqsList

def searchreqInKPI(docList,sheetList):
    # logging.info(f"In searchDataInKPI - {docList}")
    #logging.info("\n------------------------------")
    start_time = time.time()
    #logging.info(f'searchDataInKPI start time: {start_time}')

    status = True
    thematiqueList = []
    kpireqDict = {}
    tp=EI.openTestPlan()
    foundReq=[]
    try:


        for n, doc in enumerate(docList):
                with XwApp(visible=True) as app:
                    d = app.books.open(doc)
                    # sheetflag = 0
                    for sheet in d.sheets:
                        if sheet.name.find("REQ") != -1:
                            # maxCol = sheet.range('A10').end('right').last_cell.column
                            # #logging.info("maxCol - ", maxCol)
                            maxrow1 = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                            maxrow2 = sheet.range('B' + str(sheet.cells.last_cell.row)).end('up').row
                            sheet_value = sheet.used_range.value
                            if maxrow1 > maxrow2:
                                maxrow = maxrow1
                            else:
                                maxrow = maxrow2
                            for testSheet in sheetList:
                                testSh=tp.sheets[testSheet]
                                reqNotFound = []
                                reqVerNotFound = []
                                requirement=EI.getDataFromCell(testSh, (4, 3))
                                #req_list=requirement.split('|')
                                req_final=removeInterfaceReq([requirement])
                                req_final=req_final[0].split('|')
                                for r in req_final:
                                    if (r not in foundReq):
                                        with open('../Thematics_Report.txt', 'a') as f:
                                            f.writelines(
                                                "\n\n--------------------------------" + r + "--------------------------------\n")


                                        if r.find("REQ-") != -1:
                                            # logging.info("r = ", r)
                                            req = r.split("(")[0]
                                            ver = r.split("(")[1].split(")")[0]
                                        else:
                                            req = r
                                            ver = ""
                                        # logging.info("req , ver = ", req, ver)
                                        if req.find("REQ-") != -1:
                                            # result = searchDataInExcel(sheet, (maxrow, 2), req)
                                            result = searchDataInExcelCache(sheet_value, (maxrow, 2), req)
                                            # logging.info("Result =", result, result["cellPositions"])
                                            if result["count"] >= 1:
                                                # cellPositions = result["cellPositions"][0]
                                                foundVersion = False
                                                for count, cellPositions in enumerate(result["cellPositions"]):
                                                    row, col = cellPositions
                                                    version = EI.getDataFromCell(sheet, (row, col+1))
                                                    # version = EI.getDataFromCellCache(sheet_value, (row, col + 1))
                                                    logging.info("\n\nVersion in document - ", version,type(version))
                                                    logging.info("Version in ts - ", ver,"\n\n")
                                                    if type(version) is float:
                                                        version = str(int(version))
                                                        logging.info("Version in docc - ", version, "\n\n")
                                                    if ver == version:
                                                        # logging.info("fffff >> ")
                                                        foundVersion = True
                                                        foundReq.append(r)
                                                        break





                                                if foundVersion == False:
                                                    status = False
                                                    if n == (len(docList) - 1):
                                                       reqVerNotFound.append(r)


                                            else:
                                                # d.close()
                                                # app.kill()
                                                # time.sleep(2)
                                                if n == (len(docList) - 1):
                                                    # logging.info("Req not found in any sheet please check manually")
                                                    status = False
                                                    reqNotFound.append(r)
                                                    # logging.info("reqNotFound - ", reqNotFound)
                                                    effectiveExpression = "-1"
                                                    # un comment if need
                                                    # thematiqueList.append(effectiveExpression)
                                        else:
                                            # result = searchDataInExcel(sheet, (maxrow, 1), req)
                                            result = searchDataInExcelCache(sheet_value, (maxrow, 1), req)
                                            # logging.info("Result =", result, result["cellPositions"])

                                            if result["count"] == 0:
                                                if (req.find('.') != -1):
                                                    req = req.replace('.', '-')
                                                    # logging.info("req = ", req)
                                                    # result = searchDataInExcel(sheet, (maxrow, 1), req)
                                                    result = searchDataInExcelCache(sheet_value, (maxrow, 1), req)
                                            if result["count"] >= 1:
                                                # cellPositions = result["cellPositions"][0]
                                                foundReq.append(req)
                                            else:
                                                # d.close()
                                                # app.kill()
                                                # time.sleep(2)
                                                if n == (len(docList) - 1):
                                                    # logging.info("Req not found in any sheet please check manually")
                                                    status = False
                                                    req = r.split("(")[0]
                                                   # ver = r.split("(")[1].split(")")[0]
                                                    res = searchDataInExcelCache(sheet_value, (maxrow, 1), req)

                                                    #version = EI.getDataFromCell(sheet, (row, 3))
                                                    # logging.info("Result =", result, result["cellPositions"])
                                                    if res["count"] >= 1:
                                                        reqVerNotFound.append(r)
                                                    else:
                                                        reqNotFound.append(r)
                                #kpireqDict[testSheet]['foundReq']=kpireqDict[testSheet]['foundReq']+foundReq
                                if(len(reqNotFound)!=0):
                                  if(testSheet not in kpireqDict.keys()):
                                      kpireqDict[testSheet]={}
                                      kpireqDict[testSheet]['reqNotFound']=[]
                                      kpireqDict[testSheet]['reqNotFound'] = kpireqDict[testSheet]['reqNotFound'] + reqNotFound
                                  else:
                                    if 'reqNotFound' not in kpireqDict[testSheet].keys():
                                        kpireqDict[testSheet]['reqNotFound'] = []
                                        kpireqDict[testSheet]['reqNotFound'] = kpireqDict[testSheet]['reqNotFound']+reqNotFound
                                    else:
                                        kpireqDict[testSheet]['reqNotFound'] = kpireqDict[testSheet][ 'reqNotFound'] + reqNotFound
                                if len(reqVerNotFound) != 0:
                                    if testSheet not in kpireqDict.keys():
                                        kpireqDict[testSheet] = {}
                                        kpireqDict[testSheet]['reqVerNotFound']=[]
                                        kpireqDict[testSheet]['reqVerNotFound'] = kpireqDict[testSheet]['reqVerNotFound'] + reqVerNotFound
                                    else:
                                        if 'reqVerNotFound' not in kpireqDict[testSheet].keys():
                                            kpireqDict[testSheet]['reqVerNotFound'] = []
                                            kpireqDict[testSheet]['reqVerNotFound'] = kpireqDict[testSheet]['reqVerNotFound'] + reqVerNotFound
                                        else:
                                            kpireqDict[testSheet]['reqVerNotFound'] = kpireqDict[testSheet]['reqVerNotFound'] + reqVerNotFound


    except Exception as e:
        logging.info(f"Error .... {e}")
    # end_time = time.time()
    # execution_time = end_time - start_tim

    # logging.info(f'searchDataInKPI end execution time: ')
    # logging.info("\n------------------------------")
    logging.info(kpireqDict)
    UpdateHMIInfoCb(kpireqDict)




if __name__ == "__main__":
    # searchreqInKPI(["..//Input_Files//KPI//EACE_KPI_FSEE_DESP_SPA_00040 (1).xlsx","..//Input_Files//KPI//EACE_KPI_Template_FSEE_SBR.xlsx"],['VSM20_N1_20_11_0021','VSM20_N1_20_11_0025E','VSM20_N1_20_11_0012'])

    kpiDocList = ["..//Input_Files//KPI//EACE_KPI_FSEE_DESP_SPA_00040 (1).xlsx",
     "..//Input_Files//KPI//EACE_KPI_Template_FSEE_SBR.xlsx"]
    requirement = 'REQ-0736041(A)|REQ-DIAG-DTC-DESP-436(0)|REQ-0786520(A)|DCINT-00055480(1)|DCINT-00055482(1)|DCINT-00055488(1)|DCINT-00055492(1)|DCINT-00055494(1)|DCINT-00055496(1)'
    searchDataInKPI(kpiDocList, requirement)