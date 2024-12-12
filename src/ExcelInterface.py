import sys

import xlwings as xw
import re
import os
import InputConfigParser as ICF
import KeyboardMouseSimulator as KMS
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import time
import logging


def openExcel(book):
    return xw.Book(book)


def openExcelInBackground(book):
    excel_app = xw.App(visible=False)
    return excel_app.books.open(book)


def activateSheet(book, sheetId):
    sheet = book.sheets[sheetId]
    sheet.activate()


def activateSheetObj(sheet):
    sheet.activate()


def getDataFromCell(sheet, colRow):
    return sheet.range(colRow).value


def setDataFromCell(sheet, colRow, value):
    sheet.range(colRow).value = value


def setDataInCell(sheet, colRow, value):
    sheet.range(colRow).value = value


def searchDataInExcel(sheet, keyword):
    value = sheet.used_range.value
    # searchResult = {
    #     "count": 0,
    #     "cellPositions": [],
    #     "cellValue": []
    # }
    if keyword == "":
        return 0
    if value is None:
        return 0
    # x is the index of column
    # i is the value of column
    # y is the index of row
    # j is the value of cell
    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if j is not None:
                if keyword in str(j):
                    # searchResult["count"] = searchResult["count"] + 1
                    # searchResult["cellPositions"].append((x + 1, y + 1))
                    # searchResult["cellValue"].append(j)
                    return 1

    return 0


def getRangeforAnalyseDeEntrant(sheet, cell=0, direction=0):
    return sheet.range('D1').end('down').row, sheet.range('A1').end('right').column

def searchDataInCol1(sheet, cellRange, keyword):
    x, y = cellRange
    count = 0
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    path_to_sheet = sheet.book.fullname
    sheetName = sheet.name
    futures = []
    finaloutput = []
    start = 1
    with ThreadPoolExecutor(max_workers=x) as exe:
        if x > 100:
            for i in myRange(0, x, 100):
                if i > 1:
                    end = i
                    futures.append(exe.submit(threadFunCol, path_to_sheet, sheetName, (start, end, y), keyword))
                start = i + 1
        else:
            futures.append(exe.submit(threadFunCol, path_to_sheet, sheetName, (1, x, y), keyword))
        for future in concurrent.futures.as_completed(futures):
            finaloutput.append(future.result())
            dic1 = future.result()
            searchResult = mergeDict(searchResult, dic1)
    return searchResult


def searchDataInCol(sheet, specfCol, keyword, matchCase=False):
    value = sheet.used_range.value
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    if keyword == "":
        return searchResult
    # x is the index of column
    # i is the value of column
    # y is the index of row
    # j is the value of cell
    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if y == specfCol - 1:
                if j is not None:
                    # logging.info("j --- ", j)
                    if matchCase == True:
                        if keyword.lower() in str(j).lower():
                            searchResult["count"] = searchResult["count"] + 1
                            searchResult["cellPositions"].append((x + 1, y + 1))
                            searchResult["cellValue"].append(j)
                    else:
                        if keyword in str(j):
                            searchResult["count"] = searchResult["count"] + 1
                            searchResult["cellPositions"].append((x + 1, y + 1))
                            searchResult["cellValue"].append(j)
    return searchResult

def searchDataInColCache(value, specfCol, keyword, matchCase=False):
    # value = sheet.used_range.value
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    if keyword == "":
        return searchResult
    # x is the index of column
    # i is the value of column
    # y is the index of row
    # j is the value of cell
    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if y == specfCol - 1:
                if j is not None:
                    # logging.info("j --- ", j)
                    if matchCase == True:
                        if keyword.lower() in str(j).lower():
                            searchResult["count"] = searchResult["count"] + 1
                            searchResult["cellPositions"].append((x + 1, y + 1))
                            searchResult["cellValue"].append(j)
                    else:
                        if keyword in str(j):
                            searchResult["count"] = searchResult["count"] + 1
                            searchResult["cellPositions"].append((x + 1, y + 1))
                            searchResult["cellValue"].append(j)
    return searchResult


def searchDataInExcel(sheet, cellRange, keyword):
    value = sheet.used_range.value
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    if keyword == "":
        return searchResult
    # x is the index of column
    # i is the value of column
    # y is the index of row
    # j is the value of cell
    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if j is not None:
                if keyword in str(j):
                    searchResult["count"] = searchResult["count"] + 1
                    searchResult["cellPositions"].append((x + 1, y + 1))
                    searchResult["cellValue"].append(j)

    return searchResult


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


def searchDataInSpecificRows(sheet, rowRange, col, keyword):
    x1, x2 = rowRange
    y = col
    count = 0
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    if keyword == "":
        return searchResult

    for row in range(x1, x2 + 1):
        cellValue = str(sheet.range(row, y).value)
        sheetName = str(sheet) + "\n"
        if keyword in cellValue:
            searchResult["cellPositions"].append(tuple((row, y)))
            searchResult["cellValue"].append(cellValue)
            count = count + 1
    searchResult["count"] = count
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
        sheet = openExcel(path_to_sheet).sheets[sheetName]
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


def searchDataInSheet(sheet, cellRange, keyword):
    x, y = cellRange
    count = 0
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    if keyword == "":
        return searchResult

    for row in range(1, y + 1):
        for col in range(1, x + 1):
            cellValue = str(sheet.range(row, col).value)
            sheetName = str(sheet) + "\n"
            if keyword in cellValue:
                searchResult["cellPositions"].append(tuple((row, col)))
                searchResult["cellValue"].append(cellValue)
                count = count + 1
    searchResult["count"] = count
    return searchResult



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


def threadFunCol(path_to_sheet, sheetName, cellRange, keyword):
    time.sleep(1)
    # logging.info("In threadFunCol function sheetname", sheetName)
    start, end, y = cellRange
    # logging.info("searching data in col", start, end, y)
    c = 0
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    try:
        sheet = openExcel(path_to_sheet).sheets[sheetName]
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


def searchRequirementInDCI(sheet, cellRange, keyword):
    sheet_value = sheet.used_range.value
    # coords = searchDataInExcel(sheet, cellRange, keyword)
    coords = searchDataInExcelCache(sheet_value, cellRange, keyword)
    return coords


# function to return range of all different types of requiremnts
def getRequirementRange(sheet, req_type):
    rowx = 1
    i = 0
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    while (rowx < maxrow):

        if sheet.range(rowx, 1).merge_cells is True:
            rlo = sheet.range(rowx, 1).merge_area.row
            rhi = sheet.range(rowx, 1).merge_area.last_cell.row
            if ((rhi - rlo > 1) & (rhi > 3)):  # create variable for 3
                i = i + 1
                rowx = rhi
                if i == req_type:
                    break
        rowx = rowx + 1
    return (rlo, rhi)


def getReqNames(sheet):
    rqNames = []
    for i in range(1, 5):
        rlo, rhi = getRequirementRange(sheet, i)
        for j in range(rlo, rhi + 1):
            if sheet.range(j, 1).value is not None:
                rqNames.append(sheet.range(j, 1).value)
    return rqNames


def getRequirementTypes(sheet, fepsID, rqID):
    rqType = []
    rlo, rhi = getRequirementRange(sheet, rqID)
    for j in range(rlo, rhi + 1):
        cond = sheet.range(j, fepsID).value is not None
        if type(sheet.range(j, fepsID).value) is str:
            cond = sheet.range(j, fepsID).value is not None and sheet.range(j, fepsID).value.upper().find('REQUIREMENTS') == -1 and sheet.range(j, fepsID).value.upper().find('REQUIREMENT') == -1
        if cond:
            rqType.append(sheet.range(j, fepsID).value)
    logging.info(f"rqType {rqType}")

    return rqType



def getSignalExtension(signal, dciSignal):
    ext = ""
    logging.info(f"dciSignal - signal {dciSignal} - {signal}")
    if len(dciSignal) != len(signal):
        signalSplit = signal.split("_")
        ext = signalSplit[-1]
        network = slice(3)
        return ext[network]
    else:
        return ext


def isExtNwOrPc(ext):
    nwpc = True
    if ext == "P" or ext == "p" or ext == "C" or ext == "c":
        nwpc = False
    return nwpc


def getDciInfo(dciBook, requirement):
    requirement = requirement.strip()
    dciInfo = {
        "dciSignal": "",
        "network": "",
        "pc": "",
        "thm": "",
        "framename": "",
        "dciReq": "",
        "proj_param": "",
        "dciThematic": ""
    }
    # maxrow = dciBook.sheets['MUX'].range('A' + str(dciBook.sheets['MUX'].cells.last_cell.row)).end('up').row
    # logging.info(maxrow)
    for sheet in dciBook.sheets:
        logging.info("Searching in Sheet...")
        logging.info(sheet)
        # logging.info("sheet name =*" + sheet.name + "*")

        if sheet.name.strip() == "MUX":
            maxrow = (dciBook.sheets['MUX'].range('A' + str(dciBook.sheets['MUX'].cells.last_cell.row)).end(
                'up').row)
            sheet_value = sheet.used_range.value
            logging.info(maxrow)
            logging.info(
                "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            # searchResult = searchDataInExcel(sheet, (26, maxrow), requirement)
            searchResult = searchDataInExcelCache(sheet_value, (26, maxrow), requirement)
            # searchResult = searchDataInExcel(sheet,requirement)
            if searchResult["count"] > 0:
                logging.info("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Success!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                for cellPosition in searchResult["cellPositions"]:
                    logging.info("getSignal")
                    x, y = cellPosition
                    logging.info(f"123yyyyyyy {y}")
                    dciInfo["dciSignal"] = "$" + (str(getDataFromCell(sheet, (x, y + 2))))
                    dciInfo["network"] = (str(getDataFromCell(sheet, (x, y + 9))))
                    dciInfo["pc"] = (str(getDataFromCell(sheet, (x, y + 10))))
                    dciInfo["thm"] = (str(getDataFromCell(sheet, (x, y + 15))).encode('utf-8').strip())
                    dciInfo["framename"] = (str(getDataFromCell(sheet, (x, y + 9)))) + "/" + (
                        str(getDataFromCell(sheet, (x, y + 8)))) + "/" + (str(getDataFromCell(sheet, (x, y + 7))))
                    dciInfo["dciReq"] = (str(getDataFromCell(sheet, (x, 1))))
                    logging.info(dciInfo)
                    dciInfo["proj_param"] = (str(getDataFromCell(sheet, (x, y + 3))))
                    dciInfo["dciThematic"] = (str(getDataFromCell(sheet, (x, 17))).encode('utf-8').strip())
                    logging.info(dciInfo)
                    break
        if sheet.name.strip() == "FILAIRE":
            maxrow = (dciBook.sheets['FILAIRE'].range('A' + str(dciBook.sheets['FILAIRE'].cells.last_cell.row)).end(
                'up').row)
            logging.info(maxrow)
            logging.info(
                "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            # searchResult = searchDataInExcel(sheet, (26, maxrow), requirement)
            searchResult = searchDataInExcelCache(sheet_value, (26, maxrow), requirement)
            # searchResult = searchDataInExcel(sheet, requirement)
            if searchResult["count"] > 0:
                logging.info("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Success!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                for cellPosition in searchResult["cellPositions"]:
                    logging.info("getSignal")
                    x, y = cellPosition
                    logging.info(f"123yyyyyyy {y}")
                    dciInfo["dciSignal"] = "$" + (str(getDataFromCell(sheet, (x, y + 2))))
                    dciInfo["network"] = "FIL"
                    dciInfo["pc"] = (str(getDataFromCell(sheet, (x, y + 9))))
                    dciInfo["thm"] = (str(getDataFromCell(sheet, (x, y + 18))).encode('utf-8').strip())
                    logging.info(dciInfo)
                    dciInfo["dciReq"] = (str(getDataFromCell(sheet, (x, 1))))
                    dciInfo["proj_param"] = (str(getDataFromCell(sheet, (x, y + 3))))
                    dciInfo["dciThematic"] = (str(getDataFromCell(sheet, (x, 19))).encode('utf-8').strip())
                    break
        if dciInfo["dciSignal"] != "":
            break
    return dciInfo


def findTestSheet(book, dciInfo):
    sh1 = []
    sh2 = []
    paramLen = []

    dciSignal = dciInfo["dciSignal"]
    producedConsumed = dciInfo['pc']
    logging.info("Dci Signal & its length = ", dciSignal, len(dciSignal))
    for sheet in book.sheets:
        sheet_value = sheet.used_range.value
        if sheet.visible:
            if "VSM" or "BSI" in sheet.name:
                if (sheet.name.find("_0000") == -1):
                    if (sheet.name.find("_SF_") == -1):
                        maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                        if producedConsumed == 'C' or producedConsumed == 'c':
                            # searchResult = searchDataInExcel(sheet, (maxrow, 5), dciSignal)
                            searchResult = searchDataInExcelCache(sheet_value, (maxrow, 5), dciSignal)
                        else:
                            # searchResult = searchDataInExcel(sheet, (maxrow, 11), dciSignal)
                            searchResult = searchDataInExcelCache(sheet_value, (maxrow, 11), dciSignal)
                        # logging.info(f"searchResult THM DCI {searchResult}")
                        if searchResult["count"] == 0:
                            # logging.info("No match in current sheet" + str(sheet))
                            pass
                        else:
                            if sheet not in sh1:
                                sh1.append(sheet)
    if len(sh1) == 1:
        logging.info("sheet 1 = ", sh1)
        # return sh1[0]
        return sh1
    elif len(sh1) == 0:
        logging.info(" GG__GG")
        return -1
    else:
        logging.info("sheet 1 = ", sh1)
        for sheet in sh1:
            last_index = -1
            maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            sheet_value = sheet.used_range.value
            paramValue = []
            if producedConsumed == 'C' or producedConsumed == 'c':
                # searchResult = searchDataInExcel(sheet, (maxrow, 5), dciSignal)
                searchResult = searchDataInExcelCache(sheet_value, (maxrow, 5), dciSignal)
            else:
                # searchResult = searchDataInExcel(sheet, (maxrow, 11), dciSignal)
                searchResult = searchDataInExcelCache(sheet_value, (maxrow, 11), dciSignal)
            logging.info(searchResult['cellValue'])
            for cellValue in searchResult['cellValue']:
                signalExtension = getSignalExtension(cellValue, dciSignal)
                logging.info("getSignalExtension(cellValue) = ", signalExtension)
                logging.info("dciInfo[network] = " + dciInfo["network"])
                if signalExtension.lower() in dciInfo["network"].lower():
                    logging.info("sheet = ", sheet.name)
                    logging.info("cellvalue = ", cellValue)
                    logging.info("cellvalue length = ", searchResult["cellValue"].index(cellValue),
                          len(searchResult["cellValue"]))
                    # logging.info("cellvalue index = ", searchResult["cellValue"].index(cellValue, last_index+1))
                    if last_index != len(searchResult["cellValue"]) - 1:
                        last_index = searchResult["cellValue"].index(cellValue, last_index + 1)
                    else:
                        last_index = len(searchResult["cellValue"]) - 1
                    logging.info("last index = ", last_index)
                    cellpos = searchResult["cellPositions"][last_index]
                    x, y = cellpos
                    try:
                        logging.info("sheet.range(x, y+1)", x, y, str(int(float(sheet.range(x, y + 1).value))))
                        logging.info("sheet.range(x, y+1)", x, y, sheet.range(x, y + 1).value)
                        if (str(int(float(sheet.range(x, y + 1).value)))) not in paramValue:
                            paramValue.append(str(int(float(sheet.range(x, y + 1).value))))
                        if sheet not in sh2:
                            sh2.append(sheet)
                    except:
                        logging.info("sheet.range(x, y+1)", x, y, str(sheet.range(x, y + 1).value))
                        if (str(sheet.range(x, y + 1).value)) not in paramValue:
                            paramValue.append(str(sheet.range(x, y + 1).value))
                        if sheet not in sh2:
                            sh2.append(sheet)
            logging.info("Values present in testsheet", sheet, paramValue)
            if len(paramValue) != 0:
                paramLen.append(len(paramValue))
        logging.info("length", len(sh2), sh2, len(paramLen), paramLen)
        if len(sh2) == 1:
            # return sh2[0]
            logging.info("  *&*&*&*&  ")
            return sh2
        elif len(sh2) == 0:
            return -1
        else:
            # return sh2[paramLen.index(max(paramLen))]
            logging.info("   &&&  ")
            return sh2


# Find DCI signal present in test sheet
def findInterfaceReqSignal(tpBook, dciInfo):
    dciSignal = dciInfo["dciSignal"]
    produced_consumed = dciInfo['pc']
    sheet_list = []
    logging.info("Dci Signal & its length = ", dciSignal, len(dciSignal))
    for sheet in tpBook.sheets:
        if ("VSM" or "BSI" in sheet.name) and sheet.visible and sheet.name.find("_0000") == -1 and sheet.name.find("_SF_") == -1:
            maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            if produced_consumed.lower() == 'c':
                sheet_input_col = 5
                # signalSearchResult = searchDataInCol(sheet, sheet_input_col, dciSignal)
                sheet_value = sheet.used_range.value
                signalSearchResult = searchDataInColCache(sheet_value, sheet_input_col, dciSignal)
            else:
                sheet_output_col = 11
                # signalSearchResult = searchDataInCol(sheet, sheet_output_col, dciSignal)
                sheet_value = sheet.used_range.value
                signalSearchResult = searchDataInColCache(sheet_value, sheet_output_col, dciSignal)

            if signalSearchResult['count'] > 0:
                for ind, sig_coord in enumerate(signalSearchResult['cellPositions']):
                    row, col = sig_coord
                    # logging.info(f"sig_coord>> {sig_coord}")
                    if signalSearchResult['cellValue'][ind].strip() == dciSignal.strip():
                        if sheet not in sheet_list:
                            sheet_list.append(sheet)

    return sheet_list


def getNewThematics(thm, val, refEC):
    logging.info("In getNewThematics function", thm, val)
    sheet = refEC.sheets['Liste EC']
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    sheet_value = sheet.used_range.value
    # logging.info(maxrow)
    # searchResults = searchDataInExcel(sheet, (5, maxrow), thm)
    searchResults = searchDataInExcelCache(sheet_value, (5, maxrow), thm)
    logging.info("searchResults in getNewThematics function", searchResults)
    if searchResults["count"] > 0:
        for cellPosition in searchResults["cellPositions"]:
            x, y = cellPosition
            # logging.info("pos",sheet.range(x,y+1).value,x,y)
            if sheet.range(x, y + 1).value is not None:
                # logging.info("val",sheet.range(x,y+1).value,val)
                if sheet.range(x, y + 1).value == val:
                    return sheet.range(x, y + 2).value
                else:
                    result = -1
    else:
        result = -1
    return result


def getSignalInitValue(dciBook, tpBook, ssFiches, signal, value, functionName):
    for sheet in dciBook.sheets:
        if (sheet.name.strip() == "MUX") or (sheet.name.strip() == "FILAIRE"):
            logging.info("Searching in Sheet...")
            maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            sheet_value = sheet.used_range.value
            # logging.info(maxrow)
            # logging.info(sheet)
            # searchResult = searchDataInExcel(sheet, (maxrow, 3), signal)
            searchResult = searchDataInExcelCache(sheet_value, (maxrow, 3), signal)
            # logging.info("result--", searchResult)
            # logging.info("dci", signal)
            if searchResult["count"] > 0:
                logging.info("!....Success....!")
                for cellPosition in searchResult["cellPositions"]:

                    x, y = cellPosition
                    if (getDataFromCell(sheet, (x, y - 1))) == "VSM" or (getDataFromCell(sheet, (x, y - 1))) == "BSI":

                        Info = (str(getDataFromCell(sheet, (x, y + 4))))
                        # logging.info("Info",Info)
                        splitted = Info.split("\n")
                        for i in splitted:
                            if i.find("InitValue") != -1:
                                if i.find('=') != -1:
                                    a = i.split("=")[1]
                                    try:
                                        # logging.info("Initial Value = ", signal, a)
                                        # logging.info("Init Result",int(a,2))
                                        return (str(int(a, 2)))
                                    except:
                                        pass


def getSignalValueInput(dciBook, ssFiches, signal, value, functionName):
    for sheet in dciBook.sheets:
        logging.info("1", sheet, sheet.name.strip())
        if (sheet.name.strip() == "MUX") or (sheet.name.strip() == "FILAIRE"):
            logging.info("2", sheet, sheet.name.strip())
            logging.info("Searching in Sheet...")
            maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            sheet_value = sheet.used_range.value
            # logging.info(maxrow)
            # logging.info(sheet)
            # searchResult = searchDataInExcel(sheet, (maxrow, 3), signal)
            searchResult = searchDataInExcelCache(sheet_value, (maxrow, 3), signal)
            # logging.info("result--", searchResult)
            # logging.info("dci", signal)

            if searchResult["count"] > 0:
                logging.info("!....Success....!")
                for cellPosition in searchResult["cellPositions"]:
                    p = 0
                    x, y = cellPosition
                    logging.info("-------------", x, y, getDataFromCell(sheet, (x, y - 1)))
                    if (getDataFromCell(sheet, (x, y - 1))) == "VSM" or (getDataFromCell(sheet, (x, y - 1))) == "BSI":
                        producedConsumed = (getDataFromCell(sheet, (x, y + 8)))
                        if producedConsumed == 'C' or producedConsumed == 'c':
                            Info = (str(getDataFromCell(sheet, (x, y + 4))))
                            # logging.info("Info.....", Info)
                            splitted = Info.split("\n")
                            for i in splitted:
                                if value.lower() in i.lower():
                                    a = i.split("=")[1]
                                    logging.info(int(a, 2), producedConsumed)
                                    return (str(int(a, 2)))

                        else:

                            p = 1
                if p == 1:
                    logging.info("Produced signal", signal)
                    return -1
            else:
                logging.info("Signal not foundin Global Dci", signal)
                return -1


def getSignalValueOutput(dciBook, signal, value):
    for sheet in dciBook.sheets:
        if (sheet.name.strip() == "MUX") or (sheet.name.strip() == "FILAIRE"):
            logging.info("Searching in Sheet...")
            maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            sheet_value = sheet.used_range.value
            # searchResult = searchDataInExcel(sheet, (maxrow, 3), signal)
            searchResult = searchDataInExcelCache(sheet_value, (maxrow, 3), signal)
            logging.info("dci", signal)

            if searchResult["count"] > 0:
                logging.info("!....Success....!")
                for cellPosition in searchResult["cellPositions"]:
                    c = 0
                    x, y = cellPosition
                    if (getDataFromCell(sheet, (x, y - 1))) == "VSM" or (getDataFromCell(sheet, (x, y - 1))) == "BSI":
                        producedConsumed = (getDataFromCell(sheet, (x, y + 8)))
                        if producedConsumed == 'P' or producedConsumed == 'p':
                            Info = (str(getDataFromCell(sheet, (x, y + 4))))
                            splitted = Info.split("\n")
                            for i in splitted:
                                if value.lower() in i.lower():
                                    a = i.split("=")[1]
                                    logging.info(int(a, 2), producedConsumed)
                                    return (int(a, 2))
                        else:

                            c = 1
                if c == 1:
                    logging.info("Consumed signal", signal)
                    return -1
            else:
                logging.info("Signal not found in Global Dci", signal)
                return -1


# Function to get thematiques present in particular sheet


def testPlanReference():
    pass


def fillSummaryTab():
    pass


def fillImpactTab():
    pass


def addRequirement():
    pass


def findInputFiles():
    # logging.info("findInputFiles() +--->",os.listdir(ICF.getInputFolder()))
    arr = os.listdir(ICF.getInputFolder())
   # UI_task_nameEI = ICF.FetchTaskName()
    #taskname = UI_task_nameEI.split('_')[1]
    # logging.info("Documents present in input folder - ", arr)
    DCI = []
    PT = ""
    Analyse_de_entrant = []
    globalDCI = ""
    referentialEC = ""
    siAlert = ""
    ssFiches = ""
    param_global = ""
    traceability = ""
    ficher = ""
    qia_Param = ""
    MagnetoFrame = ""
    EEAD = ""
    DID_File = ""
    DNF = ""
    Analyse_DNF_kpi = ""
    QIA_Sous_Fiches = ""
    QIA = ""
    Docx = []
    CAMPAGNE = ""
    CheckList_Final_Validation = " "
    Silhouette = ""
    PLM_EE = ""
    qia_sheet_pattern = re.compile(
        r'QIA_[A-Za-z0-9]{5}_[A-Za-z0-9]{2}_[A-Za-z0-9]{5}_PARAM_GLOBAL', re.IGNORECASE)
    for i in arr:
        if i.find('DCI') != -1:
            DCI.append(i)
        elif i.find('Analyse') != -1 and i.find('~$') == -1 and i.find('Analyse_DNF_kpi') == -1:
            Analyse_de_entrant.append(i)
        elif i.find('Tests') != -1 and i.find('fiches') == -1 and i.find('~$') == -1:
            PT = i
            # logging.info("Testplan file name - ", PT)
        # elif i.find('DCI_Global') != -1:
        elif i.upper().find('DCI_GLOBAL') != -1:
            globalDCI = i
        elif (i.find('Referentiel_EC') != -1 or i.find('Referential_EC') != -1 or i.find('EC_Referential') != -1 or i.find('EC_Referentiel') != -1) and i.find('~$') == -1:
            referentialEC = i
        elif i.find('SI_ALERT') != -1:
            siAlert = i
        elif i.find('fiches') != -1 and i.find('QIA_Sous_Fiches') == -1:
            ssFiches = i
        elif i.find('PARAM_Global') != -1 and i.find('~$') == -1:
            param_global = i
        elif i.find('traceability') != -1 and i.find('~$') == -1:
            traceability = i
        elif i.find('Fichier') != -1 and i.find('~$') == -1:
            ficher = i
        elif len(re.findall(qia_sheet_pattern, i)):
            qia_Param = i
        elif i.find('MagnetoFrame') != -1 and i.find('~$') == -1:
            MagnetoFrame = i
        elif i.find('EEAD') != -1 and i.find('~$') == -1:
            EEAD = i
        elif i.upper().find('DID') != -1 and i.find('~$') == -1:
            DID_File = i
        elif i.upper().find('DNF') != -1 and i.find('~$') == -1 and i.find('Analyse_DNF_kpi') == -1:
            DNF = re.findall(r"(DNF_[A-Za-z]{2,7}|[A-Za-z]{2,7}_DNF)\.xlsx", i)[0] if len(re.findall(r"(DNF_[A-Za-z]{2,7}|[A-Za-z]{2,7}_DNF)\.xlsx", i)) > 0 else ""
        elif i.find('Analyse_DNF_kpi') != -1 and i.find('~$') == -1:
            Analyse_DNF_kpi = i
        elif i.find('QIA_Sous_Fiches') != -1 and i.find('~$') == -1:
            QIA_Sous_Fiches = i
        elif (os.path.splitext(i)[1]==".docx") and i.find('~$')==-1:
            Docx.append(i)
        elif i.find('QIA') != -1 and i.find('~$') == -1:
            QIA = i
        elif i.find('CAMPAGNE') != -1 and i.find('~$') == -1:
            CAMPAGNE = i
        elif i.find('CheckList_Final_Validation') != -1 and i.find('~$') == -1:
            CheckList_Final_Validation = i
        elif i.find('Silhouette') != -1 and i.find('~$') == -1:
            Silhouette = i
        elif i.find('PLM_EE') != -1 and i.find('~$') == -1:
            PLM_EE = i
    return [DCI, PT, Analyse_de_entrant, globalDCI, referentialEC, siAlert, ssFiches, param_global, traceability, ficher, qia_Param, MagnetoFrame, EEAD, DID_File, DNF, Analyse_DNF_kpi, QIA_Sous_Fiches, Docx, QIA, CAMPAGNE, CheckList_Final_Validation, Silhouette, PLM_EE]


def remove_dupiclates(li_st):
    new_list = []
    for a in li_st:
        if a not in new_list:
            new_list.append(a)
    return new_list


def remove_Duplicates_C4(testbook):
    for sheet in testbook.sheets:
        if (sheet.name.find('VSM') != -1 or sheet.name.find('BSI') != -1) and sheet.name.find('0000') == -1 and sheet.visible:
            # logging.info('sheet ==>>', sheet)
            test = sheet.name
            req = testbook.sheets[test].range('C4').value
            if req is not None and req != "":
                req_list = req.split('|')
                req_sorted_list = [string for string in req_list if string.strip()]
                req_set = remove_dupiclates(req_sorted_list)
                req_final = [elem + '|' for elem in req_set[:-1]] + [req_set[-1]]
                # logging.info('new_list---0--->', req_final)
                req_final_string = ''.join(req_final)
                # logging.info('req_final_string--0-->', req_final_string)
                testbook.sheets[test].api.Unprotect()
                testbook.sheets[test].range('C4').value = req_final_string
                after_adding = testbook.sheets[test].range('C4').value
                # logging.info('afteradding----0----->', after_adding)


def openAnalyseDeEntrant(taskname):
    # Open Excel
    Analyse_de_entrant = findInputFiles()[2]
    logging.info("analyse = ", Analyse_de_entrant, len(Analyse_de_entrant))
    with open('../Aptest_Tool_Report.txt', 'a') as f:
        f.writelines("\n\nanalyse = " + str(Analyse_de_entrant))
    if len(Analyse_de_entrant) != 0:
        for analyseSheet in Analyse_de_entrant:
            logging.info("Sheets=========>", analyseSheet)
            analyseDeEntrant = openExcel(ICF.getInputFolder() + "\\" + analyseSheet)
            logging.info("analyse = ", analyseDeEntrant.sheets)
            logging.info("Task Name = ", taskname)
            for sheet in analyseDeEntrant.sheets:
                logging.info("Sheet name = ", sheet.name)
                if taskname in sheet.name:
                    # Load Macro
                    logging.info("Task Name = ", taskname)
                    openExcel(ICF.getTestPlanMacro())
                    return analyseDeEntrant
            analyseDeEntrant.close()
        return -1
    else:
        logging.info("No analyse de entrant")
        return -1


def openTestPlan():
    PT = findInputFiles()[1]
    if len(PT) != 0:
        testPlan = openExcel(ICF.getInputFolder() + "\\" + PT)
        # Load Macro
        openExcel(ICF.getTestPlanMacro())
        getTestPlanAutomationMacro()
        return testPlan
    else:
        logging.info("No testplan")
        return -1


def openAna():
    PT = findInputFiles()[2]
    if len(PT) != 0:
        testPlan = openExcel(ICF.getInputFolder() + "\\" + PT)
        # Load Macro
        openExcel(ICF.getTestPlanMacro())
        getTestPlanAutomationMacro()
        return testPlan
    else:
        logging.info("No testplan")
        return -1


def openReferentialEC():
    referential = findInputFiles()[4]
    referentialEC = openExcel(ICF.getInputFolder() + "\\" + referential)
    return referentialEC


def openAlertDoc():
    alert = findInputFiles()[5]
    logging.info(ICF.getInputFolder() + "\\" + alert)
    siALert = openExcel(ICF.getInputFolder() + "\\" + alert)
    return siALert


def openGlobalDCI():
    # logging.info("Input Files", findInputFiles())
    dciBook = ""
    globalDCI = ""
    try:
        for i in findInputFiles():
            for j in i:
                if j.upper().find('DCI_GLOBAL') != -1:
                    globalDCI = j
        if globalDCI != "" and globalDCI is not None:
            dciBook = openExcel(ICF.getInputFolder() + "\\" + globalDCI)
    except Exception as e:
        logging.info(f"Error in opening the global DCI: {e}")

    return dciBook

def openDCIForInterface():
    dciExcel = openExcel(ICF.getInputFolder() + "\\" + findInputFiles()[0][0])
    return dciExcel

def openDNFKPI():
    dnfKPI = openExcel(ICF.getInputFolder() + "\\" + findInputFiles()[14])
    logging.info("dnfKPI->",dnfKPI)
    return dnfKPI

def openAnaDNF():
    analyKpi = openExcel(ICF.getInputFolder() + "\\" + findInputFiles()[15])
    return analyKpi

def openParamGlobal():
    paramGlobal = openExcel(ICF.getInputFolder() + "\\" + findInputFiles()[7])
    return paramGlobal


def openDCIExcel(DCIdoc):
    doc = findInputFiles()[0]
    logging.info("DCI = ", doc)
    pattern = r'([_|\s][vV]{1}[0-9]{1,2}.[0-9]{1,2})|([_|\s][vV]{1}[0-9]{1,2})'

    dciSheetName = re.split(pattern, DCIdoc[0])

    if re.search(" [vV]{1}[0-9]{1}.[0-9]{1}", DCIdoc[0].split(" ")[0]):
        # logging.info("found 2", str(re.findall("[vV]{1}[0-9]{1}.[0-9]{1}", DCIdoc.split(" ")[0])))
        DCIdocument = re.sub("[vV]{1}[0-9]{1}.[0-9]{1}", "", DCIdoc[0])
    elif re.search("[vV]{1}[0-9]{1}", DCIdoc[0].split(" ")[0]):
        # logging.info("found", str(re.findall("[vV]{1}[0-9]{1}", DCIdoc.split(" ")[0])))
        DCIdocument = re.sub("[vV]{1}[0-9]{1}", "", DCIdoc[0])
    else:
        DCIdocument = DCIdoc[0]

    for dci in doc:
        logging.info("in for loop dci = ", dci)
        logging.info("Comparing Before Open DCI")
        logging.info("Compare 1 From Folder", dci.split(".x")[0])
        logging.info("Compare 2 From Analyse de Entrant", DCIdocument.split(" ")[0])
        logging.info("Compare 3 From Analyse de Entrant", dciSheetName)
        if dci.split(".x")[0] in dciSheetName[0]:
            logging.info("dci = ", dci)
            dciExcel = openExcel(ICF.getInputFolder() + "\\" + dci)
            logging.info("DCI opened for interface")
            return dciExcel


def findSheetInBook(wBook, sheetName):
    logging.info(f"wBook {wBook}")
    for sheet in wBook.sheets:
        if sheet.name.lower() == sheetName.lower():
            return 1, sheetName
    return -1, sheetName


def openSousFiches():
    sousFiches = findInputFiles()[6]
    logging.info("sousFiches +++", sousFiches)
    sFiches = openExcel(ICF.getInputFolder() + "\\" + sousFiches)
    # Load Macro
    # EI.openExcel(ICF.getTestPlanMacro())
    openExcel(ICF.getTestPlanMacro())
    getTestPlanAutomationMacro()
    return sFiches


# def setTaskName(taskname):
#     globals()["taskname"] = taskname
#
# def getTaskName(key):
#     return globals()[key]


# taking the thematic lines which present in sheet
# single sheet object is input
def getTestSheetThematics(testSheet):
    ts_thematics = []
    testSheet_value = testSheet.used_range.value
    # thematic_result = searchDataInExcel(testSheet, (1, 100), "THEMATIQUE")
    thematic_result = searchDataInExcelCache(testSheet_value, (1, 100), "THEMATIQUE")
    if thematic_result['count'] > 0:
        for thm_coord in thematic_result['cellPositions']:
            row, col = thm_coord
            # proceed only the THEMATIQUE keyword present in 1st column of the test sheet
            if col == 1:
                them_val_col = 3
                thm_val = testSheet.range(row, them_val_col).value
                if thm_val != "" and thm_val is not None and thm_val != '--':
                    ts_thematics.append(thm_val)
    logging.info(f"ts_thematics {ts_thematics}")

    return ts_thematics


def getTestPlanAutomationMacro():
    return openExcel("../Macro/MacroAutomation.xlam")
