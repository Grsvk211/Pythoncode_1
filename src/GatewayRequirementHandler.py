import logging
# Processing the Gateways requirements
# Requirement text file exist under the below svn path
# svn.in.expleogroup.com/!/#AppStore/view/head/Tags/BetaVersion1.0/Tool_Requirement/Gateway_Requirement.txt
import time

import InputConfigParser as ICF
import ExcelInterface as EI
import TestPlanMacros as TPM
import re
import xlwings as xw
import os

UpdateHMIInfoCb = None


def registerInfoTextBoxGateWay(func):
    global UpdateHMIInfoCb
    UpdateHMIInfoCb = func

def saveTestPlan(tpBook):
    output_dir = os.path.abspath(r"..\Output_Files")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    if not os.path.exists(r"..\Output_Files\GatewayRequirement"):
        os.makedirs(r"..\Output_Files\GatewayRequirement")
        logging.info('new output file is created')
    time.sleep(5)
    savingPath = os.path.abspath(r'..\Output_Files\GatewayRequirement\GatewayReqTestPlan.xlsm')
    logging.info(savingPath)
    logging.info("---------------------------------")
    logging.info("Saving Testplan Sheet ", output_dir + '\\GatewayRequirement\\GatewayReqTestPlan.xlsm')
    logging.info("---------------------------------")
    tpBook.save(savingPath)
    logging.info('Testplan[sheet] is saved in output folder')
    UpdateHMIInfoCb('\n\nTestplan[sheet] is saved in output folder '+output_dir + '\GatewayRequirement')

def getGatewayBook():
    GT_Book = ''
    if os.path.isdir(ICF.getInputFolder() + "\\Gateways"):
        arr = os.listdir(ICF.getInputFolder() + "\\Gateways")
        for i in arr:
            if i.find('Gateways') != -1 and i.find('~$') == -1:
                GT_Book = i
                break
                logging.info("GT_Book - ", GT_Book)
    else:
        UpdateHMIInfoCb(
            "\n>>"+ICF.getInputFolder() + "\Gateways folder not exist, please create the folder and place Gateway sheet under this folder then run the tool.")
        return -1

    return GT_Book


# Input: path of the file where it is present
# Description: This function opening the file from the given path if present
# Output: This function returns the file object
def openGatewaySheet():
    gatewayBook = getGatewayBook()
    if gatewayBook is not None and gatewayBook != "" and gatewayBook != -1:
        return xw.Book(ICF.getInputFolder() + "\\Gateways\\" +gatewayBook)
    else:
        UpdateHMIInfoCb("\n>>Gateways file not exist, please place Gateway sheet under the folder "+ICF.getInputFolder()+"\\Gateways then run the tool.")
        return -1


# Input - name of the sheet which is used to make changes
# This function will modify the test sheet
def modifyTestSheet(sheet, type, macro):
    try:
        sheet.activate()
        TPM.selectTpWritterProfile(macro)
        TPM.unProtectTestSheet(macro)
        if type == "modify":
            TPM.selectTestSheetModify(macro)  # for changing version
            fillGatewayHistory(sheet, "Already present in FSEE_GATEWAY PT")
            TPM.TestSheetRemove(macro)
        else:
            time.sleep(3)
            logging.info("sheet.range('E5').value - ", sheet.range('E5').value)
            sheet.range('E5').value = "GATEWAY"
            logging.info("\n---------------------Added Gateway------------------")
    except:
        logging.info("Something went wrong in updating the test sheet using macro")


# Input - name of the sheet which is used to make changes and the content which is used to update
# keyword defines the content which we are going to add in the test sheet
# this function fill the history part in the given test sheet
def fillGatewayHistory(sheet, keyword):
    sheet_value = sheet.used_range.value
    try:
        maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        # cellValue = EI.searchDataInExcel(sheet, (26, maxrow), "Nature des modifications")
        cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow), "Nature des modifications")
        row, col = cellValue["cellPositions"][0]
        logging.info("In History", row, col)
        if sheet.range(row + 1, col).value is not None:
            getString = sheet.range(row + 1, col).value + keyword
        else:
            getString = keyword

        EI.setDataFromCell(sheet, (row + 1, col), getString)
        EI.setDataFromCell(sheet, (row + 1, col + 1), ICF.gettrigram())
    except:
        UpdateHMIInfoCb("\n Some went wrong in filling the history")
        return -1


# Input - name of the sheet and the value as list to search in the given sheet
# gatewaysheet defines the "GATEWAYS" sheet tab
# reqList is the requirement list which are going to treat
# This function process each requirement and return the match cases of requirement in Gateways sheet
# Output:
# - notmatchedreq returns number of not matched req as bool
# - notmatchedreq returns number of not matched req as bool
# - matchedreqList returns matched requirements as list
# - notmatchedreqList returns not matched requirements as list
def find_matched_and_not_matched_req(gatewaysheet, reqList):
    # 4 defines the requirement column in the gateway sheet
    REQ_COL = 4
    matchedreqList = []
    notmatchedreqList = []
    reqmatchResult = {
        'matchedreqList': [],
        'notmatchedreqList': []
    }
    for index, reqID in enumerate(reqList):
        logging.info("reqkeyword - ", reqID)
        # getting the requirement only
        reqId = re.sub(r'\([^)]*\)', "", str(reqID))
        # gatewaySheetSearchresult = EI.searchDataInCol(gatewaysheet, REQ_COL, reqId.strip())

        sheet_value = gatewaysheet.used_range.value
        gatewaySheetSearchresult = EI.searchDataInColCache(sheet_value, REQ_COL, reqId.strip())

        logging.info("gatewaySheetSearchresult - ", gatewaySheetSearchresult)
        if gatewaySheetSearchresult['count'] > 0:
            logging.info("Requirement " + reqID + " present in Gateways sheet")
            if len(reqList) == 1:
                matchedreqList.append(reqID)
            else:
                matchedreqList.append(reqID)
        else:
            notmatchedreqList.append(reqID)

    reqmatchResult['matchedreqList'] = matchedreqList
    reqmatchResult['notmatchedreqList'] = notmatchedreqList

    return reqmatchResult


# Input - Testplan Book object and the testsheet
# tpBook defines the testplan sheet object
# testsheet defines in which test sheet going to make modification
# this function process the requirements in C4 column in the given testsheet,
# changing the type value if requirement not present in GATEWAYS,
# Or deleting the test sheet
def changeReqType(tpBook, matriceTestSheet):
    EI.openExcel(ICF.getTestPlanMacro())
    macro = EI.getTestPlanAutomationMacro()
    reqList = tpBook.sheets[matriceTestSheet].range('C4').value
    if reqList is not None:
        reqsplit = reqList.split("|")
        filteredReqs = []
        funcGTWSheets = []
        funcGatewayBook = openGatewaySheet()
        logging.info("funcGatewayBook >> ", funcGatewayBook)
        if funcGatewayBook is not None and funcGatewayBook != -1 and funcGatewayBook != "":
            for fs in funcGatewayBook.sheets:
                funcGTWSheets.append(fs.name)
            if "GATEWAYS" not in funcGTWSheets:
                logging.info("GATEWAYS sheet not present in " + funcGatewayBook.name)
                # Send to Info box
                return -1
            gatewaySheet = funcGatewayBook.sheets['GATEWAYS']
            for reqID in reqsplit:
                # Do not consider Interface requirement
                if reqID.find("GEN-DCINT") == -1:
                    if reqID not in filteredReqs:
                        filteredReqs.append(reqID)
            logging.info("filteredReqs >> ", filteredReqs)
            logging.info("No of requirements - ", len(filteredReqs))
            if filteredReqs:
                reqresult = find_matched_and_not_matched_req(gatewaySheet, filteredReqs)
                logging.info("reqresult >> ", reqresult)
                with open('../FunctionalGateway_Report.txt', 'w') as f:
                    f.writelines("Gateway Requirement Result - \n \n \n")
                if len(filteredReqs) == len(reqresult['matchedreqList']):
                    modifyTestSheet(tpBook.sheets[matriceTestSheet], 'modify', macro)

                elif len(filteredReqs) == len(reqresult['notmatchedreqList']):
                    modifyTestSheet(tpBook.sheets[matriceTestSheet], 'add', macro)
                    UpdateHMIInfoCb("\nRequirements not present in "+matriceTestSheet+" changing the Type as Gateway in E5 column")
                    with open('../FunctionalGateway_Report.txt', 'a') as f:
                        f.writelines("\n\nRequirements " + str(
                            reqList) + " in testsheet " + matriceTestSheet + " not present in Gateways please check it manually.")

                elif len(reqresult['matchedreqList']) != 0 and len(reqresult['notmatchedreqList']) != 0:
                    logging.info("Requirements are not present in Gateways sheet")
                    with open('../FunctionalGateway_Report.txt', 'a') as f:
                        f.writelines(
                            "\n\nSome Requirements in testsheet " + matriceTestSheet + " not present in Gateways please check it manually. Matched Requirements - " + str(
                                reqresult['matchedreqList']) + " Not Matched Requirements - " + str(
                                reqresult['notmatchedreqList']))
        else:
            return -1

    return 1


# Input - Testplan Book object
# tpBook defines the Test Plan sheet object
# This function process the cell values which is having the Keyword 'Gateway'
def processGatewayRequirements(tpBook):
    logging.info("----------------Processing the Gateway Requirements----------------\n")
    tpBookSheetList = []
    for ts in tpBook.sheets:
        tpBookSheetList.append(ts.name)
    if "Matrice de tests" not in tpBookSheetList:
        UpdateHMIInfoCb("\nMatrice de tests sheet not exist in "+tpBook.name)
        return -1
    matriceSheet = tpBook.sheets["Matrice de tests"]
    matriceSheet.activate()
    DESC_COL = 2
    logging.info("Searching the Gateway keyword")
    # gatewayresult = EI.searchDataInCol(matriceSheet, DESC_COL, "Gateway", True)

    sheet_value = matriceSheet.used_range.value
    gatewayresult = EI.searchDataInColCache(sheet_value, DESC_COL, "Gateway", True)

    logging.info("gatewayresult ->> ", gatewayresult)
    if gatewayresult['count'] > 0:
        for cellPosition in gatewayresult['cellPositions']:
            row, col = cellPosition
            logging.info("\n\ncellPosition ->> ", cellPosition)
            # 1 defines the first column "N° Fiches de test"
            matriceTestSheetStr = matriceSheet.range(row, 1).value
            matriceTestSheet = re.sub(r'\_[0-9]{1,2}$',"",matriceTestSheetStr)
            # matriceTestSheetLink = matriceSheet.range(row, 1).hyperlink
            # matriceTestSheet = re.findall(r"VSM+[0-9]{1,2}\+\_[a-zA-z0-9]{1,2}\+\_[0-9]{2}\+\_[0-9]{4}|BSI+[0-9]{1,2}\+\_[a-zA-z0-9]{1,2}\+\_[0-9]{2}\+\_[0-9]{4}",
            #     matriceTestSheetLink)
            # logging.info(f"matriceTestSheetLink>> {matriceTestSheetLink}")
            logging.info(f"matriceTestSheetStr>> {matriceTestSheetStr}")
            logging.info(f"matriceTestSheet>> {matriceTestSheet}")
            if matriceTestSheet != "" and matriceTestSheet is not None:
                if matriceTestSheet not in tpBookSheetList:
                    logging.info(matriceTestSheet, " not present")
                    continue
                else:
                    logging.info(matriceTestSheet , " present in Test plan")
                    logging.info("matriceTestSheet --> ", matriceTestSheet)
                    # tpBook.sheets[matriceTestSheet].activate()
                    changeReqresult = changeReqType(tpBook, matriceTestSheet)
    else:
        logging.info("\n------------ Gateway keyword not present --------------")
        UpdateHMIInfoCb("\n\n>> Gateway keyword not present in sheet ["+matriceSheet.name+"]")
        return -1

    return changeReqresult


def gatewayReq():
    tpBook = EI.openTestPlan()
    if tpBook != -1:
        gatewayResult = processGatewayRequirements(tpBook)
        logging.info("gatewayResult >> ", gatewayResult)
        logging.info("tpBook - ", tpBook.name)
        if gatewayResult != -1:
            logging.info("\n\n-----------[Gateway Requirement process Completed!]-----------")
            saveTestPlan(tpBook)

    else:
        UpdateHMIInfoCb("\nTest plan sheet not found under folder /Input_Files")
        # exit()
