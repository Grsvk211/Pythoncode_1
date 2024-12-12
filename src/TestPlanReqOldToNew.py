# Changing the requirement name from Old (GEN-XXX) to New (REQ-xxx) in Test sheets
# Requirement text file exist under the below svn path
# svn.in.expleogroup.com/!/#AppStore/view/head/Tags/BetaVersion1.0/Tool_Requirement/Gateway_Requirement.txt


import InputConfigParser as ICF
import ExcelInterface as EI
import TestPlanMacros as TPM
import re
import xlwings as xw
import os
import time
import logging
UpdateHMIInfoCb = None


def saveTestPlan(tpBook):
    output_dir = os.path.abspath(r"..\Output_Files")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    if not os.path.exists(r"..\Output_Files\ReqNameChange"):
        os.makedirs(r"..\Output_Files\ReqNameChange")
        logging.info('new output file is created')
    time.sleep(5)
    savingPath = os.path.abspath(r'..\Output_Files\ReqNameChange\GEN_To_REQ.xlsm')
    logging.info(savingPath)
    logging.info("---------------------------------")
    logging.info("Saving Testplan Sheet ", output_dir + '\\ReqNameChange\\GEN_To_REQ.xlsm')
    logging.info("---------------------------------")
    tpBook.save(savingPath)
    logging.info('Testplan[sheet] is saved in output folder')
    UpdateHMIInfoCb('\n>>Testplan[sheet] is saved in output folder Output_Files\ReqNameChange')


def registerInfoTextBoxReqOldNew(func):
    global UpdateHMIInfoCb
    UpdateHMIInfoCb = func


def getKPI_Doc():
    PT_KPI = ''
    if os.path.isdir(ICF.getInputFolder() +"\\KPI"):
        arr = os.listdir(ICF.getInputFolder() +"\\KPI")
        for i in arr:
            # if i.find('PT_KPI') != -1 and i.find('~$') == -1:
            if i.find('KPI') != -1 or i.find('PT_KPI') != -1 and i.find('~$') == -1:
                PT_KPI = i
                break
        logging.info("PT_KPI - ", PT_KPI)
        if PT_KPI != "":
            return EI.openExcel(ICF.getInputFolder() + "\\KPI\\" + PT_KPI)

        else:
            UpdateHMIInfoCb("\nKPI Sheet not found under folder"+ICF.getInputFolder()+"\KPI.")
            return -1
    else:
        UpdateHMIInfoCb("\nKPI folder not exist in "+ICF.getInputFolder()+", please create the folder and place test sheet under this folder.")
        return -1

def OpenKPI_Sheet(filepath):
    if os.path.isfile(filepath):
        return xw.Book(filepath)
    else:
        logging.info("FA_REQ_PT_KPI3_V5__version_1_.xlsm File not found in /Input_Files/KPI folder. Please put file under this folder and run the tool")
        return -1


def replaceReqInTestSheet_and_DJNC(tpBook, sheet, searchKeyword, repREQ, repVerNum):
    if searchKeyword != "" and repREQ != "" and repVerNum != "":
        logging.info("sheet >>>>> ", sheet)
        logging.info("searchKeyword >> ", searchKeyword)
        sheet.activate()
        EI.openExcel(ICF.getTestPlanMacro())
        macro = EI.getTestPlanAutomationMacro()
        TPM.unProtectTestSheet(macro)
        UpdateHMIInfoCb("\n>> Replacing the requirement from '" + searchKeyword + "' to " + repREQ + "(" + repVerNum + ")")
        reqval = (str(sheet.range('C4').value))
        logging.info(reqval, "sheet.range('C4').value")
        replVal = re.sub(searchKeyword + r"\([^)]*\)", repREQ+"("+repVerNum+")", reqval)
        logging.info(replVal, " --> replVal")
        # #exit()
        if replVal:
            sheet.range("C4").value = replVal
            tpBookSheetList = [tps.name for tps in tpBook.sheets]
            if 'DJNC' in tpBookSheetList:
                djncSheet = tpBook.sheets['DJNC']
                djncSheet.activate()
                TYPE_JUSTIFICATION_COL = 1
                # djncReqresult = EI.searchDataInCol(djncSheet, TYPE_JUSTIFICATION_COL, searchKeyword)

                sheet_value = djncSheet.used_range.value
                djncReqresult = EI.searchDataInColCache(sheet_value, TYPE_JUSTIFICATION_COL, searchKeyword)

                if djncReqresult['count'] > 0:
                    for cellPosition in djncReqresult['cellPositions']:
                        row, col = cellPosition
                        logging.info("djnc cellPosition ->> ", cellPosition)
                        TPM.unProtectTestSheet(macro)
                        djncSheet.range(row, 1).value = repREQ
                        djncSheet.range(row, 3).value = repREQ+"("+repVerNum+")"
                else:
                    logging.info("\n--------------Requirement "+searchKeyword+" not present in DJNC sheet--------------")
                    return -1
            else:
                logging.info("\n ------ DJNC sheet not present -----")
                with open('../KPI_Replacing_GEN_to_REQ_Report.txt', 'a') as f:
                    f.writelines("\n\nDJNC sheet not present in "+tpBook.name)

    return 1


def hanlde_GEN_req(tpBook, sheet, funcName, GenReqIds):
    # kpiDocpath = "../Input_Files/KPI/FA_REQ_PT_KPI3_V5__version_1_.xlsm"
    KPIBook = getKPI_Doc()
    KPISheetList = []
    kpiReqVer = ""

    if KPIBook is not None and KPIBook != -1 and KPIBook != "":
        logging.info("-------------------------------------------")
        for kpis in KPIBook.sheets:
            KPISheetList.append(kpis.name)
        if funcName != "" and funcName is not None:
            logging.info("TP Function Name = ", funcName)
            if funcName not in KPISheetList:
                logging.info(funcName+" sheet not present")
                return -1
            KPIsheet = KPIBook.sheets[funcName]
            if KPIsheet is not None:
                KPIsheet.activate()
                REQ_PT_COL = 1
                for genreqID in GenReqIds:
                    logging.info("\nSearching the requirement '"+genreqID+"'")
                    if re.search(r'\([^)]*\)', genreqID):
                        genreqID = re.sub(r'\([^)]*\)', "", str(genreqID))
                        # kpiSearchResult = EI.searchDataInCol(KPIsheet, REQ_PT_COL, genreqID)

                        sheet_value = KPIsheet.used_range.value
                        kpiSearchResult = EI.searchDataInColCache(sheet_value, REQ_PT_COL, genreqID)

                        logging.info("kpiSearchResult ->> ", kpiSearchResult)
                        if kpiSearchResult['count'] > 0:
                            for cellPosition in kpiSearchResult['cellPositions']:
                                row, col = cellPosition
                                logging.info("cellPosition ->> ", cellPosition)
                                KPIreqCellValue = (str(KPIsheet.range(row, 3).value))
                                KPIverCellValue = (str(KPIsheet.range(row, 4).value))
                                logging.info("KPIreqCellValue ->> ", KPIreqCellValue)
                                logging.info("KPIverCellValue ->> ", KPIverCellValue)
                                if (KPIreqCellValue is not None and KPIreqCellValue != "" and KPIreqCellValue != "None") and (KPIverCellValue is not None and KPIverCellValue !="" and KPIverCellValue != "None"):
                                    kpiReqVer = KPIreqCellValue+"("+KPIverCellValue+")"
                                    replaceReqInTestSheet_and_DJNC(tpBook, sheet, genreqID, KPIreqCellValue, KPIverCellValue)
                                else:
                                    if (KPIreqCellValue is not None or KPIreqCellValue != "" or KPIreqCellValue != "None") and (KPIverCellValue is not None and KPIverCellValue != "" and KPIverCellValue != "None"):
                                        with open('../KPI_Replacing_GEN_to_REQ_Report.txt', 'a') as f:
                                            f.writelines(
                                                "\n\n"+genreqID+" in sheet "+sheet.name+" having the new version not having new requirement")
                                    elif (KPIverCellValue is not None or KPIverCellValue != "" or KPIverCellValue != "None") and (KPIreqCellValue is not None and KPIreqCellValue != "" and KPIreqCellValue != "None"):
                                        with open('../KPI_Replacing_GEN_to_REQ_Report.txt', 'a') as f:
                                            f.writelines(
                                                "\n\n"+genreqID+" in sheet "+sheet.name+" having new requirement not having the new version.")
                                    else:
                                        with open('../KPI_Replacing_GEN_to_REQ_Report.txt', 'a') as f:
                                            f.writelines(
                                                "\n\n"+genreqID+" in sheet "+sheet.name+" not having new requirement and new version")

                        else:
                            logging.info("\n!!!!! Reuirement "+genreqID+" not present in KPI Doc")
                            with open('../KPI_Replacing_GEN_to_REQ_Report.txt', 'a') as f:
                                f.writelines(
                                    "\n\n" + genreqID + " in sheet "+sheet.name+" not present in KPI Doc sheet "+KPIsheet.name)
                    else:
                        logging.info("\n!!!!! Reuirement " + genreqID + " not having the version in sheet "+sheet.name)
                        with open('../KPI_Replacing_GEN_to_REQ_Report.txt', 'a') as f:
                            f.writelines(
                                "\n\n " + genreqID + " in Test sheet " + sheet.name + " not having the version")

            else:
                logging.info("\n!!!!!!'"+KPIsheet.name+"' sheet not found!!!!!!")
                return -1
        else:
            logging.info("\n!!!!!!!!! Problem in finding the Function Name from TestPlan sheet !!!!!!!!!")
            UpdateHMIInfoCb("\n>>"+funcName+" sheet not present in KPI File")
            return -1
    else:
        return -1

    return 1


def processGEN_Requirements(tpBook):
    funcName = tpBook.sheets['Sommaire'].range(4, 3).value
    logging.info("tpBook.sheets >> ", tpBook.sheets)
    with open('../KPI_Replacing_GEN_to_REQ_Report.txt', 'w') as f:
        f.writelines("---------------   Old requirement to New requirement name change report   ---------------\n\n")
    GenReq_replace = ''
    for sheet in tpBook.sheets:
        GenReqIDs = []
        if (sheet.name.find("VSM") != -1 or sheet.name.find("BSI") != -1) and sheet.visible:
            reqColValue = sheet.range("C4").value
            if reqColValue is not None:
                if reqColValue.find("GEN") != -1:
                    logging.info("\n\n!!!!!!!!!!!!!!!!Processing the sheet - ", sheet.name, "!!!!!!!!!!!!!!!!")
                    logging.info("Processing Requirements - ", reqColValue)
                    UpdateHMIInfoCb("\n\n!!!!!Processing the sheet - "+ sheet.name+ "!!!!!")
                    reqs = reqColValue.split("|")
                    for reqId in reqs:
                        if reqId is not None and reqId.upper().find("DCINT") == -1 and reqId.upper().find("DCI") == -1 and reqId.upper().find("REQ") == -1:
                            GenReqIDs.append(reqId)
                    logging.info("GenReqIDs ->> ", GenReqIDs)
                    GenReq_replace = hanlde_GEN_req(tpBook, sheet, funcName, GenReqIDs)
                    logging.info("GenReq_replace ->> ", GenReq_replace)
                else:
                    logging.info("\nTest sheet "+sheet.name+" not having old requirements")
                    UpdateHMIInfoCb("\nTest sheet "+sheet.name+" not having old requirements")
            else:
                logging.info("\nRequirements not present in sheet "+sheet.name)
                UpdateHMIInfoCb("\n\nRequirements not present in sheet " + sheet.name)

    return GenReq_replace

def getTestPlanSheet():
    PT = ''
    if os.path.isdir(ICF.getInputFolder() +"\\Testsheet"):
        arr = os.listdir(ICF.getInputFolder() +"\\Testsheet")
        for i in arr:
            if i.find('Tests') != -1 and i.find('~$') == -1:
                PT = i
                break
    else:
        UpdateHMIInfoCb("\n Input_Files/Testsheet folder not exist, please create the folder and place test sheet under this folder.")
        return -1
    return PT

def reqNameChange():
    # tpBook = EI.openExcel("../Input_Files/Testsheet/Tests_20_27_01272_18_01101_FSEE_PCGA_V31_VSM.xlsm")
    PT = getTestPlanSheet()
    if PT != "" and PT is not None and PT != -1:
        tpBook = EI.openExcel(ICF.getInputFolder() + "\\Testsheet\\" + PT)
        logging.info("tpBook - ", tpBook.name)
        genReqResult = processGEN_Requirements(tpBook)
        logging.info("\n\ngenReqResult >> ", genReqResult)
        if genReqResult != -1 and genReqResult != '':
            UpdateHMIInfoCb("\n---[Changing the name of requirements from old to new process completed]---")
            saveTestPlan(tpBook)
    else:
        UpdateHMIInfoCb("\nTest plan sheet not found under folder "+ICF.getInputFolder()+"\Testsheet, please place the file and run the tool.")