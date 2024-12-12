import ExcelInterface as xi
import xlwings as xw
import re
import pywintypes
import InputConfigParser as ICF
import TestPlanMacros as TPM
import os
import time
import logging
# ICF.loadConfig()
# tpBook = xi.openTestPlan()
# sheets = tpBook.sheets


missingReqTrcmtx = []

# this function is used for to save the Gen to Dic excel sheet in the ReqNameChange folder output folder
def saveTestPlan(tpBook):
    # output_dir = os.path.abspath(r"..\Output_Files")
    output_dir = os.path.isfile(ICF.getOutputFiles())
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    if not os.path.exists(ICF.getOutputFiles()+"\\ReqNameChange"):
        os.makedirs(ICF.getOutputFiles()+"\\ReqNameChange")
        logging.info('new output file is created')
    # time.sleep(5)
    savingPath = os.path.abspath(ICF.getOutputFiles()+"\\ReqNameChange\\GEN_To_DCI.xlsm")
    logging.info(savingPath)
    logging.info("---------------------------------")
    logging.info("Saving Testplan Sheet ", ICF.getOutputFiles() + '\\ReqNameChange\\GEN_To_DCI.xlsm')
    logging.info("---------------------------------")
    tpBook.save(savingPath)
    logging.info('Testplan[sheet] is saved in output folder')

def filterVSM(trcBookSheet):
    # Input: TraceabilityMatrix Sheet (Object) taken as input
    # Description: function filters the data by "VSM" and creates a dictionary of
    #              (old_requirement_name=>new_requirement_name)
    # Output: Returns a map of old requirement namews to new requirement names (Only VSMs)

    xi.activateSheetObj(trcBookSheet)

    old2NewMap = dict()

    nrows = trcBookSheet.used_range.last_cell.row

    for i in range(1, nrows + 1):
        # logging.info(i)
        componentValue = xi.getDataFromCell(trcBookSheet, f"A{i}")
        if componentValue==None:
            continue
        isVsm = componentValue.lower()=="vsm"
        if isVsm:
            old2NewMap[xi.getDataFromCell(trcBookSheet, f"B{i}")[:-3]] = xi.getDataFromCell(trcBookSheet, f"C{i}")

    return old2NewMap


def filterInterfaceReqKPI(kpiFile, functionName):
    # Input: kpiFile (object) and functionName (string) taken as input
    # Description: For a given function name, the function 
    # Output: returns an array of all interface requirement names

    interfaceReq = []

    functionSheet = kpiFile.sheets[functionName]

    nrows = functionSheet.used_range.last_cell.row

    for i in range(1, nrows + 1):
        reqValue = xi.getDataFromCell(functionSheet, f"B{i}")

        if reqValue==None:
            continue

        isInterfaceReq = reqValue.lower()=="interface req"
        if isInterfaceReq:
            interfaceReq.append(xi.getDataFromCell(functionSheet, f"A{i}"))

    return interfaceReq


def getSheetsWithReqName(workbook, reqName):
    # Input: Takes test file (object), and requirement name (string) as input
    # Description: Function searches for sheet with requirement name
    # Output: Returns an empty error if no sheet contains requirement name, otherwise returns an array of sheet names

    sheets = workbook.sheets

    sheetsWithReqName = []

    for sheet in sheets:
        sheetReqStr = xi.getDataFromCell(sheet, "C4")
        if reqName in str(sheetReqStr):
            sheetsWithReqName.append(sheet.name)
    return sheetsWithReqName


def log(msg):
    # Input: Takes msg string
    # Description:  Outputs inputted msg as log
    # Output: Outputs "LOG: LOG_MSG_EXAMPLE"

    logging.info(f"LOG: {msg}")


def findRowInSheet(sheet, keyword, col):
    # Input: Takes a sheet object, keyword to search for, and column to search in, as input
    # Description: Function iterates over each row of a column in a particular sheet and 
    #              searches for specified keyword
    # Output: Returns None if keyword is not found, otherwise returns row number

    nrows = sheet.used_range.last_cell.row
    for i in range(1, nrows+1):
        cellVal = xi.getDataFromCell(sheet, f"{col}{i}")
        if cellVal and cellVal.strip()==keyword:
            return i
    return None


def extractArchFromConfig(configSheet, thematicStr):
    # Input: Takes configSheet (object), thematic string as input
    # Description: Thematics may not always be in form of LVM, or LYQ. This function
    #              converts thematic string of type:  LYQ_02|DXD_04|IBM_07|LRQ_01
    #              to an equivalent array of LVM_XX LYQ_XX: [LYQ_02, LVM_01, LVM_02, LVM_03]
    # Output: returns the generated array

    thematicArr = set()

    thematicStrSplit = thematicStr.split("|")

    for thematic in thematicStrSplit:
        if "LVM" in thematic or "LYQ" in thematic:
            thematicArr.add(thematic)
            continue
        rowNum = findRowInSheet(configSheet, thematic.strip(), "A")

        if rowNum==None:
            logging.info(f"{thematic.strip()} not found in configSheet")
            continue

        gRowValue = xi.getDataFromCell(configSheet, f"G{rowNum}")
        hRowValue = xi.getDataFromCell(configSheet, f"H{rowNum}")

        isG_X, isH_X = False, False

        if gRowValue!=None:
            isG_X = gRowValue.strip().lower()=='x'

        if hRowValue!=None:
            isH_X = hRowValue.strip().lower()=='x'

        if isG_X:
            thematicArr.add("LVM_01")

        if isH_X:
            thematicArr.add("LVM_02")
            thematicArr.add("LVM_03")
    return list(thematicArr)


def replaceOldReq2NewReq(sheet, oldReq, newReq):
    # Input: Takes testplan function sheet (object), old requirement name, new requirement name as input 
    # Description: Function gets the current requirement name from the sheet, replaces old requirement 
    #              with the new name, make the changes in file, and save the changes.
    # Output: logging.infos updated sheet's name

    sheetReqStr = xi.getDataFromCell(sheet, "C4")
    sheetReqStr = sheetReqStr.replace(oldReq.strip(), newReq.strip())

    sheetReqStr = sheetReqStr.replace("||","|")

    if sheetReqStr[-1]=="|":
        sheetReqStr = sheetReqStr[:-1]

    logging.info(sheetReqStr)

    sheet.activate()
    xi.openExcel(ICF.getTestPlanMacro())
    macro = xi.getTestPlanAutomationMacro()
    TPM.unProtectTestSheet(macro)
    xi.setDataFromCell(sheet, "C4", sheetReqStr)
    log(f"{sheet.name} UPDATED")


def getReqsNotPresentInTrcmtx():
    # Input:
    # Description: Returns missing traceability matrix requirements as a list
    # Output
    return missingReqTrcmtx


def logMissingReq():
    # Input: 
    # Description: Function writes all requirements that are not in traceability matrix
    # Output: Logs it in log.txt file
    with open("log.txt", "w") as f:
        f.write("Missing requirements in traceability matrix:\n")
        for req in missingReqTrcmtx:
            f.write(str(req)+"\n")


def renameInterfaceRequirements(traceablityMatrixFileName, dciFileName, kpiFileName, testFile, configFile):
    # removed paraeters: functionName
    # Input: takes file names of: traceability matrix, dci file, kpi file, testplan file, and config file
    # Description: Renames interface requirement names if they specifiy certain conditions
    # Output:
    # with xw.App(visible=False, add_book=False) as app:
    try:
        trcBook = xi.openExcel(ICF.getInputFolder()+"\\Interface_Old_To_New\\"+traceablityMatrixFileName)
        kpiBook = xi.openExcel(ICF.getInputFolder()+"\\Interface_Old_To_New\\"+kpiFileName)
        dciBook = xi.openExcel(ICF.getInputFolder()+"\\Interface_Old_To_New\\"+dciFileName)
        testBook = xi.openExcel(ICF.getInputFolder()+"\\Interface_Old_To_New\\"+testFile)
        configBook = xi.openExcel(ICF.getInputFolder()+"\\Interface_Old_To_New\\"+configFile)
    except Exception as e:
        logging.info(f"File(s) not found: {e}")

    log("OPENED ALL WORKBOOKS")

    try:
        trcBookSheet = trcBook.sheets["Sheet1"]
    except Exception as e:
        logging.info(f"Sheet 1 not found \neError: {e}")

    old2NewMap = filterVSM(trcBookSheet)

    log("FILTERED FOR VSM")
    functionName = getFunctionName(testBook)
    interfaceReqArr = filterInterfaceReqKPI(kpiBook, functionName)
    log("FILTERED FOR INTERFACE REQ NAMES")

    for interfaceReqName in interfaceReqArr:
        try:
            # ============ Get new requirement names ============
            replaceReqWithThis = ""

            if interfaceReqName[:-3] not in old2NewMap:
                missingReqTrcmtx.append(interfaceReqName)

            newReqNames = old2NewMap[interfaceReqName[:-3]]  # "Reqid(1)Reqid(2)Reqid(3)""
            newReqNamesArrWithoutVersions = re.split("\(\d+\)", newReqNames)[:-1] # (Reqid, Reqid, Reqid)

            parts = re.split("(\(\d+\))", newReqNames)
            newReqVersions = [parts[i][1:-1] for i in range(1, len(parts), 2)] # (1, 2, 3)

            # ============ Search requirement name in DCI file and get thematics ============

            try:
                dciMUXSheet = dciBook.sheets["MUX"]
            except Exception as e:
                logging.info(f"MUX sheet not found\n Error: {e}")
                continue

            dciRow, dciCol = dciMUXSheet.used_range.last_cell.row, 26

            reqSheetNames = getSheetsWithReqName(testBook, interfaceReqName)

            if(len(reqSheetNames)==0):
                logging.info(f"There's no sheet in test file that contains {interfaceReqName} as a requirement")
                continue

            for reqSheetName in reqSheetNames:
                for nameIdx, reqName in enumerate(newReqNamesArrWithoutVersions):

                    # Search requirement name in DCI and get row number
                    row = 0
                    for i in range(1, dciRow+1):
                        if reqName.strip() == xi.getDataFromCell(dciMUXSheet, f"A{i}").strip()[:-3]:
                            row = i
                            break

                    thematicsStr = xi.getDataFromCell(dciMUXSheet, f"P{row}")  # "LLL_02 AND LRW_01 AND LVM_01 AND LYQ_01"

                    pattern = r'\b(?:LVM|LQY)_\d{2}\b'
                    dciThematicArr = re.findall(pattern, thematicsStr) # ['LVM_01', 'LYQ_01']

                    reqSheetThematics = xi.getDataFromCell(testBook.sheets[reqSheetName], "C8")

                    try:
                        reqSheetThematics = extractArchFromConfig(configBook.sheets["Thématiques"], reqSheetThematics)
                    except Exception as e:
                        logging.info(f"Config file doesn't contain \"Thématiques\" sheet \nError:{e}")
                        exit() # exit because it'll fail for all cases


                    # Check if thematics from DCI is present in thematics of test plan reqSheet
                    isTrue = True

                    logging.info(f"DCI THEMATICS ARR: {dciThematicArr}")
                    logging.info(f"REQ SHEET THEMATICS ARR: {reqSheetThematics}")
                    for thematic in dciThematicArr:
                        isThemInReqThem = False
                        for reqThematic in reqSheetThematics:
                            if thematic in reqThematic:
                                isThemInReqThem = True
                                break
                        isTrue = isTrue and isThemInReqThem
                        if not isTrue:
                            break

                    if isTrue:
                        # logging.info(newReqVersions)
                            replaceReqWithThis += f"{reqName.strip()}({newReqVersions[nameIdx]})|"

                if len(replaceReqWithThis)>0:
                    replaceOldReq2NewReq(testBook.sheets[reqSheetName], interfaceReqName, replaceReqWithThis)
                    logging.info(f"INTERFACE REQ NAME CHANGED: {interfaceReqName} => {replaceReqWithThis}")

            log(f"{interfaceReqName} COMPLETED")

        except KeyError:
            logging.info(f"Interface req not found in traceability matrix: {KeyError}")
        except pywintypes.com_error as e:
            logging.info(f"Program ended as files were either closed or the task is completed: {e}")
        except Exception as e:
            logging.info(f"Unknown error occured: {e}")

    xi.remove_Duplicates_C4(testBook)
    logMissingReq()
    saveTestPlan(testBook)
    logging.info("the Process for treating the old to new Completed.")

    trcBook.close()
    kpiBook.close()
    dciBook.close()
    testBook.close()
    configBook.close()

def getFunctionName(testFile):
    funcName = testFile.sheets['Sommaire'].range(4, 3).value
    functionName = funcName.split("-")[1].strip()
    return funcName

def reqNameChanging():
    traceablityMatrixFileName = xi.findInterfaceOldNewInputFiles()[3]
    dciFileName = xi.findInterfaceOldNewInputFiles()[2]
    kpiFileName = xi.findInterfaceOldNewInputFiles()[1]
    testFile = xi.findInterfaceOldNewInputFiles()[0]
    configFile = xi.findInterfaceOldNewInputFiles()[4]
    # functionName = getFunctionName(testFile)

    logging.info("traceablityMatrixFileName - ", traceablityMatrixFileName)
    logging.info("kpiFileName - ", kpiFileName)
    logging.info("configFile - ", configFile)
    logging.info("dciFileName - ", dciFileName)
    logging.info("testFile - ", testFile)

    renameInterfaceRequirements(traceablityMatrixFileName, dciFileName, kpiFileName, testFile, configFile)




