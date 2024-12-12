import xlwings as xw
import re
import os, sys
from web_interface import startDocumentDownload
import InputConfigParser as ICF
import KeyboardMouseSimulator as KMS
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import time
from os import listdir
from os.path import isfile, join
import BusinessLogic as BL
import ExcelInterface as EI
import copy
import TestPlanMacros as TPM
from datetime import date
import WordDocInterface as WDI
from QIA_Updater import UpdateQiaParamGlobal
import logging
# confirmationPop = None

didTypeVal = ""


def confirmationPopup(func):
    global confirmationPop
    confirmationPop = func
    return confirmationPop


def getreqTestSheets(tpBook, rowOfInterface, keyword):
    maxrow = tpBook.sheets['Impact'].range('A' + str(tpBook.sheets['Impact'].cells.last_cell.row)).end('up').row
    logging.info("ppmaxrow- ", maxrow)
    col = 4
    rowList = []
    rqList = []
    testSheetList = []
    sheet = tpBook.sheets['Impact']
    logging.info("refTs-", sheet)
    logging.info("type(rowOfInterface)", type(rowOfInterface))
    logging.info("rowOfInterface12 -", rowOfInterface)
    logging.info("testvalue1 - ", sheet.range(18, 4).value)
    bMultipleTs = False

    for i in range(rowOfInterface, maxrow + 1):
        cellValue = str(sheet.range(i, 1).value)
        if keyword in cellValue:
            TPcellValue = sheet.range(i, 4).value.split("\n")
            if len(TPcellValue) > 2:
                bMultipleTs = True
            logging.info("TPcellValue - ", TPcellValue)
            for t in TPcellValue:
                if len(t) != 0:
                    testSheetList.append(t)
            rowList.append(i)

    result = {'testSheetList': testSheetList, 'rowList': rowList, 'multipleTS': bMultipleTs}
    return result


def openParamGlobalSheet():
    param_global = EI.findInputFiles()[7]
    if len(param_global) != 0 and param_global != "" and param_global is not None:
        ParamBook = EI.openExcel(ICF.getInputFolder() + "\\" + param_global)
        return ParamBook
    else:
        logging.info("\nParam Global Sheet not exist!!!\n")
        return -1


def removeRefAndVer(filename):
    finalDocName = filename
    if re.search(r"\_[0-9]{5}\_[0-9]{2}\_[0-9]{5}", filename):
        finalDocName = re.sub(r"\_[0-9]{5}\_[0-9]{2}\_[0-9]{5}", "", filename)
    if re.search(r"_(V|v)+[0-9+]{1,2}", finalDocName):
        finalDocName = re.sub(r"_(V|v)+[0-9+]{1,2}", "", finalDocName)

    return finalDocName


def getInpDocPath(docName, version=""):  # version is optional
    logging.info("In getDocPath Function")
    logging.info("Docname = ", docName)
    logging.info("version = ", version)
    global oldVersion
    path = -1
    pat = ICF.getInputFolder() + "\\"
    onlyfiles = [f for f in listdir(pat) if isfile(join(pat, f))]
    logging.info("+++++++++++", pat, onlyfiles)

    logging.info(f"docName before {docName}")
    docName = BL.removeRefVerFromFilename(docName)
    if re.search("\(.*?\)", docName):
        docName = re.sub("\(.*?\)","", docName)
    logging.info(f"docName after1 {docName}")
    docName = docName.replace("-", "")
    logging.info(f"docName after {docName}")
    for fileName in onlyfiles:
        logging.info("Filename & docName = ", fileName, docName)
        fileName1 = removeRefAndVer(fileName)
        logging.info(f"fileName after - {fileName1}")
        if docName.strip() in fileName1:
            logging.info(">>>>>>>>>>>>>>>>.<<<<<<<<<<<<<<<<<<")
            fileVer = re.search("[(V|v)]+[0-9+]+", fileName)
            logging.info(f"fileVer {fileVer}")
            if fileVer is not None:
                fileVerRes = fileVer.group()
                logging.info("fileVer = ", fileVerRes)
                if (fileVerRes.upper().split("V")[1]) == version:
                    if os.path.splitext(fileName)[1] == ".docx":
                        path = pat + fileName
                        logging.info("Path found === ", path)
                        break
                    elif (os.path.splitext(fileName)[1] == ".doc") or (os.path.splitext(fileName)[1] == ".docm") or (
                            os.path.splitext(fileName)[1] == ".rtf"):
                        logging.info(".doc Name ", fileName)
                        path = BL.save_as_docx(pat + fileName)
                        oldVersion = 1
                        WDI.oldVersion = 1
                        break
                    else:
                        path = -2
                else:
                    path = -2
            else:
                logging.info("Document " + fileName + " is not having version in ipnut folder")
                path = -1
        else:
            path = -1
    return path


def searchSignalInCol(sheet, cellRange, keyword):
    x, y = cellRange
    count = 0
    rowcount = 0
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    if keyword == "":
        return searchResult

    for row in range(3, y + 1):
        cellValue = str(sheet.range(row, 3).value)
        sheetName = str(sheet) + "\n"
        if keyword in cellValue:
            searchResult["cellPositions"].append(tuple((row, 9)))
            searchResult["cellValue"].append(cellValue)
            count = count + 1
    searchResult["count"] = count
    logging.info("SearchResult- ", searchResult)
    return searchResult


def GetParamGlobalData(paramBook, pgData):
    ParamData = {
        "FLUX_MESSAGERIE_NEA": "",
        "PARAM_SIGNAL": ""
    }
    ParamDataList = []
    maxrow = paramBook.sheets['Paramètre'].range('A' + str(paramBook.sheets['Paramètre'].cells.last_cell.row)).end(
        'up').row
    logging.info("\npmaxrow- ", maxrow)
    sheet = paramBook.sheets['Paramètre']
    dciSignal = pgData['dciInfo']['dciSignal'].replace("$", "")
    if sheet.name.strip() == "Paramètre":
        searchResult = searchSignalInCol(sheet, (26, maxrow), dciSignal)
        if searchResult["count"] == 1:
            for cellPosition in searchResult["cellPositions"]:
                x, y = cellPosition
                ParamData["FLUX_MESSAGERIE_NEA"] = (str(EI.getDataFromCell(sheet, (x, y))))
                ParamData["PARAM_SIGNAL"] = (str(EI.getDataFromCell(sheet, (x, 3))))
                ParamDataList.append(ParamData.copy())
        elif searchResult["count"] > 1:
            QIAConditionResult = findQIACondition(sheet, searchResult, pgData)
            if QIAConditionResult:
                raiseQIAResult = raiseQIA(QIAConditionResult)
                logging.info("raiseQIAResult - ", raiseQIAResult)
                if raiseQIAResult:
                    if raiseQIAResult[0] != 0:
                        generateQIAResponse = generateQIAData(sheet, raiseQIAResult, pgData)
                        if generateQIAResponse == 1:
                            x, y = raiseQIAResult[1]
                            ParamData["FLUX_MESSAGERIE_NEA"] = (str(EI.getDataFromCell(sheet, (x, y))))
                            ParamData["PARAM_SIGNAL"] = (str(EI.getDataFromCell(sheet, (x, 3))))
                            ParamDataList.append(ParamData.copy())
                    else:
                        BL.displayInformation(f"\nRequirement, Frame and Version are matched for the requirement '{pgData['dciInfo']['dciReq']}', no need to raise QIA")
                        x, y = raiseQIAResult[1]
                        ParamData["FLUX_MESSAGERIE_NEA"] = (str(EI.getDataFromCell(sheet, (x, y))))
                        ParamData["PARAM_SIGNAL"] = (str(EI.getDataFromCell(sheet, (x, 3))))
                        ParamDataList.append(ParamData.copy())
        else:
            pgData['modified_req'] = pgData['dciInfo']['dciReq']
            pgData['comment'] = "Adding new flow"
            pgData['requestType'] = 1
            pgData['paramType'] = 0
            pgData['actual_req'] = pgData['dciInfo']['dciReq']
            addQIAresult = callQIAFunc(pgData)

    logging.info("ParamDataList -", ParamDataList)
    return ParamDataList


def removeDuplicates_old(reslist):
    newlist = []
    listwithpos = []
    for lt, pos, pc in reslist:
        if lt not in newlist:
            newlist.append(lt)
            listwithpos.append((lt, pos))
    return listwithpos


def ComparePC_and_removeDuplicates(reslist, dcipc):
    newlist = []
    listwithpos = []
    matchedBoolval = []
    for lt, pos, pc in reslist:
        logging.info(lt)
        if lt not in newlist:
            newlist.append(lt)
            listwithpos.append((lt, pos, pc))
        else:
            logging.info(lt, pos, pc, " =>lt,pos,pc")
            for lwp, lwpos, lwpc in listwithpos:
                if lwp == lt:
                    if (lwp, lwpos, lwpc) not in matchedBoolval:
                        matchedBoolval.append((lwp, lwpos, lwpc))
                    if pc != lwpc:
                        if dcipc != lwpc:
                            logging.info("\nNot Match")
                            if dcipc == pc:
                                newlist.append(lt)
                                listwithpos.append((lt, pos, pc))
                                # removing the existing value which is not matched with dci PC
                                newlist.remove(lwp)
                                listwithpos.remove((lwp, lwpos, lwpc))
                        logging.info(lwp, lwpos, lwpc, "- lwp,lwpos,lwpc")
                        matchedBoolval.append((lt, pos, pc))

    logging.info("\n newlist - ", newlist)
    logging.info("\n listwithpos - ", listwithpos)
    return listwithpos


def raiseQIA(reslist):
    found = 0
    for ind, lst in enumerate(reslist):
        req, frame, ver = lst[0]
        if (req == True and frame == True and ver == True):
            found = 1
            return 0, lst[1]
    if found == 0:
        for ind, lst in enumerate(reslist):
            req, frame, ver = lst[0]
            if (req == True and frame == True and ver == False):
                found = 1
                return 1, lst[1]
    if found == 0:
        for ind, lst in enumerate(reslist):
            req, frame, ver = lst[0]
            if (req == False and frame == True and ver == False):
                found = 1
                return 2, lst[1]
    if found == 0:
        for ind, lst in enumerate(reslist):
            req, frame, ver = lst[0]
            if (req == True and frame == False and ver == True) or (req == True and frame == False and ver == False):
                found = 1
                return 3, lst[1]
    if found == 0:
        for ind, lst in enumerate(reslist):
            req, frame, ver = lst[0]
            if (req == False and frame == False and ver == False):
                found = 1
                return 4, lst[1]


def findQIACondition(sheet, searchResult, pgData):
    reqDCINTVer = ''
    reqFORMERVer = ''
    qiaConditionFinalList = []
    if searchResult["count"] > 0:
        logging.info("\n!!!!!!!!!!!! Got signal values !!!!!!!!!!!!\n")
        cellpositionval = searchResult["cellPositions"]
        logging.info("cellpositionval - ", cellpositionval)
        for cellPosition in searchResult["cellPositions"]:
            isrequirement = 0
            isframe = 0
            isreqver = 0
            qiaConditionIterationList = []
            x, y = cellPosition
            logging.info("Signal Process x, y  - ", x, y)
            reqDCINT = re.sub(r'\([^)]*\)', "", str(sheet.range(x, 4).value))
            reqFORMER = re.sub(r'\([^)]*\)', "", str(sheet.range(x, 5).value))
            paramFrameName = str(sheet.range(x, 9).value)
            logging.info("\n\n------------ Processing the cell position" + str(cellPosition) + "------------")
            logging.info("pgData['reqIdrep'] - ", pgData['reqIdrep'])
            if pgData['reqIdrep'] in reqDCINT or pgData['reqIdrep'] in reqFORMER:
                logging.info("Requirement Matched")
                qiaConditionIterationList.append(True)
            else:
                logging.info("Requirement not matched")
                qiaConditionIterationList.append(False)

            if pgData['dciInfo']['framename'] == paramFrameName:
                logging.info("Frame Matched")
                qiaConditionIterationList.append(True)
            else:
                logging.info("Frame not matched")
                qiaConditionIterationList.append(False)

            if (reqDCINT != "'--" and reqDCINT != "--" and reqDCINT is not None):
                dcintreq = str(sheet.range(x, 4).value).split("|")
                if len(dcintreq) > 0:
                    for dcreq in dcintreq:
                        logging.info("dci_req -", dcreq)
                        if re.sub(r'\([^)]*\)', "", dcreq) == pgData['reqIdrep']:
                            if dcreq.find('(') != -1:
                                reqDCINTVer = dcreq.split("(")[1].split(")")[0]
                                reqDCINTVer = int(reqDCINTVer)
                            else:
                                reqDCINTVer = ''

            if (reqFORMER != "'--" and reqFORMER != "--" and reqFORMER is not None):
                formerreq = str(sheet.range(x, 5).value).split("|")
                if len(formerreq) > 0:
                    for freq in formerreq:
                        logging.info("former_req -", freq)
                        if re.sub(r'\([^)]*\)', "", freq) == pgData['reqIdrep']:
                            if freq.find('(') != -1:
                                reqFORMERVer = freq.split("(")[1].split(")")[0]
                                reqFORMERVer = int(reqFORMERVer)
                            else:
                                reqFORMERVer = ''


            if reqDCINTVer == pgData['reqversion'] or reqFORMERVer == pgData['reqversion']:
                logging.info("Version Matched")
                qiaConditionIterationList.append(True)

            else:
                logging.info("Version Not Matched")
                qiaConditionIterationList.append(False)
            qiaConditionFinalList.append((qiaConditionIterationList, cellPosition, str(sheet.range(x, 10).value)))
            logging.info("qiaConditionFinalList", qiaConditionFinalList)

    return ComparePC_and_removeDuplicates(qiaConditionFinalList, pgData['dciInfo']['pc'])


def generateQIAData(sheet, raiseQIAResult, dataDic):
    x, y = raiseQIAResult[1]
    missedParamValueFlag = raiseQIAResult[0]
    fver = 0
    dciver = 0
    modified_req = ''
    reqDCINTVer = ''
    reqFORMERVer = ''
    reqDCINT = str(sheet.range(x, 4).value)
    reqFORMER = str(sheet.range(x, 5).value)

    logging.info("missedParamValueFlag - ", missedParamValueFlag)
    if missedParamValueFlag == 1:
        if (reqDCINT != "'--" and reqDCINT != "--" and reqDCINT is not None):
            dcintreq = reqDCINT.split("|")
            if len(dcintreq) > 0:
                for dcreq in dcintreq:
                    logging.info("dci_req -", dcreq)
                    if re.sub(r'\([^)]*\)', "", dcreq) == dataDic['reqIdrep']:
                        reqDCINTVer = dcreq.split("(")[1].split(")")[0]
                        reqDCINTVer = int(reqDCINTVer)
                        dciver = 1

        if (reqFORMER != "'--" and reqFORMER != "--" and reqFORMER is not None):
            formerreq = reqFORMER.split("|")
            if len(formerreq) > 0:
                for freq in formerreq:
                    logging.info("former_req -", freq)
                    if re.sub(r'\([^)]*\)', "", freq) == dataDic['reqIdrep']:
                        reqFORMERVer = freq.split("(")[1].split(")")[0]
                        reqFORMERVer = int(reqFORMERVer)
                        fver = 1
        logging.info("reqversion-", dataDic['reqversion'])
        logging.info("reqDCINTVer-", reqDCINTVer)
        logging.info("reqFORMERVer-", reqFORMERVer)
        if dciver == 1:
            modified_req = re.sub(r'\([^)]*\)', "", str(sheet.range(x, 4).value)) + "(" + str(
                dataDic['reqversion']) + ")"
        else:
            modified_req = re.sub(r'\([^)]*\)', "", str(sheet.range(x, 5).value)) + "(" + str(
                dataDic['reqversion']) + ")"

        dataDic['modified_req'] = modified_req
        dataDic['comment'] = "Updating the requirement version"
        dataDic['requestType'] = 2
        dataDic['paramType'] = missedParamValueFlag
        dataDic['actual_req'] = dataDic['dciInfo']['dciReq']
        BL.displayInformation("\nRaising QIA for Updating the requirement version")
        addQIAresult = callQIAFunc(dataDic)

    if missedParamValueFlag == 2:
        if reqDCINT == "'--" or reqDCINT == "--" or reqDCINT is None:
            modified_req = dataDic['dciInfo']['dciReq']
        else:
            modified_req = reqDCINT + "|" + dataDic['dciInfo']['dciReq']
        logging.info("modified_req - ", modified_req)

        dataDic['modified_req'] = modified_req
        dataDic['comment'] = "Updating the requirement"
        dataDic['requestType'] = 2
        dataDic['paramType'] = missedParamValueFlag
        dataDic['actual_req'] = dataDic['dciInfo']['dciReq']
        BL.displayInformation("\nRaising QIA for Updating the requirement")
        addQIAresult = callQIAFunc(dataDic)

    if missedParamValueFlag == 3:
        if dataDic['reqIdrep'] in reqDCINT:
            modified_req = reqDCINT
        else:
            modified_req = reqFORMER

        # dataDic['modified_req'] = modified_req
        dataDic['modified_req'] = dataDic['dciInfo']['dciReq']
        dataDic['comment'] = "Adding the new frame"
        dataDic['requestType'] = 1
        dataDic['paramType'] = missedParamValueFlag
        dataDic['actual_req'] = dataDic['dciInfo']['dciReq']
        BL.displayInformation("\nRaising QIA for Adding the new frame")
        addQIAresult = callQIAFunc(dataDic)

    if missedParamValueFlag == 4:
        dataDic['modified_req'] = dataDic['dciInfo']['dciReq']
        dataDic['comment'] = "Adding the new requirement and new frame"
        dataDic['requestType'] = 1
        dataDic['paramType'] = missedParamValueFlag
        dataDic['actual_req'] = dataDic['dciInfo']['dciReq']
        BL.displayInformation("\nRaising QIA for Adding the new frame and new requirement")
        addQIAresult = callQIAFunc(dataDic)

    return addQIAresult


def callQIAFunc(QIAData):
    logging.info("\nQIAData = ", QIAData)
    reqtype = ""
    if QIAData['dciInfo']['proj_param'] != "" and QIAData['dciInfo']['proj_param'] is not None and QIAData['dciInfo'][
        'proj_param'] != "None":
        Nom_du_SO = getDCIProjParam(QIAData['dciInfo']['proj_param'])
    else:
        Nom_du_SO = "NEA"
    logging.info("Nom_du_SO - ", Nom_du_SO)

    if QIAData['requestType'] == 1:
        reqType = "Creation"
    else:
        reqType = "Modification"

    qia_dci_signal = QIAData['dciInfo']['dciSignal']
    dci_networkk = ""

    if QIAData['paramType'] == 3 or QIAData['paramType'] == 4:
        splitted_ntw = QIAData['dciInfo']['network'].split("_")
        if str(QIAData['dciInfo']['network']).upper().find('CAN') != -1:
            dci_networkk = splitted_ntw[1] if len(splitted_ntw) > 1 else ""
        elif str(QIAData['dciInfo']['network']).upper().find('LIN') != -1:
            dci_networkk = splitted_ntw[0] if len(splitted_ntw) > 0 else ""
        # qia_dci_signal = str(QIAData['dciInfo']['dciSignal'])+"_"+str(QIAData['dciInfo']['network'])
        qia_dci_signal = f"{QIAData['dciInfo']['dciSignal']}_{dci_networkk}"

    qia_comment = f"Information can be found in the DCI document ({QIAData['dciInfo']['dci_ref_num']}) ({QIAData['dciInfo']['dci_ver']})."

    QIA_Data = {"TP_Refnum": QIAData['TPRef_num'], "taskName": QIAData['FuncName'],
                "Req_type": reqType, "Expl": QIAData['comment'],
                "Nom_du_SO": Nom_du_SO, "columnG": "--",
                "signal": qia_dci_signal, "newreq": QIAData['modified_req'],
                "dciframe": QIAData['dciInfo']['framename'], "flowtype": QIAData['dciInfo']['pc'],
                "trigram": ICF.gettrigram(), 'qiacomment': qia_comment}
    # "dci_ref":str(QIAData['dciInfo']['dci_ref_num']),"dci_ver": str(QIAData['dciInfo']['dci_ver'])
    QIA_Resp = addQIADataInQIASheet(None, QIA_Data)

    if QIA_Resp == 1:
        BL.displayInformation(f"\n## QIA raised successfully for the requirement {QIAData['actual_req']}, File saved in Input Folder ##")
    else:
        return -1

    return 1


archi_formats = [{'R1': 'R1.0,R1,R1_0'}, {'R2': 'R2,R1.1,R1_1,R2.0'}, {'R3': 'R1.2,R1_2,R3.0,R3'}]
archi_values = [{'R1R2': 'NEA_R1|NEA_R1_1', 'R1R2R3': 'NEA_R1_X', 'R2R3': 'NEA_R1_X', 'R1': 'NEA_R1', 'R2': 'NEA_R1_1',
                 'R3': 'NEA_R1_2', 'R1R3': 'NEA_R1|NEA_R1_2'}]
architectures = ['R1', 'R2', 'R3']


def get_archi_val(archi, arch_keys, archi_formats):
    for key in arch_keys:
        for arch in archi_formats:
            if key in arch.keys():
                split_archi = arch[key].split(',')
                logging.info("split_archi --> ", split_archi)
                if archi in split_archi:
                    archi_key = key
    return {"archii": archi_key}


def getDCIProjParam(projParam):
    splitProjparam = projParam.split(",")
    reqProjParam = ''
    lpcount = 0
    fvalsplit = ''
    archiList = []
    if len(splitProjparam) > 1:
        for sp in splitProjparam:
            logging.info("\n\nsp - " + sp + "\n")
            logging.info("lpcount - ", lpcount)
            x = re.findall(
                "_NEAR[1-2]{1}.[0-2]{1}$|_NEAR[1-3]{1}$|_NEA_R[1-3]{1}$|_NEA_R[1-2]{1}.[0-2]{1}$|_NEA_R[1-2]{1}_[0-2]{1}$",
                sp)
            logging.info("\n \n x >>", x)
            if x:
                split_archi = re.findall(r"R[1-3]{1}.[0-3]{1}$|R[1-3]{1}_[0-3]{1}$|R[1-3]{1}", x[0])
                logging.info("split_archi -->> ", split_archi)
                if split_archi:
                    archi_data = get_archi_val(split_archi[0], architectures, archi_formats)
                    if archi_data['archii'] not in archiList:
                        archiList.append(archi_data['archii'])
                    logging.info("\narchiList --> ", archiList)
                    # reqProjParam += "|" + archi_data['archi_val'] if lpcount > 0 else archi_data['archi_val']
            lpcount = lpcount + 1
        if archiList:
            sorted_archiList = archiList.sort()
            combine_archi = "".join(archiList)
            logging.info("combine_archi -- ", combine_archi)
            for val in archi_values:
                if combine_archi in val.keys():
                    reqProjParam += val[combine_archi]
        logging.info("final1 - ", reqProjParam)

    else:
        x = re.findall(
            "_NEAR[1-2]{1}.[0-2]{1}$|_NEAR[1-3]{1}$|_NEA_R[1-3]{1}$|_NEA_R[1-2]{1}.[0-2]{1}$|_NEA_R[1-2]{1}_[0-2]{1}$",
            splitProjparam[0])
        logging.info("\n\n x >> ", x, "\n")
        if x:
            split_archi = re.findall(r"R[1-3]{1}.[0-3]{1}$|R[1-3]{1}_[0-3]{1}$|R[1-3]{1}", x[0])
            logging.info("split_archi -->> ", split_archi)
            if split_archi:
                archi_data = get_archi_val(split_archi[0], architectures, archi_formats)
                if archi_data['archii'] not in archiList:
                    archiList.append(archi_data['archii'])
                logging.info("\narchiList2 --> ", archiList)

            if archiList:
                sorted_archiList = archiList.sort()
                combine_archi = "".join(archiList)
                logging.info("combine_archi -- ", combine_archi)
                for val in archi_values:
                    if combine_archi in val.keys():
                        reqProjParam += val[combine_archi]

        logging.info("final2 - ", reqProjParam)

    return reqProjParam


def addQIADataInQIASheet(param_refnum, QIAData):
    logging.info("\nProcessing the QIA Sheet!!!\n")
    logging.info("QIAData - ", QIAData)
    path = ICF.getInputFolder() + "\\" + EI.findInputFiles()[10]
    isQAFileExist = os.path.isfile(path)
    Arch = BL.getArch(ICF.FetchTaskName())
    if not isQAFileExist:
        if (Arch == 'BSI'):
            BL.displayInformation("Please wait downloading the QIA of Param Global...")
            startDocumentDownload([("01642_23_00140", "")])
            BL.displayInformation("Download Completed...")
        else:
            BL.displayInformation("Please wait downloading the QIA of Param Global...")
            startDocumentDownload([("00949_11_06142", "")])
            BL.displayInformation("Download Completed...")
    qia_sheet_name = 'NEW_QIA'
    path = ICF.getInputFolder() + "\\" + EI.findInputFiles()[10]
    isQAFileExist = os.path.isfile(path)
    logging.info("QA File Exist- ", isQAFileExist)
    Today = date.today()
    curr_date = Today.strftime("%m/%d/%y")
    # 0 means no problem with saving file
    # 1 means problem with saving file
    saving_flag = 0
    if isQAFileExist:
        QIABook = EI.openExcel(path)
        if QIABook:
            qia_sheet = ''
            is_sheet_exist, qia_sheet = EI.findSheetInBook(QIABook, qia_sheet_name)
            logging.info(f"is_sheet_exist -> {is_sheet_exist}")
            if is_sheet_exist == 1 and qia_sheet != '':
                QIAsheet = QIABook.sheets["NEW_QIA"]
                QIAsheet.activate()
                QIAmaxrow = QIAsheet.range('H' + str(QIAsheet.cells.last_cell.row)).end('up').row
                logging.info("QIAmaxrow - ", QIAmaxrow)
                maxrow_slno = QIAsheet.range(QIAmaxrow, 1).value
                logging.info("maxrow_slno - ", maxrow_slno)
                try:
                    trigram = QIAData["trigram"]
                    if trigram.upper().find('EXPLEO'):
                        trigram = trigram.upper().replace('LEO', "")
                    logging.info("-------Adding the row in QIA for requirement " + QIAData["newreq"] + "-------")
                    # QIAmaxrow + 1 is the new row next to max row
                    QIAsheet.range(QIAmaxrow + 1, 1).value = int(maxrow_slno+1) if maxrow_slno is not None and maxrow_slno != "" else ""
                    QIAsheet.range(QIAmaxrow + 1, 2).value = QIAData["TP_Refnum"]
                    QIAsheet.range(QIAmaxrow + 1, 3).value = QIAData["taskName"]
                    QIAsheet.range(QIAmaxrow + 1, 4).value = QIAData["Req_type"]
                    QIAsheet.range(QIAmaxrow + 1, 5).value = QIAData["Expl"]
                    QIAsheet.range(QIAmaxrow + 1, 6).value = QIAData["Nom_du_SO"]
                    QIAsheet.range(QIAmaxrow + 1, 7).value = QIAData["columnG"]
                    QIAsheet.range(QIAmaxrow + 1, 8).value = QIAData["signal"].replace("$", "")
                    QIAsheet.range(QIAmaxrow + 1, 9).value = QIAData["newreq"]
                    QIAsheet.range(QIAmaxrow + 1, 10).value = "--"
                    QIAsheet.range(QIAmaxrow + 1, 11).value = "--"
                    QIAsheet.range(QIAmaxrow + 1, 12).value = "--"
                    QIAsheet.range(QIAmaxrow + 1, 13).value = "--"
                    QIAsheet.range(QIAmaxrow + 1, 14).value = QIAData["dciframe"]
                    QIAsheet.range(QIAmaxrow + 1, 15).value = QIAData["flowtype"]  # P|C
                    QIAsheet.range(QIAmaxrow + 1, 16).value = trigram
                    QIAsheet.range(QIAmaxrow + 1, 17).value = curr_date
                    QIAsheet.range(QIAmaxrow + 1, 18).value = 'Open'
                    QIAsheet.range(QIAmaxrow + 1, 21).value = f"{trigram} {curr_date}: {QIAData['qiacomment']}"
                    try:
                        QIABook.save()
                    except:
                        print("\nConflict with file while saving, please save QIA Param Global file manually.")
                        saving_flag = 1
                except:
                    BL.displayInformation("\n----------!!!!!!Error in adding new row!!!!!!----------\n")
                    pass
                    if saving_flag != 1:
                        QIABook.close()
                    return -1
            else:
                logging.info("#NEW_QIA sheet not present in QIA Param Global....")
                BL.displayInformation(f"\n### sheet name '{qia_sheet_name}' not present... ###")
                if saving_flag != 1:
                    QIABook.close()
                return -1

            if saving_flag != 1:
                QIABook.close()
    else:
        BL.displayInformation("QIA Sheet not exist in input folder..")
        return -1

    return 1


# Function to search & replace the signal|flow in TestSheet of the respective requirement
# type 1 means replace the existing signal with param global signal
# type 2 means concatenate the frame with existing signal

def searchSignalandReplaceinTS(TestSheetname, reqIdrep, dciInfo, type, paramdata):
    logging.info("\n-------------Replacing the signal in Test Sheet (" + TestSheetname + ")-------------\n")
    data_dic = {
        "new_signal": "",
        "test_sheet_name": "",
    }
    tsreqIDs = []
    reptext = ''
    tpBook = EI.openTestPlan()
    sheet = tpBook.sheets[TestSheetname]
    sheet.activate()
    macro = EI.getTestPlanAutomationMacro()
    TPM.selectTpWritterProfile(macro)
    TPM.unProtectTestSheet(macro)
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end(
        'up').row
    getreq = sheet.range(4, 3).value
    signal = str(dciInfo['dciSignal']).strip()
    logging.info("TS Requirements - ", getreq)
    logging.info("DCI signal for test sheet compare- ", signal)
    if reqIdrep in getreq:
        logging.info("Flow Type ==>", type)
        if type==1:
            logging.info("\n*******Replacing the Flow*******\n")
            reptext = paramdata["PARAM_SIGNAL"]
            data_dic['new_signal'] = str("$" + reptext)
            data_dic['test_sheet_name'] = str(TestSheetname)
            return data_dic
        else:
            logging.info("\n*******Appending the Network with Flow*******\n")
            if paramdata["Signal_Exist_NoFrame"]==1:
                reptext = dciInfo['network']
            else:
                reptext = paramdata["FLUX_MESSAGERIE_NEA"].split("/")[0]
            newvalue = signal + "_" + reptext
            data_dic['new_signal'] = str("$" + newvalue)
            data_dic['test_sheet_name'] = str(TestSheetname)
            return data_dic
        logging.info("DCI N/w- ", dciInfo['network'])
        logging.info("Replace Text- ", reptext)
    else:
        logging.info("-----------------Requirement (" + reqIdrep + ") not present in Testsheet-----------------")
    return 1


# def searchSignalandReplaceinTS(TestSheetname, reqIdrep, dciInfo, type, paramdata):
#     logging.info("\n-------------Replacing the signal in Test Sheet (" + TestSheetname + ")-------------\n")
#     data_dic = {
#         "new_signal": "",
#         "test_sheet_name": "",
#     }
#     tsreqIDs = []
#     reptext = ''
#     tpBook = EI.openTestPlan()
#     sheet = tpBook.sheets[TestSheetname]
#     sheet.activate()
#     macro = EI.getTestPlanAutomationMacro()
#     TPM.selectTpWritterProfile(macro)
#     TPM.unProtectTestSheet(macro)
#     maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end(
#         'up').row
#     getreq = sheet.range(4, 3).value
#     signal = dciInfo['dciSignal']
#     logging.info("TS Requirements - ", getreq)
#     logging.info("DCI signal for test sheet compare- ", signal)
#     if reqIdrep in getreq:
#         logging.info('fvf')
#         for row in range(11, maxrow + 1):
#             if (dciInfo["pc"] == 'C'):
#                 cellValue = str(sheet.range(row, 5).value)
#             else:
#                 cellValue = str(sheet.range(row, 11).value)
#             if signal in cellValue:
#                 logging.info("Flow Type ==>", type)
#                 if type == 1:
#                     logging.info("\n*******Replacing the Flow*******\n")
#                     reptext = paramdata["PARAM_SIGNAL"]
#                     # if (dciInfo["pc"] == 'C'):
#                     data_dic['new_signal'] = str("$" + reptext)
#                     data_dic['test_sheet_name'] = str(TestSheetname)
#                     return data_dic
#                     #     sheet.range(row, 5).value = "$" + reptext
#                     # else:
#                     #     sheet.range(row, 11).value = "$" + reptext
#                 else:
#                     logging.info("\n*******Appending the Network with Flow*******\n")
#                     if paramdata["Signal_Exist_NoFrame"] == 1:
#                         reptext = dciInfo['network']
#                     else:
#                         reptext = paramdata["FLUX_MESSAGERIE_NEA"].split("/")[0]
#                     newvalue = cellValue + "_" + reptext
#                     # if (dciInfo["pc"] == 'C'):
#                     data_dic['new_signal'] = str(newvalue)
#                     data_dic['test_sheet_name'] = str(TestSheetname)
#                     return data_dic
#                     #     sheet.range(row, 5).value = newvalue
#                     # else:
#                     #     sheet.range(row, 11).value = newvalue
#                     # sheet.range(row, 5).value = newvalue
#                 logging.info("DCI N/w- ", dciInfo['network'])
#                 logging.info("Replace Text- ", reptext)
#     else:
#         logging.info("-----------------Requirement (" + reqIdrep + ") not present in Testsheet-----------------")
#     return 1


def getArchi(them):
    try:
        findThem = re.findall("[A-Z]{3}_[0-9]{2}", them)
        logging.info(f"findThem {findThem}")
        them_archi = []
        final_archi = ""
        archi_values = {'R1R2': 'NEA_R1|NEA_R1_1', 'R1R2R3': 'NEA_R1_X', 'R2R3': 'NEA_R1_1|NEA_R1_2', 'R1': 'NEA_R1',
                        'R2': 'NEA_R1_1', 'R3': 'NEA_R1_2', 'R1R3': 'NEA_R1|NEA_R1_2'}
        if len(findThem) == 1:
            if "LVM_01" in findThem or "LYQ_01" in findThem:
                them_archi.append(archi_values['R1'])
            elif "LVM_02" in findThem:
                them_archi.append(archi_values['R2'])
            elif "LVM_03" in findThem:
                them_archi.append(archi_values['R3'])
            elif "LYQ_02" in findThem:
                them_archi.append(archi_values['R2R3'])
        elif len(findThem) > 1:
            if "LVM_01" in findThem or "LYQ_01" in findThem:
                them_archi.append(archi_values['R1'])
            if "LVM_02" in findThem or "LYQ_02":
                them_archi.append(archi_values['R2'])
            if "LVM_03" in findThem:
                them_archi.append(archi_values['R3'])

        if them_archi:
            if archi_values['R1'] in them_archi and archi_values['R2'] in them_archi and archi_values[
                'R3'] in them_archi:
                final_archi = 'NEA_R1_X'
            else:
                if len(them_archi) > 1:
                    final_archi = "|".join(set(them_archi))
                else:
                    final_archi = them_archi[0]
        else:
            final_archi = 'NEA'

        logging.info(f"final_archi {final_archi}")
    except Exception as ex:
        logging.info(f"Error in finding archi for DID: {ex}")
    return final_archi


def getThemArchi(thematic_data):
    logging.info(f"Finding thematic architechture...... {thematic_data}")
    archi = ""
    if thematic_data != -1 and thematic_data != -2:
        if thematic_data['effectivity'] != "":
            archi = getArchi(thematic_data['effectivity'])
        elif thematic_data['lcdv'] != "":
            archi = getArchi(thematic_data['lcdv'])
        elif thematic_data['diversity'] != "":
            archi = getArchi(thematic_data['diversity'])
        elif thematic_data['target'] != "":
            archi = getArchi(thematic_data['target'])

    return archi


def getReqVer(req):
    if req.find('(') != -1:
        new_reqName = req.split("(")[0].split()[0] if len(req.split("(")) > 0 else ""
        new_reqVer = req.split("(")[1].split(")")[0] if len(req.split("(")) > 1 else ""
    else:
        new_reqName = req.split()[0] if len(req.split()) > 0 else ""
        new_reqVer = req.split()[1] if len(req.split()) > 1 else ""
    return new_reqName, new_reqVer


def convertDID(extracted_DID):
    DID_formats = {'P0': 0, 'P1': 1, 'P2': 2, 'P3': 3, 'C0': 4, 'C1': 5, 'C2': 6, 'C3': 7, 'B0': 8, 'B1': 9, 'B2': 'A',
                   'B3': 'B', 'U0': 'C', 'U1': 'D', 'U2': 'E', 'U3': 'F'}
    splitted_DID = extracted_DID.split("-")
    didVal = splitted_DID[1]
    converted_DID = ""
    if didVal != '':
        code = didVal[:2]
        finalDID = DID_formats[code]
        logging.info(f"finalDID {finalDID}")
        if finalDID != "":
            logging.info(f"splitted_DID before {splitted_DID}")
            splitted_DID.remove(didVal)
            didVal = didVal.replace(code, str(finalDID))
            splitted_DID.insert(1, didVal)
            converted_DID = '-'.join(splitted_DID)

    return converted_DID


def findFlowFromContent(content):
    # finding the Flow/Signal from the content
    # if signal not present in content text then it will check the nested flow table inside the content

    splitted_content = content.split('\n')
    logging.info(f"splitted_content {splitted_content}")
    findFlow = []
    flow_name = ""
    for sp_cont in splitted_content:
        if sp_cont != "":
            if sp_cont.lower().find('involved flow:') != -1 or sp_cont.lower().find(
                    'involved flow :') != -1 or sp_cont.lower().find('flow :') != -1:
                logging.info(f"sp_cont {sp_cont}")
                findFlow = re.findall(":(.*)", sp_cont)
                logging.info(f"findFlow {findFlow}")
    if findFlow:
        flow_name = findFlow[0]
    else:
        logging.info("######## Flow not present in content need to check in nested flow table............")

    return flow_name


def identify_did_type():
    # Identifying the type Read or Write from the DID Excel file
    logging.info("Identifying the type for DID")
    try:
        DID_File = EI.findInputFiles()[13]
        if DID_File != "" and DID_File != None:
            DID_Book = EI.openExcel(ICF.getInputFolder() + "\\" + DID_File)

    except Exception as e:
        logging.info(f"Error in identify DID type: {e}")


def split_did_with_dot(DIDVal):
    # separating the every two characters of DID value with dot(.)
    didVal = '-'.join(DIDVal)
    didVal = didVal.replace('-', "")
    didVal = didVal.replace('VSM', "")
    didVal = '.'.join(didVal[i:i + 2] for i in range(0, len(didVal), 2))

    return didVal


def get_ref_num_from_doc(docName):
    # it returns the reference number which present in the input doc
    # return empty if not present
    refNumber = re.findall(r"\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]|[0-9]{5}\_[0-9]{2}\_[0-9]{5}", docName)
    refNum = refNumber[0].strip("[]") if refNumber else ""
    logging.info("refNum ->> ", refNum)
    return refNum


def sample(value):
    global didTypeVal
    didTypeVal = value
    logging.info(f">> didTypeVal>> {didTypeVal}")


def qia_for_DID(docName, requirement, reqData, testPlanReference, functionName):
    # It will get the all respective data for each column and raise QIA of param global
    # It returns 1 if successfully raise else returns -1 any error means or returns -2 if qia file not exist

    logging.info("\nProcessing the qia for DID.......")
    slNos = []
    DID_value = ""
    DID_type = ''
    DIDVal = ''
    try:
        # 0 means no problem with saving file
        # 1 means problem with saving file
        saving_flag = 0
        did_qia_col_value = {'A': "", 'B': "", "C": "", "D": "Creation", "E": "New DIAG for NEA", "F": "", "G": "--",
                             "H": "", "I": "", "J": "--", "K": "", "L": "", "M": "--", "N": "--", "O": "--", "Q": "",
                             "R": "Open", "S": "", "T": "", "U": ""}

        logging.info(f"EI.findInputFiles()[10] {EI.findInputFiles()[10]}")
        isQIAParamExist = os.path.isfile(ICF.getInputFolder() + "\\" + EI.findInputFiles()[10])
        logging.info("isQIAParamExist- ", isQIAParamExist)
        if isQIAParamExist:
            Today = date.today()
            curr_date = Today.strftime("%m/%d/%y")
            trigram = ICF.gettrigram()
            flowPrefix = ['REQ_', 'REP_']
            did_qia_col_value['B'] = testPlanReference
            did_qia_col_value['C'] = functionName
            did_qia_col_value['P'] = trigram
            did_qia_col_value['Q'] = curr_date
            req, ver = getReqVer(requirement)
            logging.info(f"reqData123 --> {reqData}")
            if type(reqData) is not dict:
                reqDataCond = "reqData.strip() != -1 and reqData.strip() != -2 and reqData.strip() != "" and reqData.strip() is not None"
            else:
                reqDataCond = True
            if reqDataCond:
                QIA_Param_Book = EI.openExcel(ICF.getInputFolder() + "\\" + EI.findInputFiles()[10])
                logging.info(f"QIA_Param_Book123 {QIA_Param_Book}")
                if QIA_Param_Book:
                    qia_sheet = ''
                    is_sheet_exist, qia_sheet = EI.findSheetInBook(QIA_Param_Book, 'NEW_QIA')
                    logging.info(f"is_sheet_exist -> {is_sheet_exist}")
                    if is_sheet_exist == 1 and qia_sheet != '':
                        QIAsheet = QIA_Param_Book.sheets[qia_sheet]
                        QIAsheet.activate()
                        them_arch = getThemArchi(reqData)
                        did_qia_col_value['F'] = them_arch
                        logging.info(f"\n\n.................them_arch {them_arch}..................")
                        req_flow = findFlowFromContent(reqData['content'])
                        did_qia_col_value['I'] = f"{req}({ver})"

                        if re.search("VSM-[A-Za-z0-9]{4,8}", reqData['comment']) and not re.search("VSM-[A-Z]{1,2}[A-Za-z0-9]{1,8}-[0-9]{2}", reqData['comment']):
                            # VSM-[A-Z]{3}[0-9]{2}
                            extracted_DID = re.findall("VSM-[A-Za-z0-9]{4,8}", reqData['comment'])
                            logging.info(f"extracted_DID1 {extracted_DID}")
                            DID_value = extracted_DID[0]
                        else:
                            # VSM-U0131-81
                            extracted_DID = re.findall("VSM-[A-Z]{1,2}[A-Za-z0-9]{1,8}-[0-9]{2}", reqData['comment'])
                            logging.info(f"extracted_DID2 {extracted_DID}")
                            if extracted_DID:
                                DID_value = convertDID(extracted_DID[0])

                        if DID_value != "" and DID_value is not None:
                            DIDVal = split_did_with_dot(DID_value)

                            docRefNum = get_ref_num_from_doc(docName)
                            # DID_type = 'Read'
                            didTypeValFlag = False
                            didInfo = {'req': requirement, 'didVal': DID_value, 'cerebro': "cerebro.inetpsa.com/data-identifier-list"}
                            DID_type = confirmationPop(didInfo)
                            # loop until the result is available
                            while not didTypeVal:
                                # block for a moment
                                logging.info("sleep.....")
                                time.sleep(2)
                            DID_type = didTypeVal
                            logging.info(f"DID_type>>> {DID_type}  ..................")
                            if DID_type != "" and DID_type != None:
                                didTypeValFlag = True
                            if didTypeValFlag:
                                if docRefNum != '' and docRefNum is not None:
                                    did_qia_col_value[
                                        'U'] = f"{trigram} {curr_date} Information can be found on SSD Ref : {docRefNum}."

                        if req_flow != "" and req_flow is not None:
                            for flow_req_res in flowPrefix:
                                if DID_type != '' and DIDVal != '':
                                    if DID_type.lower() == 'read':
                                        if flow_req_res.strip() == 'REQ_':
                                            did_qia_col_value['K'] = f"22.{DIDVal}"
                                        else:
                                            did_qia_col_value['K'] = f"62.{DIDVal}"
                                    else:
                                        if flow_req_res.strip() == 'REQ_':
                                            did_qia_col_value['K'] = f"2E.{DIDVal}"
                                        else:
                                            did_qia_col_value['K'] = f"6E.{DIDVal}"
                                did_qia_col_value['H'] = f"{flow_req_res}{str(req_flow).strip()}"
                                QIAmaxrow = QIAsheet.range('B' + str(QIAsheet.cells.last_cell.row)).end('up').row
                                logging.info("QIAmaxrow - ", QIAmaxrow)
                                if QIAsheet.range(QIAmaxrow, 1).value is not None and QIAsheet.range(QIAmaxrow, 1).value != '':
                                    did_qia_col_value['A'] = int(QIAsheet.range(QIAmaxrow, 1).value) + 1
                                UpdateQiaParamGlobal(QIAsheet, did_qia_col_value)
                                slNos.append(did_qia_col_value['A'])
                        try:
                            QIA_Param_Book.save()
                        except:
                            print("\nConflict with file while saving, please save QIA Param Global file manually.")
                            saving_flag = 1

                        if saving_flag != 1:
                            QIA_Param_Book.close()
                        logging.info(f"did_qia_col_value, slNos -->  {did_qia_col_value},, <{slNos}>.")

                        if slNos:
                            return 1, slNos
                    else:
                        logging.info(f"#New QIA sheet not present in QIA Param Global.....")

                # else:
                #     logging.info(f"############ Not able to get the content from document ############")
                #     return -3
        else:
            logging.info(f"############ QIA Param Global File not Exist ############")
            return -2
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        logging.info(f"\nError in QIA for DID: {ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
        return -1


def findDIDReq(reqDocData):
    # It finds out the processing requirement is DID or not
    # If it is 'DID' return 1 else return -1
    logging.info(f"\n\nFinding the DID ....")
    content = reqDocData['content']
    comment = reqDocData['comment']
    isDID = 0
    if content.lower().find('diagnostic') != -1:
        isDID = 1
    elif comment.upper().find('DID') != -1:
        isDID = 1

    return isDID


def handle_qia_of_did(inpdocList, requirement, testPlanReference, functionName):
    status_msg = ""
    dcoName = ""
    response, slno = "", ""
    try:
        for inputdoc in inpdocList:
            logging.info(f"inputdoc>> {inputdoc}")
            currVer = re.search("([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})", inputdoc)
            if currVer.group() is not None:
                version = currVer.group().upper()
                if "V" in version:
                    version = version.replace("V", "")
                    logging.info(str(int(float(version))))
                    currVer = str(int(float(version)))
            else:
                currVer = ""
            logging.info("Current Version = ", currVer)
            currDocPath = getInpDocPath(inputdoc, currVer)
            logging.info(f"currDocPath1 - {currDocPath}")
            # currDocPath = ICF.getInputFolder() + "\\" + inputdoc
            # currDocPath = ICF.getInputFolder() + "\\" + "[V21.0]CABIN_SSD_LIC_GEN2_00998_16_02483_v20_22Q2.docx"
            # logging.info(f"currDocPath2 - {currDocPath}")
            req, ver = getReqVer(requirement)
            if currDocPath != -1:
                newReqContent = WDI.getReqContent(currDocPath, req, ver)
                if newReqContent != -1 and newReqContent is not None and newReqContent != "":
                    dcoName = inputdoc
                    break
            else:
                logging.info(f"##### Could not find the doc path {inputdoc} #####")
                status_msg = f"##### Could not find the doc path {inputdoc} #####"
        if dcoName != "" and newReqContent != "":
            isDID = findDIDReq(newReqContent)
            if isDID == 1:
                logging.info(f"{requirement} is DID................")
                qia_response = qia_for_DID(dcoName, requirement, newReqContent, testPlanReference, functionName)
                if type(qia_response) is tuple:
                    response, slno = qia_response
                    logging.info(f"qia_response --> {qia_response}")
                    if response == 1:
                        status_msg = f"QIA of Param raised successfully..... SlNo's: {slno}"
                elif qia_response == -2:
                    status_msg = f"QIA file not exist......"
                elif qia_response == -3:
                    status_msg = f"No content from document......"

                return {'response': response, 'slno': slno, 'status_msg': status_msg}
            else:
                logging.info(f"Requirement {requirement} is not DID.....")
                return 2
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"Error in handle_qia_of_did:{ex}{exc_tb.tb_lineno}")
        return -1


if __name__ == "__main__":
    ICF.loadConfig()
    # qia_for_DID("", "", "")
