import shutil

import datetime as datetime

date_time = datetime.datetime.now()
import KeyboardMouseSimulator as KMS
import InputConfigParser as ICF
import ExcelInterface as EI
import WordDocInterface as WDI
import AnaLyseThematics as AT
import AnalyseTestSheet as ATS
import InputDocLinkPopup as IDLP
import TestPlanMacros as TPM
import time
import ctypes
import os
import threading
import re
import difflib
from datetime import date, datetime
from os import listdir
from os.path import isfile, join
import difflib
from pathlib import Path
import sys
from web_interface import startDocumentDownload, configChromeVersion
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import shutil
from Backlog_Handler import getCombinedThematicLines
import copy
import QIA_Param as QP
from GatewayRequirementHandler import gatewayReq
from TestPlanReqOldToNew import reqNameChange
from renameReq import reqNameChanging
import NewRequirementHandler as NRH
import QIAParamCreateNewFrame as QIACNF
import QIAParamDTC as QPD
import QIA_Calibration as QCU
import QIA_PT as QPT
import DTC_Frame_Wired as DFW
import Diag_Req_Handler as DRH
import SS_fiche_evolved as SFE
import Check_report_generate as CRG
from QIA_PT_Interface_Requirements import execute_interface_requirement_treatment
import DocumentSearch as DS
import interface_replacing_or_adding_signal as IRAS
import logging
import ThematicApplicability as TA

# VSM and BSI pattern Reference link and input Documnet
pattren_ref = "([A-Z0-9]{4,5})+(_[0-9]{2})+(_[A-Z0-9]{4,5})+|([A-Z]{4})+(_[A-Z0-9]{5})+(_[0-9]{4})+|([A-Z0-9]{4})+_([A-Z0-9]{4})+_([A-Z0-9]{3})+_([0-9]{3})+"

# For pattern version
pattren_ver = "([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})"

ref_num_pattern = "([A-Z0-9]{4,5})+(_[0-9]{2})+(_[A-Z0-9]{4,5})+|([A-Z]{4})+(_[A-Z0-9]{5})+(_[0-9]{4})+|([A-Z0-9]{4})+_([A-Z0-9]{4})+_([A-Z0-9]{3})+_([0-9]{3})+|([A-Z0-9]{4})+_([A-Z0-9]{6})+_([0-9]{4})"

# Input doc reference number
inpdoc_refNum = "(([A-Z]+|[0-9]+){4,5})+_(([0-9]+){2})+_(([0-9]+){4,5})"

inpdoc_refNum_grp = "([A-Z0-9]{4,5})+(_[0-9]{2})+(_[0-9]{4,5})+|([A-Z]{4})+(_[A-Z0-9]{5})+(_[0-9]{4})+|([A-Z0-9]{4})+_([A-Z0-9]{4})+_([A-Z0-9]{3})+_([0-9]{3})+ "

Thematics = "[a-zA-Z]{3}_[0-9]{2}"

rx_refential_VSM = "[A-Za-z0-9]+\s([0-9]{4})+_([0-9]{2})\s([0-9]{5})+(_[0-9]{2})+(_[0-9]{5})+\s+[A-Za-z0-9]+\.[0-9]+"
# Example for the VSM---> Referential 2022_10 00949_17_00725 V51.0

# rx_refential_BSI = "[[A-Za-z]+\s([A-Za-z0-9]+(_[A-Za-z0-9]+)+)\s([A-Za-z0-9]+(_[A-Za-z0-9]+)+)\s[A-Za-z0-9]*\.[0-9]+"
# Example for the BSI---> Referential 2022_10 AEEV_LEV07_0086 V102.0
rx_refential_BSI = r"[A-Za-z]+\s([A-Za-z0-9]+(_[A-Za-z0-9]+)+)\s([A-Za-z0-9]+(_[A-Za-z0-9]+)+)\sV[0-9]+(\.[0-9]+)?"

DeleteReqPattern = r'([a-zA-Z]+(\s+-->|-->|\s+--->|--->|\s+!|!)|\([a-zA-Z]+\)|DELETE|delete|[A-Za-z]{6})'

impactRow = 18
oldVersion = 0
WDI.oldVersion = 0


# funcImpactComment = 0
def save_as_docx(path):
    # Opening MS Word
    try:
        import win32com.client as win32
        from win32com.client import constants

        word = win32.gencache.EnsureDispatch('Word.Application')

    except AttributeError:
        logging.info("e")
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\\.gen_py\\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        import win32com.client as win32
        from win32com.client import constants
        word = win32.gencache.EnsureDispatch('Word.Application')

    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
    doc.Close(False)

    logging.info("DOCX File Path", os.path.realpath(new_file_abs))
    return os.path.realpath(new_file_abs)


def removeMisc(txt):
    op = re.sub("\([a-zA-Z0-9_\-\s]+\)", "", txt.strip())
    logging.info("OP Before = ", op)
    misc = re.findall("\([a-zA-Z0-9_\-\s]+\)", op)
    logging.info(len(misc))
    while len(misc) != 0:
        op = re.sub("\([a-zA-Z0-9_\-\s]+\)", "", op.strip())
        misc = re.findall("\([a-zA-Z0-9_\-\s]+\)", op)
    logging.info(op)
    return op


def removeRefVerFromFilename(fileName):
    del_ref = re.sub("([0-9]{5})+(_[0-9]{2})+(_[0-9]{5})+", "", fileName.strip())
    del_Ver = re.sub("[vV]{1}[0-9]{1,2}\.[0-9]{1,2}", "", del_ref.strip())
    del_Ver_1 = re.sub("[vV]{1}[0-9]{1,2}", "", del_Ver.strip())
    del_misc = removeMisc(del_Ver_1.strip())
    del_start_hyp = re.sub("^[_-]+", "", del_misc.strip())
    del_end_hyp = re.sub("[_-]+$", "", del_start_hyp.strip())
    final = del_end_hyp.strip()
    return final


def getDocPathQIA(docName, version=""):  # version is optional
    logging.info("\n\n\nIn getDocPath Function QIA")
    logging.info("Docname = ", docName)
    logging.info("version = ", version)
    global oldVersion
    path = -1
    pat = ICF.getInputFolder() + "\\"
    onlyfiles = [f for f in listdir(pat) if isfile(join(pat, f)) and (
                os.path.splitext(f)[1] == ".docx" or os.path.splitext(f)[1] == ".doc" or os.path.splitext(f)[
            1] == ".docm" or os.path.splitext(f)[1] == ".rtf")]
    logging.info("+++++++++++555 >>", pat)

    # logging.info("onlyfilesonlyfiles ", onlyfiles)
    # docName = removeRefVerFromFilename(docName)

    # splittedDocName = docName.split("-")[0]
    # if splittedDocName.find(".") != -1:
    #     docName = splittedDocName.split(".")[0]
    # else:
    #     docName = splittedDocName.strip()

    # for fileName in onlyfiles:
    if docName in onlyfiles:
        # if docName in fileName:
        logging.info("----------------===========----------")
        fileVer = re.search("[V]+[0-9+]+", docName)
        logging.info("fileVer ? ", fileVer)
        if fileVer is not None:
            fileVerRes = fileVer.group()
            logging.info("fileVerRes.upper().split()[1] = ", fileVerRes.upper().split("V")[1])
            if (fileVerRes.upper().split("V")[1]) == version:
                if os.path.splitext(docName)[1] == ".docx":
                    path = pat + docName
                    logging.info("Path found === ", path)
                    # break
                elif (os.path.splitext(docName)[1] == ".doc") or (os.path.splitext(docName)[1] == ".docm") or (
                        os.path.splitext(docName)[1] == ".rtf"):
                    logging.info(".doc Name ", docName)
                    path = save_as_docx(pat + docName)
                    oldVersion = 1
                    WDI.oldVersion = 1
                    # break
                else:
                    path = -2
            else:
                path = -2
        else:
            logging.info("Document " + docName + " is not having version in ipnut folder")
            path = -1
        # else:
        #     path = -1

    return path


def getDocPath(docName, version=""):  # version is optional
    logging.info("In getDocPath Function")
    logging.info("Docname = ", docName)
    logging.info("version = ", version)
    global oldVersion
    path = -1
    pat = ICF.getInputFolder() + "\\"
    onlyfiles = [f for f in listdir(pat) if isfile(join(pat, f))]
    logging.info("+++++++++++", pat, onlyfiles)

    docName = removeRefVerFromFilename(docName)
    if docName.find("-  (") != -1:
        docName = re.sub(r'\-.*', "", docName)
        logging.info(docName, "docName122")

    if re.search("\(.*?\)", docName):
        docName = re.sub("\(.*?\)", "", docName)

    docName = re.sub("(\s\-\s$)", "", docName)
    docName = re.sub("(\s\-$)", "", docName)
    logging.info(f"docName after removing junks-  {docName}")
    for fileName in onlyfiles:
        if os.path.splitext(fileName)[1] == ".docx" or os.path.splitext(fileName)[1] == ".doc" or os.path.splitext(fileName)[1] == ".docm" or os.path.splitext(fileName)[1] == ".rtf":
            if docName.strip() in fileName:
                logging.info("---Input Document Mactched---")
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
                        elif (os.path.splitext(fileName)[1] == ".doc") or (
                                os.path.splitext(fileName)[1] == ".docm") or (
                                os.path.splitext(fileName)[1] == ".rtf"):
                            logging.info(".doc Name ", fileName)
                            path = save_as_docx(pat + fileName)
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


def getDocPath_old(docName, version=""):  # version is optional
    logging.info("In getDocPath Function")
    logging.info("Docname = ", docName)
    logging.info("version = ", version)
    global oldVersion
    path = -1
    pat = ICF.getInputFolder() + "\\"
    onlyfiles = [f for f in listdir(pat) if isfile(join(pat, f))]
    logging.info("+++++++++++", pat, onlyfiles)

    docName = removeRefVerFromFilename(docName)
    if docName.find("-  (") != -1:
        docName = re.sub(r'\-.*', "", docName)
        logging.info(docName, "docName122")

    # splittedDocName = docName.split("-")[0]
    # if splittedDocName.find(".") != -1:
    #     docName = splittedDocName.split(".")[0]
    # else:
    #     docName = splittedDocName.strip()
    if re.search("\(.*?\)", docName):
        docName = re.sub("\(.*?\)", "", docName)
    logging.info(f"docName after1112 {docName}")
    # docName = docName.replace("-", "")
    docName = re.sub("(\s\-\s$)", "", docName)
    docName = re.sub("(\s\-$)", "", docName)
    logging.info(f"docName after1211 {docName}")
    for fileName in onlyfiles:
        # logging.info("\n\nFilename & docName = ", fileName,"\n", docName)
        if docName.strip() in fileName:
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
                        path = save_as_docx(pat + fileName)
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


def createStartThread(fName):
    thread = threading.Thread(target=fName)
    thread.start()


def getAddress(ref, ver=None):
    address = None
    if ref is not None:
        if ver is None:
            ver = "vc"
        else:
            pass
        address = "http://docinfogroupe.inetpsa.com/ead/doc/ref." + ref + "/v." + ver + "/fiche"
    return address


def getTaskDetails():
    return ICF.getTaskDetails()


def getReqDatafromImpact(tpBook, rowOfInterface, keyword):
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
            if sheet.range(i, 4).value is not None:
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


def getReferentiel():
    referentiel = []
    for task_detail in ICF.getTaskDetails():
        logging.info(task_detail)
        referentiel.append(task_detail["referentiel"])
        logging.info("Referentiel = " + str(referentiel))
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nReferentiel = " + str(referentiel))
    return referentiel


def getTaskName():
    taskName = []
    for task_detail in ICF.getTaskDetails():
        taskName.append(task_detail["taskName"])
        logging.info("Task Name = " + str(taskName))
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nTask Name = " + str(taskName))
    return taskName


def getTrigram():
    trigram = []
    for task_detail in ICF.getTaskDetails():
        trigram.append(task_detail["trigram"])
        logging.info("Trigram = " + str(trigram))
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nTrigram = " + str(trigram))
    return trigram


def launchDocInfo():
    logging.info("Launching Browser")
    path = '"' + ICF.getIEPath() + '"' + " " + ICF.getDocInfoUrl()
    os.system(path)


def launchAnalyseDeEntrant():
    path = '"' + ICF.getExcelPath() + '"' + " " + ICF.getAnalyseDeEntrant()
    logging.info(path)
    os.system(path)


def findFunctionSheet():
    return getTaskName()


def dontUpdateAnalyse():
    KMS.rightArrow()
    time.sleep(1)
    KMS.pressEnter()
    time.sleep(1)
    KMS.showWindow("Excel")


def getPTReference(sheet, colRow):
    return EI.getDataFromCell(sheet, colRow)


def fillSummary(tpBook, ipDocList, allFeps, referentiel, trigram):
    time.sleep(2)
    logging.info("fillSummary ipDocList = ", ipDocList)
    newIpDocList = ipDocList

    logging.info("tpBook.name.split('.')[0] = ", tpBook.name)
    logging.info("tpBook.name.split('.')[0] = ", tpBook.name.split('.')[0])

    KMS.showWindow(tpBook.name.split('.')[0])
    time.sleep(1)

    time.sleep(2)
    EI.activateSheet(tpBook, 'Sommaire')

    time.sleep(2)
    logging.info(EI.getDataFromCell(tpBook.sheets['Sommaire'], 'B6'))

    time.sleep(1)
    EI.setDataFromCell(tpBook.sheets['Sommaire'], 'B6', trigram)

    time.sleep(1)
    EI.setDataFromCell(tpBook.sheets['Sommaire'], 'C6', date.today().strftime('%d/%m/%Y'))

    time.sleep(1)
    ipDocName = ""
    ipDocString = ""
    if tpBook.sheets['Sommaire'].range(6, 1).merge_cells is True:
        cellRange = tpBook.sheets['Sommaire'].range(6, 1).merge_area
        rlo = cellRange.row
        rhi = cellRange.last_cell.row
        logging.info("length ", len(newIpDocList), rhi - rlo + 1, rlo, rhi)
    logging.info("list ", newIpDocList)
    for i in range(rlo, rhi):
        if tpBook.sheets['Sommaire'].range(i, 5).merge_cells is True:
            tpBook.sheets['Sommaire'].range(i, 5).unmerge()
        if tpBook.sheets['Sommaire'].range(i, 6).merge_cells is True:
            tpBook.sheets['Sommaire'].range(i, 6).unmerge()
        if tpBook.sheets['Sommaire'].range(i, 7).merge_cells is True:
            tpBook.sheets['Sommaire'].range(i, 7).unmerge()

    if (rhi - rlo + 1) < (len(newIpDocList)):
        logging.info("goin in...")
        for i in range((rhi - rlo + 1), len(newIpDocList)):
            tpBook.sheets['Sommaire'].range('7:7').insert(shift='down')
    elif (rhi - rlo + 1) == (len(newIpDocList)):
        pass
    else:
        logging.info("In else =--==--")
        for i in range(len(newIpDocList), (rhi - rlo + 1)):
            for j in range(1, 10):
                # logging.info("Deleting 6, ", j)
                tpBook.sheets['Sommaire'].range(6, j).delete()

        time.sleep(2)
        logging.info(EI.getDataFromCell(tpBook.sheets['Sommaire'], 'B6'))
        row = rhi - rlo + 1
        verCol = 'A' + str(6 + row - len(newIpDocList))
        logging.info(EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol))

        if EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol) is not None:
            EI.setDataFromCell(tpBook.sheets['Sommaire'], 'A6',
                               int(EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol)) + 1)

        logging.info("verCol = ", verCol)
        time.sleep(1)
        EI.setDataFromCell(tpBook.sheets['Sommaire'], 'B6', trigram)

        time.sleep(1)
        EI.setDataFromCell(tpBook.sheets['Sommaire'], 'C6', date.today().strftime('%d/%m/%Y'))

        verCol = 'A7'
        logging.info(EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol))

        if EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol) is not None:
            EI.setDataFromCell(tpBook.sheets['Sommaire'], 'A6',
                               int(EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol)) + 1)

    if tpBook.sheets['Sommaire'].range(6, 1).merge_cells is True:
        cellRange = tpBook.sheets['Sommaire'].range(6, 1).merge_area
        rlow = cellRange.row
        rhigh = cellRange.last_cell.row
        if tpBook.sheets['Sommaire'].range((rlow, 4), (rhigh, 4)).merge_cells is False:
            tpBook.sheets['Sommaire'].range((rlow, 4), (rhigh, 4)).merge()
        for i in range(8, 10):
            if tpBook.sheets['Sommaire'].range((rlow, i), (rhigh, i)).merge_cells is False:
                tpBook.sheets['Sommaire'].range((rlow, i), (rhigh, i)).merge()
            if tpBook.sheets['Sommaire'].range((rlow, i), (rhigh, i)).merge_cells is False:
                tpBook.sheets['Sommaire'].range((rlow, i), (rhigh, i)).merge()
        verCol = 'A' + str(rhigh + 1)
        if EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol) is not None:
            EI.setDataFromCell(tpBook.sheets['Sommaire'], 'A6',
                               int(EI.getDataFromCell(tpBook.sheets['Sommaire'], verCol)) + 1)
    else:
        logging.info("Not a Merge cell")
    rowx = rlo
    for ipDoc in newIpDocList:
        ipDocName = ipDocName + "-" + ipDoc + " \n"
        logging.info("ipDocName = ", ipDocName)
        logging.info("without group = ",
              #       pattren_ref="([0-9]{5})+(_[0-9]{2})+(_[0-9]{5})+"
              # pattren_ver = "([-|_|\s][vV]{1}[0-9]{1,2}.[0-9]{0,2})|([-|_|\s][vV]{1}[0-9]{1,2})"
              (re.search(inpdoc_refNum, ipDoc.strip())))
        regexGroup = (re.search(inpdoc_refNum_grp, ipDoc.strip()))

        EI.setDataFromCell(tpBook.sheets['Sommaire'], (rowx, 5), ipDoc.split(" -")[0])
        if regexGroup is not None:
            EI.setDataFromCell(tpBook.sheets['Sommaire'], (rowx, 6), regexGroup.group())
        else:
            EI.setDataFromCell(tpBook.sheets['Sommaire'], (rowx, 6), ipDocName)
        logging.info("!!! ipDoc !!! = ", ipDoc)
        verNumRe = re.search("[vV]{1}[0-9]{1,2}\.[0-9]{1,2}", ipDoc)
        verNum = verNumRe.group()
        # logging.info("!!! ipDoc Version split!!! = ", (ipDoc.upper().split("V")))
        # docVersionList = re.split(pattren_ver, ipDoc)
        # docVersionList[1] = docVersionList[1].replace("_", "")
        # docVersionList[1] = docVersionList[1].strip()
        # docVersionList[1] = docVersionList[1].replace("V", "")
        # docVersionList[1] = docVersionList[1].replace("v", "")
        # # docVersion = re.search("[0-9]{1,2}.[0-9]", (ipDoc.upper().split(" V")[1].split(" ")[0]))
        # docVersion = docVersionList[1]

        if verNum is not None:
            logging.info("!!! Version !!! = ", verNum)
            EI.setDataFromCell(tpBook.sheets['Sommaire'], (rowx, 7), (verNum))
        else:
            logging.info("document Version is None")
        # Add Hyperlink to latest version of input document
        refNum = re.search(inpdoc_refNum_grp, ipDoc.strip())

        if refNum is not None:
            tpBook.sheets['Sommaire'].range(rowx, 6).add_hyperlink(
                address=getAddress(refNum.group()),
                # address=getAddress(refNum),
                text_to_display=refNum.group(),
                screen_tip=None)
            # Add Hyperlink to specified version of input document logging.info("(ipDoc.upper().split('V')[1].split()[0]))
            # =>>>>>?????", (ipDoc.upper().split(" V")[1].split(" ")[0]))
            versionList = re.split(pattren_ver, ipDoc)
            tpBook.sheets['Sommaire'].range(rowx, 7).add_hyperlink(
                address=getAddress(
                    refNum.group(), ver=(re.sub(r'(v|V)', "", verNum))),
                text_to_display=(verNum),
                screen_tip=None)

        if rowx < (rlo + len(newIpDocList)):
            rowx = rowx + 1

    ipDocString = "Update of Test Plan according to:\n" + ipDocName + "-" + referentiel + "\n-" + allFeps
    logging.info("REFERENCIAL----->>>", referentiel)

    EI.setDataFromCell(tpBook.sheets['Sommaire'], 'D6', ipDocString)
    try:
        EI.setDataFromCell(tpBook.sheets['Sommaire'], (6, 8), (referentiel.split("-")[0]).split(" ")[1])
        EI.setDataFromCell(tpBook.sheets['Sommaire'], (6, 9), ("V" + str(float(referentiel.split(" V")[1]))))
        EI.setDataFromCell(tpBook.sheets['Sommaire'], (6, 10), (referentiel.split("-")[0]).split(" ")[1])

        # Add Hyperlink to latest version of referentiel
        # changed 1st parameter of getAddress() from referentiel.split("-")[1] to reff - cz
        # previous one was giving index outage error
        logging.info("/?????????/", re.search(rx_refential_VSM, referentiel))
        logging.info("/?????????/", re.search(rx_refential_BSI, referentiel))
        rx_reff = re.search(rx_refential_VSM, referentiel)
        rx_reff_BSI = re.search(rx_refential_BSI, referentiel)
        if rx_reff is not None:
            logging.info("rx_reff", rx_reff)
            reff = rx_reff.group()
            logging.info("Referential reference = ", reff)
            GUI_reff = re.search(pattren_ref, reff)
            GUI_ref = GUI_reff.group()
            logging.info("GUI_Referential reference = ", GUI_ref)
            tpBook.sheets['Sommaire'].range(6, 8).add_hyperlink(
                # address=getAddress(reff, (str(float(referentiel.split(" V")[1])))),
                address=getAddress(GUI_ref, (str(float(referentiel.split(" V")[1])))),
                text_to_display=referentiel.split("-")[0].split(" ")[1], screen_tip=None)

            # Add Hyperlink to specified version of referentiel
            tpBook.sheets['Sommaire'].range(6, 9).add_hyperlink(
                # address=getAddress(reff, (str(float(referentiel.split(" V")[1])))),
                address=getAddress(GUI_ref, (str(float(referentiel.split(" V")[1])))),
                text_to_display=("V" + str(float(referentiel.split(" V")[1]))), screen_tip=None)
            time.sleep(1)
        elif rx_reff_BSI is not None:
            logging.info("rx_reff_BSI", rx_reff_BSI)
            reff = rx_reff_BSI.group()
            logging.info("Referential reference = ", reff)
            GUI_reff = re.search(pattren_ref, reff)
            GUI_ref = GUI_reff.group()
            logging.info("GUI_Referential reference = ", GUI_ref)
            tpBook.sheets['Sommaire'].range(6, 8).add_hyperlink(
                # address=getAddress(reff, (str(float(referentiel.split(" V")[1])))),
                address=getAddress(GUI_ref, (str(float(referentiel.split(" V")[1])))),
                text_to_display=referentiel.split("-")[0].split(" ")[1], screen_tip=None)

            # Add Hyperlink to specified version of referentiel
            tpBook.sheets['Sommaire'].range(6, 9).add_hyperlink(
                # address=getAddress(reff, (str(float(referentiel.split(" V")[1])))),
                address=getAddress(GUI_ref, (str(float(referentiel.split(" V")[1])))),
                text_to_display=("V" + str(float(referentiel.split(" V")[1]))), screen_tip=None)
            time.sleep(1)
    except Exception as e:
        logging.info("Exception in fillSummary = ", e)
        time.sleep(1)
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines(
                "\n\nReferential link cannot be uploaded in summary tab as Referential Version is not mentioned in "
                "UserInput")
        pass
    logging.info("END")


def fillSheetHistory(sheet, keyword):
    logging.info(f"keywordHistory {keyword}")
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    sheet_value = sheet.used_range.value
    try:
        # cellValue = EI.searchDataInExcel(sheet, (26, maxrow), "Nature des modifications")
        cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow), "Nature des modifications")
    except:
        # cellValue = EI.searchDataInExcel(sheet, (26, maxrow), "Nature des modifications")
        cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow), "Nature des modifications")
    row, col = cellValue["cellPositions"][0]
    logging.info("In History", row, col)
    logging.info("sheet.range(row + 1, col).value", sheet.range(row + 1, col).value)
    if sheet.range(row + 1, col).value is not None:
        getString = sheet.range(row + 1, col).value + keyword
    else:
        getString = keyword

    EI.setDataFromCell(sheet, (row + 1, col), getString)


def fillHistoryAndTrigram(sheet, keyword):
    try:
        maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        sheet_value = sheet.used_range.value
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




def fillImpact(tpBook, listOfRequirements, fepsNum, rqIDs, fepsForDuplicateReqs):
    logging.info("fillImpact interface function rqIDs----->",rqIDs, fepsForDuplicateReqs)
    maxrow = tpBook.sheets['Impact'].range('A' + str(tpBook.sheets['Impact'].cells.last_cell.row)).end('up').row
    i = 0
    length = len(listOfRequirements)
    rowOfInterface = maxrow + 1
    time.sleep(2)
    EI.activateSheet(tpBook, 'Impact')
    logging.info("++++++++Value of ROW = " + str(maxrow))
    logging.info("++++++++Value of length = " + str(length))
    logging.info("++++++++listOfRequirements = ", listOfRequirements)
    condition = maxrow + length
    impactComment = ""
    while maxrow <= condition:
        A = 1
        B = 2
        F = 6
        E = 5
        while i < length:
            isDeleteReq = findDeleteReq(listOfRequirements[i])
            if isDeleteReq['is_delete_req'] == 1:
                if isDeleteReq['Ver'] == "":
                    listOfRequirements[i] = isDeleteReq['ModifiedReq']
                impactComment = "Deleted Requirement."

            if listOfRequirements[i].find("("):
                splitRequirements = listOfRequirements[i].split("(")
                Requirement = splitRequirements[0]
                logging.info(Requirement)
                if len(listOfRequirements[i].split("(")) > 1:
                    splitVersion = splitRequirements[1]
                    vers = splitVersion.split(")")
                    version = int(vers[0])
                    logging.info(version)
                else:
                    splitVersion = ""
                    vers = ""
                    version = ""
                    logging.info(version)

            else:
                splitRequirements = listOfRequirements[i].split(" ")
                Requirement = splitRequirements[0]
                logging.info(Requirement)
                if len(listOfRequirements[i].split(" ")) > 1:
                    vers = splitRequirements[1]
                    version = int(vers[0])
                    logging.info(version)
                else:
                    vers = ""
                    version = ""
                    logging.info(version)

            time.sleep(1)

            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, A), Requirement)
            time.sleep(1)
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, B), version)

            # if version == 0:
            #     time.sleep(1)
            #     if isDeleteReq['is_delete_req'] != 1:
            #         impactComment = "New interface requirement. No functional impact.\nQIA Param Global- Point No- "
            #     EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, E), impactComment)

            if version == 0:
                time.sleep(1)
                if isDeleteReq['is_delete_req'] != 1:
                    impactComment = "New interface requirement.\nRaised QIA Param Global- Point No-\n"
                    # Raised QIA PT No-.
                    # " Not present in TP."
                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, E), impactComment)
            else:
                time.sleep(1)
                if isDeleteReq['is_delete_req'] != 1:
                    impactComment = "Added interface requirement. No functional impact.\nQIA Param Global- Point No- "
                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, E), impactComment)

            time.sleep(1)
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, F), fepsNum[1:])
            # for remove duplicates and add the feps in one cell in impact tab
            if fepsForDuplicateReqs:
                addDuplicateFEPS(tpBook, maxrow, F, fepsNum, fepsForDuplicateReqs, rqIDs, Requirement, version)

            else:
                logging.info("fepsForDuplicateReqs else condition interface req-->")
                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, F), fepsNum[1:])
            break

        maxrow = maxrow + 1
        i = i + 1
        logging.info(maxrow)

    logging.info("END")
    return rowOfInterface


def modifyImpact(tpBook, listOfRequirements, rowOfInterface, fepsNumber, requirmentlist, TP_Ref, FuncName):
    logging.info("In modifyImpact function - ", listOfRequirements, rowOfInterface)
    maxrow = tpBook.sheets['Impact'].range('A' + str(tpBook.sheets['Impact'].cells.last_cell.row)).end('up').row
    testSheetList = []
    sheet = tpBook.sheets['Impact']
    global impactRow
    macro = EI.getTestPlanAutomationMacro()
    col = 4
    rowList = []
    rqList = []
    isDeleted = 0
    logging.info("maxrow", maxrow)
    for i in range(rowOfInterface, maxrow + 1):
        if sheet.range(i, col).value is not None:
            cellValue = sheet.range(i, col).value.split("\n")
            for t in cellValue:
                if len(t) != 0:
                    testSheetList.append(t)
            rowList.append(i)
    logging.info("List of Testsheets to be modified::::::", testSheetList)
    logging.info("List of rowList to be modified::::::", rowList)
    if len(testSheetList) == 0:
        logging.info("Modification not required12321321")
        rqList = listOfRequirements

        for req in rqList:
            logging.info(f"reqIM {req}")
            if req.find('(') != -1:
                reqName = req.split("(")[0].split()[0] if len(req.split("(")) > 0 else ""
                reqVer = req.split("(")[1].split(")")[0] if len(req.split("(")) > 1 else ""
            else:
                reqName = req.split()[0] if len(req.split()) > 0 else ""
                reqVer = req.split()[1] if len(req.split()) > 1 else ""
            # searchResult = EI.searchDataInCol(tpBook.sheets['Impact'], 1, reqName)

            sheet_value = tpBook.sheets['Impact'].used_range.value
            searchResult = EI.searchDataInColCache(sheet_value, 1, reqName)
            if searchResult['count'] > 0:
                for cellPos in searchResult['cellPositions']:
                    row, col = cellPos
                    if tpBook.sheets['Impact'].range(row, 5).value.upper().find("DELETED") != -1:
                        isDeleted = 1
                        logging.info("isDeleted No TS>> ", isDeleted)
                        modifyDeleteReq(tpBook, "", macro, reqName, reqVer, fepsNumber, requirmentlist, TP_Ref,FuncName, "")
    else:
        logging.info("Modification required")
        versionList = []
        reqList = []
        col = 2
        for i in rowList:
            if sheet.range(i, col).value is not None:
                versionList.append(str(sheet.range(i, col).value)[0])
                reqList.append(sheet.range(i, col - 1).value)

        logging.info("reqList -- ", reqList)
        for testSheet in testSheetList:
            logging.info("------------testSheet---------", testSheet)
            for ver, reqIm, r in zip(versionList, reqList, rowList):
                logging.info("\n\nr - ", r, reqIm, ver)
                yellowModified = 0
                greenModified = 0
                if tpBook.sheets['Impact'].range(r, 5).value.upper().find("DELETED") != -1:
                    isDeleted = 1
                logging.info("isDeleted >> ", isDeleted)
                # logging.info("from IM",testSheet,ver,reqIm,row)
                for i in listOfRequirements:
                    if reqIm in i:
                        listOfRequirements.remove(i)
                # logging.info("tpBook.sheets[testSheet] - ", tpBook.sheets[testSheet])
                getReqList = tpBook.sheets[testSheet].range('C4').value
                reqSplit = getReqList.split("|")
                logging.info("reqSplit - ", reqSplit)
                for i in range(len(reqSplit)):
                    if len(reqSplit[i]) != 0:
                        openBracket = reqSplit[i].find("(")
                        closeBracket = reqSplit[i].find(")")
                        reqName = ""
                        reqVer = ""
                        if openBracket != -1 and closeBracket != -1:
                            reqName = reqSplit[i].split("(")[0]
                            reqVer = reqSplit[i].split("(")[1].split(")")[0]
                        else:
                            logging.info("Do not split")
                        # logging.info("From TS", reqName, reqVer)
                        if reqName == reqIm:
                            if reqVer == ver:
                                if greenModified == 0:
                                    logging.info("no modification in test sheet required")
                                    logging.info("Modifying Impact Tab")
                                    if isDeleted != 1:
                                        tpBook.sheets['Impact'].range(r,
                                                                      5).value = "Interface Requirement.Requirement already present in PT. Associating FEPS with it."
                                        if tpBook.sheets[testSheet].range('C7').value == 'VALIDEE':
                                            EI.activateSheet(tpBook, tpBook.sheets[testSheet])
                                            time.sleep(1)
                                            TPM.selectTestSheetModify(macro)  # for changing version
                                            fillSheetHistory(tpBook.sheets[testSheet],
                                                             "Associating the interface requirement  " + reqIm + "(" + ver + ")" + " with FEPS" + fepsNumber)
                                    else:
                                        logging.info("modifyDeleteReq2")
                                        modifyDeleteReq(tpBook, testSheet, macro, reqIm, ver, fepsNumber,
                                                        requirmentlist, TP_Ref, FuncName, "")
                                    # Associating the requirement GEN-VHL-DCINT-Ssy_IHV_ASS.0610(0) with FEPS_104555
                                    greenModified = 1
                            else:
                                if yellowModified == 0 and reqVer > ver:
                                    logging.info(" Dont take this requirement into account")
                                    if isDeleted != 1:
                                        tpBook.sheets['Impact'].range(r,
                                                                      5).value = " Interface requirement is already present with higher version in Test Plan."
                                else:
                                    if yellowModified == 0 and reqVer < ver:
                                        logging.info("Modying test sheet", reqVer, ver)
                                        # if encour or validee
                                        logging.info("Modifying Imapct Tab")
                                        # EI.activateSheet(tpBook,tpBook.sheets['Impact'])
                                        if isDeleted != 1:
                                            tpBook.sheets['Impact'].range(r,
                                                                          5).value = "Interface Requirement.Incrementing version from " + reqVer + " to " + ver + ". No functional impact. \nRaised QIA Param Global- Point No- "
                                    if isDeleted != 1:
                                        if tpBook.sheets[testSheet].range('C7').value == 'VALIDEE':
                                            EI.activateSheet(tpBook, tpBook.sheets[testSheet])
                                            time.sleep(1)
                                            TPM.selectTestSheetModify(macro)  # for changing version
                                        fillSheetHistory(tpBook.sheets[testSheet],
                                                         "Interface Requirement.Incrementing version of requirement " + reqIm + " from version " + reqVer + " to " + ver + ". No functional impact. \nRaised QIA Param Global- Point No- ")
                                    else:
                                        logging.info("modifyDeleteReq3")
                                        modifyDeleteReq(tpBook, testSheet, macro, reqIm, ver, fepsNumber,
                                                        requirmentlist, TP_Ref, FuncName, "")

                                    reqSplit[i] = reqName + "(" + ver + ")"
                                    reqString = ""
                                    for req in reqSplit:
                                        reqString = reqString + "|" + req
                                    tpBook.sheets[testSheet].range('C4').value = reqString
                                    logging.info("Modified Test Sheet")
                                    yellowModified = 1
        rqList = listOfRequirements
    impactRow = maxrow + 1
    logging.info("Row ", impactRow)
    return rqList


def fillHistoryForDeleteReq(sheet, keyword):
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    sheet_value = sheet.used_range.value
    try:
        # cellValue = EI.searchDataInExcel(sheet, (26, maxrow), "Nature des modifications")
        cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow), "Nature des modifications")
    except:
        # cellValue = EI.searchDataInExcel(sheet, (26, maxrow), "Nature des modifications")
        cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow), "Nature des modifications")
    row, col = cellValue["cellPositions"][0]
    logging.info("In History Del - ", row, col)
    logging.info("sheet.range(row + 1, col).value - ", sheet.range(row + 1, col).value)
    if sheet.range(row + 1, col).value is not None and keyword not in sheet.range(row + 1, col).value:
        getString = sheet.range(row + 1, col).value + keyword
    else:
        getString = keyword

    EI.setDataFromCell(sheet, (row + 1, col), getString)


def checkReqInParamGlobal(tpBook, testSheet, reqIm, ver):
    paramBook = QP.openParamGlobalSheet()
    paramSheet = paramBook.sheets['Paramètre']
    REQ_DCINT_COL = 4
    REQ_FORMER_COL = 5
    REQ_FLUX_SPEC_COL = 3
    FLUX_MESSAGERIE_COL = 9
    PC_NEA_COL = 10
    PC_PROJECT_COL = 1
    PC_ARCHI_COL = 2
    paramResult = {
        'DCINT_ReqValue': [],
        'FORMER_ReqValue': [],
        'paramSignal': '',
        'paramFrame': '',
        'DelReqComment': '',
        'Project': '',
        'Archi': '',
        'PC_NEA': '',
        'CommentValue': ''

    }
    # commentValue 1- remove requirement, 2- remove flow
    # searchResult_DCINT = EI.searchDataInCol(paramSheet, REQ_DCINT_COL, reqIm)
    sheet_value_dci = paramSheet.used_range.value
    searchResult_DCINT = EI.searchDataInColCache(sheet_value_dci, REQ_DCINT_COL, reqIm)

    logging.info("searchResult_DCINT - ", searchResult_DCINT)
    if searchResult_DCINT['count'] != 0:
        if searchResult_DCINT['cellPositions']:
            for cellPosition in searchResult_DCINT['cellPositions']:
                row, col = cellPosition
                logging.info("paramSheet.range(row, REQ_DCINT_COL).value ", paramSheet.range(row, REQ_DCINT_COL).value)
                paramResult['DCINT_ReqValue'] = paramSheet.range(row, REQ_DCINT_COL).value
                # paramResult['FORMER_ReqValue'] = paramSheet.range(row, REQ_FORMER_COL).value
                paramResult['FORMER_ReqValue'] = ""
                paramResult['paramSignal'] = paramSheet.range(row, REQ_FLUX_SPEC_COL).value
                paramResult['paramFrame'] = paramSheet.range(row, FLUX_MESSAGERIE_COL).value
                paramResult['PC_NEA'] = paramSheet.range(row, PC_NEA_COL).value
                paramResult['Project'] = paramSheet.range(row, PC_PROJECT_COL).value
                paramResult['Archi'] = paramSheet.range(row, PC_ARCHI_COL).value

                reqLen = len(paramSheet.range(row, REQ_DCINT_COL).value.split('|'))
                if reqLen > 1:
                    paramResult['DelReqComment'] = "Remove the requirement from Param Global."
                    paramResult['CommentValue'] = 1
                elif reqLen == 1 and (paramSheet.range(row, REQ_FORMER_COL).value == "" or paramSheet.range(row,
                                                                                                            REQ_FORMER_COL).value.find(
                        "--") != -1):
                    paramResult['DelReqComment'] = "Remove the flow from Param Global sheet"
                    paramResult['CommentValue'] = 2
                elif reqLen == 1 and (paramSheet.range(row, REQ_FORMER_COL).value != "" or paramSheet.range(row,
                                                                                                            REQ_FORMER_COL).value.find(
                        "--") == -1):
                    paramResult['DelReqComment'] = "Remove the requirement from Param Global sheet"
                    paramResult['CommentValue'] = 1
    else:
        # searchResult_FORMER = EI.searchDataInCol(paramSheet, REQ_FORMER_COL, reqIm)

        sheet_value_former = paramSheet.used_range.value
        searchResult_FORMER = EI.searchDataInColCache(sheet_value_former, REQ_FORMER_COL, reqIm)

        logging.info("searchResult_FORMER - ", searchResult_FORMER)
        if searchResult_DCINT['count'] != 0:
            if searchResult_FORMER['cellPositions']:
                for cellPosition in searchResult_FORMER['cellPositions']:
                    row, col = cellPosition
                    logging.info("paramSheet.range(row, REQ_FORMER_COL).value ", paramSheet.range(row, REQ_FORMER_COL).value)
                    # paramResult['DCINT_ReqValue'] = paramSheet.range(row, REQ_DCINT_COL).value
                    paramResult['DCINT_ReqValue'] = ""
                    paramResult['FORMER_ReqValue'] = paramSheet.range(row, REQ_FORMER_COL).value
                    paramResult['paramSignal'] = paramSheet.range(row, REQ_FLUX_SPEC_COL).value
                    paramResult['paramFrame'] = paramSheet.range(row, FLUX_MESSAGERIE_COL).value
                    paramResult['PC_NEA'] = paramSheet.range(row, PC_NEA_COL).value
                    paramResult['Project'] = paramSheet.range(row, PC_PROJECT_COL).value
                    paramResult['Archi'] = paramSheet.range(row, PC_ARCHI_COL).value

                    reqLen = len(paramSheet.range(row, REQ_FORMER_COL).value.split('|'))
                    if reqLen > 1:
                        paramResult['DelReqComment'] = "Remove the requirement from Param Global."
                        paramResult['CommentValue'] = 1
                    elif reqLen == 1 and (paramSheet.range(row, REQ_DCINT_COL).value == "" or paramSheet.range(row,
                                                                                                               REQ_DCINT_COL).value.find(
                            "--") != -1):
                        paramResult['DelReqComment'] = "Remove the flow from Param Global sheet"
                        paramResult['CommentValue'] = 2
                    elif reqLen == 1 and (paramSheet.range(row, REQ_DCINT_COL).value != "" or paramSheet.range(row,
                                                                                                               REQ_DCINT_COL).value.find(
                            "--") == -1):
                        paramResult['DelReqComment'] = "Remove the requirement from Param Global sheet"
                        paramResult['CommentValue'] = 1
    return paramResult


def findDciFileAndGetDciData(fepsNumber, requirmentlist, reqIm, ver):
    for feps in requirmentlist:
        logging.info("feps >>>>.>>>> ", feps)
        logging.info("feps NUM >>>>.>>>> ", "FEPS" + fepsNumber)
        if feps == "FEPS" + fepsNumber or feps.find("FEPS" + fepsNumber) != -1:
            DCIdoc = []
            for inputdoc in requirmentlist[feps]['Input_Docs']:
                dci_ref_num = ""
                dci_ver = ""
                logging.info("inputdoc >>> ", inputdoc)
                if inputdoc.find('DCI') != -1:
                    logging.info("DCI Input Document- ", inputdoc)
                    DCIdoc.append(inputdoc)
                    logging.info(DCIdoc)
                    dci = EI.openDCIExcel(DCIdoc)
                    if dci is not None:
                        if re.search(pattren_ref, inputdoc):
                            dci_ref = re.findall(pattren_ref, inputdoc)
                            dci_ref_num = "".join(dci_ref[0])
                            logging.info("refnm1 ", "".join(dci_ref[0]))
                        if re.search(pattren_ver, inputdoc):
                            dci_version = re.findall(pattren_ver, inputdoc)
                            dci_ver = "".join(dci_version[0])
                            logging.info("ver ", "".join(dci_version[0]))
                        dciInfo = EI.getDciInfo(dci, reqIm)
                        dciInfo['dci_ref_num'] = dci_ref_num
                        dciInfo['dci_ver'] = dci_ver

                        logging.info("dciInfo Val ->>>> ", dciInfo)

    return dciInfo


def delete_or_replace_flow(tpBook, testSheet, dciData, flowSearchResult):
    flowWithNetwork = []
    flowWithoutNetwork = []
    if flowSearchResult['count'] > 0:
        if len(dciData['framename'].split('/')[0].split('_')) == 2:
            dciFrame = dciData['framename'].split('/')[0].split('_')[1]
        elif len(dciData['framename'].split('/')[0].split('_')) == 3:
            dciFrame = dciData['framename'].split('/')[0].split('_')[1] + "_" + \
                       dciData['framename'].split('/')[0].split('_')[2]
        logging.info("dciFrame>> ", dciFrame)
        for cellPosition in flowSearchResult['cellPositions']:
            logging.info("TScellPosition >> ", cellPosition)
            row, col = cellPosition
            cellValue = tpBook.sheets[testSheet].range(row, col).value
            logging.info("TScellValue >> ", cellValue)
            if cellValue is not None:
                if dciFrame in cellValue:
                    tpBook.sheets[testSheet].activate()
                    logging.info("===============Signal append with n/w============")
                    flowWithNetwork.append((cellValue, (cellPosition)))
                else:
                    logging.info("===============Signal not append with n/w============")
                    flowWithoutNetwork.append((cellValue, (cellPosition)))

        logging.info("flowWithNetwork>> ", flowWithNetwork)
        logging.info("flowWithoutNetwork>> ", flowWithoutNetwork)
        logging.info("len(flowWithNetwork) ", len(flowWithNetwork))
        logging.info("len(flowWithoutNetwork) ", len(flowWithoutNetwork))
        if len(flowWithNetwork) != 0 and len(flowWithoutNetwork) == 0:
            for fwn in flowWithNetwork:
                logging.info("cellPos Mod>> ", fwn[1])
                x, y = fwn[1]  # getting the cellposition
                logging.info("xy - ", x, y)
                modifiedVal = cellValue.replace("_" + dciFrame, "")
                logging.info("modifiedVal >>> ", modifiedVal)
                tpBook.sheets[testSheet].range(x, y).value = modifiedVal
        if len(flowWithNetwork) != 0 and len(flowWithoutNetwork) != 0:
            for fwn in flowWithNetwork:
                if dciFrame in fwn[0]:
                    # flowPosition = EI.searchDataInExcel(tpBook.sheets[testSheet], '', fwn[0])
                    sheet_value = tpBook.sheets[testSheet].used_range.value
                    flowPosition = EI.searchDataInExcelCache(sheet_value, '', fwn[0])

                    logging.info("flowPosition -- ", flowPosition)
                    if flowPosition['count'] > 0:
                        for cellPos in flowPosition['cellPositions']:
                            x1, y1 = cellPos
                            logging.info("Removing the step")
                            logging.info("x1, y1 -- ", x1, y1)
                            tpBook.sheets[testSheet].range(str(x1) + ":" + str(x1)).delete()
                            break
    else:
        logging.info(">>>>>>>>>>>>>>>> Signal not present in Testsheet " + testSheet + " >>>>>>>>>>>>>>>>")
        return 1

    return 1


def modifyDeleteReq(tpBook, testSheet, macro, reqIm, ver, fepsNumber, requirmentlist, TP_Ref, FuncName, nover):
    logging.info("-------------modifyDeleteReq " + str(reqIm) +" "+ str(ver) +" " + str(testSheet) + "------------")

    if reqIm != "" and reqIm is not None:
        if testSheet != "" and testSheet is not None:
            logging.info("tpBook.sheets[testSheet].range('C7').value - ", tpBook.sheets[testSheet].range('C7').value)
            reqList = tpBook.sheets[testSheet].range('C4').value
            if reqList is not None and reqList != "" and reqList.find("--") == -1:
                if reqIm in reqList:
                    logging.info("----------------++++++++++-------------")
                    try:
                        logging.info(f"ffffffffffff {ver}")
                        if ver is None or ver == '':
                            ts_requirement = re.findall(reqIm + r'\([^)]*\)', reqList)
                            logging.info(f"ts_requirement {ts_requirement}")
                            if ts_requirement:
                                ts_req, ver = find_req_ver(ts_requirement[0])
                                logging.info(f"ts_req, ver {ts_req, ver}")
                        if tpBook.sheets[testSheet].range('C7').value == 'VALIDEE':
                            logging.info("mmmmmmmmmmmmmmmmmmmmmmmmm")
                            EI.activateSheet(tpBook, tpBook.sheets[testSheet])
                            time.sleep(1)
                            TPM.selectTestSheetModify(macro)  # for changing version
                            fillHistoryForDeleteReq(tpBook.sheets[testSheet],
                                                    "Removed the Interface requirement " + reqIm + "(" + ver + ")" + ".No functional imapct.")
                        elif tpBook.sheets[testSheet].range('C7').value == 'EN COURS':
                            logging.info("nnnnnnnnnnnnnnnnnnnnnnnnn")
                            EI.activateSheet(tpBook, tpBook.sheets[testSheet])
                            time.sleep(1)
                            if nover != "" and nover is not None:
                                TPM.selectTestSheetModify(macro)  # for changing version
                            fillHistoryForDeleteReq(tpBook.sheets[testSheet],
                                                    "Removed the Interface requirement " + reqIm + "(" + ver + ")" + ".No functional imapct.")
                        modifiesReqs = re.sub(reqIm + r'\([^)]*\)', "", reqList)
                        logging.info("modifiesReqs-->> ", modifiesReqs)
                        modifiesReqsList = modifiesReqs.replace("||", "|")
                        logging.info("modifiesReqsList - ", modifiesReqsList)
                        tpBook.sheets[testSheet].range('C4').value = modifiesReqsList
                    except:
                        displayInformation("______________Something wrong in deleting the requirement______________")

        paramGlobalResponse = checkReqInParamGlobal(tpBook, testSheet, reqIm, ver)
        logging.info("\n\nparamGlobalResponse >>> ", paramGlobalResponse)
        if len(paramGlobalResponse['FORMER_ReqValue']) != 0 or len(paramGlobalResponse['DCINT_ReqValue']) != 0:
            logging.info("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
            qicomment = ""
            dciData = findDciFileAndGetDciData(fepsNumber, requirmentlist, reqIm, ver)
            if paramGlobalResponse['CommentValue'] == 1:
                qicomment = f"Please delete the requirement {reqIm}({ver}) as it is removed from the DCI document ({dciData['dci_ref_num']}) ({dciData['dci_ver']})."
            else:
                qicomment = f"Please delete the flow as it is removed from the DCI document ({dciData['dci_ref_num']}) ({dciData['dci_ver']})."

            logging.info(f"qicomment {qicomment}")
            if testSheet != "" and testSheet is not None:
                # flowSearchResult = EI.searchDataInExcel(tpBook.sheets[testSheet], '',
                #                                         "$" + paramGlobalResponse['paramSignal'])
                sheet_value =tpBook.sheets[testSheet].used_range.value
                flowSearchResult = EI.searchDataInExcelCache(sheet_value, '',
                                                        "$" + paramGlobalResponse['paramSignal'])

                logging.info("flowSearchResult - ", flowSearchResult)
                if dciData is not None and dciData != -1:
                    FlowDeleteReplaceResponse = delete_or_replace_flow(tpBook, testSheet, dciData, flowSearchResult)
                    logging.info("FlowDeleteReplaceResponse >> ", FlowDeleteReplaceResponse)

            # raising QIA param
            if dciData['dciSignal'] != "" and dciData['dciSignal'] is not None:
                logging.info("dciData['proj_param'] - ", dciData['proj_param'])
                if dciData['proj_param'] is not None and dciData['proj_param'] != "" and dciData[
                    'proj_param'].upper().find("NEAR") != -1:
                    Nom_du_SO = QP.getDCIProjParam(dciData['proj_param'])
                else:
                    Nom_du_SO = dciData['proj_param']
                QIA_Data = {"TP_Refnum": TP_Ref, "taskName": FuncName,
                            "Req_type": 'Modification', "Expl": paramGlobalResponse['DelReqComment'],
                            "Nom_du_SO": Nom_du_SO, "columnG": "--",
                            "signal": dciData['dciSignal'], "newreq": dciData['dciReq'],
                            "dciframe": dciData['framename'], "flowtype": dciData['pc'],
                            "trigram": ICF.gettrigram(), "qiacomment": qicomment}
            else:
                logging.info("============{{{{{{{{}}}}}}}============")
                if paramGlobalResponse['DCINT_ReqValue'] != "" and paramGlobalResponse['FORMER_ReqValue'] == "":
                    param_req = paramGlobalResponse['DCINT_ReqValue']
                elif paramGlobalResponse['DCINT_ReqValue'] == "" and paramGlobalResponse['FORMER_ReqValue'] != "":
                    param_req = paramGlobalResponse['DCINT_ReqValue']
                QIA_Data = {"TP_Refnum": TP_Ref, "taskName": FuncName,
                            "Req_type": 'Modification', "Expl": paramGlobalResponse['DelReqComment'],
                            "Nom_du_SO": paramGlobalResponse['Project'], "columnG": paramGlobalResponse['Archi'],
                            "signal": paramGlobalResponse['paramSignal'], "newreq": param_req,
                            "dciframe": paramGlobalResponse['paramFrame'], "flowtype": paramGlobalResponse['PC_NEA'],
                            "trigram": ICF.gettrigram(), "qiacomment": qicomment}
            # "dci_ref": str(dciData['dci_ref_num']),
            # "dci_ver": str(dciData['dci_ver'])
            QP.addQIADataInQIASheet(None, QIA_Data)
        else:
            displayInformation(
                "-------[Requirement " + reqIm + "(" + ver + ")" + " not present in param global sheet]------")

        return 1


def addRequirement(tps, nrqs):
    rqList = ""
    orq = EI.getDataFromCell(tps, 'C4')
    logging.info("Old Requirements = " + str(orq))
    logging.info("New Requirements = " + str(nrqs))
    isOldRq = False
    if orq is not None:
        isOldRq = True
        rqList = orq
    # for nrq in nrqs:
    if isOldRq:
        # nrq = "|" + nrqs
        rqList = orq + "|" + nrqs
    else:
        # nrq = nrqs + "|"
        rqList = nrqs
    EI.setDataFromCell(tps, 'C4', rqList)


def handle_req_with_same_ver(tpBook, ts, macro, reqName, reqVer, fepsNum):
    if tpBook.sheets[ts].range('C7').value == 'VALIDEE':
        EI.activateSheet(tpBook, tpBook.sheets[ts])
        time.sleep(1)
        logging.info("hi1--->", ts)
        TPM.selectTestSheetModify(macro)  # for changing version
        logging.info("hi12--->", ts)
        fillSheetHistory(tpBook.sheets[ts],
                         "Associating the requirement  " + reqName + "(" + reqVer + ")" + " with FEPS" + fepsNum + ".")


def treatBackLog(tpBook, tsList):
    display_info = UpdateHMIInfoCb
    logging.info("\n----Treating backlog for req with same version-----\n")
    # opened_bks = EI.getOpenedBooks()
    # logging.info("opened_bks -- ", opened_bks)
    for testSheet in tsList:
        try:
            logging.info(f"ICF.getBackLog()  {ICF.getBackLog() } {testSheet}")
            if ICF.getBackLog() is True:
                display_info(f"\n========={testSheet}===========")
                logging.info("Backlog Selected")
                logging.info("Backlog Selected")
                # referential = EI.findInputFiles()[4]
                refEC = EI.openReferentialEC()
                kpiDocList = getKPIDocPath(ICF.getInputFolder() + "\\KPI")
                logging.info("kpiDocList --> ", kpiDocList)
                logging.info("kpiDocList --> ", kpiDocList)
                # rawReqs = analyseThematics.getTsRawReq()
                rawReqs = tpBook.sheets[testSheet].range('C4').value
                logging.info("rawReqs --> ", rawReqs)
                logging.info("rawReqs --> ", rawReqs)
                ts_reqs = [rawReqs]
                currArch = getArch(ICF.FetchTaskName())
                logging.info("*** kpiDocList ***", kpiDocList)
                logging.info("*** reqs ***", ts_reqs)
                logging.info("*** refEC ***", str(refEC))
                logging.info("*** currArch ***", currArch)
                ts_req_list = removeInterfaceReq(ts_reqs)

                start_time_bl = time.time()
                logging.info(f'\n BL start time: {start_time_bl}')
                status, combinedThemLines = getCombinedThematicLines(kpiDocList, ts_req_list,
                                                                     refEC, currArch)

                end_time_bl = time.time()
                execution_time_bl = end_time_bl - start_time_bl
                logging.info(f'\n BL end execution time: {execution_time_bl}')

                refEC.close()
                logging.info("***Thematic Combinations = ***", status, combinedThemLines)
                if status:
                    logging.info("Update the Thematic Combinations in the View")
                    # display_info(f"========={tpBook.sheets[testSheet].name} - Backlog Output===========")
                    display_info(f"Backlog Output:\n")
                    display_info(str(combinedThemLines))
                    display_info("====================")
                else:
                    logging.info(f"Requirement Not available in KPI Sheet. Proceed Manually - {ts_reqs}")
                    #display_info(
                     #   "Requirement Not available in KPI Sheet. Make sure all the requirement in the test sheet #is mentioned in KPI Sheet")
        except Exception as e:
            logging.info(f"!!!!!Something went wrong in treating backlog for requirement with same version!!!!! {e}")
            exc_type, exc_obj, exc_tb = sys.exc_info()
            logging.info(f"!!!!!Something went wrong in treating backlog for requirement with same version!!!!! {e} {exc_tb.tb_lineno}")


def fillImpactEvolved(tpBook, oldReq, fepsNum, flag, Arch, rqIDs, fepsForDuplicateReqs, newReq=""):
    verInfo = []
    reqSplit = ''
    logging.info("In fillImpactEvolved function - ", oldReq, fepsNum, newReq)
    maxrow = tpBook.sheets['Impact'].range('A' + str(tpBook.sheets['Impact'].cells.last_cell.row)).end('up').row
    sheet = tpBook.sheets['Impact']
    macro = EI.getTestPlanAutomationMacro()
    i = 0
    length = 1
    time.sleep(2)
    EI.activateSheet(tpBook, 'Impact')
    if oldReq.find('(') != -1:
        oldReq = "".join(oldReq.split())
    logging.info("++++++++Value of ROW = " + str(maxrow))
    logging.info("++++++++Value of length = " + str(length))
    logging.info("++++++++Value of Req = ", oldReq, newReq)
    condition = maxrow + 1

    colOfRequirement = 1
    colOfVer = 2
    colOfFT = 4
    colOfComment = 5
    colOfFeps = 6
    logging.info("True Or False = ", len(newReq) != 0)
    if len(newReq) != 0:
        logging.info("old to new conversion", oldReq, newReq)
        if oldReq.find('(') != -1:
            oldReqName = oldReq.split("(")[0].split()[0] if len(oldReq.split("(")) > 0 else ""
            oldReqVer = oldReq.split("(")[1].split(")")[0] if len(oldReq.split("(")) > 1 else ""
        else:
            oldReqName = oldReq.split()[0] if len(oldReq.split()) > 0 else ""
            oldReqVer = oldReq.split()[1] if len(oldReq.split()) > 1 else ""

        logging.info("oldReqVer = ", oldReqVer)
        logging.info("oldReqName = ", oldReqName)
        # time.sleep(1)

        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfRequirement), oldReqName)
        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), oldReqVer)
        # for remove duplicates and add the feps in one cell in impact tab
        if fepsForDuplicateReqs:
            addDuplicateFEPS(tpBook, maxrow, colOfFeps, fepsNum, fepsForDuplicateReqs, rqIDs, oldReqName, oldReqVer)
        else:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfFeps), fepsNum[1:])
        logging.info("filled old req")

        if flag == -1:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, 5),
                               "The thematic lines of the requirement are NA for" + Arch + ".\n Proceed Manually.")
        elif flag == -2:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, 5),
                               "Input doc is not in the correct format.\n Proceed Manually.")

        # Run impact
        time.sleep(2)
        TPM.selectTPImpact(macro)
        time.sleep(10)
        # Get data in D column and check any one of the test sheet for requirement
        # If old req not present repalce it with new req
        if flag == 1:
            if sheet.range(maxrow + 1, colOfFT).value is None:

                if newReq.find('(') != -1:
                    newReqName = newReq.split("(")[0].split()[0] if len(newReq.split("(")) > 0 else ""
                    newReqVer = newReq.split("(")[1].split(")")[0] if len(newReq.split("(")) > 1 else ""
                else:
                    newReqName = newReq.split()[0] if len(newReq.split()) > 0 else ""
                    newReqVer = newReq.split()[1] if len(newReq.split()) > 1 else ""
                logging.info("newReqVer = ", newReqVer)
                logging.info("newReqName = ", newReqName)
                time.sleep(5)

                # to check if req already present in previous feps
                checkreq = EI.searchDataInSpecificRows(tpBook.sheets['Impact'], (18, 100), colOfRequirement, newReqName)
                logging.info(f"\n---> checkreq {checkreq} --->")
                if checkreq['count'] > 0:
                    x, y = checkreq["cellPositions"][0]
                    v = EI.getDataFromCell(tpBook.sheets['Impact'], (x, colOfVer))
                    if v == newReqVer:
                        logging.info("*********2222222222222222222")
                        logging.info(f"x, y 124 {x, y}")
                        f = EI.getDataFromCell(tpBook.sheets['Impact'], (x, colOfFeps))
                        f = str(int(f)) + "\n" + str(fepsNum[1:])
                        EI.setDataFromCell(tpBook.sheets['Impact'], (x, colOfFeps), f)
                        comnt = EI.getDataFromCell(tpBook.sheets['Impact'], (x, 5))
                        comnt = str(comnt) + "\nAssociating it with FEPS " + str(fepsNum[1:])
                        EI.setDataFromCell(tpBook.sheets['Impact'], (x, 5), comnt)
                        return -2, None
                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfRequirement), newReqName)
                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), newReqVer)

                # for remove duplicates and add the feps in one cell in impact tab
                if fepsForDuplicateReqs:
                    addDuplicateFEPS(tpBook, maxrow, colOfFeps, fepsNum, fepsForDuplicateReqs, rqIDs, newReqName, newReqVer)
                else:
                    EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfFeps), fepsNum[1:])
                time.sleep(2)
                TPM.selectTPImpact(macro)
                time.sleep(10)

                # testSheetList = EI.getDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfFT))
                if sheet.range(maxrow + 1, colOfFT).value is not None:
                    testSheetList = sheet.range(maxrow + 1, colOfFT).value.split("\n")
                    logging.info("Test sheet list(1) = ", testSheetList)
                    with open('../Aptest_Tool_Report.txt', 'a') as f:
                        f.writelines("\n\nTest sheet list(1) = " + str(testSheetList))
                    logging.info("testSheetListBF2 --> ", testSheetList)
                    temp = []
                    temp1 = [temp.append(ts) if ts != "" and ts.find("_SF_") == -1 else "" for ts in testSheetList]
                    testSheetList_new = copy.deepcopy(temp)

                    logging.info(testSheetList_new, " ===>TestSheetListAF2")
                    reqVer = ""
                    for index, t in enumerate(testSheetList_new):
                        if t != '':
                            if t.find("DJNC") == -1:
                                getReqList = tpBook.sheets[t].range('C4').value
                                reqSplit = getReqList.split("|")
                                for i in range(len(reqSplit)):
                                    if len(reqSplit[i]) != 0:
                                        reqName = reqSplit[i].split("(")[0]
                                        reqVer = reqSplit[i].split("(")[1].split(")")[0]
                                        logging.info("From TS====>>>>>>",reqName,reqVer)
                                        if reqName == newReqName:
                                            logging.info("reqName, newReqName TS====>>>>>>", reqName, newReqName)
                                            logging.info("reqVer, newReqVer TS====>>>>>>", reqVer, newReqVer)
                                            if reqVer == newReqVer:
                                                time.sleep(1)
                                                # EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),"Evolved Requirement.\nRequirement already present in testplan. Associating FEPS with" + fepsNum)

                                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),"Evolved Requirement.\nRequirement already present with same version (" + str(
                                                                       reqVer) + ") in testplan. Associating FEPS with" + fepsNum + ".")
                                                # return -2
                                                handle_req_with_same_ver(tpBook, t, macro, reqName, reqVer, fepsNum)
                                                logging.info("sheet index1-->", len(testSheetList_new) - 1, index)
                                                if len(testSheetList_new) - 1 == index:
                                                    logging.info("sdafsdf2--->", len(testSheetList_new) - 1, index)
                                                    treatBackLog(tpBook, testSheetList_new)
                                                    return -2, None
                                                else:
                                                    logging.info("continuous to else part")
                                                    continue

                                            # To check test req version is greater
                                            elif (ord(reqVer) > ord(newReqVer)):
                                                time.sleep(1)
                                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),"Requirement already present in testplan with Higher Version " + str(
                                                                       reqVer) + ".")
                                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), reqVer)
                                                return -2, None

                                        else:
                                            logging.info("reqName, newReqName else====>>>>>>", reqName, newReqName)
                                            logging.info("reqVer, newReqVer else====>>>>>>", reqVer, newReqVer)
                                            if reqName.find('.') != -1:
                                                reqName = reqName.replace('.', '-')
                                            if reqName.find('_') != -1:
                                                reqName = reqName.replace('_', '-')
                                            if newReqName.find('.') != -1:
                                                newReqName = newReqName.replace('.', '-')
                                            if newReqName.find('_') != -1:
                                                newReqName = newReqName.replace('_', '-')
                                            if reqName == newReqName:
                                                if reqVer == newReqVer:
                                                    time.sleep(1)
                                                    # EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),"Evolved Requirement.\nRequirement already present in testplan. Associated with FEPS" + fepsNum + ".")
                                                    EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                                                       "Evolved Requirement.\nRequirement already present with same version (" + str(
                                                                           reqVer) + ") in testplan. Associating FEPS with" + fepsNum + ".")
                                                    # return -2
                                                    handle_req_with_same_ver(tpBook, t, macro, reqName, reqVer, fepsNum)
                                                    logging.info("sheet index2-->", len(testSheetList_new) - 1, index)
                                                    if len(testSheetList_new) - 1 == index:
                                                        logging.info("sdafsdf3--->", len(testSheetList_new) - 1, index)
                                                        treatBackLog(tpBook, testSheetList_new)
                                                        return -2, None
                                                    else:
                                                        logging.info("continuous to else part")
                                                        continue
                                                # To check test req version is greater
                                                elif (ord(reqVer) > ord(newReqVer)):
                                                    logging.info("reqVer, newReqVer elif1====>>>>>>", reqVer, newReqVer)
                                                    time.sleep(1)
                                                    EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                                                       "Requirement already present in testplan with version " + str(
                                                                           reqVer) + ". Associating FEPS with it.")
                                                    EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer),
                                                                       reqVer)
                                                    return -2, None
                    # Get data in D column and check any one of the test sheet for requirement
                    verInfo = getLastVerInfo(tpBook, newReqName, reqVer, testSheetList)
                    logging.info("verInfo(1) = ", verInfo)
                else:
                    return -1, None
            else:
                testSheetList = sheet.range(maxrow + 1, colOfFT).value.split("\n")
                logging.info("Test sheet list(2) = ", testSheetList)
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nTest sheet list(2) = " + str(testSheetList))
                verInfo = getLastVerInfo(tpBook, oldReqName, oldReqVer, testSheetList)
                logging.info("verInfo(2) = ", verInfo)
    else:
        logging.info("only 1 req name1")
        if oldReq.find('(') != -1:
            oldReqName = oldReq.split("(")[0].split()[0] if len(oldReq.split("(")) > 0 else ""
            oldReqVer = oldReq.split("(")[1].split(")")[0] if len(oldReq.split("(")) > 1 else ""
        else:
            oldReqName = oldReq.split()[0] if len(oldReq.split()) > 0 else ""
            oldReqVer = oldReq.split()[1] if len(oldReq.split()) > 1 else ""
        logging.info("ReqVer = ", oldReqVer)
        logging.info("ReqName = ", oldReqName)
        # to check if req already present in previous feps
        checkreq = EI.searchDataInSpecificRows(tpBook.sheets['Impact'], (18, 100), colOfRequirement, oldReqName)
        logging.info(f"\n---> checkreq1 {checkreq} --->")
        if checkreq['count'] > 0:
            x, y = checkreq["cellPositions"][0]
            v = EI.getDataFromCell(tpBook.sheets['Impact'], (x, colOfVer))
            if v == oldReqVer:
                logging.info("*********111111111")
                logging.info(f"x, y 123 {x, y}")
                f = EI.getDataFromCell(tpBook.sheets['Impact'], (x, colOfFeps))
                f = str(int(f)) + "\n" + str(fepsNum[1:])
                EI.setDataFromCell(tpBook.sheets['Impact'], (x, colOfFeps), f)
                comnt = EI.getDataFromCell(tpBook.sheets['Impact'], (x, 5))
                comnt = str(comnt) + "\nAssociating it with FEPS " + str(fepsNum[1:])
                EI.setDataFromCell(tpBook.sheets['Impact'], (x, 5), comnt)
                return -2, None
        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfRequirement), oldReqName)
        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), oldReqVer)

        # for remove duplicates and add the feps in one cell in impact tab
        if fepsForDuplicateReqs:
            addDuplicateFEPS(tpBook, maxrow, colOfFeps, fepsNum, fepsForDuplicateReqs, rqIDs, oldReqName, oldReqVer)
        else:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfFeps), fepsNum[1:])

        if flag == -1:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, 5),
                               "The thematic lines of the requirement are NA for" + Arch + ".\n Proceed Manually")
        elif flag == -2:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, 5),
                               "Input doc is not in the correct format.\n Proceed Manually.")

        time.sleep(2)
        TPM.selectTPImpact(macro)
        time.sleep(10)
        if flag == 1:
            if sheet.range(maxrow + 1, colOfFT).value is not None:
                testSheetList = sheet.range(maxrow + 1, colOfFT).value.split("\n")
                logging.info("Test sheet list(3) = ", testSheetList)
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nTest sheet list(3) = " + str(testSheetList))
                logging.info("testSheetListBF --> ", testSheetList)
                temp = []
                temp1 = [temp.append(ts) if ts != "" and ts.find("_SF_") == -1 else "" for ts in testSheetList]
                testSheetList_new = copy.deepcopy(temp)

                logging.info(testSheetList_new, " ===>TestSheetListAF1")
                logging.info(testSheetList_new, " ===>TestSheetListAF1")

                for index, t in enumerate(testSheetList_new):
                    if t != '':
                        if t.find("DJNC") == -1:
                            getReqList = tpBook.sheets[t].range('C4').value
                            reqSplit = getReqList.split("|")
                            logging.info("reqSplit1212 ", reqSplit)
                            for i in range(len(reqSplit)):
                                if len(reqSplit[i]) != 0:
                                    # logging.info("reqSplit[i] = ", reqSplit[i])
                                    if reqSplit[i].find("(") != -1:
                                        reqName = reqSplit[i].split("(")[0]
                                        reqVer = reqSplit[i].split("(")[1].split(")")[0]
                                    else:
                                        reqName = reqSplit[i]
                                        reqVer = ""
                                    logging.info("From TS", reqName, reqVer)
                                    logging.info("From INP", oldReqName, oldReqVer)
                                    if reqName.strip() == oldReqName.strip():
                                        if reqVer == oldReqVer:
                                            logging.info("||||||||________|||||||")
                                            time.sleep(1)
                                            # EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                            #                    "Evolved Requirement.\nRequirement already present in testplan. Associated with FEPS" + fepsNum + ".")
                                            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                                               "Evolved Requirement.\nRequirement already present with same version (" + str(
                                                                   reqVer) + ") in testplan. Associating FEPS with" + fepsNum + ".")
                                            # return -2

                                            handle_req_with_same_ver(tpBook, t, macro, reqName, reqVer, fepsNum)
                                            logging.info("sheet index3-->", len(testSheetList_new) - 1, index)
                                            if len(testSheetList_new) - 1 == index:
                                                logging.info("sdafsdf1--->", len(testSheetList_new) - 1, index)
                                                treatBackLog(tpBook, testSheetList_new)
                                                return -2, None
                                            else:
                                                logging.info("continuous to else part")
                                                continue
                                        # To check test req version is greater
                                        # elif ord(reqVer) > ord(oldReqVer):
                                        #     time.sleep(1)
                                        #     EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                        #                        "Requirement already present in testplan with version " + str(
                                        #                            reqVer) + ". Associating FEPS with it.")
                                        #     EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), reqVer)
                                        #     return -2
                                        # if reqVer > oldReqVer:
                                        #     logging.info(" Dont take this requirement into account")
                                        #     EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                        #                        "Evolved requirement is already present with higher version in Test Plan.")
                                        elif re.search(r'^[A-Za-z]$', str(reqVer).strip()) and re.search(r'^[A-Za-z]$', str(oldReqVer).strip()):
                                            if ord(str(reqVer).strip()) > ord(str(oldReqVer).strip()):
                                                time.sleep(1)
                                                # EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),"Requirement already present in testplan with version " + str(reqVer) + ". Associating FEPS with it.")
                                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                                                   "Evolved requirement is already present with higher version in Test Plan.")

                                                # EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), reqVer)
                                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), oldReqVer)
                                                return -2, None
                                        else:
                                            if str(reqVer).isnumeric() and str(oldReqVer).isnumeric():
                                                if reqVer > oldReqVer:
                                                    # EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),"Requirement already present in testplan with version " + str(reqVer) + ". Associating FEPS with it.")
                                                    EI.setDataFromCell(tpBook.sheets['Impact'],(maxrow + 1, colOfComment),"Evolved requirement is already present with higher version in Test Plan.")
                                                    # EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer),reqVer)
                                                    EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer),oldReqVer)
                                                    return -2, None

                                        Evo_Req_pattern = '([A-Za-z]{1})|([0-9]{1})'
                                        if Evo_Req_pattern == reqVer:
                                            # if re.search(Evo_Req_pattern, reqVer):
                                            logging.info("hhhhhhhhh")
                                            if re.search("[A-Za-z]", str(reqVer)) and re.search("[A-Za-z]", str(oldReqVer)):
                                                req_cond = "ord(reqVer) > ord(oldReqVer)"
                                            else:
                                                req_cond = "reqVer > oldReqVer"
                                            logging.info(f"ord(reqVer) {ord(reqVer)}")
                                            logging.info(f"ord(oldReqVer) {ord(oldReqVer)}")
                                            exit()
                                            if req_cond:
                                                time.sleep(1)
                                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                                                   "Requirement already present in testplan with Higher version " + str(
                                                                       reqVer) + ".")
                                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), reqVer)
                                                return -2, None

                                            else:
                                                num1 = reqVer
                                                num2 = oldReqVer
                                                if type(num1) is int:
                                                    converted_num1 = int(float(num1))
                                                    converted_num2 = int(float(num2))
                                                else:
                                                    converted_num1 = num1.strip()
                                                    converted_num2 = num2.strip()
                                                logging.info(converted_num1, converted_num2)
                                                if converted_num1 > converted_num2:
                                                    time.sleep(1)
                                                    EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfComment),
                                                                       "Requirement already present in testplan with Higher version " + converted_num1 + ".")
                                                    EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer),
                                                                       converted_num1)
                                                    return -2, None

                # Get data in D column and check any one of the test sheet for requirement
                verInfo = getLastVerInfo(tpBook, oldReqName, oldReqVer, testSheetList)
                logging.info("verInfo(3) = ", verInfo)
            else:
                return -1, None
    impactRow = maxrow + 1
    return verInfo, reqSplit


def getLastVerInfo(tpBook, reqName, reqVer, testsheets):
    logging.info("Requirement - ", "*" + reqName + "*", "*" + reqVer + "*")
    vPT = -1
    verInfo = []
    count = 0
    for t in testsheets:
        if t == '':
            testsheets.remove(t)
    logging.info("testsheets - ", testsheets)
    for sheet in testsheets:
        logging.info(sheet, " --- ", testsheets)
        if sheet:
            if sheet.find("DJNC") == -1:
                logging.info("------------------------" + sheet + "-------------------------")
                maxrow = tpBook.sheets[sheet].range('A' + str(tpBook.sheets[sheet].cells.last_cell.row)).end('up').row
                logging.info("maxrow & Sheet - ", maxrow, sheet)
                try:
                    # cellValue = EI.searchDataInExcel(tpBook.sheets[sheet], (26, maxrow), "Nature des modifications")
                    sheet_value = tpBook.sheets[sheet].used_range.value
                    cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow), "Nature des modifications")

                except:
                    # cellValue = EI.searchDataInExcel(tpBook.sheets[sheet], (26, maxrow), "Nature des modifications")
                    sheet_value = tpBook.sheets[sheet].used_range.value
                    cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow), "Nature des modifications")

                x, y = cellValue["cellPositions"][0]
                logging.info("In History", x, y)
                # ft = EI.activateSheet(tpBook, tpBook.sheets[sheet])
                logging.info(reqName, x + 100, y)

                # need to check range only for history section of testfile.
                try:
                    # reqInCell = EI.searchDataInExcel(tpBook.sheets[sheet], (maxrow, y), reqName)

                    sheet_value = tpBook.sheets[sheet].used_range.value
                    reqInCell = EI.searchDataInExcelCache(sheet_value, (maxrow, y), reqName)

                except:
                    # reqInCell = EI.searchDataInExcel(tpBook.sheets[sheet], (maxrow, y), reqName)
                    sheet_value = tpBook.sheets[sheet].used_range.value
                    reqInCell = EI.searchDataInExcelCache(sheet_value, (maxrow, y), reqName)

                logging.info("Cell of Requirement = ", reqInCell, len(reqInCell["cellPositions"]))
                try:
                    if len(reqInCell["cellPositions"]) > 1:
                        reqInCell["cellPositions"].sort()
                        row, col = reqInCell["cellPositions"][1]
                        logging.info(row, col - 1)
                        vPT = int(EI.getDataFromCell(tpBook.sheets[sheet], (row, col - 1)))
                        logging.info("Version of PT", vPT)
                        with open('../Aptest_Tool_Report.txt', 'a') as f:
                            f.writelines("\n\nVersion of PT" + str(vPT))
                        # verInfo.append((vPT, tpBook.sheets[sheet], reqName, reqVer))
                    else:
                        maxrow = tpBook.sheets[sheet].range('A' + str(tpBook.sheets[sheet].cells.last_cell.row)).end(
                            'up').row
                        sheet_value = tpBook.sheets[sheet].used_range.value
                        try:
                            # cellValue = EI.searchDataInExcel(tpBook.sheets[sheet], (26, maxrow),
                            #                                  "Nature des modifications")
                            cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow),
                                                             "Nature des modifications")

                        except:
                            # cellValue = EI.searchDataInExcel(tpBook.sheets[sheet], (26, maxrow),
                            #                                  "Nature des modifications")
                            cellValue = EI.searchDataInExcelCache(sheet_value, (26, maxrow),
                                                                "Nature des modifications")

                        x, y = cellValue["cellPositions"][0]
                        logging.info("In History", x, y)
                        vPT = int(EI.getDataFromCell(tpBook.sheets[sheet], (maxrow, y - 1)))
                except Exception as ex:
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    logging.info(f"\nError: {ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
                verInfo.append((vPT, tpBook.sheets[sheet], reqName, reqVer))
            else:
                logging.info("DJNC -- >>")
                count = count + 1
                if len(testsheets) == 1:
                    return -1
                else:
                    pass
    return verInfo


def get_All_PrevDocs_Sheets_RefVers(tpBook, vPT, testSheet):
    print("In getPrevDocs function = ", vPT, testSheet)
    rowx = 6
    refList = []
    verList = []
    result_tuples =[]
    sheet = tpBook.sheets['Sommaire']

    # print("1--->handling empty string in requirement list")
    try:
        maxrow = tpBook.sheets['Sommaire'].range('A' + str(tpBook.sheets['Sommaire'].cells.last_cell.row)).end('up').row
        # KMS.showWindow(tpBook.name.split('.')[0])
        time.sleep(1)
        tpBook.sheets['Sommaire'].activate()
        while rowx <= maxrow:
            if sheet.range(rowx, 1).value is not None:
                # print(sheet.range(rowx, 1).value, str(int(float(sheet.range(rowx, 1).value))), vPT)
                # if str(vPT) in str(sheet.range(rowx, 1).value):
                if str(sheet.range(rowx, 1).value).lower() == str(vPT).lower():
                # Your code here

                    print("HHere")
                    rlo = sheet.range(rowx, 1).merge_area.row
                    rhi = sheet.range(rowx, 1).merge_area.last_cell.row
                    # print("I---->", rlo, rhi)

                    greppedTypeIndex = []
                    greppedVersionIndex = []

                    # for i in range(rlo, rhi):
                    for i in range(rlo, rhi + 1):
                        # print(i)

                        if sheet.range(i, 6).merge_cells is False:
                            if sheet.range(i, 6).value is not None:
                                refList.append(sheet.range(i, 6).value)

                        elif sheet.range(i, 6).merge_cells is True:
                            rlo = sheet.range(i, 6).merge_area.row
                            rhi = sheet.range(i, 6).merge_area.last_cell.row
                            mergedCount = rhi - rlo
                            # print("Values of rlo,rhi,mergedcount--->", rlo, rhi, mergedCount)
                            # sheet.range(i, 6).unmerge()
                            # tpBook.sheets['Sommaire'].range(i, 6).unmerge()
                            try:
                                if sheet.range(i, 6).value is not None:
                                    refList.append(sheet.range(i, 6).value)

                                    for k in range(1, mergedCount + 1):
                                        # print("value of k--->", k)
                                        if sheet.range(i + k, 6).value is None:
                                            # print("value of i+k--->", sheet.range(i + k, 6).value)
                                            sheet.range(i + k, 6).value = sheet.range(i, 6).value
                                            refList.append(sheet.range(i + k, 6).value)
                            except:
                                if sheet.range(i, 6).value is not None:
                                    refList.append(sheet.range(i, 6).value)


                        if i not in greppedVersionIndex:
                            if sheet.range(i, 7).merge_cells is False:
                                if sheet.range(i, 7).value is not None:
                                    version = re.search('[0-9]{1,2}', str(sheet.range(i, 7).value))
                                    if version is not None:
                                        # print("adding Version to version list1 =" + str(version) + "i = " + str(i))
                                        # print("adding Version to version list1  and value of i", version, i)
                                        verList.append(version.group())
                                    else:
                                        verList.append('')
                            elif sheet.range(i, 7).merge_cells is True:
                                rlo = sheet.range(i, 7).merge_area.row
                                rhi = sheet.range(i, 7).merge_area.last_cell.row
                                mergedCount = rhi - rlo
                                # print("Values of rlo,rhi,mergedcount--->", rlo, rhi, mergedCount)
                                # sheet.range(i, 7).unmerge()
                                # tpBook.sheets['Sommaire'].range(i, 7).unmerge()
                                if sheet.range(i, 7).value is not None:
                                    version = re.search('[0-9]{1,2}', str(sheet.range(i, 7).value))
                                    if version is not None:
                                        # print("adding Version to version list2 =", + str(version) + "i = " + str(i) )
                                        # print("adding Version to version list2  and value of i", version, i)
                                        verList.append(version.group())
                                    else:
                                        verList.append('')

                                    for k in range(1, mergedCount + 1):
                                        # print("value of k--->", k)
                                        if sheet.range(i + k, 7).value is None:
                                            # print("value of i+k--->", sheet.range(i + k, 7).value)
                                            sheet.range(i + k, 7).value = sheet.range(i, 7).value
                                            # print("")
                                            version = re.search('[0-9]{1,2}', str(sheet.range(i + k, 7).value))
                                            if version is not None:
                                                # print("adding Version to version list3 =" + str(version) + "i = " + str(i))
                                                # print("adding Version to version list3  and value of i", version, i)
                                                verList.append(version.group())
                                                # print("adding i+k to greppedVersionIndex", i + k)
                                                greppedVersionIndex.append(i + k)
                                            else:
                                                verList.append('')
                        else:
                            print("Already added the version to list", i)

            rowx = rowx + 1
        print("Reflist", refList)
        print("VerList", verList)
        # Create a list of tuples
        result_tuples = list(zip(refList, verList))
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"\nSomething went wrong {ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
    return result_tuples


def getPrevDocs(tpBook, vPT, reqName, reqVer, testSheet):
    logging.info("In getPrevDocs function = ", vPT, reqName, reqVer, testSheet)
    prevDoc = -1
    rowx = 6
    ipDocList = []
    refList = []
    verList = []
    sheet = tpBook.sheets['Sommaire']
    getReqList = tpBook.sheets[testSheet].range('C4').value.split("|")
    # logging.info("req list = ", getReqList)
    logging.info("before req list = ", getReqList)
    # Use list comprehension to remove empty elements
    getReqList = [item for item in getReqList if item]

    # Print the filtered list
    logging.info("after req list = ", getReqList)

    testReqName = ""
    testReqVer = ""
    logging.info("1--->handling empty string in requirement list")
    for req in getReqList:
        if req != "":
            if req.find("(") != -1:
                tempReqName = req.split("(")[0]
                logging.info("taking 1 element....")
                tempReqVer = req.split("(")[1].split(")")[0]
                if tempReqName == reqName.strip():
                    testReqName = tempReqName
                    testReqVer = tempReqVer
            else:
                logging.info("req = ", req)
                try:
                    tempReqName = req.split()[0]
                    logging.info("taking 1 element....1")
                    tempReqVer = req.split()[1]
                except:
                    tempReqVer = ''
                    tempReqName = ''
                if tempReqName == reqName.strip():
                    testReqName = tempReqName
                    testReqVer = tempReqVer

    maxrow = tpBook.sheets['Sommaire'].range('A' + str(tpBook.sheets['Sommaire'].cells.last_cell.row)).end('up').row
    KMS.showWindow(tpBook.name.split('.')[0])
    time.sleep(1)
    tpBook.sheets['Sommaire'].activate()
    while rowx < maxrow:
        if sheet.range(rowx, 1).value is not None:
            logging.info(sheet.range(rowx, 1).value, str(int(float(sheet.range(rowx, 1).value))), vPT)
            if str(vPT) in str(sheet.range(rowx, 1).value):
                logging.info("HHere")
                rlo = sheet.range(rowx, 1).merge_area.row
                rhi = sheet.range(rowx, 1).merge_area.last_cell.row
                logging.info("I---->", rlo, rhi)

                greppedTypeIndex = []
                greppedVersionIndex = []

                # for i in range(rlo, rhi):
                for i in range(rlo, rhi + 1):
                    logging.info(i)

                    if i not in greppedTypeIndex:
                        if sheet.range(i, 5).merge_cells is False:
                            if sheet.range(i, 5).value is not None:
                                ipDocList.append(sheet.range(i, 5).value)

                        elif sheet.range(i, 5).merge_cells is True:
                            rlo = sheet.range(i, 5).merge_area.row
                            rhi = sheet.range(i, 5).merge_area.last_cell.row
                            mergedCount = rhi - rlo
                            logging.info("Values of rlo,rhi,mergedcount--->", rlo, rhi, mergedCount)
                            sheet.range(i, 5).unmerge()
                            # tpBook.sheets['Sommaire'].range(i, 6).unmerge()
                            if sheet.range(i, 5).value is not None:
                                ipDocList.append(sheet.range(i, 5).value)

                                for h in range(1, mergedCount + 1):
                                    logging.info("value of k--->", h)
                                    if sheet.range(i + h, 5).value is None:
                                        logging.info("value of i+k--->", sheet.range(i + h, 5).value)
                                        sheet.range(i + h, 5).value = sheet.range(i, 5).value
                                        ipDocList.append(sheet.range(i + h, 5).value)
                                        greppedTypeIndex.append(i + h)

                    if sheet.range(i, 6).merge_cells is False:
                        if sheet.range(i, 6).value is not None:
                            refList.append(sheet.range(i, 6).value)

                    elif sheet.range(i, 6).merge_cells is True:
                        rlo = sheet.range(i, 6).merge_area.row
                        rhi = sheet.range(i, 6).merge_area.last_cell.row
                        mergedCount = rhi - rlo
                        logging.info("Values of rlo,rhi,mergedcount--->", rlo, rhi, mergedCount)
                        sheet.range(i, 6).unmerge()
                        # tpBook.sheets['Sommaire'].range(i, 6).unmerge()
                        if sheet.range(i, 6).value is not None:
                            refList.append(sheet.range(i, 6).value)

                            for k in range(1, mergedCount + 1):
                                logging.info("value of k--->", k)
                                if sheet.range(i + k, 6).value is None:
                                    logging.info("value of i+k--->", sheet.range(i + k, 6).value)
                                    sheet.range(i + k, 6).value = sheet.range(i, 6).value
                                    refList.append(sheet.range(i + k, 6).value)

                    if i not in greppedVersionIndex:
                        if sheet.range(i, 7).merge_cells is False:
                            if sheet.range(i, 7).value is not None:
                                version = re.search('[0-9]{1,2}', str(sheet.range(i, 7).value))
                                if version is not None:
                                    # logging.info("adding Version to version list1 =" + str(version) + "i = " + str(i))
                                    logging.info("adding Version to version list1  and value of i", version, i)
                                    verList.append(version.group())
                                else:
                                    verList.append('')
                        elif sheet.range(i, 7).merge_cells is True:
                            rlo = sheet.range(i, 7).merge_area.row
                            rhi = sheet.range(i, 7).merge_area.last_cell.row
                            mergedCount = rhi - rlo
                            logging.info("Values of rlo,rhi,mergedcount--->", rlo, rhi, mergedCount)
                            sheet.range(i, 7).unmerge()
                            # tpBook.sheets['Sommaire'].range(i, 7).unmerge()
                            if sheet.range(i, 7).value is not None:
                                version = re.search('[0-9]{1,2}', str(sheet.range(i, 7).value))
                                if version is not None:
                                    # logging.info("adding Version to version list2 =", + str(version) + "i = " + str(i) )
                                    logging.info("adding Version to version list2  and value of i", version, i)
                                    verList.append(version.group())
                                else:
                                    verList.append('')

                                for k in range(1, mergedCount + 1):
                                    logging.info("value of k--->", k)
                                    if sheet.range(i + k, 7).value is None:
                                        logging.info("value of i+k--->", sheet.range(i + k, 7).value)
                                        sheet.range(i + k, 7).value = sheet.range(i, 7).value
                                        logging.info("")
                                        version = re.search('[0-9]{1,2}', str(sheet.range(i + k, 7).value))
                                        if version is not None:
                                            # logging.info("adding Version to version list3 =" + str(version) + "i = " + str(i))
                                            logging.info("adding Version to version list3  and value of i", version, i)
                                            verList.append(version.group())
                                            logging.info("adding i+k to greppedVersionIndex", i + k)
                                            greppedVersionIndex.append(i + k)
                                        else:
                                            verList.append('')
                    else:
                        logging.info("Already added the version to list", i)

        rowx = rowx + 1
    logging.info("ipDoc", ipDocList)
    logging.info("VerList", verList)
    # if len(ipDocList) == len(verList):
    if (len(ipDocList) != 0) and (len(verList) != 0):
        if testReqName != "":
            for i in range(len(ipDocList)):

                if i > len(verList) - 1:
                    prevDoc = getDocPath(ipDocList[i], '')
                    logging.info("ipDocList[i], verList[i] = ", ipDocList[i], '')
                else:
                    prevDoc = getDocPath(ipDocList[i], verList[i])
                    logging.info("ipDocList[i], verList[i] = ", ipDocList[i], verList[i])

                logging.info(f"prevDoc {prevDoc}")

                if (type(prevDoc)) == str:
                    time.sleep(15)
                    # prevTableList = WDI.getTables(prevDoc)
                    # logging.info(currTableList)
                    logging.info("Testsheet")
                    # oldRqTable = WDI.findTable(prevTableList, testReqName + "(" + str(testReqVer) + ")")
                    oldRqTable = DS.find_requirement_content(prevDoc, testReqName + "(" + str(testReqVer) + ")")
                    logging.info(f"oldRqTable {oldRqTable}")

                    if oldRqTable != -1 and oldRqTable:

                        return prevDoc, testReqName, testReqVer
                    else:
                        # prevDoc=-2
                        # oldRqTable = WDI.findTable(prevTableList, testReqName + " " + str(testReqVer))
                        oldRqTable = DS.find_requirement_content(prevDoc, testReqName + " " + str(testReqVer))
                        logging.info(f"oldRqTable {oldRqTable}")
                        if oldRqTable != -1 and oldRqTable:

                            return prevDoc, testReqName, testReqVer
                        else:
                            # oldRqTable = WDI.findTable(prevTableList, testReqName + "  " + str(testReqVer))
                            oldRqTable = DS.find_requirement_content(prevDoc, testReqName + "  " + str(testReqVer))
                            logging.info("3rd", oldRqTable)
                            if oldRqTable != -1 and oldRqTable:
                                logging.info("Sucess123")
                                return prevDoc, testReqName, testReqVer
                            else:
                                # oldRqTable = WDI.findTable(prevTableList, testReqName + " (" + str(testReqVer) + ")")
                                oldRqTable = DS.find_requirement_content(prevDoc, testReqName + " (" + str(testReqVer) + ")")
                                logging.info("3rdd", oldRqTable)
                                if oldRqTable != -1 and oldRqTable:
                                    logging.info("Sucess")
                                    return prevDoc, testReqName, testReqVer
                                else:
                                    prevDoc = -1
                                # testReqName = -2
                                # testReqVer = -2
                else:
                    prevDoc = -1
                    logging.info("Files shown in summary tab of test plan NOT FOUND in Input Folder")
                    # ctypes.windll.user32.MessageBoxW(0, "Files shown in summary tab of test plan NOT FOUND in Input Folder","Evolved Requirement", 1)

        else:
            prevDoc = -2
            testReqName = -2
            testReqVer = -2
    else:
        prevDoc = -1
    # else:
    #     logging.info("VerList in the summary is not in numerics.Please change Manually and run tool again")
    #     displayInformation("\n\n#### VerList in the summary is not in numerics.Please change Manually and run tool again.  ####")
    # prevDoc = -2
    return prevDoc, testReqName, testReqVer


def parseThematics(thms):
    parsedThms = re.findall(Thematics, thms)
    thmStr = ""
    lastElem = len(parsedThms)
    count = 1
    for thm in parsedThms:
        if count == lastElem:
            thmStr = thmStr + "|" + thm
        else:
            thmStr = thmStr + thm
        count += 1
    return thmStr


def t_FindReqVer(path, testReqName, testReqVer):
    logging.info("In t_FindReqVer function = ", "*" + testReqName + "*", "*" + testReqVer + "*")
    try:
        prevTableList = WDI.getTables(path)
    except:
        path = -2
        return path, testReqName, testReqVer
    # logging.info(currTableList)
    # oldRqTable = WDI.threading_findTable(prevTableList, testReqName)
    oldRqTable = DS.find_requirement_content(path, testReqName)
    if oldRqTable == -1:
        logging.info("req name not -1")
        if (testReqName.find('.') != -1) | (testReqName.find('_') != -1):
            testReqName = testReqName.replace('.', '-')
        # oldRqTable = WDI.threading_findTable(prevTableList, testReqName + "(" + testReqVer + ")")
        oldRqTable = DS.find_requirement_content(path, testReqName + "(" + testReqVer + ")")
    else:
        # oldRqTable = WDI.threading_findTable(prevTableList, testReqName + "(" + testReqVer + ")")
        oldRqTable = DS.find_requirement_content(path, testReqName + "(" + testReqVer + ")")
    # newRqTable=WDI.findTable(currTableList, requirement)
    # logging.info(oldRqTable())
    # oldContent=WDI.searchTable(oldRqTable, req)
    if oldRqTable != -1 and oldRqTable:
        logging.info("********Scan All Doc(1) - Docx Found", path)
        return path, testReqName, testReqVer
    else:
        # prevDoc=-2
        # oldRqTable = WDI.findTable(prevTableList, testReqName + " " + str(testReqVer))
        oldRqTable = DS.find_requirement_content(path, testReqName + " " + str(testReqVer))
        if oldRqTable != -1 and oldRqTable:
            logging.info("********Scan All Doc(2) - Docx Found", path)
            return path, testReqName, testReqVer
        else:
            # oldRqTable = WDI.findTable(prevTableList, testReqName + "  " + str(testReqVer))
            oldRqTable = DS.find_requirement_content(path, testReqName + "  " + str(testReqVer))
            logging.info("3rd", oldRqTable)
            if oldRqTable != -1 and oldRqTable:
                logging.info("Sucess")
                logging.info("********Scan All Docx(1) - Doc Found", path)
                return path, testReqName, testReqVer
            else:
                # oldRqTable = WDI.findTable(prevTableList, testReqName + " (" + str(testReqVer) + ")")
                oldRqTable = DS.find_requirement_content(path, testReqName + " (" + str(testReqVer) + ")")
                logging.info("3rd", oldRqTable)
                if oldRqTable != -1 and oldRqTable:
                    logging.info("Sucess")
                    logging.info("********Scan All Docx(2) - Doc Found", path)
                    return path, testReqName, testReqVer
                else:
                    logging.info("********Scan All Docx(3) -Requirement not found in ", path)
                    path = -2
                    logging.info("3rd output", testReqName, testReqVer)
                    return path, testReqName, testReqVer
                # testReqName = -2
                # testReqVer = -2


def scanAllDocs(tpBook, reqName, reqVer, testSheet):
    logging.info("**********Scanning All Docs in Input Folder******************")
    getReqList = tpBook.sheets[testSheet].range('C4').value.split("|")
    # logging.info("req list = ", getReqList)
    logging.info("before req list = ", getReqList)
    # Use list comprehension to remove empty elements
    getReqList = [item for item in getReqList if item]

    # Print the filtered list
    logging.info("after req list = ", getReqList)
    testReqName = ""
    testReqVer = ""
    logging.info("2--->handling empty string in requirement list")

    # for req in getReqList:
    #     if req.find("(")!=-1:
    #         tempReqName = req.split("(")[0]
    #         tempReqVer = req.split("(")[1].split(")")[0]
    #         if tempReqName==reqName.strip():
    #             testReqName = tempReqName
    #             testReqVer = tempReqVer
    #     else:
    #         logging.info("req = ", req)
    #         tempReqName = req.split()[0]
    #         try:
    #             tempReqVer = req.split()[1]
    #         except:
    #             tempReqVer = ''
    #         if tempReqName==reqName.strip():
    #             testReqName = tempReqName
    #             testReqVer = tempReqVer

    for req in getReqList:
        if req != "":
            if req.find("(") != -1:
                tempReqName = req.split("(")[0]
                tempReqVer = req.split("(")[1].split(")")[0]
                if tempReqName == reqName.strip():
                    testReqName = tempReqName
                    testReqVer = tempReqVer
            else:
                logging.info("req = ", req)
                try:
                    tempReqName = req.split()[0]
                    tempReqVer = req.split()[1]
                except:
                    tempReqVer = ''
                    tempReqName = ''
                if tempReqName == reqName.strip():
                    testReqName = tempReqName
                    testReqVer = tempReqVer

    if testReqName != "":
        pat = ICF.getInputFolder() + "\\"
        onlyfiles = [f for f in listdir(pat) if isfile(join(pat, f))]
        logging.info("+++++++testReqName != """, pat, onlyfiles)
        for fileName in onlyfiles:
            logging.info("Filename & docName = ", fileName)
            futures = []
            finaloutput = ()
            nThreads = len(onlyfiles)
            logging.info("no. of threads scan doc", nThreads)
            with ThreadPoolExecutor(max_workers=nThreads) as exe:
                # for files in onlyfiles:
                logging.info("inside thread files")
                if os.path.splitext(fileName)[1] == ".docx":
                    path = pat + fileName
                    logging.info("Path found === ", path)
                    logging.info("req and ver", testReqName, testReqVer)
                    futures.append(exe.submit(t_FindReqVer, path, testReqName, testReqVer))
                elif (os.path.splitext(fileName)[1] == ".doc") or (os.path.splitext(fileName)[1] == ".docm") or (
                        os.path.splitext(fileName)[1] == ".rtf"):
                    logging.info(".doc Name ", fileName)
                    path = save_as_docx(pat + fileName)
                    futures.append(exe.submit(t_FindReqVer, path, testReqName, testReqVer))
                else:
                    path = -2

            for future in concurrent.futures.as_completed(futures):
                path, testReqName, testReqVer = future.result()
                if type(path) == str:
                    logging.info("REQ Found")
                    for f in futures:
                        f.cancel()
                    # concurrent.futures.Future.cancel()
                    # future.result()
                    logging.info("futur.result", future.result())
                    # output.append(future.result())
                    # return path, testReqName, testReqVer
                    return future.result()
    else:
        path = -1
    logging.info("***output", path, testReqName, testReqVer)
    return path, testReqName, testReqVer


def modifySignalName(req_id, dci_sheet, ):
    # searchR
    pass


def getInterfaceReqList(rqIDs):
    interfaceReqList = []
    for feps in rqIDs:
        for reqname in dict(rqIDs[feps]):
            rqList = rqIDs[feps][reqname]
            if reqname == 'Interfaces':
                for item in rqList:
                    interfaceReqList.append(item)
    return interfaceReqList


def getReference():
    reference = []
    for ref in ICF.getDocToDownload():
        reference.append(ref["Reference"])
        logging.info("Reference = " + str(reference))
    return reference


def getVersion():
    version = []
    for ref in ICF.getDocToDownload():
        version.append(ref["Version"])
        logging.info("Version = " + str(version))
    return version


def removeDuplicates(lst):
    return [t for t in (set(tuple(i) for i in lst))]


UpdateHMIInfoCb = None


def registerInfoTextBox(func):
    global UpdateHMIInfoCb
    UpdateHMIInfoCb = func


errorPopupCb = None


def registerErrorPopup(func1):
    global errorPopupCb
    errorPopupCb = func1


def displayErrorPopup():
    UpdateHMIInfoCb("Invalid InputDocument Name. Correct the document name and rerun the tool")
    errorPopupCb(
        'Input Document Name in Analyse_de_entrant sheet is not in the right format. '
        'Please change the format to "Filename -referenceName versionNumber"')


# def getKPIDocPath(path):
#     docList = []
#     documents = os.listdir(path)
#     logging.info("--", documents)
#     for d in documents:
#         a = (path + "\\" + d)
#         docList.append(a)
#     return docList

def getKPIDocPath(path):
    docList = []
    documents = os.listdir(path)
    logging.info("--", documents)
    for d in documents:
        a = (path + "\\" + d)
        if d.find("") != -1:
            docList.append(a)
    return docList


def getArch(taskname):
    arch = "VSM"
    x = re.findall("^F_", taskname)
    if x:
        arch = "BSI"
    return arch


# Input - requirement name
# Description - finding whether the given requirement is having the DELETE keyword or not
# Output - returns a dictionary with values OrgReq(requirement with keyword DELETE), ModifiedReq(requirement without keyword DELETE), is_delete_req(0 - not having DELETE keyword, 1 - having DELETE keyword)
def findDeleteReq(reqId):
    delReqResult = {
        'OrgReq': '',
        'ModifiedReq': '',
        'Ver': '',
        'is_delete_req': 0
    }
    if reqId is None or reqId == "":
        return delReqResult
    DelReqModified = ''
    is_deleted_req = 0
    delReqResult['OrgReq'] = reqId
    version = ""
    logging.info('reqId >> ', reqId)
    if reqId.find("DELETE") != -1:
        if reqId.find("("):
            splitRequirements = reqId.split("(")
            Requirement = splitRequirements[0]
            logging.info(Requirement)
            if len(reqId.split("(")) > 1:
                splitVersion = splitRequirements[1]
                vers = splitVersion.split(")")
                version = "(" + vers[0] + ")"
                logging.info(version)
            else:
                splitVersion = ""
                vers = ""
                version = ""
                logging.info(version)

        else:
            splitRequirements = reqId.split(" ")
            Requirement = splitRequirements[0]
            logging.info(Requirement)
            if len(reqId.split(" ")) > 1:
                vers = splitRequirements[1]
                version = " " + vers[0]
                logging.info(version)
            else:
                vers = ""
                version = ""
                logging.info(version)
        DelReqModified = removeDeleteKeyword(reqId)
        is_deleted_req = 1
    delReqResult['ModifiedReq'] = DelReqModified
    delReqResult['is_delete_req'] = is_deleted_req
    delReqResult['Ver'] = version
    logging.info("delReqResult >> ", delReqResult)
    return delReqResult


# Input - requirement name
# Description - removing the keyword DELETE and the special characters
# Output - returns a requirement without keyword DELETE(eg: I/P: REQ-XXX DELETE !, O/P: REQ-XXX)
def removeDeleteKeyword(reqId):
    modified_reqId = re.sub(DeleteReqPattern, "", str(reqId)).lstrip().rstrip()
    modified_reqId = re.sub(DeleteReqPattern, "", str(modified_reqId)).strip()
    logging.info("modified_reqId >> ", modified_reqId)
    return modified_reqId


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

def getCurrentDocPath(feps, rqIDs):
    currDoc = ''
    for inputdoc in rqIDs[feps]['Input_Docs']:
        logging.info("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
        logging.info("inputdoc --123 ", inputdoc)
        currVer = re.search(pattren_ver, inputdoc)
        if currVer.group() is not None:
            version = currVer.group().upper()
            if "V" in version:
                version = version.replace("V", "")
                logging.info(str(int(float(version))))
                currVer = str(int(float(version)))
            else:
                logging.info("OOPs invalid Version format. Character 'V' not found")
                currVer = -1
        else:
            logging.info("No Version found in inputdoc", inputdoc)
        logging.info("Current Version = ", currVer)
        currDoc = getDocPath(inputdoc, currVer)  # get path of current doc from input folder
        if currDoc != -1:
            return currDoc
    return currDoc


# # these function is used to search the previous req in the searching logic folder. If doc is present.
# def search_requirement(requirement_id, file_path):
#     reqName, reqVer = NRH.getReqVer(requirement_id)
#     con = DS.find_requirement_content(file_path, reqName + "(" + reqVer + ")")
#     print("con0 $---->", con)
#
#     if not con:  # Check if con is empty
#         con = DS.find_requirement_content(file_path, reqName + " (" + reqVer + ")")
#         print("con1 $---->", con)
#
#     if not con:  # Check if con is empty
#         con = DS.find_requirement_content(file_path, reqName + "  (" + reqVer + ")")
#         print("con1.1 $---->", con)
#
#     if not con:  # Check if con is still empty
#         con = DS.find_requirement_content(file_path, reqName + " " + reqVer)
#         print("con2 $---->", con)
#
#     if not con:  # Check if con is still empty
#         con = DS.find_requirement_content(file_path, reqName + "  " + reqVer)
#         print("con3 $---->", con)
#     return con

def search_requirement(requirement_id, file_path):
    searchResult = ''
    table = ''
    reqName, reqVer = NRH.getReqVer(requirement_id)
    variations = [
        reqName + "(" + reqVer + ")",
        reqName + " (" + reqVer + ")",
        reqName + "  (" + reqVer + ")",
        reqName + " " + reqVer,
        reqName + "  " + reqVer,
    ]

    searchResult = None  # Initialize con to None

    for variation in variations:
        print("variation--------->", variation)
        searchResult, table, file_path = DS.find_requirement_content(file_path, variation)
        # print(f"con $----> {variation}: {con}")
        print(f"con $----> {variation}: {searchResult}")
        if searchResult and searchResult[0]:  # Check if the first element of con is not empty
            break  # Break the loop if content is found

    return searchResult, table, file_path



def findReqinSearchLogicDoc(req, reqSplits):
    display_info = UpdateHMIInfoCb
    logging.info("req, reqSplits------>",req, reqSplits)
    if "-->" in req or "==>" in req or "->" in req or "=>" in req:
        req = req.split("-->")[-1].split("==>")[-1].split("->")[-1].split("=>")[-1].strip()
    else:
        req = req
    logging.info("req----->", req)
    req = req.split(' ')[0]
    reqSplit = []
    for i in reqSplits:
        if req in i:
            reqSplit.append(i)
    Search_flag = 0
    reqName = ''
    reqVer = ''
    logging.info("reqSplit1212 ", reqSplit)
    for i in range(len(reqSplit)):
        if len(reqSplit[i]) != 0:
            # logging.info("reqSplit[i] = ", reqSplit[i])
            if reqSplit[i].find("(") != -1:
                reqName = reqSplit[i].split("(")[0]
                reqVer = reqSplit[i].split("(")[1].split(")")[0]
            else:
                reqName = reqSplit[i]
                reqVer = ""
            logging.info("From TS", reqName, reqVer)
    requirement_id = reqName + " " + reqVer
    logging.info("requirement_id--->", requirement_id)
    file_path = os.path.abspath(r'..\Input_Files\Search_Logic')

    if os.path.exists(file_path) and os.path.isdir(file_path):
        # List all files in the folder
        files = os.listdir(file_path)

        # Filter for .docx files
        docx_files = [f for f in files if f.lower().endswith('.docx') or f.lower().endswith('.doc')]

        # Check if any .docx files are found
        if docx_files:
            for docx_file in docx_files:
                file_path = os.path.join(file_path, docx_file)
                logging.info("Found .docx file:", file_path)
                try:
                    # c = DS.find_requirement_content(file_path, requirement_id)
                    c = search_requirement(requirement_id, file_path)
                    logging.info("c----->", c)
                    if c:
                        Search_flag = 1
                        logging.info("req is present in the doc")
                        display_info(f"{reqSplit[0]} is present in the SearchLogic Doc.")
                    else:
                        Search_flag = -1
                        logging.info("req is not present in the doc")
                        display_info(f"{reqSplit[0]} is not present in the SearchLogic Doc searching in vPT to get Previous Version.")
                except:
                    logging.info("SearchLogic Output file not present in the folder")
                    display_info(f"Search_Logic Output file not present in the folder.")
                    Search_flag = -1

        else:
            logging.info("No .docx or .doc files found in the folder.")
            display_info(f"No .docx or .doc files found in the folder.")
            Search_flag = -1
            # Add your logic here for when no .docx files are found
    else:
        logging.info("Search_Logic Folder does not exist in Input Folder.")
        display_info(f"Search_Logic Folder does not exist in Input Folder.")
        Search_flag = -1
        # Add your logic here for when the folder doesn't exist
    return file_path, reqName, reqVer, Search_flag


def evolPreviousFun(tpBook,currDoc,prevDocc,listOfTestSheets,reqName,reqVer,Arch,newReq,req,macro,listOfSF_TestSheets,testReqName,testReqVer, oldReq,req_ver_sf,flag):
    testSheets = []
    if newReq == "":
        newReq = req
    logging.info("newReq-****>", newReq)
    display_info = UpdateHMIInfoCb
    dciBook = EI.openGlobalDCI()
    logging.info("SUCCESS")
    alertDoc = EI.openAlertDoc()
    logging.info("SUCCESS")
    ssFiches = EI.openSousFiches()
    logging.info("SUCCESS")
    refEC = EI.openReferentialEC()
    logging.info("SUCCESS")
    logging.info("ALL DOCUMENTS OPENED")
    time.sleep(1)
    KMS.showWindow(tpBook.name.split('.')[0])

    # commented on 21-4-2023 start
    logging.info("reqName, reqVer, Arch,newReq=newReq1->",reqName, reqVer, Arch, newReq)
    analyseThematics = AT.AnalyseThematics(tpBook, refEC, currDoc, prevDocc,
                                           listOfTestSheets, reqName, reqVer, Arch,
                                           newReq=newReq)

    logging.info(f"currDoc,prevDoc1 {currDoc, prevDocc}")
    logging.info("reqName, reqVer, Arch,newReq=newReq2->", reqName, reqVer, Arch, newReq)
    analyseContents = ATS.AnalyseTestSheet(dciBook, tpBook, alertDoc, ssFiches, currDoc,
                                           prevDocc, listOfTestSheets, reqName, reqVer,
                                           newReq=newReq)

    try:
        analyseThematics.Analyse()
        logging.info("closed  analyseThematics")
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"Exception in analyse thematics = {ex} {exc_tb.tb_lineno}")

        flag = 1
        logging.info("Unable to Analyse Thematics")
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines(
                "\n\nUnable to analyse thematique of " + req + ". Please proceed manually")
        time.sleep(2)
    try:
        analyseContents.Analyse()
        logging.info("closed  analyseContents")
        # alertDoc.close()
        # refEC.close()
        # logging.info("Documents closed after analyseContents")

    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"Exception in analyse content = {ex} {exc_tb.tb_lineno}")

        flag = 1
        logging.info("Unable to Analyse Content")
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines(
                "\n\nUnable to analyse content of " + req + ". Please proceed manually")
        time.sleep(2)

    # commented on 21-4-2023 end

    dciBook.close()
    alertDoc.close()
    ssFiches.close()
    refEC.close()
    logging.info("ALL DOCUMENTS CLOSED")

    # commented on 21-4-2023 start

    themImpactComment = AT.themImpactComment.copy()
    logging.info("themImpactComment in main(1) = ", themImpactComment)
    themImpactComment = list(dict.fromkeys(themImpactComment))
    logging.info("themImpactComment in main(2) = ", themImpactComment)
    contentFuncImpact = ATS.contentFuncImpact.copy()
    logging.info("contentFuncImpact in main(1) = ", contentFuncImpact)
    contentFuncImpact = list(dict.fromkeys(contentFuncImpact))
    logging.info("contentFuncImpact in main(2) = ", contentFuncImpact)

    # commented on 21-4-2023 end

    if flag == 0:
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nSuccessfully Filled Impact Sheet")
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nRequirement completely Analysed")
        time.sleep(2)

    else:
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nRequirement not treated sucessfully")
        time.sleep(2)
    for testSheet in listOfTestSheets:  # testSheet is string and not tpBook object
        logging.info("For adding requirement", testSheet, tpBook)

        testSheets.append(testSheet)
        logging.info("b------>", testSheets)
        EI.activateSheet(tpBook, testSheet)
        time.sleep(1)
        if tpBook.sheets[testSheet].range('C7').value == 'VALIDEE':
            TPM.selectTestSheetModify(macro)  # for changing version

        # defined empty list on 21-4-2023 start
        # themImpactComment = []
        # contentFuncImpact = []

        # Document thematic result
        logging.info("reqName, reqVer, Arch,newReq=newReq3->", reqName + "(" + reqVer + ")", testReqName + " " + testReqVer)
        # EI.addEvovledReq(themImpactComment, contentFuncImpact, tpBook, testSheet, reqName + " " + reqVer, newReq)
        # EI.addEvovledReq(themImpactComment, contentFuncImpact, tpBook, testSheet, testReqName + " " + testReqVer, reqName + " " + reqVer)
        EI.addEvovledReq(themImpactComment, contentFuncImpact, tpBook, testSheet, testReqName + " " + testReqVer, newReq)

        # if ICF.getBackLog() is True:
        #     logging.info("Backlog Selected")
        #     display_info(f"\n========={testSheet.name}===========")
        #     refEC = EI.openReferentialEC()
        #     kpiDocList = getKPIDocPath(ICF.getInputFolder() + "/KPI")
        #     # rawReqs = analyseThematics.getTsRawReq()
        #     rawReqs = testSheet.range('C4').value
        #     ts_reqs = [rawReqs]
        #     currArch = getArch(ICF.FetchTaskName())
        #     logging.info("*** kpiDocList ***", kpiDocList)
        #     logging.info("*** reqs ***", ts_reqs)
        #     logging.info("*** vvvvvvvvvvvvv ***", ts_reqs)
        #     logging.info("*** refEC ***", str(refEC))
        #     logging.info("*** currArch ***", currArch)
        #     ts_req_list = removeInterfaceReq(ts_reqs)
        #     logging.info("vvvvvvvv2342521525---->",ts_req_list)
        #
        #     start_time_bl = time.time()
        #     logging.info(f'\n BL start time: {start_time_bl}')
        #     status, combinedThemLines = getCombinedThematicLines(kpiDocList, ts_req_list,
        #                                                          refEC, currArch)
        #     end_time_bl = time.time()
        #     execution_time_bl = end_time_bl - start_time_bl
        #     logging.info(f'\n BL end execution time: {execution_time_bl}')
        #
        #     refEC.close()
        #     logging.info("***Thematic Combinations = ***", status, combinedThemLines)
        #     if status:
        #         logging.info("Update the Thematic Combinations in the View")
        #         # display_info(f"========={testSheet.name} - Backlog Output===========")
        #         display_info(f"Backlog Output:\n")
        #         display_info(str(combinedThemLines))
        #     else:
        #         logging.info(f"Requirement Not available in KPI Sheet. Proceed Manually - {ts_reqs}")
        #     #    display_info(
        #      #       "Requirement Not available in KPI Sheet. Make sure all the requirement in the test sheet is #mentioned in KPI Sheet")
        # else:
        #     logging.info('Backlog not selected')
    ssfiches = EI.openSousFiches()
    if (len(listOfSF_TestSheets) > 0):
        # logging.info("index of requirement",l)
        # flg=1
        SFE.treat_SF_evolved_req(tpBook, ssfiches, listOfSF_TestSheets, themImpactComment,
                                 contentFuncImpact, testReqName, testReqVer, oldReq, newReq
                                 )
        # if (newReq == ""):
        if (newReq == oldReq):
            comm = f"Evolved requirement.Incremented {reqName} from version {testReqVer} to {reqVer}"
        else:
            comm = f"Evolved requirement.Changed requirement name from {oldReq} to {newReq}"
        SFE.QIA_ssfiche_dict(reqName, reqVer, listOfSF_TestSheets, comm, req_ver_sf)
    ssfiches.save()
    ssfiches.close()
    return testSheets


def evolReq(tpBook, fepsNumber, macro, Arch, feps, rqIDs, reqname, req_ver_sf, fepsForDuplicateReqs, requirement=""):
    testSheets = ""
    display_info = UpdateHMIInfoCb
    # rqList = rqIDs[feps][reqname]
    # for req in rqList:
    req = requirement
    oldReq = ""
    newReq = ""
    flag = 0
    logging.info(
        "-------------------------------------" + req + "-------------------------------------")
    with open('../Aptest_Tool_Report.txt', 'a') as f:
        f.writelines(
            "\n\n------------------------------------------------" + req + " ----------------------------------------------------")
    if req.find("->") != -1:
        if len(req.split("-->")) >= 2:
            logging.info("old to new conversion", req)
            tempReq = req.split("-->")[0]
            oldReq = "".join(tempReq.split())
            logging.info("oldReq = ", oldReq)
            newReq = req.split("-->")[1]
            logging.info("newReq = ", newReq)
        else:
            logging.info("only 1 req name2")
            oldReq = req
            logging.info("ReqName = ", oldReq)
    else:
        logging.info("IN EXCEPT")
        if len(req.split("==>")) >= 2:
            logging.info("old to new conversion2", req)
            tempReq = req.split("==>")[0]
            oldReq = "".join(tempReq.split())
            logging.info("oldReq = ", oldReq)
            newReq = req.split("==>")[1]
            logging.info("newReq = ", newReq)
        elif len(req.split("=>")) >= 2:
            logging.info("old to new conversion", req)
            tempReq = req.split("=>")[0]
            oldReq = "".join(tempReq.split())
            logging.info("oldReq = ", oldReq)
            newReq = req.split("=>")[1]
            logging.info("newReq = ", newReq)
        else:
            logging.info("only 1 req name3")
            oldReq = req
            logging.info("ReqName = ", oldReq)
    with open('../Aptest_Tool_Report.txt', 'a') as f:
        f.writelines("\n\nFilling Impact Sheet")
    with open('../Aptest_Tool_Report.txt', 'a') as f:
        f.writelines("\n\nFilled Impact sheet and got the VPT and test sheets to be treated")

    currDocEvol = getCurrentDocPath(feps, rqIDs)
    logging.info("currDocEvol1------->", currDocEvol)
    flag = TA.checkReq(currDocEvol, oldReq, Arch, newReq)
    if flag == -3:
        display_info("\n\nAnalysis de entrance input file not present in Input_Files folder")
        errorPopupCb(
                f'Analysis de entrance input file name present under the {feps} not Match in Input_Files folder. Please change the document name and re run the tool.')
        print(f'Analysis de entrance input file name present under the {feps} not Match in Input_Files folder. Please change the document name and re run the tool.')
        return -2
    # verInfo, oldReqName, oldReqVer = fillImpactEvolved(tpBook, oldReq, fepsNumber, flag, Arch, newReq)
    verInfo, reqSplit = fillImpactEvolved(tpBook, oldReq, fepsNumber, flag, Arch, rqIDs, fepsForDuplicateReqs, newReq)
    # verInfo = fillImpactEvolved(tpBook, oldReq, fepsNumber, flag, Arch, newReq)
    # reqSplit = ''
    logging.info("verInfo, oldReqName, oldReqVer----->",verInfo, reqSplit)
    logging.info("----dc1111111", verInfo)
    if flag == 1:
        if (verInfo != -1) and (verInfo != -2):
            vPT = verInfo[0][0]
            logging.info("Version of PT", vPT)
            prevDocc, testReqName, testReqVer, Search_flag = findReqinSearchLogicDoc(req, reqSplit)
            logging.info("prevDocc------->", prevDocc)
            if Search_flag == -1:
                if vPT != -1:
                    try:
                        # documentLinks = IDLP.getDocLinks(tpBook, str(vPT))
                        olddocs = IDLP.getDocRefandVersion(tpBook, str(vPT))
                        logging.info(f"olddocs {olddocs}")
                        logging.info("Starting download of previous version ")
                        if ICF.getAutoDownloadStatusPreviousDocument():
                            startDocumentDownload(olddocs)
                            display_info("Valid reference number found, Downloading the Previous documents")

                        # IDLP.showLinkPopUp(documentLinks, str(vPT))
                        # with open('../Aptest_Tool_Report.txt', 'a') as f:
                        #     f.writelines("\n\nAll documents of PT version " + str(
                        #         vPT) + " downloaded for requirement - " + req)
                        time.sleep(2)
                    except Exception as ex:
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        logging.info(f"Exception caused while downloading the old documents {ex} line: {exc_tb.tb_lineno}.")
                else:
                    logging.info("No last version info found so no links to display")
                logging.info("loop over vPT1 ------>", vPT)
            for inputdoc in rqIDs[feps]['Input_Docs']:
                logging.info("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
                logging.info("inputdoc --123 ", inputdoc)
                currVer = re.search(pattren_ver, inputdoc)
                if currVer.group() is not None:
                    version = currVer.group().upper()
                    if "V" in version:
                        version = version.replace("V", "")
                        logging.info(str(int(float(version))))
                        currVer = str(int(float(version)))
                    else:
                        logging.info("OOPs invalid Version format. Character 'V' not found")
                        currVer = -1
                else:
                    logging.info("No Version found in inputdoc", inputdoc)
                logging.info("Current Version = ", currVer)
                # currVer = str(int(float((re.search("[0-9]{Exception caused while downloading the old documents1,2}.[0-9]",(inputdoc.upper().split(" V")[1].split(" ")[0])).group()))))
                currDoc = getDocPath(inputdoc, currVer)  # get path of current doc from input folder
                prevDoc = -1
                logging.info("Input Doc ", inputdoc)
                logging.info("Current Doc Version = ", currVer)
                logging.info("Current Doc Path = ", currDoc)
                listOfTestSheets = []
                listOfSF_TestSheets = []
                hPT = []
                if (type(currDoc) == str):
                    logging.info("11111111111111111111111111111111111")
                    for vPT, testSheet, reqName, reqVer in verInfo:
                        if vPT != -1:
                            versionPT = vPT
                            logging.info("loop over vPT3 ------>", vPT)
                            c = hPT.append(vPT)
                        if testSheet.name.find("SF") == -1:
                            listOfTestSheets.append(testSheet)
                            logging.info("loop over vPT4 ------>", vPT)
                            c = hPT.append(vPT)
                        else:
                            # with open('../Aptest_Tool_Report.txt', 'a') as f:
                            #     f.writelines(
                            #         "\n\nRequirement - " + req + " is present in SF sheet so according modify the sous fiches manually")
                            # time.sleep(2)
                            listOfSF_TestSheets.append(testSheet.name)
                            logging.info("loop over vPT5 ------>", vPT)
                            hPT.append(vPT)
                        reqName = reqName
                        reqVer = reqVer
                    logging.info("hPT--c--------->",hPT)
                    d = list(set(hPT))
                    logging.info("d------>",d)
                    if Search_flag == -1:
                        if vPT != -1:
                            logging.info("loop over vPT and testSheetList--->", vPT)
                            prevDoc, testReqName, testReqVer = getPrevDocs(tpBook, vPT, reqName, reqVer,
                                                                           testSheet)  # get previous version of the ipDoc
                            logging.info("Path of old version doc ", prevDoc)
                            logging.info("Path of new version doc ", currDoc, currVer)
                            if prevDoc == -1:
                                prevDoc, testReqName, testReqVer = scanAllDocs(tpBook, reqName, reqVer,
                                                                               testSheet)  # get previous version of the ipDoc
                                logging.info("Path of old version doc ", prevDoc)
                                logging.info("Path of new version doc ", currDoc, currVer)
                            # else:
                            #     # version or refernce not mentioned correctly in summary tab
                            #     if prevDoc == -2:
                            #         return -2
                        else:
                            prevDoc, testReqName, testReqVer = scanAllDocs(tpBook, reqName, reqVer,
                                                                           testSheet)  # get previous version of the ipDoc
                            logging.info("Path of old version doc ", prevDoc)
                            logging.info("Path of new version doc ", currDoc, currVer)
                        break
            if (len(listOfTestSheets) == 0) and (type(currDoc) == str):
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines(
                        "\n\n" + req + " cannot be treated as requirement in not present in GC or N1 sheet")

            if Search_flag == -1:
                logging.info("Search_flag-----111111111")
                prevDoc, testReqName, testReqVer = scanAllDocs(tpBook, reqName, reqVer,
                                                               testSheet)  # get previous version of the ipDoc
                logging.info("testReqName, testReqVer77777777--->",testReqName, testReqVer)
                logging.info("Path of old version doc ", prevDoc)
                if((type(prevDoc) == str) and (type(currDoc) == str)):
                    logging.info("Search_flag---------11111111111111111")
                    logging.info("reqName, reqVer, Arch, newReq000--------->", reqName, reqVer, Arch, newReq)
                    logging.info("newReq, req000--------->", newReq, req)
                    testSheets = evolPreviousFun(tpBook, currDoc, prevDoc, listOfTestSheets, reqName, reqVer, Arch, newReq, req, macro,
                             listOfSF_TestSheets, testReqName, testReqVer, oldReq, req_ver_sf,flag)

            if Search_flag == 1:
                logging.info("flagg1111111111111111")
                logging.info("reqName, reqVer, Arch, newReq1111--------->", reqName, reqVer, Arch, newReq)
                logging.info("newReq, req1111--------->", newReq, req)
                if ((type(prevDocc) == str) and (type(currDoc) == str)):
                    testSheets = evolPreviousFun(tpBook, currDoc, prevDocc, listOfTestSheets, reqName, reqVer, Arch, newReq, req, macro,
                             listOfSF_TestSheets, testReqName, testReqVer, oldReq, req_ver_sf,flag)


            elif (currDoc == -1) or (prevDoc == -1):
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines(
                        "\n\n" + req + " cannot be treated as the input documents given in Analyse_de_entrant are not found in input folder")
            elif (currDoc == -2) or (prevDoc == -2):
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines(
                        "\n\n" + req + " cannot be treated as the input documents in input folder have different format or the version is not same as mentioned in summary tab")
            else:
                logging.info("files not found in input folder")
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines(
                        "\n\nUnable to analyse content of " + req + " because files not found in input folder with .doc or .docx format. Please proceed manually")
                time.sleep(2)

                # add pop up provide link to docInfo
        elif (verInfo == -2):
            time.sleep(2)
            pass
        else:
            # display_info(str(req) + " cannot be treated as it is not impacted in testplan." + "\n")
            # ctypes.windll.user32.MessageBoxW(0, req + " cannot be treated as it is not impacted in testplan \n OR \n If " + req + " impacted in testplan last version information of requirement not found", "Evolved Requirements", 1)
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines("\n\n" + req + " cannot be treated as it is not impacted in testplan.")
            time.sleep(2)
    elif flag == -1:
        if newReq != '':
            oldReq = newReq
        display_info(f'The thematic lines of the "{oldReq}" are NA for {Arch}, Proceed Manually.')
    return testSheets

def treate_backlog(testSheets):
    display_info = UpdateHMIInfoCb
    new_list = []
    seen_sheets = []
    if ICF.getBackLog() is True:
        logging.info("Backlog Selected")
        logging.info("testsheets impacted for all reqs--->",testSheets)
        testSheets = list(set([sheet for sublist in testSheets for sheet in sublist]))
        logging.info("bbbbbb------33333",testSheets)
        for testSheet in testSheets:
            logging.info("testsheet0000000------>",testSheet)
            testSheet.activate()
            display_info(f"\n========={testSheet.name}===========")
            refEC = EI.openReferentialEC()
            kpiDocList = getKPIDocPath(ICF.getInputFolder() + "/KPI")
            # rawReqs = analyseThematics.getTsRawReq()
            rawReqs = testSheet.range('C4').value
            ts_reqs = [rawReqs]
            currArch = getArch(ICF.FetchTaskName())
            logging.info("*** kpiDocList ***", kpiDocList)
            logging.info("*** reqs ***", ts_reqs)
            logging.info("*** vvvvvvvvvvvvv ***", ts_reqs)
            logging.info("*** refEC ***", str(refEC))
            logging.info("*** currArch ***", currArch)
            ts_req_list = removeInterfaceReq(ts_reqs)
            logging.info("vvvvvvvv2342521525---->", ts_req_list)

            start_time_bl = time.time()
            logging.info(f'\n BL start time: {start_time_bl}')
            status, combinedThemLines = getCombinedThematicLines(kpiDocList, ts_req_list,
                                                                 refEC, currArch)
            end_time_bl = time.time()
            execution_time_bl = end_time_bl - start_time_bl
            logging.info(f'\n BL end execution time: {execution_time_bl}')

            refEC.close()
            logging.info("***Thematic Combinations = ***", status, combinedThemLines)
            if status:
                logging.info("Update the Thematic Combinations in the View")
                # display_info(f"========={testSheet.name} - Backlog Output===========")
                display_info(f"Backlog Output:\n")
                display_info(str(combinedThemLines))
            else:
                logging.info(f"Requirement Not available in KPI Sheet. Proceed Manually - {ts_reqs}")
            #    display_info(
            #       "Requirement Not available in KPI Sheet. Make sure all the requirement in the test sheet is #mentioned in KPI Sheet")
    else:
        logging.info('Backlog not selected')

def findDCI_Applicability(themCodes):
    # finding the applicability for LVM or LYQ if present in Q colum in dci file
    archiList = []

    combined_archi = ""
    for thm_code in themCodes:
        if 'LVM_01' == thm_code or 'LYQ_01' == thm_code:
            if 'R1' not in archiList:
                archiList.append('R1')
        if 'LVM_02' == thm_code or 'LYQ_02' == thm_code or thm_code == 'LVM_03':
            if 'R2' not in archiList:
                archiList.append('R2')
        # if thm_code == 'LVM_03':
        #     if 'R3' not in archiList:
        #         archiList.append('R3')
    logging.info(f"archiList BF {archiList}")
    combined_archi = "".join(sorted(archiList))
    logging.info(f"combined_archi {combined_archi}")
    return combined_archi


def finalizeSheetArchi(testSheetThmList):
    # finalize the architecture for the sheet
    # eg. ['R1R2', 'R1']
    logging.info(f"\n\ntestSheetThmList {testSheetThmList}")
    # ('R1' in testSheetThmList and ('R1R2R3' in testSheetThmList or 'R1R2' in testSheetThmList)) or ('R2' in testSheetThmList and ('R1R2R3' in testSheetThmList or 'R1R2' in testSheetThmList)) or ('R2' in testSheetThmList and 'R1' in testSheetThmList) or
    if (
            'R1R2' in testSheetThmList or 'R1R2R3' in testSheetThmList) and 'R1' not in testSheetThmList and 'R2' not in testSheetThmList:
        logging.info("iffff...")
        return 'R1R2'
    elif (
            'R1' in testSheetThmList and 'R1R2' not in testSheetThmList and 'R1R2R3' not in testSheetThmList and 'R2' not in testSheetThmList) or (
            'R1' in testSheetThmList and (
            'R1R2' in testSheetThmList or 'R1R2R3' in testSheetThmList) and 'R2' not in testSheetThmList):
        logging.info("elif...")
        return 'R1'
    elif (
            'R2' in testSheetThmList and 'R1R2' not in testSheetThmList and 'R1R2R3' not in testSheetThmList and 'R1' not in testSheetThmList) or (
            'R2' in testSheetThmList and (
            'R1R2' in testSheetThmList or 'R1R2R3' in testSheetThmList) and 'R1' not in testSheetThmList):
        logging.info("eliffff...")
        return 'R2'
    else:
        return -1


def filterThemArch(thematicLine, refEC, ARCH):
    logging.info("In filterThemArch === ", thematicLine)
    ListOfThematics = thematicLine.split("|")
    time.sleep(1)
    sheet = refEC.sheets['Liste EC']
    logging.info("sheet = ", sheet)
    maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    logging.info(maxrow, ListOfThematics)
    ListOfThematicsCopy = ListOfThematics.copy()
    thm_archi_list = []
    finalThm = ""
    logging.info("ARCH -->> ", ARCH)
    logging.info("ListOfThematicsCopy -->> ", ListOfThematicsCopy)
    for i in ListOfThematicsCopy:
        flagR1 = 0
        flagR2 = 0
        # searchResults = EI.searchDataInCol(sheet, 7, i)

        sheet_value = sheet.used_range.value
        searchResults = EI.searchDataInColCache(sheet_value, 7, i)

        logging.info("\n\nsearchResults------->aa1:", searchResults)
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
                    logging.info("not applicable for BSI but its present in req\n")
                    return -1
            # for VSM Arch
            elif ARCH == "VSM":
                if (applicableR1 == "Y") and (applicableR2 == "Y"):
                    thm_archi_list.append('R1R2')
                    logging.info("NEA R1 R2 applicable")
                elif (applicableR1 == "Y") or (applicableR2 == "Y"):
                    if (applicableR1 == "Y"):
                        thm_archi_list.append('R1')
                        logging.info("NEA R1 applicable")
                    elif (applicableR2 == "Y"):
                        thm_archi_list.append('R2')
                        logging.info("NEA R2 applicable")
                else:
                    logging.info("not applicable for VSM but its present in req\n")
                    return -1
            else:
                logging.info("arch not found\n")
        else:
            logging.info("Thematique not found in referntial EC")
            return -1
    logging.info(f"thm_archi_list BF {thm_archi_list}")
    if thm_archi_list:
        # thm_archi_list = set(thm_archi_list)
        logging.info(f"thm_archi_list BF {thm_archi_list}")
        finalized_arch = finalizeArchi(thm_archi_list)
        finalThm = "".join(finalized_arch)
    return finalThm


def finalizeArchi(archiList):
    if ('R1R2' and 'R1' in archiList and 'R2' not in archiList) or (
            'R1' in archiList and 'R1R2' not in archiList and 'R2' not in archiList):
        return 'R1'
    elif 'R1R2' and 'R2' in archiList and 'R1' not in archiList:
        return 'R2'
    elif 'R1R2' in archiList and 'R1' not in archiList and 'R2' not in archiList:
        return 'R1R2'


def getDCI_Thematic_Archi(dciInfo):
    thm_applicability = ""
    dciThm = dciInfo['dciThematic']
    try:
        if dciThm != "" and dciThm is not None:
            # taking the LVM or LYQ thematic code from thematic content
            themCodes = re.findall("LVM_[0-9]{2}|LYQ_[0-9]{2}", str(dciThm))
            if themCodes:
                logging.info(f"** themCodes {themCodes} **")
                thm_applicability = findDCI_Applicability(themCodes)
            else:
                # taking the Nom_du_SO value from dci if no LVM or LYQ present in thematic
                logging.info("Both LYQ and LVM not preset....")
                logging.info(f"dciInfo['proj_param'] {dciInfo['proj_param']}.....")

                # taking architecture as all(R1R2R3)if ends with NEA
                if dciInfo['proj_param'].endswith('_NEA'):
                    logging.info(r"///////R1R2R3//////")
                    thm_applicability = "R1R2R3"
                else:
                    # finding the architecture
                    logging.info(f"Nom_du_SO {dciInfo['proj_param']}")
                    Nom_du_SO = QP.getDCIProjParam(dciInfo['proj_param'])
                    archi_values = [
                        {'R1R2': 'NEA_R1|NEA_R1_1', 'R1R2R3': 'NEA_R1_X', 'R2R3': 'NEA_R1_X', 'R1': 'NEA_R1',
                         'R2': 'NEA_R1_1', 'R3': 'NEA_R1_2', 'R1R3': 'NEA_R1|NEA_R1_2'}]
                    for i in archi_values:
                        for key in i:
                            if i[key] == Nom_du_SO.strip():
                                logging.info("archi_values", key)
                                thm_applicability = key
                                break
                    logging.info(f"thm_applicability {thm_applicability}.....")
    except Exception as exp:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        logging.info(f"\nError ................. {exp} line no. {exc_tb.tb_lineno} file name: {exp_fname}")

    return thm_applicability


# find and compare the thematic architecture of the test sheet with dci req thematic
# input single sheet object and the dci information
def check_interface_req_thematic_with_sheet(testSheet, dciInfo):
    result = {'archi_matched': 0, 'finalizedSheetArchi': "", 'dciThemArchi': "", 'thm_not_found': 0, 'is_exp': 0}
    # archi_matched = 0
    testSheetThematics = EI.getTestSheetThematics(testSheet)
    thm_line_applicablity = []
    try:
        if testSheetThematics:
            refEC = EI.openReferentialEC()
            for themLine in testSheetThematics:
                logging.info(f"themLine1234 --> {themLine}")
                logging.info("\n\n")
                # thm_lvm_lyq_val = [1 if re.search("LVM_[0-9]{2}|LYQ_[0-9]{2}", x) else 0 for x in themLine]
                applicableThemLine = ""
                if re.search("LVM_[0-9]{2}|LYQ_[0-9]{2}", themLine):
                    if (themLine.find('LVM_01') != -1 or themLine.find(
                            'LYQ_01') != -1) and (
                            themLine.find('LVM_02') == -1 and themLine.find(
                        'LVM_03') == -1 and themLine.find(
                        'LYQ_02') == -1):
                        applicableThemLine = 'R1'
                    elif (themLine.find('LVM_02') != -1 or themLine.find(
                            'LYQ_02') != -1) and (
                            themLine.find('LVM_01') == -1 and themLine.find(
                        'LVM_03') == -1 and themLine.find(
                        'LYQ_01') == -1):
                        applicableThemLine = 'R2'
                    elif themLine.find('LVM_03') != -1:
                        applicableThemLine = 'R3'
                    else:
                        logging.info(f"Contrary in thematic line...{testSheet} - {themLine}")
                        break
                    logging.info(f"applicableThemLine {applicableThemLine}")
                else:
                    KMS.showWindow(refEC.name.split('.')[0])
                    applicableThemLine = filterThemArch(themLine, refEC, getArch(ICF.FetchTaskName()))
                    logging.info(f"applicableThemLine else {applicableThemLine}")
                if applicableThemLine not in thm_line_applicablity and applicableThemLine != "":
                    thm_line_applicablity.append(applicableThemLine)
            # passing each line applicability in a list to find-out the exact architecture
            logging.info(f"\n\nthm_line_applicablity {thm_line_applicablity}\n\n")
            if thm_line_applicablity:
                finalizedSheetArchi = finalizeSheetArchi(thm_line_applicablity)
                logging.info(f"\n\nfinalizedSheetArchi {finalizedSheetArchi}\n\n")

            if finalizedSheetArchi != -1:
                dciThemArchi = getDCI_Thematic_Archi(dciInfo)
                logging.info(f"dciThemArchidciThemArchi {dciThemArchi}")
                if finalizedSheetArchi.strip() == dciThemArchi.strip():
                    logging.info("?><??")
                    result['archi_matched'] = 1
                elif finalizedSheetArchi == 'R1R2' and dciThemArchi == 'R1R2R3':
                    result['archi_matched'] = 1
                else:
                    logging.info("23@@@@@@23")
                    if (finalizedSheetArchi == 'R1' and dciThemArchi == 'R2') or (
                            finalizedSheetArchi == 'R2' and dciThemArchi == 'R1'):
                        result['archi_matched'] = 2
                    elif (finalizedSheetArchi == 'R1R2' and (dciThemArchi == 'R1' or dciThemArchi == 'R2')):
                        result['archi_matched'] = 1
                result['finalizedSheetArchi'] = finalizedSheetArchi
                result['dciThemArchi'] = dciThemArchi
            logging.info(f"archi_matched {result['archi_matched']}")
            refEC.close()
        else:
            logging.info(f"Thematic not found in {testSheet}")
            result['thm_not_found'] = 1

    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        logging.info(
            f"\nError in func check_interface_req_thematic_with_sheet: {ex} \nError line no. {exc_tb.tb_lineno} file name: {exp_fname}")
        result['is_exp'] = 1

    return result


def isfolderPathExists(folderPath):
    if os.path.exists(folderPath):
        return True
    return False


def isMacroFileExists(MacroFilePath):
    if os.path.isfile(MacroFilePath):
        return True
    return False


def validateSommarieSheetData(tpBook):
    displayInformation("\n..........Validating summary sheet data..........")
    sheet = tpBook.sheets['Sommaire']
    # sum_max_row = sheet.range('F' + str(sheet.cells.last_cell.row)).end('up').row
    sum_max_row = tpBook.sheets['Sommaire'].range('A' + str(tpBook.sheets['Sommaire'].cells.last_cell.row)).end('up').row
    sheet.api.Unprotect()
    if sum_max_row:
        REF_COL = 6
        VER_COL = 7
        refList = []
        verList = []
        rowx = 6
        while rowx < sum_max_row:
            if sheet.range(rowx, 1).value is not None:
                logging.info("HHere")
                rlo = sheet.range(rowx, 1).merge_area.row
                rhi = sheet.range(rowx, 1).merge_area.last_cell.row
                logging.info("I---->", rlo, rhi)

                greppedTypeIndex = []
                greppedVersionIndex = []

                # for i in range(rlo, rhi):
                for i in range(rlo, rhi + 1):
                    logging.info(i)
                    if sheet.range(i, 6).merge_cells is False:
                        if sheet.range(i, 6).value is not None and sheet.range(i, 6).value != '' and sheet.range(i, 6).value != '--' and sheet.range(i, 6).value !='-':
                            refList.append(sheet.range(i, 6).value)

                    elif sheet.range(i, 6).merge_cells is True:
                        rlo = sheet.range(i, 6).merge_area.row
                        rhi = sheet.range(i, 6).merge_area.last_cell.row
                        mergedCount = rhi - rlo
                        logging.info("Values of rlo,rhi,mergedcount--->", rlo, rhi, mergedCount)
                        sheet.range(i, 6).unmerge()
                        if sheet.range(i, 6).value is not None and sheet.range(i, 6).value != '' and sheet.range(i, 6).value != '--' and sheet.range(i, 6).value !='-':
                            refList.append(sheet.range(i, 6).value)

                            for k in range(1, mergedCount + 1):
                                logging.info("value of k--->", k)
                                if sheet.range(i + k, 6).value is None:
                                    logging.info("value of i+k--->", sheet.range(i + k, 6).value)
                                    sheet.range(i + k, 6).value = sheet.range(i, 6).value
                                    if sheet.range(i + k, 6).value is not None and sheet.range(i + k, 6).value != '' and sheet.range(i + k, 6) != '--' and sheet.range(i + k, 6) != '-':
                                        refList.append(sheet.range(i + k, 6).value)

                    if i not in greppedVersionIndex:
                        if sheet.range(i, 7).merge_cells is False:
                            if sheet.range(i, 7).value is not None:
                                version = re.search('[0-9]{1,2}', str(sheet.range(i, 7).value))
                                if sheet.range(i, 7).value is not None and sheet.range(i, 7).value != '' and sheet.range(i, 7).value != '--' and sheet.range(i, 7).value != '-':
                                    verList.append(sheet.range(i, 7).value)



                        elif sheet.range(i, 7).merge_cells is True:
                            rlo = sheet.range(i, 7).merge_area.row
                            rhi = sheet.range(i, 7).merge_area.last_cell.row
                            mergedCount = rhi - rlo
                            logging.info("Values of rlo,rhi,mergedcount--->", rlo, rhi, mergedCount)
                            sheet.range(i, 7).unmerge()
                            if sheet.range(i, 7).value is not None:
                                if str(sheet.range(i, 7).value) is not None and str(sheet.range(i, 7).value) != '' and str(sheet.range(i, 7).value) != '--' and str(sheet.range(i, 7).value) != '-':
                                    verList.append(str(sheet.range(i, 7).value))


                                for k in range(1, mergedCount + 1):
                                    logging.info("value of k--->", k)
                                    if sheet.range(i + k, 7).value is None:
                                        logging.info("value of i+k--->", sheet.range(i + k, 7).value)
                                        sheet.range(i + k, 7).value = sheet.range(i, 7).value
                                        if str(sheet.range(i + k, 7).value) is not None and str(sheet.range(i + k, 7).value) != '' and str(sheet.range(i + k, 7).value) != '--' and str(sheet.range(i + k, 7).value) != '-':
                                            greppedVersionIndex.append(i + k)
                                            verList.append(str(sheet.range(i + k, 7).value))
                    else:
                        logging.info("Already added the version to list", i)

            rowx = rowx + 1

    logging.info(f"verList {verList}")
    logging.info(f"refList {refList}")
    ref_ver = list(zip(refList, verList))
    logging.info(f"ref_ver {ref_ver}")
    logging.info("-------------------------------------------------------------------------------------")
    ref_ver = [item for item in ref_ver if item != (None, '')]
    ref_ver = [[str(item[0]).strip(), str(item[1])] for item in ref_ver]
    for ref, ver in ref_ver:
        ver = str(ver).strip()
        if ref is not None and ref != '' and ref != '--' and ref != '-':
            if not re.search(ref_num_pattern, str(ref).strip()) or not re.search('^(([0-9]{1,2})|([0-9]{1,2}\.[0-9]{1,2}))$',str(ver).strip('V|v').strip()):
                logging.info(f", -{ref}-, -{ver}-")
                displayInformation(f"Invalid reference number or version present in test plan -{ref}- -{ver}-")
                return False
    return True


def validateFilePaths():
    macro_flag, inp_flag, op_flag, download_flag = 0, 0, 0, 0
    if not isMacroFileExists(ICF.getTestPlanMacro()):
        errorPopupCb("Invalid Test Plan Macro Path!\nPlease give valid test plan macro path")
        displayInformation("Invalid Test Plan Macro Path!")
        displayInformation("Process Terminated")
        macro_flag = 1

    if not isfolderPathExists(ICF.getInputFolder()):
        errorPopupCb("Invalid Input Folder Path!\nPlease give valid input folder path")
        displayInformation("Invalid Input Folder Path")
        displayInformation("Process Terminated")
        inp_flag = 1

    if not isfolderPathExists(ICF.getOutputFiles()):
        errorPopupCb("Invalid Output Folder Path!\nPlease give valid output folder path")
        displayInformation("Invalid Output Folder Path")
        displayInformation("Process Terminated")
        op_flag = 1

    if not isfolderPathExists(ICF.getDownloadFolder()):
        errorPopupCb("Invalid Download Folder Path!\nPlease give valid download folder path")
        displayInformation("Invalid Download Folder Path")
        displayInformation("Process Terminated")
        download_flag = 1

    return macro_flag, inp_flag, op_flag, download_flag

logging.basicConfig(level=logging.CRITICAL)

def find_req_ver(requirement):
    reqName = ''
    reqVer = ''
    try:
        if requirement.find('(') != -1:
            reqName = requirement.split("(")[0]
            reqVer = requirement.split("(")[1].split(")")[0]
        else:
            reqName = requirement.split()[0]
            reqVer = requirement.split()[1] if len(requirement.split()) > 1 else ""
    except:
        logging.info("requirement not present")
        pass

    return reqName, reqVer


# These function is Used to get the duplicate Req FEPS to add in the impact sheet FEPS Column.
# Example Output--> {'REQ-0499102 (F)': ['FEPS_114040', 'FEPS_114406'], 'REQ-0499124(E)': ['FEPS_114040', 'FEPS_114406'], 'REQ-0499125 F': ['FEPS_114040', 'FEPS_114406']}
def findFEPSWithSameRequirements(rqIDs):
    # An empty dictionary to store the result
    result = {}
    # A dictionary to keep track of which FEPS contain each requirement
    req_to_feps = {}

    # Loop through the outer dictionary
    for key, value in rqIDs.items():
        # Loop through the inner dictionary (which may contain 'New Requirements' key)
        for subkey, subvalue in value.items():
            if (subkey == 'New Requirements ') or (subkey == 'Evolved Requirements') or (subkey == 'Interfaces'):
                # Loop through the list of requirements
                for req in subvalue:
                    # Add the current FEPS key to the list of FEPS containing this requirement
                    req_to_feps.setdefault(req, []).append(key)

    # Iterate through the requirements and find FEPS with the same requirements
    for req, feps_list in req_to_feps.items():
        if len(feps_list) > 1:
            # If more than one FEPS contains this requirement, add it to the result
            result[req] = feps_list

    return result


# these function is used to remove the duplicates reqs from the feps &
# it will take only the higher version of the duplicate reqs from the FEPS.
def removeDuplicateReqFeps(rqIDs):
    result = {}
    latest_versions = {}

    for key, value in rqIDs.items():
        for subkey, subvalue in value.items():
            logging.info("subkey000---->", subkey)
            logging.info("subvalue000---->", subvalue)
            if subkey in ['Interfaces', 'New Requirements ', 'Evolved Requirements']:
                for item in subvalue:
                    req_id, version = find_req_ver(item)
                    if req_id in latest_versions:
                        if version > latest_versions[req_id]:
                            latest_versions[req_id] = version
                            if req_id in result.setdefault(key, {}).setdefault(subkey, []):
                                result[key][subkey].remove(req_id)  # Remove previous version
                            result[key][subkey].append(item)  # Add current version
                    else:
                        latest_versions[req_id] = version
                        result.setdefault(key, {}).setdefault(subkey, []).append(item)
            else:
                result.setdefault(key, {}).setdefault(subkey, subvalue)

    logging.info("result7777----->", result)
    filtered_rq_ids = {}

    for key, value in result.items():
        if 'Input_Docs' in value.keys() and len(value.keys()) == 1:
            continue
        else:
            filtered_rq_ids[key] = value

    return filtered_rq_ids


# Used to get the reqs from the key elements.
def get_req_from_FEPS(rqIDs,feps_list,req,column6_value,result):
    # Check if the requirement exists in rqIDs
    if 'New Requirements ' in rqIDs[feps_list[0]].keys():
        if req in rqIDs[feps_list[0]]['New Requirements ']:
            result[req] = column6_value
            logging.info("I")

    if 'Evolved Requirements' in rqIDs[feps_list[0]].keys():
        if req in rqIDs[feps_list[0]]['Evolved Requirements']:
            result[req] = column6_value
            logging.info("II")

    if 'Interfaces' in rqIDs[feps_list[0]].keys():
        if req in rqIDs[feps_list[0]]['Interfaces']:
            result[req] = column6_value
            logging.info("III")
    return result


def addDuplicateFEPS(tpBook, maxrow, col, fepsNum, fepsForDuplicateReqs, rqIDs, Requirement, version):
    value, req = addFepsinImpact(fepsForDuplicateReqs, rqIDs, Requirement, version)
    logging.info("interface value, req --------->", value, req)
    Requirements = Requirement + " " + str(version)
    logging.info("fepsForDuplicateReqs if condition interface req-->", req)
    logging.info("Requirement-->", Requirements)
    if req == Requirements:
        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, col), value)

    elif req != Requirements:
        logging.info("@@@@@@@@@@")
        # match = re.match(r'(\S+)\s+(\d+)', Requirements)
        if Requirements:
            Requirements = Requirement + "(" +str(version) + ")"
            logging.info("reqq11--------->", Requirements)
            if req == Requirements:
                logging.info("1111111111111111111")
                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, col), value)
            else:
                logging.info("$$$$$$$$$$$$$")
                # if match:
                if Requirements:
                    Requirements = Requirement + " " + "("+str(version)+")"
                    logging.info("reqq22--------->", Requirements)
                    if req == Requirements:
                        logging.info("2222222221111111")
                        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, col), value)
                    else:
                        logging.info("%%%%%%%%%%%%%%")
                        # if match:
                        if Requirements:
                            Requirements = Requirement + "  " + "(" + str(version) + ")"
                            logging.info("reqq33--------->", Requirements)
                            if req == Requirements:
                                logging.info("3333333333")
                                EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, col), value)
                            else:
                                logging.info("^^^^^^^^^^^^^^^^^^")
                                # if match:
                                if Requirements:
                                    Requirements = Requirement + "  " +  str(version)
                                    logging.info("reqq44--------->", Requirements)
                                    if req == Requirements:
                                        logging.info("4444444444444444444")
                                        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, col), value)
                                    else:
                                        logging.info("ooppooppoopoopopo")
                                        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, col), fepsNum[1:])
    else:
        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, col), fepsNum[1:])


# Used to set the FEPS in the impact sheet
def addFepsinImpact(fepsForDuplicateReqs,rqIDs, reqname='', reqver=''):
    value = ''
    req = ''
    # Create a dictionary to store the result
    result = {}
    # Iterate through the requirements in fepsForDuplicateReqs
    for req, feps_list in fepsForDuplicateReqs.items():
        # Initialize the value for column 6 as a string with FEPS entries on separate lines
        column6_value = ', \n'.join(feps_list)

        for req, feps_list in fepsForDuplicateReqs.items():
            # Initialize the value for column 6 as a string with FEPS entries on separate lines
            # column6_value = '\n'.join(feps_list)
            column6_value = '\n'.join(fep.replace('FEPS_', '') for fep in feps_list)
            reqq = reqname + " " + str(reqver)
            logging.info("req--->",req)
            logging.info("reqq--->",reqq)
            if reqq == req:
                logging.info("hi")
                result = get_req_from_FEPS(rqIDs, feps_list, req, column6_value, result)
            elif reqq != req:
                logging.info("llllllllllll")
                if reqq:
                    logging.info("reqname & reqver elif reqq != req--------->",reqname, reqver)
                    reqq = reqname + "("+str(reqver)+")"
                    logging.info("reqq00--------->",reqq)
                    if reqq == req:
                        result = get_req_from_FEPS(rqIDs, feps_list, req, column6_value, result)
                    else:
                        logging.info("oooooooooooohhhhhhjjjjjjjjj")
                        if reqq:
                            logging.info("reqname & reqver elif reqq--------->", reqname, reqver)
                            reqq =  reqname + " " + "("+str(reqver)+")"
                            logging.info("reqq11--------->", reqq)
                            if reqq == req:
                                logging.info("hi")
                                result = get_req_from_FEPS(rqIDs, feps_list, req, column6_value, result)
                            else:
                                logging.info("oooooooooooohhhhhhjjjjjjjjj")
                                if reqq:
                                    logging.info("reqname & reqver elif reqq--------->", reqname, reqver)
                                    reqq = reqname + "  " + "(" + str(reqver) + ")"
                                    logging.info("reqq22--------->", reqq)
                                    if reqq == req:
                                        logging.info("hii")
                                        result = get_req_from_FEPS(rqIDs, feps_list, req, column6_value, result)
                                    else:
                                        logging.info("oooooooooooohhhhhhjjjjjjjjj")
                                        if reqq:
                                            logging.info("reqname & reqver elif reqq--------->", reqname, reqver)
                                            reqq = reqname + "  " + str(reqver)
                                            logging.info("reqq33--------->", reqq)
                                            if reqq == req:
                                                logging.info("hiii")
                                                result = get_req_from_FEPS(rqIDs, feps_list, req, column6_value, result)

        # logging.info the result
        for req, value in result.items():
            logging.info(f"Requirement: {req}, Column 6: \n{value}")
    return value, req


# if __name__ == "__main__":
def main1():
    global reqName
    print("\n######## App Test Tool Started ########\n")
    start_time = time.time()
    logging.info(f"start_time: {start_time}")
    Testplan_reference = []
    display_info = UpdateHMIInfoCb
    logging.info(display_info)
    display_info("App Test Tool Started")
    ICF.loadConfig()

    # checking by configuring the Chrome
    chrome_version_status = configChromeVersion()
    if chrome_version_status == -1:
        display_info(
            "\n\n###### Please check the chrome version and keep the latest chrome driver version inside ./config folder ######")
        return -1

    # Validate given paths in GUI
    macro_path, inp_path, op_path, download_path = validateFilePaths()
    if macro_path != 0 or inp_path != 0 or op_path != 0 or download_path != 0:
        return -1

    logging.info("ICF.getGatewayReq()  ", ICF.getGatewayReq())
    logging.info("ICF.getReqNameChange()", ICF.getReqNameChange())
    logging.info("ICF.getInterfaceReqNameChange()", ICF.getInterfaceReqNameChange())
    if ICF.getGatewayReq() is True:
        display_info('Gateway Requirement process started')
        gatewayReq()
        display_info('Process Completed...')
        return 1
    elif ICF.getReqNameChange() is True:
        display_info('Requirement name changing from GEN to REQ process started')
        reqNameChange()
        display_info('Process Completed....')
        return 1
    elif ICF.getGatewayReq() is True and ICF.getReqNameChange() is True:
        gatewayReq()
        reqNameChange()
        return 1

    if ICF.getInterfaceReqNameChange() is True:
        display_info('Interface Requirement name changing process started')
        logging.info("Interface req old to new")
        reqNameChanging()
        display_info('Process Completed....')
        return 1

    # check report start
    if ICF.getCheckReportStatus() is True:
        try:
            display_info("\n*** Running the check report ***")
            CRG.check_sf_report()
            display_info("\nGenerating check report porcess completed...")
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            logging.info(
                f"\nSomething went wrong in processing check report: {e} line no. {exc_tb.tb_lineno} file name: {exp_fname}********************")
            display_info('\nSomething went wrong in processing check report....')
        return 1
    # check report end

    reference = getReference()
    version = getVersion()
    macro = EI.getTestPlanAutomationMacro()
    # download reference from JSON (Analyse_de_entrant and QIA_Param)
    inputDoc = list(zip(reference, version))
    logging.info("inputDoc = ", inputDoc)
    logging.info("ICF.getAutoDownloadStatusAnalyzeDeEntrant() = ", ICF.getAutoDownloadStatusAnalyzeDeEntrant())
    if ICF.getAutoDownloadStatusAnalyzeDeEntrant():
        logging.info("Auto download in progress !!!")
        display_info("Automatic Download In Progress !!!")
        startDocumentDownload(inputDoc, True)
    else:
        display_info("Auto Download Option Disabled - Directly checking the Input documents")

    taskFoundFlag = 1
    with open('../Aptest_Tool_Report.txt', 'w') as f:
        f.writelines("Aptest result - \n \n \n")
    logging.info(ICF.getExcelPath())
    taskList = getTaskName()
    referentielList = getReferentiel()
    triList = getTrigram()
    for i in triList:
        if len(i):
            logging.info("trigram is present")
        else:
            logging.info("trigram not present")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines("\n\nTrigram is not mentioned in UserInput,please enter the Trigram")
    logging.info(taskList, len(taskList))
    logging.info(referentielList, len(referentielList))
    logging.info(triList, len(triList))
    logging.info("Date_Tested - ", date.today().strftime('%d/%m/%Y'))
    rqIDs = {}
    ipDocName = []
    fepsString = ""
    requirmentlist = {}
    TSList_Arr = []
    logging.info("ZIP list = ", taskList, referentielList, triList)
    for taskname, referentiel, trigram in zip(taskList, referentielList, triList):
        logging.info("loop started", taskname, referentiel, trigram)
        time.sleep(1)
        # Step2
        logging.info("opening analyse_des_entrant_sheet")
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nopening Analyse_de_entrant Sheet")
        analyseDeEntrant = EI.openAnalyseDeEntrant(taskname)
        logging.info("analyseDeEntrant = ", analyseDeEntrant)
        display_info("Analyse Sheet Available - Parsing the Input document Reference")
        if analyseDeEntrant != -1:
            time.sleep(1)
            KMS.showWindow("Analyse")
            time.sleep(1)
            # Step 3
            try:
                time.sleep(2)
                EI.activateSheet(analyseDeEntrant, taskname)
                # Step 4
                testPlanReference = getPTReference(analyseDeEntrant.sheets[taskname], (7, 5))
                logging.info("PT Refernce =", testPlanReference)
                Testplan_reference.append(testPlanReference)

                time.sleep(1)
                # Step 10
                inputDocumentReference = EI.parseIpDocId(analyseDeEntrant.sheets[taskname])
                if inputDocumentReference == -1:
                    displayErrorPopup()
                    return -108
                logging.info("Input Doc Reference =", inputDocumentReference)

                inputDocumentVersion = EI.parseIpDocId_ver(analyseDeEntrant.sheets[taskname])
                if inputDocumentVersion == -1:
                    displayErrorPopup()
                    return -108
                logging.info("Input Doc Version =", inputDocumentVersion)

                time.sleep(1)
                for i in EI.getFepsDoc(analyseDeEntrant.sheets[taskname]):
                    logging.info("i - ", i)
                    if (re.search(pattren_ref, i)) is not None:
                        b = re.search(pattren_ver, i).group()
                        i = i[:((i.index(b)) + 5)]
                        i = i.strip()
                        logging.info("i(1) - ", i)
                        if i not in ipDocName:
                            ipDocName.append(i)
                    else:
                        pass
                logging.info("Input Doc Name =", ipDocName)
                rqIDs.update(EI.getRequirementIDs(analyseDeEntrant.sheets[taskname]))
                requirmentlist = copy.deepcopy(rqIDs)
                logging.info("rqIDs =", rqIDs)
                fepsString = fepsString + EI.getFepsString(analyseDeEntrant.sheets[taskname])
                KMS.showWindow("Analyse")
                time.sleep(5)
                if analyseDeEntrant is not None:
                    analyseDeEntrant.close()

                # docToDownload.append() = inputDocumentReference + inputDocumentVersion

                docToDownload = inputDocumentReference.copy()
                docToDownload.append(testPlanReference)

                versionToDownload = inputDocumentVersion.copy()
                testPlan_tup = (testPlanReference, "")
                versionToDownload.append(testPlan_tup)

                logging.info("docToDownload = ", docToDownload)
                logging.info("versionToDownload = ", versionToDownload)

                # download reference from JSON (Analyse_de_entrant and QIA_Param)

                if ICF.getAutoDownloadStatusInputDocument():
                    startDocumentDownload(removeDuplicates(versionToDownload))
                    display_info("Valid reference number found, Downloading the Input documents")

                logging.info("Closed analyse_deentrant_sheet")
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nclosed Analyse_de_entrant Sheet")
            except Exception as e:
                if analyseDeEntrant is not None:
                    analyseDeEntrant.close()
                taskFoundFlag = 0
                logging.info("Exception in main(1) = ", e)
                # display_info("Exception in main(1) = " + str(e))
                # ctypes.windll.user32.MessageBoxW(0, "Entered task not found. \nPlease enter the correct task", "Aptest",
                # 1)
                # display_info(0, "Entered task not found. \nPlease enter the correct task" + "Aptest" + str(1))
                break
        else:
            taskFoundFlag = 0
            logging.info("Analyse sheet not present in input folder")
            display_info("Analyse sheet not present in input folder")
            # ctypes.windll.user32.MessageBoxW(0, "Analyse sheet not present in input folder", "Aptest", 1)
            break
        time.sleep(2)

    fepsForDuplicateReqs = findFEPSWithSameRequirements(rqIDs)
    logging.info("fepsForDuplicateReqs---->", fepsForDuplicateReqs)
    rqIDs = removeDuplicateReqFeps(rqIDs)
    logging.info("remove Duplicate Req from rqIDs =", rqIDs)


    if taskFoundFlag == 1:
        tpBook = EI.openTestPlan()
        summary_valid = validateSommarieSheetData(tpBook)
        if not summary_valid:
            errorPopupCb("Reference number or Version not valid in summary sheet.\nPlease correct it and Re-run the tool.")
            display_info(f"Invalid Reference or Version in Testplan summary sheet")
            display_info("Process Terminated")
            return -1

        logging.info("Opening Test Plan")
        display_info("Opening Test Plan")
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nOpening Test Plan Sheet")
        if tpBook != -1:
            KMS.showWindow(tpBook.name.split('.')[0])
            # Step 15
            taskArch = taskname.split("_")[0]
            if taskArch == "F":
                Arch = "BSI"
            else:
                Arch = "VSM"
            if Arch == "VSM":
                time.sleep(1)
                TPM.selectArch(macro)
            else:
                time.sleep(1)
                TPM.selectArch(macro)
                time.sleep(1)
                KMS.mouseClick()
            time.sleep(1)
            TPM.selectTpWritterProfile(macro)
            time.sleep(1)
            TPM.selectTPInit(macro)
            # Step 16
            fillSummary(tpBook, ipDocName, fepsString, referentielList[0], triList[0])
            # Step 17
            funcName = tpBook.sheets['Sommaire'].range(4, 3).value
            functionName = funcName.split("-")[1].strip()
            logging.info("Function Name = ", functionName)
            logging.info("Filling the Summary")
            display_info("Filling the Summary")

            # Step 18
            time.sleep(2)
            # TPM.selectToolbox()
            TPM.selectToolbox(macro)
            time.sleep(2)
            req_ver_sf = {'req': [], 'ver': [], 'sf_sheet': [], 'flow': [], 'req_comment': []}
            for feps in rqIDs:
                fepsNumber = (re.search("_[0-9]+", feps)).group()
                logging.info("fepsNumber-------------------->", fepsNumber)
                currDoc = getCurrentDocPath(feps, rqIDs)
                for reqname in dict(rqIDs[feps]):
                    if reqname == 'Evolved Requirements':
                        rqList = rqIDs[feps][reqname]
                        all_test_sheets = []
                        for req in rqList:
                            evolReqRes = evolReq(tpBook, fepsNumber, macro, Arch, feps, rqIDs, reqname, req_ver_sf, fepsForDuplicateReqs, req)
                            if evolReqRes == -2:
                                return -1
                            if evolReqRes:
                                all_test_sheets.append(evolReqRes)
                                time.sleep(2)
                        treate_backlog(all_test_sheets)
                        time.sleep(2)
                    elif reqname.strip() == "New Requirements":
                        logging.info("\n\n_______________Processing the new requirements_______________")
                        new_rqList = rqIDs[feps][reqname]
                        logging.info("new_rqList---fepsForDuplicateReqs-->", fepsForDuplicateReqs)
                        new_req_response = handle_new_requirements(tpBook, macro, new_rqList, fepsNumber, Arch, feps,rqIDs, reqname, testPlanReference, functionName, req_ver_sf,fepsForDuplicateReqs)
                        if new_req_response == -2:
                            return -1
                        with open('../Aptest_Tool_Report.txt', 'a') as f:
                            f.writelines("\n\nFilled the Impact sheet and test sheets to be treated")

                logging.info("Array inputdoc !!!!!!!!!!!!!!!!!", rqIDs[feps]['Input_Docs'])
                for inputdoc in rqIDs[feps]['Input_Docs']:
                    DCIdoc = []
                    count = 0
                    if inputdoc.find('DCI') != -1:
                        logging.info("DCI name in Analyse_De_Entrant ========>", inputdoc)
                        DCIdoc.append(inputdoc)
                        logging.info("DCIdoc = ", DCIdoc)
                        dci = EI.openDCIExcel(DCIdoc)
                        time.sleep(2)
                        # KMS.showWindow("DCI")
                        if dci is not None:
                            for reqname in dict(rqIDs[feps]):
                                rqList = rqIDs[feps][reqname]
                                if reqname == 'Interfaces':
                                    KMS.showWindow(tpBook.name.split('.')[0])
                                    time.sleep(1)
                                    logging.info("rqIDs interface oooo------->", rqIDs, fepsForDuplicateReqs)
                                    rowOfInterface = fillImpact(tpBook, rqList, fepsNumber, rqIDs, fepsForDuplicateReqs)
                                    time.sleep(2)
                                    TPM.selectTPImpact(macro)
                                    time.sleep(10)
                                    logging.info("rowOfInterface - ", rowOfInterface)
                                    filteredReqs = modifyImpact(tpBook, rqIDs[feps][reqname], rowOfInterface,fepsNumber, requirmentlist, testPlanReference,
                                                                functionName)
                                    logging.info(f"filteredReqs {filteredReqs}")
                                    for reqs in filteredReqs:
                                        logging.info(reqs, "|||||||||||reqs")
                                        isDeleted = 0
                                        EI.activateSheet(tpBook, tpBook.sheets['Impact'])
                                        time.sleep(1)
                                        dciInfo = EI.getDciInfo(dci, reqs)
                                        testPlanSheet = EI.findTestSheet(tpBook, dciInfo)

                                        # checking for delete req
                                        sheet_value = tpBook.sheets['Impact'].used_range.value
                                        logging.info(f"reqs {reqs}")
                                        reqName, reqVer = find_req_ver(reqs)
                                        logging.info(f"reqName, reqVer {reqName, reqVer}")
                                        searchReqResult = EI.searchDataInColCache(sheet_value, 1, str(reqName).strip())
                                        logging.info("\n\n ??????????????searchReqResult ??????? ", searchReqResult)
                                        req_ts = ''
                                        if searchReqResult['count'] > 0:
                                            for cellPos in searchReqResult['cellPositions']:
                                                r, c = cellPos
                                                if tpBook.sheets['Impact'].range(r, 5).value.upper().find(
                                                        "DELETED") != -1:
                                                    isDeleted = 1
                                                    req_ts = tpBook.sheets['Impact'].range(r, 4).value

                                        logging.info("isDeleted >> ", isDeleted)
                                        logging.info(f"testPlanSheet {testPlanSheet}")
                                        if isDeleted == 1:
                                            logging.info(f"req_ts {req_ts}")
                                            if req_ts != '' and req_ts is not None:
                                                logging.info("modifyDeleteReq1")
                                                modifyDeleteReq(tpBook, str(req_ts).strip(), macro, reqName, reqVer, fepsNumber, requirmentlist, testPlanReference, functionName, 'nover')
                                        else:
                                            if testPlanSheet != -1:
                                                for thm_ind, testSheet in enumerate(testPlanSheet):
                                                    logging.info(f"\n\n@@@@@@@@ Processing the sheet {testSheet} @@@@@@@@")
                                                    # check thematic applicability
                                                    archiResultValue = check_interface_req_thematic_with_sheet(
                                                        testSheet, dciInfo)
                                                    logging.info(f"'''archiResultValue''' {archiResultValue}")
                                                    if archiResultValue['archi_matched'] == 1:
                                                        if testSheet.range('C7').value == "VALIDEE":
                                                            EI.activateSheetObj(testSheet)
                                                            TPM.selectTestSheetModify(macro)
                                                            time.sleep(3)
                                                        EI.activateSheet(tpBook, testSheet)
                                                        fillSheetHistory(testSheet,
                                                                         "Added requirement " + reqs + ". No functional impact. \n")
                                                        # Step 24
                                                        # rqIDs = EI.getRequirementIDs(analyseDeEntrant.sheets[taskname])
                                                        logging.info("vvvvbbbbnnnnnmmmmm")
                                                        addRequirement(testSheet, reqs)

                                                        KMS.showWindow(tpBook.name.split('.')[0])
                                                        # Step 23
                                                        time.sleep(5)
                                                        if ICF.getBackLog() is True:
                                                            logging.info("Backlog Selected")
                                                            logging.info(f"testSheet {testSheet}")
                                                            display_info(f"========={testSheet.name}===========")
                                                            kpiDocList = getKPIDocPath(ICF.getInputFolder() + "/KPI")
                                                            rawReqs = testSheet.range('C4').value
                                                            logging.info(f"rawReqs {rawReqs}")
                                                            ts_reqs = [rawReqs]
                                                            refEC = EI.openReferentialEC()
                                                            currArch = getArch(ICF.FetchTaskName())
                                                            ts_req_list = removeInterfaceReq(ts_reqs)
                                                            status, combinedThemLines = getCombinedThematicLines(
                                                                kpiDocList, ts_req_list, refEC, currArch)
                                                            refEC.close()
                                                            logging.info("***Thematic Combinations = ***", status,
                                                                  combinedThemLines)
                                                            if status:
                                                                logging.info("Update the Thematic Combinations in the View")
                                                                display_info(
                                                                    f"=========Backlog Output - {testSheet.name}===========")
                                                                display_info(str(combinedThemLines))
                                                                display_info("====================")
                                                            else:
                                                                logging.info(
                                                                    f"Requirement Not available in KPI Sheet. Proceed Manually - {ts_reqs}")
                                                                display_info(

                                                                    "Requirement Not available in KPI Sheet. Make sure all the requirement in the test sheet is mentioned in KPI Sheet")
                                                        break
                                                    elif archiResultValue['archi_matched'] == 2:
                                                        logging.info(
                                                            f"\n\n##################### The sheet {testSheet} no need to be split in order to add this requirement... #####################")
                                                    else:
                                                        logging.info(
                                                            f"\n************* Thematic archi {archiResultValue['finalizedSheetArchi']} in sheet {testSheet} not matched with DCI thematic archi {archiResultValue['dciThemArchi']} *************")
                                                        if len(testPlanSheet) - 1 == thm_ind:
                                                            if archiResultValue['archi_matched'] == 0:
                                                                logging.info(
                                                                    f"\n******* Information: Thematic architecture not matched with any of the sheets {testPlanSheet} *******")
                                                        else:
                                                            continue
                                                if archiResultValue['archi_matched'] != 1:
                                                    if reqs.find("("):
                                                        splitRequirements = reqs.split("(")
                                                        Requirement = splitRequirements[0]
                                                    else:
                                                        splitRequirements = reqs.split(" ")
                                                        Requirement = splitRequirements[0]
                                                    logging.info(f"reqComment {reqs}")
                                                    reqPos = EI.searchDataInColCache(
                                                        tpBook.sheets['Impact'].used_range.value, 1,
                                                        str(Requirement).strip())
                                                    logging.info("reqPos->", reqPos)
                                                    if reqPos['count'] > 0:
                                                        commentCol = 5
                                                        cell_x, cell_y = reqPos['cellPositions'][0]
                                                        impactComment = "Interface requirement.\nRaised QIA of PT no.\nRaised QIA of Param Global no."
                                                        EI.setDataFromCell(tpBook.sheets['Impact'],
                                                                           (cell_x, commentCol), impactComment)

                                                # thm = parseThematics(dciInfo["thm"])
                                                # addThematics(testPlanSheet, thm)

                                            else:
                                                time.sleep(1)
                                                sheet_value = tpBook.sheets['Impact'].used_range.value
                                                if reqs.find("("):
                                                    # cellValue = EI.searchDataInExcel(tpBook.sheets['Impact'], (1000, 1),reqs.split("(")[0])
                                                    cellValue = EI.searchDataInExcelCache(sheet_value, (1000, 1), reqs.split("(")[0])
                                                else:
                                                    # cellValue = EI.searchDataInExcel(tpBook.sheets['Impact'], (1000, 1),reqs.split(" ")[0])
                                                    cellValue = EI.searchDataInExcelCache(sheet_value, (1000, 1), reqs.split("(")[0])

                                                logging.info("++", cellValue, reqs)
                                                for cellCoords in cellValue["cellPositions"]:
                                                    x, y = cellCoords
                                                    if x > 17:
                                                        EI.setDataFromCell(tpBook.sheets['Impact'], (x, y + 4),
                                                                           "Interface Requirement.\n" + reqs + " cannot be tested as signal not yet impacted in testplan.\nQIA PT - \nQIA Param Global- Point No- ")
                                                        break
                                                # ctypes.windll.user32.MessageBoxW(0, "Requirement " + reqs + " cannot be tested as signal not yet impacted in testplan", "Interface Requirement", 1)
                                                # display_info(
                                                #     "Requirement " + str(reqs) + " cannot be tested as signal not yet impacted in testplan" + "\n")
                                                with open('../Aptest_Tool_Report.txt', 'a') as f:
                                                    f.writelines(
                                                        "\n\nRequirement " + reqs + " cannot be tested as signal not yet impacted in testplan")
                                                time.sleep(2)
                                            logging.info("******************* Requirement Treated *******************")
                                            display_info("Requirement Treated")
                            dci.close()
                        else:
                            logging.info("DCI file name present in Analyse_de_entrant sheet is not found in input folder.")
                            display_info(
                                "DCI file name present in Analyse_de_entrant sheet is not found in input folder.")
                            # ctypes.windll.user32.MessageBoxW(0, "DCI name present in Analyse_de_entrant for - " + str(feps) + " is not found in input folder", "Interface Requirement", 1)
                            with open('../Aptest_Tool_Report.txt', 'a') as f:
                                f.writelines("\n\nDCI name present in Analyse_de_entrant for - " + str(
                                    feps) + " is not found in input folder")
                    else:
                        count = count + 1
                        logging.info("inputdoc !!!!!!!!!!!!!!!!!>", inputdoc, count)
                        # if count == len(rqIDs[feps]['Input_Docs']):
                        if count == len(rqIDs[feps]['Input_Docs']) - 1:
                            for reqname in dict(rqIDs[feps]):
                                if reqname == 'Interfaces':
                                    logging.info(
                                        "No DCI document name present in Analyse_de_entrant for - " + str(feps) + "\n")
                                    # display_info(
                                    #     "No DCI document name present in Analyse_de_entrant for - " + str(feps) + "\n")
                                    # ctypes.windll.user32.MessageBoxW(0, "No DCI document name present in Analyse_de_entrant for - " + str(feps), "Interface Requirement", 1)
                                    with open('../Aptest_Tool_Report.txt', 'a') as f:
                                        f.writelines(
                                            "\n\nNo DCI document name present in Analyse_de_entrant for - " + str(feps))
            # Step 25
            if (len(req_ver_sf['sf_sheet']) != 0):
                display_info("\nProcessing SF sheets....")
                sf_sheets, req_ver_sf = SFE.QIA_ssfiche_update(tpBook, Arch, taskname, trigram, display_info,
                                                               req_ver_sf)
                ssfiche = EI.openSousFiches()
                sheets = SFE.ss_fiche_update(sf_sheets, ssfiche)
                ssfiche.activate()
                # if (flag == 1):
                # ssfiche = EI.openSousFiches()
                # sheets=SFE.ss_fiche_update(sf_sheets, ssfiche)
                TPM.selectSynthUpdateFor_SF(macro)
                time.sleep(3)
                # ssfiche = EI.openSousFiches()
                ssfiche.save()
                ssfiche.close()
                tpBook.activate()
                TPM.synchronizeSubSheet(macro)
                time.sleep(2)
            EI.activateSheet(tpBook, tpBook.sheets['Impact'])
            TPM.selectTPImpact(macro)
            time.sleep(10)

            end_time = time.time()
            execution_time = end_time - start_time
            logging.info(f"end_time {end_time}")
            logging.info(f"execution_time {execution_time}")

            # QIA PT for interface process starts
            display_info("\nChecking for interface requirement QIA PT points....")
            execute_interface_requirement_treatment(ICF.FetchTaskName(), tpBook, rqIDs)

            # DCI signal process
            ReqData = {"testPlanReference": testPlanReference, "requirmentlist": requirmentlist,
                       "functionName": functionName, "tpBook": tpBook}
            logging.info("ReqDataNew -- ", ReqData)
            display_info("\nChecking for param global...")
            SearchandModifySignal(ReqData)

            EI.remove_Duplicates_C4(tpBook)
            time.sleep(10)
            TPM.selectSynthUpdate(macro)
            display_info("\nSynthesis update in progress...")
            # logging.info("All requirements treated")

            # check report
            # display_info("All requirements treated")
            end_time1 = time.time()
            execution_time1 = end_time - start_time
            logging.info(f"end_time after synthesis {end_time1}")
            logging.info(f"execution_time after synthesis {execution_time1}")

            time.sleep(100)
            logging.info('****************************************************************')
            logging.info(tpBook)
            logging.info('****************************************************************')
            # display_info("All treatable requirements treated.")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines("\n\nAll requirements treated.")
                display_info(f"\n\nAll requirements treated.")
            pa = ICF.getOutputFiles()
            logging.info('OutputPath = ', pa)
            pat1 = os.path.abspath(r"..\Output_Files")
            logging.info(pat1)
            if not os.path.exists(pat1):
                os.makedirs(pat1)
                logging.info('new output file is created')
                display_info('new output file is created')
            time.sleep(5)
            pat2 = os.path.abspath(r'..\Output_Files\Testplan.xlsm')
            logging.info(pat2)
            logging.info("---------------------------------")
            logging.info("Saving Testplan Sheet ", pat1 + '\\Testplan.xlsm')
            logging.info("---------------------------------")
            tpBook.save(pat2)
            logging.info('Testplan[sheet] is saved in output folder')
            display_info('Testplan[sheet] is saved in output folder')
            tpBook.close()
            print("\n\nTask Completed......")

            # ctypes.windll.user32.MessageBoxW(0,
            #                                  "All treatable requirements treated.\nFor non treated requirements please check the Aptest_tool_Report.txt file",
            #                                  "Aptest", 1)
            # +++++++++++++++++++++++++
        else:
            logging.info("Test plan not present in input folder", "Aptest", 1)
            display_info("Test plan not present in input folder")

            # ctypes.windll.user32.MessageBoxW(0, "Testplan not present in input folder", "Aptest", 1)


def handle_qia_requirements(tpBook, QIA_Data_List, file_with_signal):
    qia_pt_inp_doc_res = ""
    try:
        if QIA_Data_List:
            result = QPT.getDataAsDict(QIA_Data_List)
            if result["qiaDict"]:
                logging.info("len(result['qiaDict']['SPNFR']) ", len(result["qiaDict"]['SPNFR']), "----",
                      len(result["qiaDict"]['SPFNM']))

                if file_with_signal:
                    combine_qia_pt_data = QPT.combineQiaPtInpDocData(file_with_signal)
                    logging.info("combine_qia_pt_data123 ", combine_qia_pt_data)
                    if combine_qia_pt_data:
                        for qia_inp_doc_data in combine_qia_pt_data:
                            qia_pt_inp_doc_res = QPT.raiseQIA_InputDoc(tpBook, qia_inp_doc_data)
                            qia_inp_doc_data['type'] = qia_inp_doc_data['type']
                            qia_inp_doc_data['reqs'] = qia_inp_doc_data['req']
                            qia_pt_inp_doc_res['reqs'] = qia_inp_doc_data['req']
                            logging.info("\n\nqia_pt_inp_doc_res11 ", qia_pt_inp_doc_res)
                            if qia_pt_inp_doc_res['QIA_PT_input_doc_res'] == -2:
                                result["qiaDict"] = QPT.removeReq(result["qiaDict"], qia_inp_doc_data)

                qia_comment_list = QPT.getQiaRemarks(result["qiaDict"])
                logging.info("qia_comment_list1233 ", qia_comment_list)
                if qia_comment_list:
                    # and qia_pt_inp_doc_res['QIA_PT_input_doc_res'] == 1
                    final_clubbed_comment = QPT.club_all_qia_data(qia_comment_list)
                    logging.info(f"final_clubbed_comment --> {final_clubbed_comment}")
                    qia_comment_data = QPT.getQiaComment(QIA_Data_List, final_clubbed_comment, qia_pt_inp_doc_res)
                    logging.info(f"qia_comment_data --> {qia_comment_data}")
                    qiaResponse = QPT.raiseQIA_PT(tpBook, final_clubbed_comment, qia_comment_data)
                    if qiaResponse['status'] == 1 and qiaResponse['reqData'] != "" and qiaResponse[
                        'reqData'] is not None:
                        logging.info("qiaResponse qia--> ", qiaResponse)
                        for req, slno in qiaResponse['reqData']:
                            if slno != "" and slno is not None:
                                split_req = req.split('\n')
                                split_req = req.split(',')
                                logging.info("split_req ", split_req)
                                for ReqId in split_req:
                                    logging.info(f"ReqId {ReqId}")
                                    # getReqCoords = EI.searchDataInCol(tpBook.sheets['Impact'], 1, ReqId.strip())

                                    sheet_value = tpBook.sheets['Impact'].used_range.value
                                    getReqCoords = EI.searchDataInColCache(sheet_value, 1, ReqId.strip())

                                    logging.info("getReqCoords >> ", getReqCoords)
                                    if getReqCoords['count'] > 0:
                                        tpBook.sheets['Impact'].activate()
                                        commentCol = 5
                                        for cellPos in getReqCoords['cellPositions']:
                                            row, col = cellPos
                                            tpBook.sheets['Impact'].activate()
                                            existCmt = tpBook.sheets['Impact'].range(row, commentCol).value
                                            newcomment = f"{existCmt} QIA PT raised for this requirement\nNo. {int(slno)}."
                                            EI.setDataFromCell(tpBook.sheets['Impact'], (row, commentCol),
                                                               newcomment)
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        logging.info(
            f"\n******************** Something went wrong in handling the QIA PT Error: {e} line no. {exc_tb.tb_lineno} file name: {exp_fname}********************")


def SearchandModifySignal(ReqData):
    logging.info("Interface Data- ", ReqData["requirmentlist"])
    logging.info("TP Ref_num- ", ReqData["testPlanReference"])
    QIA_Data_List = []
    file_with_signal = []
    isInterface = 0
    for feps in ReqData['requirmentlist']:
        if "Interfaces" in ReqData["requirmentlist"][feps].keys():
            isInterface = 1
            break
    logging.info(f"isInterface {isInterface}")
    if isInterface == 1:
        downloading_docs = QPT.downloadDocsForQIA_PT(ReqData['tpBook'])
        logging.info(f"downloading_docs {downloading_docs}")
        for feps in ReqData["requirmentlist"]:
            logging.info("DCI FEPS-", feps)
            for inputdoc in ReqData["requirmentlist"][feps]['Input_Docs']:
                dci_ref_num = ""
                dci_ver = ""
                DCIdoc = []
                count = 0
                if inputdoc.find('DCI') != -1:
                    logging.info("DCI Input Document- ", inputdoc)
                    DCIdoc.append(inputdoc)
                    logging.info(DCIdoc)
                    dci = EI.openDCIExcel(DCIdoc)
                    time.sleep(2)
                    if dci is not None:
                        if re.search(pattren_ref, inputdoc):
                            dci_ref = re.findall(pattren_ref, inputdoc)
                            dci_ref_num = "".join(dci_ref[0])
                            logging.info("refnm1 ", "".join(dci_ref[0]))
                        # elif re.search(pattern_ref_BSI_1, inputdoc):
                        #     dci_ref = re.findall(pattern_ref_BSI_1, inputdoc)
                        #     dci_ref_num = "".join(dci_ref[0])
                        #     logging.info("refnm2 ", "".join(dci_ref[0]))
                        if re.search(pattren_ver, inputdoc):
                            dci_version = re.findall(pattren_ver, inputdoc)
                            dci_ver = "".join(dci_version[0])
                            logging.info("ver ", "".join(dci_version[0]))

                        for reqName in dict(ReqData["requirmentlist"][feps]):
                            if reqName == 'Interfaces':
                                rqList = ReqData["requirmentlist"][feps][reqName]
                                logging.info("Requirement List- ", rqList)
                                for reqId in rqList:
                                    logging.info("reqId - ", reqId)
                                    splitRequirements = reqId.split("(")
                                    reqIdrep = splitRequirements[0]
                                    logging.info("Requirement ID- ", reqIdrep)
                                    if len(reqId.split("(")) > 1:
                                        splitVersion = splitRequirements[1]
                                        vers = splitVersion.split(")")
                                        reqversion = int(vers[0])
                                        logging.info("reqversion- ", reqversion)
                                    else:
                                        splitVersion = ""
                                        vers = ""
                                        reqversion = ""
                                        logging.info("reqversion- ", reqversion)
                                    testsheetList = getReqDatafromImpact(ReqData['tpBook'], 18, reqIdrep)
                                    logging.info("TestSheets- ", testsheetList)
                                    # if len(testsheetList['testSheetList']) != 0:
                                    dciInfo = EI.getDciInfo(dci, reqId)
                                    logging.info("DCI Info Response- ", dciInfo)
                                    if dciInfo['dciSignal'] is not None and dciInfo['dciSignal'] != '':
                                        logging.info("DCI-Signal- ", dciInfo['dciSignal'])
                                        logging.info("DCI-Frame- ", dciInfo["framename"])
                                        # if not paramBook:
                                        paramBook = QP.openParamGlobalSheet()
                                        dcisignall = dciInfo['dciSignal']
                                        time.sleep(2)
                                        dciInfo['dci_ref_num'] = dci_ref_num
                                        dciInfo['dci_ver'] = dci_ver
                                        if paramBook is not None and paramBook != "" and paramBook != -1:
                                            pgData = {"dciInfo": dciInfo, "reqIdrep": reqIdrep,
                                                      "TPRef_num": ReqData["testPlanReference"],
                                                      "FuncName": ReqData["functionName"], "reqversion": reqversion}
                                            # ParamData = EI.searchDciSignal_and_GetPGResponse(paramBook, dciInfo, reqIdrep,ReqData["testPlanReference"],ReqData["functionName"],reqversion)
                                            ParamData = QP.GetParamGlobalData(paramBook, pgData)
                                            logging.info("Param Global Response ", ParamData)
                                            if ParamData:
                                                # for index, fmsg in enumerate(ParamData):
                                                for fmsg in ParamData:
                                                    fluxmsg = fmsg["FLUX_MESSAGERIE_NEA"].split("/")
                                                    newparamsignal = fmsg["PARAM_SIGNAL"]
                                                    logging.info("Param Frame N/w- ", fluxmsg[0])
                                                    if fluxmsg[0] == dciInfo['network']:
                                                        logging.info(
                                                            "\n-------------------Param frame and DCI network matched------------------\n")
                                                        if len(testsheetList['testSheetList']) > 0:
                                                            for tslist in testsheetList['testSheetList']:
                                                                logging.info("Test Sheet name- ", tslist)
                                                                data_dic = QP.searchSignalandReplaceinTS(tslist, reqIdrep, dciInfo, 1,
                                                                                              fmsg)
                                                                logging.info('data_dic--->', data_dic)
                                                                if data_dic != 1:
                                                                    IRAS.dci_value_update(ParamData, dciInfo, data_dic)
                                                                    dci = EI.openDCIExcel(DCIdoc)
                                                                else:
                                                                    logging.info('no new name found from QIA process |***| matched |***| no need to do any thing |***| ')

                                                    else:
                                                        logging.info(
                                                            "\n-------------------Param frame and DCI network not matched------------------\n")
                                                        if len(testsheetList['testSheetList']) > 0:
                                                            for tslist in testsheetList['testSheetList']:
                                                                logging.info("Test Sheet name- ", tslist)
                                                                data_dic = QP.searchSignalandReplaceinTS(tslist, reqIdrep, dciInfo,
                                                                                              1, fmsg)
                                                                logging.info('data_dic--->', data_dic)
                                                                if data_dic != 1:
                                                                    IRAS.dci_value_update(ParamData, dciInfo, data_dic)
                                                                    dci = EI.openDCIExcel(DCIdoc)
                                                                else:
                                                                    logging.info('no new name found from QIA process |***| matched |***| no need to do any thing |***| ')
                                            else:
                                                logging.info("\n ----------Data for signal [" + dciInfo[
                                                    'dciSignal'] + "]+ not present----------\n")
                                        else:
                                            logging.info(
                                                "\n ----------Please check Param Global File exist or not----------\n")

                                    else:
                                        logging.info(
                                            "\n ----------Signal not present in DCI File for requirement [" + reqId + "]----------\n")

                                    # else:
                                    #     logging.info(
                                    #         "\n ----------Test[Sheet] not present requirement [" + reqId + "]----------\n")
                                    # For raising QIA PT
                                    # signalTestlanSheet = EI.findTestSheet(ReqData['tpBook'], dciInfo)

                                    testPlanSheet = EI.findTestSheet(ReqData['tpBook'], dciInfo)
                                    if testPlanSheet == -1:
                                        dciInfo['feps'] = feps
                                        try:
                                            if downloading_docs != -1:
                                                qia_data_res = QPT.processInerfaceReqSignal(ReqData['tpBook'], dciInfo)
                                                # logging.info("\n\nQIAPT_Response ", qia_data_res['resultResponse'])
                                                logging.info("\nQIA_Data123 ", qia_data_res)
                                                if qia_data_res != '' and qia_data_res is not None:
                                                    if qia_data_res['resultResponse'] != -1:
                                                        QIA_Data_List.append(qia_data_res['reqData'])
                                                    if qia_data_res['signal_exist_file']:
                                                        file_with_signal.append(qia_data_res['signal_exist_file'])
                                                    logging.info("file_with_signal3 ", file_with_signal)
                                                    logging.info("QIA_Data_List ", QIA_Data_List)
                                        except Exception as ex:
                                            exc_type, exc_obj, exc_tb = sys.exc_info()
                                            logging.info(f"\nError occurred QIA PT: {ex} line: {exc_tb.tb_lineno}")
                        dci.close()
                    else:
                        logging.info("\n ----------Something went wrong in opening the DCI file----------\n")
        handle_qia_requirements(ReqData['tpBook'], QIA_Data_List, file_with_signal)
    return 1


def create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps, dtc='', flow='', frame='', ckt='',defectCodeDNFKPI='', reqData=''):
    tpBook.activate()
    TPM.selectTestSheetAdd(macro)
    time.sleep(5)
    logging.info("tpBook.sheets.active ", tpBook.sheets.active)
    created_FT = tpBook.sheets.active
    filling_FT_res = NRH.fill_FT(tpBook, created_FT.name, new_req, macro, rqIDs, feps, dtc, flow, frame, ckt,
                                 defectCodeDNFKPI, reqData)
    if filling_FT_res != -1:
        try:
            fillHistoryAndTrigram(created_FT, "Created new sheet")
            reqq, verr = NRH.getReqVer(new_req)
            logging.info(f"reqq {reqq, verr}")
            # new_req_pos = EI.searchDataInCol(tpBook.sheets['Impact'], 1, reqq)
            sheet_value = tpBook.sheets['Impact'].used_range.value
            new_req_pos = EI.searchDataInColCache(sheet_value, 1, reqq)

            logging.info("new_req_pos ", new_req_pos)
            if new_req_pos['count'] > 0:
                ts_col = 4
                comment_col = 5
                for cell_pos in new_req_pos['cellPositions']:
                    x, y = cell_pos
                    logging.info(f"cell_pos {cell_pos}")
                    EI.setDataFromCell(tpBook.sheets['Impact'], (x, comment_col), "New Requirement.")
        except Exception as e:
            logging.info(f"\nError in filling history for new requirement.. {e}")


# def handle_new_requirements(tpBook, macro, new_rqList, fepsNumber, Arch, feps, rqIDs, reqname, testPlanReference, functionName):
def handle_new_requirements(tpBook, macro, new_rqList, fepsNumber, Arch, feps, rqIDs, reqname, testPlanReference, functionName, req_ver_sf,fepsForDuplicateReqs):
    all_test_sheets = []
    logging.info("handle_new_requirements--fepsForDuplicateReqs",fepsForDuplicateReqs)
    display_info = UpdateHMIInfoCb
    for new_req in new_rqList:
        currDoc = getCurrentDocPath(feps, rqIDs)
        flag = TA.checkReq(currDoc, new_req, Arch)
        if flag == -3:
            display_info("\n\nAnalysis de entrance input file not present in Input_Files folder")
            errorPopupCb(
                f'Analysis de entrance input file name present under the {feps} not Match in Input_Files folder. Please change the document name and re run the tool.')
            print(f'Analysis de entrance input file name present under the {feps} not Match in Input_Files folder. Please change the document name and re run the tool.')
            return -2
        isDTCReq = 0
        if new_req.lower().find('requirement') == -1:
            logging.info(f"\n************** requirement {new_req} **************")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines(
                    f"\n\n------------------ New Requirement {new_req} ---------------------")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines("\nFilling Impact Sheet")
            impact_new_req_res = NRH.fillImpactNewReq(tpBook, new_req, fepsNumber, flag, Arch, rqIDs, fepsForDuplicateReqs)
            logging.info("impact_new_req_res --> ", impact_new_req_res)
            if flag == 1:
                trigram = ICF.gettrigram()
                logging.info('Trigram--->', trigram)
                # Reference_of_SSD = '02017_19_02187'
                val_map = {"B": testPlanReference, "C": functionName, "D": "Creation",
                           "E": "New Calibration for NEA", "F": "NEA", "G": "--", "H": "",
                           "I": "", "J": "--", "K": functionName, "L": "*********",
                           "M": "", "N": "--", "O": "--", "P": trigram,
                           "Q": date_time.strftime("%m/%d/%y"), "R": "Open", "U": str(trigram + ' ' + (
                        date_time.strftime(
                            "%m/%d/%y")) + " : " + "Information can be found on SSD_PArameter_board Ref : ")}
                logging.info('val_map------>', val_map)
                # input_docum = rqIDs[feps]['Input_Docs']
                Req_name = impact_new_req_res["new_reqName"]
                logging.info('Req_name----->', Req_name)
                Req_vers = impact_new_req_res["new_reqVer"]
                logging.info('Req_vers---->', Req_vers)
                QCU.UpdateQiaParamGlobal(tpBook, val_map, Req_name, Req_vers, trigram)
                logging.info('completed QIA of Calibration^^^!!!!!^^^!!!!!!^^^')

                k = rqIDs[feps]['Input_Docs']
                for i in k:
                    # Here i is the input document details
                    if i.find('EEAD') != -1:
                        EEAD = EI.findInputFiles()[12]
                        path = ICF.getInputFolder() + "\\" + EEAD
                        new_reqName, new_reqVer = NRH.getReqVer(new_req)
                        reqData = WDI.getReqContent(path, new_reqName, new_reqVer)
                        logging.info(f"reqData 0--> {reqData}")
                        try:
                            celltext = reqData['celltext']
                            dtc_pattern = r"record a DTC"
                            match = re.search(dtc_pattern, celltext)
                            if match:
                                isDTCReq = 1
                                dtc = match.group()
                                logging.info("Extracted DTC:", dtc)
                                QPD.creatQIAParamDTC(testPlanReference, functionName, new_req, trigram)
                        except:
                            logging.info("requirement not contains record a DTC")
                            pass
                        try:
                            QIACNF.creatNewFrameQIA(testPlanReference, functionName, new_req, trigram,
                                                    i)
                        except:
                            logging.info("requirement not contains frame")
                            pass
                # Process the QIA Param for DID
                qia_result = QP.handle_qia_of_did(rqIDs[feps]['Input_Docs'], new_req, testPlanReference, functionName)
                logging.info(f"qia_result   >> {qia_result}")
                if qia_result != -1 and qia_result != -2 and qia_result != 2 and qia_result is not None:
                    logging.info(f"QIA Response for requirement {new_req}")
                    logging.info(f"Status Message  {qia_result['status_msg']}")

                if impact_new_req_res['status'] == 2:
                    logging.info(".......Not a new requirement, taking as evolved.......")
                    testSheets = evolReq(tpBook, fepsNumber, macro, Arch, feps, rqIDs, reqname, req_ver_sf, fepsForDuplicateReqs,new_req)
                    all_test_sheets.append(testSheets)
                    time.sleep(2)
                else:
                    logging.info(".......Not present in FT, creating the new FT......")
                    # if isDTCReq == 1:
                    EEAD = EI.findInputFiles()[12]
                    if EEAD is not None and EEAD != "":
                        path = ICF.getInputFolder() + "\\" + EEAD
                        logging.info("path------>", path)
                        new_reqName, new_reqVer = NRH.getReqVer(new_req)
                        reqData = WDI.getReqContent(path, new_reqName, new_reqVer)
                        logging.info(f"reqData 0--> {reqData}")
                        flows, frame, dtc = DFW.getFlow(new_req, reqData)
                        circuitArr, flowArr_E_Col, WRArr_I_Col, circuit = DFW.getCircuits(new_req, flows, reqData)
                        defectCodeDNFKPI = ""
                        if dtc == "record a DTC":
                            if flows:
                                defectCodeDNFKPI = DFW.getDefectCode(flows, reqData)
                            try:
                                logging.info("dtc_pattern--->", dtc)
                                try:
                                    if dtc == 'record a DTC':
                                        if circuit:
                                            for flow in flows:
                                                for ckt in circuit:
                                                    logging.info("flow-->", flow)
                                                    logging.info("ckt-->", ckt)
                                                    create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps, dtc,
                                                                                flow,
                                                                                frame,
                                                                                ckt, defectCodeDNFKPI, reqData)
                                        elif flows:
                                            for flow in flows:
                                                logging.info("flow1")
                                                ckt = ""
                                                create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps, dtc, flow,frame,ckt, defectCodeDNFKPI, reqData)
                                    else:
                                        logging.info("'record a DTC' text is not present in the requirement table.")
                                except:
                                    logging.info("flow not present1")
                            except Exception as ex:
                                exc_type, exc_obj, exc_tb = sys.exc_info()
                                logging.info(f"dtc_pattern or flow is not present{ex}{exc_tb.tb_lineno}")
                                pass
                            try:
                                try:
                                    frames, dtc, flowframe = DFW.getFrame(new_req, reqData)
                                    if dtc == 'record a DTC':
                                        flow = ""
                                        ckt = ""
                                        if flowframe and (flows is None or flows == ""):
                                            for flowframe in flowframe:
                                                logging.info("frame2-->", flowframe)
                                                create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps, dtc, flow,
                                                                            frame, ckt, defectCodeDNFKPI, reqData)
                                        elif frames and (flows is None or flows == ""):
                                            for frame in frames:
                                                logging.info("frame3-->", frame)
                                                create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps, dtc, flow,
                                                                            frame,
                                                                            ckt, defectCodeDNFKPI, reqData)
                                    else:
                                        logging.info("'record a DTC' text is not present in the requirement table.")
                                except Exception as ex:
                                    exc_type, exc_obj, exc_tb = sys.exc_info()
                                    logging.info(f"frame not present1   {ex}  {exc_tb.tb_lineno}")
                            except Exception as ex:
                                exc_type, exc_obj, exc_tb = sys.exc_info()
                                logging.info(f"dtc_pattern or frame is not present{ex}{exc_tb.tb_lineno}")
                        else:
                            create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps)
                            logging.info("It is not a DTC Requirement")
                    elif new_req:
                        # -------------------------------------------------------- Diagnostics --------------------------------------------------------------- #

                        logging.info('2222')
                        new_reqName, new_reqVer = NRH.getReqVer(new_req)
                        isDiag = DRH.extract_diag_req(new_reqName, new_reqVer, rqIDs, feps)
                        logging.info(f"isDiag {isDiag}")
                        if isDiag == -1:
                            logging.info("Not a diag requiremnt.....")
                            create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps)
                        else:
                            logging.info('###### COMPLETED DIAGNOSTICS REQUIREMENT ######')

                        # -------------------------------------------------------- Diagnostics --------------------------------------------------------------- #
                    else:
                        # flow = ''
                        # frame = ''
                        # ckt = ''
                        # dtc = ''
                        # defectCodeDNFKPI = ''
                        # reqData = ''
                        create_new_ft_and_fill_data(tpBook, new_req, macro, rqIDs, feps)
                        logging.info("'EEAD' file is not present")
            else:
                logging.info(f'"{new_req}" req is not added.')
                display_info(
                    f'Due to the thematics lines of "{new_req}" not being applicable for {Arch}, Proceed Manually.')
    treate_backlog(all_test_sheets)
    time.sleep(2)


def displayInformation(msg):
    display_info = UpdateHMIInfoCb
    return display_info(msg)
