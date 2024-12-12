import xlwings as xw
import re
import time
import InputConfigParser as ICP
import ExcelInterface as EI
import web_interface as WI
import datetime
import WordDocInterface as WDI
from datetime import datetime as dt
import BusinessLogic as BL
# from tkinter import *
import tkinter as tk
from tkinter import messagebox
import os
import InputConfigParser as ICF
from datetime import date
import sys
import xlsxwriter
import logging
#


# dciInfo = {
#     "dciSignal": "REGUL_BASR",
#     "pc": "C"
# }
# dciSignal = "REGUL_BASR"
# dciPC = "C"

UpdateHMIInfoCb = None

# SNP - signal not present in the input document
# SPNFR  - signal present no functional requirement in input document
# SPFM  - signal and functional requirement present in input document
# SPFNM  - signal present but functional requirement p/c not matched in input document


# if not os.path.exists("..\QIA_PT_LOG"):
#     os.makedirs("..\QIA_PT_LOG")
#     QIA_PT_LogFile = dt.now().strftime('../QIA_PT_LOG/QIA_PT_Log_%d_%m_%Y_%H_%M.log')
# else:
#     QIA_PT_LogFile = dt.now().strftime('../QIA_PT_LOG/QIA_PT_Log_%d_%m_%Y_%H_%M.log')


#####################################
# I/P: message which is going to display in GUI information box
# Desc: display the information in the GUI
# O/P: message
#####################################
def showSignalStatusConfirmation(func):
    global UpdateHMIInfoCb
    UpdateHMIInfoCb = func


#####################################
# I/P: excel sheet object which want to save and filetype to diferentiate the file name to save
# Desc: save the Excel sheet in output folder
# O/P: -
#####################################
def saveQIA_PT(excelBook, fileType):
    output_dir = os.path.abspath(r"..\Output_Files")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    if not os.path.exists(r"..\Output_Files\QIA_PT_Sheet"):
        os.makedirs(r"..\Output_Files\QIA_PT_Sheet")
        logging.info('new output file is created')
    time.sleep(5)
    if fileType == 1:
        saveFileNameAs = "QIA_PT.xlsx"
        savingPath = os.path.abspath(r'..\Output_Files\QIA_PT_Sheet\QIA_PT.xlsm')
    else:
        saveFileNameAs = "QIA_PT_InputDoc.xlsx"
        savingPath = os.path.abspath(r'..\Output_Files\QIA_PT_Sheet\QIA_PT_InputDoc.xlsm')
    logging.info(savingPath)
    logging.info("---------------------------------")
    logging.info("Saving QIA PT Sheet ", output_dir + '\\QIA_PT_Sheet\\' + saveFileNameAs)
    logging.info("---------------------------------")
    excelBook.save(savingPath)
    logging.info('QIA PT sheet is saved in output folder')
    # UpdateHMIInfoCb('\n>>QIA PT sheet is saved in output folder Output_Files\ReqNameChange')


#####################################
# I/P: file name to write the information
# Desc: add the given message in the text file
# O/P: -
#####################################
def writeLog(logFile, type, msg):
    logging.info("logFile ", logFile)
    with open(logFile, type) as f:
        f.writelines(f"\n{msg}")


# writeLog(QIA_PT_LogFile, 'w', "------------------- QIA PT Process Report -------------------\n \n")


#####################################
# I/P: -
# Desc: open the testplan book from the input folder
# O/P: test plan book object
#####################################
def openTestPlanSheet():
    try:
        tpBook = EI.openExcel(ICP.getInputFolder() + "\\" + EI.findInputFiles()[1])
        logging.info(f"\n Opening the Test Plan - {EI.findInputFiles()[1]}")
        return tpBook
    except Exception as exp:
        logging.info(f"\nException in opening test plan :- {exp}")
        return -1


#####################################
# I/P: testplan book abject
# Desc: finding the dci signal is present in the test sheet or not
# O/P: - test sheets as list
#####################################
def findSignalInTestSheet(tpBook, dciInfo):
    logging.info("\n ------- Finding the Signal in Test Plan Sheets -------")
    signalTestSheets = []
    if tpBook is not None:
        for tpSheet in tpBook.sheets:
            tpSheet_value = tpSheet.used_range.value
            try:
                if (tpSheet.name.find("VSM") != -1 or tpSheet.name.find("BSI") != -1) and tpSheet.name.find(
                        '_0000') == -1:
                    logging.info("dciInfo['dciSignal'].strip  ", dciInfo['dciSignal'].strip("$"))
                    # searchSignalResult = EI.searchDataInExcel(tpSheet, "", dciInfo['dciSignal'].strip("$"))
                    searchSignalResult = EI.searchDataInExcelCache(tpSheet_value, "", dciInfo['dciSignal'].strip("$"))
                    if searchSignalResult['count'] > 0:
                        logging.info("searchSignalResult -> ", searchSignalResult)
                        signalTestSheets.append(tpSheet.name)
            except Exception as exp:
                logging.info(f"\nException in searching signal in test sheets :- {exp}")
                # writeLog(QIA_PT_LogFile, 'a', f"\nException in searching signal in test sheets :- {exp}")

        logging.info("signalTestSheets -> ", signalTestSheets)
    return signalTestSheets


#####################################
# I/P: dci sheet information for the respective requirement
# Desc: Comparing the input given by user with value in pc column in DCI file
# O/P: returns boolean
#####################################
def compareSignalStatus(dciInfo, userInput):
    logging.info("dciInfo['pc'] ", dciInfo['pc'])
    if dciInfo['pc'] == userInput:
        return True
    return False


#####################################
# I/P: reference number to download the document
# Desc: download the file using the given reference number from docinfo portal
# O/P: return numeric either 1 or -1
#####################################
def downloadQIA_PT(QIA_PT_Reference):
    try:
        docRefVer = []
        docRefVer.append((QIA_PT_Reference, ""))
        logging.info("docRefVer >> ", docRefVer)
        WI.startDocumentDownload(docRefVer)
        return 1
    except Exception as exp:
        logging.info(f"Error in downloading the QIA PT file {exp}")
        return -1


#####################################
# I/P: testplan book object
# Desc: get the reference number of QIA PT from summary tab
# O/P: return reference number if present else return -1
#####################################
def getQIA_PT_Reference(tpBook):
    try:
        sheet = tpBook.sheets['Sommaire']
        if sheet:
            qiaPT_Ref = sheet.range(1, 6).value
            return qiaPT_Ref
    except Exception as exp:
        logging.info(f"Error in finding summary sheet {exp}")
        return -1


#####################################
# I/P: qia_pt refernce number
# Desc: checking the QAI PT Excel file is present under the input folder or not using the reference number
# not present then it will download the file
# O/P: return QIA PT file object
#####################################
def getQIA_PT_File(QIA_PT_Reference):
    QIA_PT = ''
    if os.path.isdir(ICF.getInputFolder()):
        arr = os.listdir(ICF.getInputFolder())
        logging.info("QIA_PT_Reference>> ", QIA_PT_Reference)
        for i in arr:
            not_found = 0
            if re.search(r'QIA+\_[0-9]{2}_[0-9]{2}_' + QIA_PT_Reference, i) and i.find('~$') == -1:
                QIA_PT = i
                break
            else:
                not_found = 1
        logging.info("QIA_PT - ", QIA_PT)
        if not_found == 1:
            getQIA_PT_File = downloadQIA_PT(QIA_PT_Reference)
            if getQIA_PT_File != -1:
                for i in os.listdir(ICF.getInputFolder()):
                    logging.info("iiii ", i)
                    if re.search(r'QIA+\_[0-9]{2}_[0-9]{2}_' + QIA_PT_Reference, i) and i.find('~$') == -1:
                        return i

        else:
            return QIA_PT


#####################################
# I/P: testplan book object
# Desc: find the version of the test plan
# O/P: return testplan version if present else return -1
#####################################
def getTPVersion(tpBook):
    TP_Name = ""
    try:
        sheet = tpBook.sheets['Sommaire']
        if sheet:
            verCol = 1
            rowHeading = 5
            rowContentStart = rowHeading + 1
            return tpBook.sheets['Sommaire'].range(rowContentStart, verCol).value
    except Exception as exp:
        logging.info(f"Error in finding summary sheet {exp}")
        return -1


def getDocReferenceVer(docName):
    DocName = docName
    refPattern = ["\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]", "[0-9]{5}\_[0-9]{2}\_[0-9]{5}\_"]
    verPattern = ['\[V[0-9]{1,2}\.[0-9]{1,2}\]', 'V[0-9]{1,2}\.[0-9]{1,2}\_', '\_V[0-9]{1,2}\.[0-9]{1,2}',
                  '\_V[0-9]{1,2}', 'V[0-9]{1,2}\_']
    for refPat in refPattern:
        if re.search(refPat, docName):
            logging.info(";;;;;;;;;;;;;")
            DocName = re.sub(refPat, "", docName)
            break
    for verPat in verPattern:
        if re.search(verPat, DocName):
            logging.info("mmmmmmmmmmmmmmmmm")
            DocName = re.sub(verPat, "", DocName)
            break

    version = re.findall(r"\[V[0-9]{1,2}\.[0-9]{1,2}\]|V[0-9]{1,2}\.[0-9]{1,2}|V[0-9]{1,2}", docName)
    refNumber = re.findall(r"\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]|[0-9]{5}\_[0-9]{2}\_[0-9]{5}", docName)
    refNum = refNumber[0].strip("[]") if refNumber else ""
    verNum = version[0].strip("[]") if version else ""
    DocName = DocName.split('.')[0] if DocName else docName
    logging.info("refNum1 ->> ", refNum)
    logging.info("version ->> ", verNum)
    logging.info("DocName1 ->> ", DocName)

    return DocName, refNum, verNum


# def getQiaComment_old(QIA_Data_List, qia_pt_inp_doc):
#     qiaComment = {
#         'remark_cmt': " signal is not present in the test plan.",
#         'qiaPT':'',
#         'qiaPT_InpDoc':''
#     }
#     trigram = ICF.gettrigram()
#     if trigram.upper().find('EXPLEO'):
#         trigram = trigram.upper().replace('LEO', "")
#     Today = date.today()
#     curr_date = Today.strftime("%m/%d/%y")
#     logging.info("trigram ", trigram)
#     flag_qiaPT = 0
#     flag_qiaPT_InpDoc = 0
#     if QIA_Data_List:
#         for feps, req, signal, qiaType, docName in QIA_Data_List:
#             if qiaType == 2 and flag_qiaPT != 1:
#                 doc_name, refnum = getDocReference(docName)
#                 qiaComment['qiaPT'] = f"{trigram} ({curr_date}): Functional requirement {req} available in the {doc_name} Ref_No.{refnum} and it will take into the account in next update."
#                 flag_qiaPT = 1
#             elif (qiaType == 3 or qiaType == 4) and flag_qiaPT_InpDoc != -1:
#                 qiaComment['qiaPT_InpDoc'] = f"{trigram} ({curr_date}): Functional requirement not available and raised QIA of input document raised --QIAPT refnum--"
#                 flag_qiaPT_InpDoc = 1
#     logging.info("\nqiaComment >>> ", qiaComment)
#
#     return qiaComment

def getQiaComment(QIA_Data_List, final_clubbed_data, qia_pt_inp_doc=None):
    try:
        qiaComment = {
            'remark_txt': " signal is not present in the test plan.",
            'qiaPT': '',
            'qiaPT_InpDoc': '',
            'qia_comment_combined': ''
        }
        trigram = ICF.gettrigram()
        if trigram.upper().find('EXPLEO'):
            trigram = trigram.upper().replace('LEO', "")
        Today = date.today()
        curr_date = Today.strftime("%m/%d/%y")
        logging.info("trigram ", trigram)
        flag_qiaPT = 0
        flag_qiaPT_InpDoc = 0
        signal_req = ""
        if QIA_Data_List:
            for feps, req, signal, qiaType, docName, signalCont in QIA_Data_List:
                for final_club_data in final_clubbed_data:
                    if signalCont != "" and signalCont is not None:
                        splitted_signal = signalCont.split('==>')
                        if splitted_signal:
                            signal_req = splitted_signal[0]
                    logging.info(f"signal_req {signal_req}")
                    if final_club_data['combinedType'] != 'SNP':
                        if qiaType == 2 and flag_qiaPT != 1 and final_club_data['combinedType'].find('SPFM') != -1:
                            doc_name, refnum, vernum = getDocReferenceVer(docName)
                            qiaComment[
                                'qiaPT'] = f"{trigram} ({curr_date}): Functional requirement {signal_req} available in the {doc_name} Ref_No.{refnum} {vernum} and it will take into the account in next update."
                            flag_qiaPT = 1
                        elif (qiaType == 3 or qiaType == 4) and flag_qiaPT_InpDoc != -1:
                            if (final_club_data['combinedType'].find('SPNFR') != -1 or final_club_data[
                                'combinedType'].find(
                                    'SPFNM') != -1) and final_club_data['combinedType'].find('SPFM') != -1:
                                qiaComment[
                                    'qiaPT_InpDoc'] = f" \nFunctional requirement {qia_pt_inp_doc['reqs']} not available and raised QIA of input document ({qia_pt_inp_doc['ref_number'][0]}) (No.{int(qia_pt_inp_doc['raised_slno'])})."
                            elif (final_club_data['combinedType'].find('SPNFR') != -1 or final_club_data[
                                'combinedType'].find('SPFNM') != -1) and final_club_data['combinedType'].find(
                                'SPFM') == -1:
                                qiaComment[
                                    'qiaPT_InpDoc'] = f"{trigram} ({curr_date}): Functional requirement {qia_pt_inp_doc['reqs']} not available and QIA of input document raised ({qia_pt_inp_doc['ref_number'][0]}) (No.{int(qia_pt_inp_doc['raised_slno']) if qia_pt_inp_doc['raised_slno'] is not None else '-'})."

                            flag_qiaPT_InpDoc = 1
                    # else:
                    #     qiaComment[
                    #         'qiaPT'] = "Please proceed QIA PT for input document manually."
        qiaComment['qia_comment_combined'] = qiaComment['qiaPT'] + qiaComment['qiaPT_InpDoc']
        logging.info("\nqiaComment >>> ", qiaComment['qia_comment_combined'])
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"\nError:........ {ex} line no.{exc_tb.tb_lineno}.")
    return qiaComment


def getDataAsDict(QIA_Data_List):
    # SNP - signal not present in the input document
    # SPNFR  - signal present no functional requirement in input document
    # SPFM  - signal and functional requirement present in input document
    # SPFNM  - signal present but functional requirement p/c not matched in input document

    qiaDict = {
        'SNP': [],
        'SPNFR': [],
        'SPFM': [],
        'SPFNM': [],
    }

    CommentList = []
    if QIA_Data_List:
        for feps, req, signal, qiaType, docName, signalCont in QIA_Data_List:
            logging.info("123qiaType ", qiaType)
            if qiaType == 1:
                clubbedData = clubQIA_Data(qiaDict['SNP'], feps, req, signal)
                logging.info("clubQIA_Data SPN ", clubbedData)
            elif qiaType == 2:
                clubbedData = clubQIA_Data(qiaDict['SPFM'], feps, req, signal)
                logging.info("clubQIA_Data SPFM ", clubbedData)
            elif qiaType == 3:
                clubbedData = clubQIA_Data(qiaDict['SPFNM'], feps, req, signal)
                logging.info("clubQIA_Data SPFNM ", clubbedData)
            elif qiaType == 4:
                clubbedData = clubQIA_Data(qiaDict['SPNFR'], feps, req, signal)
                logging.info("clubQIA_Data SPNFR ", clubbedData)
        logging.info("qiaDict - ", qiaDict)

        return {"qiaDict": qiaDict}
    else:
        logging.info("\nNo data to raise the QIA PT...........")
        return -1


def combineTxt(qia_dic_list):
    cmt = ""
    if qia_dic_list:
        for ind, qia_data in enumerate(qia_dic_list):
            logging.info(qia_data, " --> qia_data")
            if ind > 0:
                cmt += f"\n{''.join(qia_data)}"
            else:
                cmt += f"{''.join(qia_data)}"
    logging.info("cmt123>> ", cmt)
    return cmt


def getQiaRemarks(qiaDataDict):
    logging.info("qiaDataDictqiaDataDict123 ", qiaDataDict)
    qia_comment_list = []
    for qia in qiaDataDict:
        if qiaDataDict[qia]:
            combine_qia_txt = combineTxt(qiaDataDict[qia])
            # qia_comment_list.append(combine_qia_txt)
            qia_comment_list.append({"qia_cmt": combine_qia_txt, "qia": qia})

    logging.info("qia_comment_list--> ", qia_comment_list)

    return qia_comment_list


def getDataAsTuple(data, docName, qiaType=None, fileType=None, signalContent=""):
    tupledata = (data['feps'], data['dciReq'], data['dciSignal'], qiaType, docName, signalContent)
    return tupledata


def clubQIA_Data(dictList, feps, req, signal):
    signal = str(signal.strip('$'))
    if len(dictList) > 0:
        for ind in range(len(dictList)):
            logging.info(ind)
            if feps + ": " == dictList[ind][0]:
                logging.info("---------")
                dictList[ind].append(", " + req + "(" + signal + ")")
            else:
                if ind == len(dictList) - 1:
                    logging.info("+++++++++++++")
                    dictList.insert(ind + 1, [feps + ": ", req + "(" + signal + ")"])
                else:
                    continue
    else:
        logging.info("***************")
        dictList.append([feps + ": ", req + "(" + signal + ")"])

    logging.info("dictList -- ", dictList)
    return dictList


#####################################
# I/P: testplan book object , dci data, document name and comment type
# Desc: raising the QIA PT in QIA PT Excel File and saving the File in output folder
# O/P: return number if no error else return -1
#####################################
def raiseQIA_PT(tpBook, clubbed_qia_data, qiaCmt):
    #  dciInfo, docName, cmtType=""
    logging.info("\n################# Raise QIA PT #################")
    logging.info("\n\nqiaDocumentData >> ", clubbed_qia_data)
    QIA_PT_Reference = getQIA_PT_Reference(tpBook)
    logging.info("QIA_PT_Reference >> ", QIA_PT_Reference)
    QIA_PT_File = getQIA_PT_File(QIA_PT_Reference)
    successFlag = 0
    if QIA_PT_File != "" and QIA_PT_File != None:
        logging.info("QIA_PT_File >> ", QIA_PT_File)
        QIA_PT_Book = EI.openExcel(ICF.getInputFolder() + "\\" + QIA_PT_File)
        QIAsheet = QIA_PT_Book.sheets["QIA"]
        QIAsheet.activate()
        # maxrow_slno = QIAsheet.range(QIAmaxrow, 1).value
        # row = QIAmaxrow + 1
        sslNoCol, TP_VerCol, refCol, ETapeCol, remarqueCol, criticiteCol, categoryCol, emetteurCol, dateCol, statusCol, commentCol = 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 13
        getTPVer = getTPVersion(tpBook)
        logging.info("||||||||||getTPVersion >> ", getTPVer)
        Today = date.today()
        # if prevDoc is not None and prevDoc != "" and prevDoc != -1:
        #     prevDocSplit = os.path.basename(prevDoc)
        # logging.info("docName >> ", docName)
        # qiaComment = ""
        # if cmtType == 1:
        #     qiaComment = "The signal " + dciInfo[
        #         'dciSignal'] + " is not present in test plan and the functional requirement for the signal is present in the document " + docName + "."
        # elif cmtType == 3:
        #     qiaComment = "The signal " + dciInfo[
        #         'dciSignal'] + " is not present in the test plan and there is no functional requirement available for this signal."

        logging.info("qia_comment_list123 ", clubbed_qia_data)
        reqData = []
        for qiadata in clubbed_qia_data:
            if qiadata['remark_data'] != "" and qiadata['remark_data'] is not None:
                logging.info("\n\n type1111 ", qiadata['combinedType'])
                raisedSlNo = ""
                commentVal = ""
                if qiadata['combinedType'] == 'SNP':
                    commentVal = "Please proceed QIA PT for input document manually."
                else:
                    commentVal = qiaCmt['qia_comment_combined']
                try:
                    QIAmaxrow = QIAsheet.range('E' + str(QIAsheet.cells.last_cell.row)).end('up').row
                    logging.info("QIAmaxrow - ", QIAmaxrow)
                    maxrow_slno = QIAsheet.range(QIAmaxrow, 1).value
                    row = QIAmaxrow + 1
                    QIAsheet.range(row, sslNoCol).value = (int(maxrow_slno + 1))
                    QIAsheet.range(row, TP_VerCol).value = getTPVer if getTPVer is not None and getTPVer != -1 else ""
                    QIAsheet.range(row, refCol).value = "General"
                    QIAsheet.range(row, ETapeCol).value = ""
                    QIAsheet.range(row, remarqueCol).value = qiadata['remark_data']+" "+qiaCmt['remark_txt']
                    QIAsheet.range(row, criticiteCol).value = "Majeur"
                    QIAsheet.range(row, categoryCol).value = "Fond"
                    QIAsheet.range(row, emetteurCol).value = ICF.gettrigram().upper().replace('LEO', "")
                    QIAsheet.range(row, dateCol).value = Today.strftime("%m/%d/%y")
                    QIAsheet.range(row, statusCol).value = 'Open'
                    QIAsheet.range(row, commentCol).value = commentVal
                    raisedSlNo = QIAsheet.range(row, sslNoCol).value
                    logging.info("raisedSlNo2 .. ", raisedSlNo)
                    requirement = re.sub(r'\(.*?\)', "", qiadata['remark_data'])
                    req = re.sub(r'FEPS\_[0-9]{1,9}\:', "", requirement.strip())
                    logging.info("reqqnew ", req.strip())
                    reqData.append((req.strip(), raisedSlNo))
                    successFlag = 1
                except Exception as exp:
                    logging.info(f"\n***************** Error in raising the QIA PT1 {exp} *****************")
                    continue
                    # return -1, raisedSlNo
                    successFlag = 0
        logging.info("successFlag ", successFlag)
        if successFlag == 1:
            QIA_PT_Book.save()
            saveQIA_PT(QIA_PT_Book, 1)
            QIA_PT_Book.close()
    return {'status': 1, 'reqData': reqData}


#####################################
# I/P: testplan book object , document name , signal content
# Desc: process the content of the requirement where the signal present
# - popuping the content to get the user response to find the signal is produced or consumend
# O/P: return response (P/C/cancelled) and the numeric value to find signal status are mismatched or not
#####################################
def processMatchedSignalContent(tpBook, findSignalInDoc, docName, dciInfo):
    logging.info("************ processMatchedSignalContent **************")
    top = tk.Tk()
    top.geometry("100x100")
    inputStatusFound = 0
    getDataAsTupleRes = ""
    signalContent = ""
    if findSignalInDoc:
        for signalData in findSignalInDoc:
            # showSignalStatusConfirmation(signalData)
            response = messagebox.askyesnocancel(
                "Is the signal [" + dciInfo['dciSignal'] + "] Produced or Consumed",
                "Click 'Yes' if it is produced.\nClick 'No' if it is consumed.\n\n" + signalData.split('==>')[1])
            if response and response is not None:
                logging.info("---------------Produced----------------")
                isProduced = compareSignalStatus(dciInfo, 'P')
                if isProduced:
                    logging.info("+++++++++++Signal status is produced raise QIA PT+++++++++++++++")
                    signalContent = signalData
                    inputStatusFound = 0
                    getDataAsTupleRes = getDataAsTuple(dciInfo, docName, 2, 1, signalContent)
                    logging.info("getDataAsTupleRes1 >> ", getDataAsTupleRes)
                    break
                    # QIA_PT_Response, raisedSlNo = raiseQIA_PT(tpBook, dciInfo, docName, cmtType=1)
                    # if QIA_PT_Response != -1:
                    #     logging.info(f"✔Raising QIA PT done for the signal {dciInfo['dciSignal']}...............")
                    #     inputStatusFound = 0
                    #     break
                else:
                    logging.info("+++++++++++Signal status is not same+++++++++++++++")
                    inputStatusFound = 1

            elif not response and response is not None:
                logging.info("---------------Consumed----------------")
                isConsumed = compareSignalStatus(dciInfo, 'C')
                if isConsumed:
                    logging.info("+++++++++++Signal status is produced raise QIA PT+++++++++++++++")
                    signalContent = signalData
                    inputStatusFound = 0
                    getDataAsTupleRes = getDataAsTuple(dciInfo, docName, 2, 1, signalContent)
                    logging.info("getDataAsTupleRes2 >> ", getDataAsTupleRes)
                    break
                    # QIA_PT_Response, raisedSlNo = raiseQIA_PT(tpBook, dciInfo, docName, cmtType=1)
                    # if QIA_PT_Response != -1:
                    #     logging.info(f"✔Raising QIA PT done for the signal {dciInfo['dciSignal']}...............")
                    #     inputStatusFound = 0
                    #     break
                else:
                    logging.info("+++++++++++Signal status is not same+++++++++++++++")
                    inputStatusFound = 1

            else:
                logging.info("---------------Cancelled go to next requirement----------------")

            logging.info("\n\nresponse>> ", response)

    # return inputStatusFound, response, raisedSlNo
    return inputStatusFound, getDataAsTupleRes, response


#####################################
# I/P: testplan book object , QIA PT book object , input data and signal response
# Desc: add the new line with input data in QIA PT Excel
# O/P: return numeric value 1 or -1
#####################################
def addDataInInputDoc_QIA_PT(tpBook, QIA_PT_Book, qia_inp_data):
    # qiaInpDocData, signalResponse, dciInfo

    try:
        raised_slno = ""
        result = {'response': "", 'raised_slno ': raised_slno}
        if QIA_PT_Book is not None:
            notesSheet = QIA_PT_Book.sheets["Suivi des Remarques"]
            notesSheet.activate()
            notesSheet_value = notesSheet.used_range.value
            # noteSheetMaxRow = notesSheet.range('G' + str(notesSheet.cells.last_cell.row)).end('up').row
            # question_remark_col_pos = EI.searchDataInExcel(notesSheet, "", "Question / Remark")
            question_remark_col_pos = EI.searchDataInExcelCache(notesSheet_value, "", "Question / Remark")
            q_row, q_col = question_remark_col_pos['cellPositions'][0]
            logging.info("qx, qy --> ", q_row, q_col)
            col_alp = xlsxwriter.utility.xl_col_to_name(q_col - 1)
            logging.info("col_alp ... ", col_alp)
            noteSheetMaxRow = notesSheet.range(col_alp + str(notesSheet.cells.last_cell.row)).end('up').row

            logging.info("QIAmaxrow2 - ", noteSheetMaxRow)
            row = noteSheetMaxRow + 1
            maxrow_slno = notesSheet.range(noteSheetMaxRow, 1).value
            logging.info("maxrow_slno1 - ", maxrow_slno)
            if maxrow_slno == "" or maxrow_slno is None:
                maxrow_slno_prev = notesSheet.range(row - 1, 1).value
                logging.info("maxrow_slno_prev >> ", maxrow_slno_prev)
                if maxrow_slno_prev is not None and maxrow_slno_prev != "":
                    maxrow_slno = (int(maxrow_slno_prev + 1))
            logging.info("maxrow_slno2 >> ", maxrow_slno)
            sslNoCol, HCol, docCol, localizationCol, typologyCol, criticityCol, questionCol, emitterCol, statusCol = 1, 2, 3, 4, 5, 6, 7, 8, 12
            qiaComment = ""
            docColValue = ""
            docFileName = ""
            trigram = ICF.gettrigram()
            if trigram.upper().find('EXPLEO'):
                trigram = trigram.upper().replace('LEO', "")
            try:
                notesSheet.range(row, sslNoCol).value = maxrow_slno
                notesSheet.range(row, q_col).value = f"Major - {trigram} {qia_inp_data['inp_doc']} {qia_inp_data['remark']}."
                raised_slno = notesSheet.range(row, sslNoCol).value
                QIA_PT_Book.save()
                saveQIA_PT(QIA_PT_Book, 2)
                QIA_PT_Book.close()
                result['response'] = 1
                result['raised_slno'] = raised_slno
            except Exception as exp:
                logging.info(f"\n***************** Error in raising the QIA PT2 {exp} *****************")
                result['response'] = -1
                result['raised_slno'] = raised_slno
    except Exception as exp:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"\nError in adding data in QIA PT Input Document........ {exp} line no.{exc_tb.tb_lineno}.")
    return result


#####################################
# I/P: testplan book object , dci data , input data and signal response
# Desc: it will download the QIA PT of input document from docinfo protal and fill the data in new row
# O/P: return numeric value 1 or -1
#####################################
def raiseQIA_InputDoc(tpBook, qia_inp_data):
    # dciInfo, qiaInpDocData, signalResponse
    logging.info("Finding the QIA of Input Document")
    # logging.info("\nqiaInpDocData = ", qiaInpDocData)
    logging.info("\ninp_doc_name = ", qia_inp_data)
    ref_number_arr = []
    QIA_PT_input_doc_res = ""
    result = {"QIA_PT_input_doc_res": 0, 'ref_number': []}
    try:
        # for qia_inp_data in qiaInpDocData:
        docRefVer = []
        qia_ref_num = ""
        refnum = getDocNameRefVer(qia_inp_data['inp_doc'], 'ref')
        docRefVer.append((refnum, ''))  #'02017_19_02191'
        logging.info("\ndocRefVer -> ", docRefVer)
        qiaInputDocFile, refNumber = WI.startDocumentDownloadFilesFromDossier(docRefVer, True)
        # qiaInputDocFile = ["QIA_SSD_HMIF_MOBY_HMI.xlsm"]
        logging.info("qiaInputDocFile>> ", qiaInputDocFile)
        if qiaInputDocFile and qiaInputDocFile != -1:
            time.sleep(3)
            QIA_InputDocBook = EI.openExcel(ICF.getInputFolder() + "\\" + qiaInputDocFile[0])
            logging.info("QIA_InputDocBook ", QIA_InputDocBook)
            # QIA_PT_input_doc_res = addDataInInputDoc_QIA_PT(tpBook, QIA_InputDocBook, qiaInpDocData, signalResponse, dciInfo)
            QIA_PT_input_doc_res = addDataInInputDoc_QIA_PT(tpBook, QIA_InputDocBook, qia_inp_data)
            logging.info(f"QIA_PT_input_doc_res >>> {QIA_PT_input_doc_res}")
            if QIA_PT_input_doc_res['response'] == 1:
                ref_number_arr.append(refNumber)
            result["QIA_PT_input_doc_res"] = QIA_PT_input_doc_res['response']
            result['ref_number'] = ref_number_arr
            result['raised_slno'] = QIA_PT_input_doc_res['raised_slno']
        else:
            # display popup to user to download the QIA input document manualy
            logging.info("\n\n------------- Dossier Tab not found for the reference " + refnum + " please do it manually ------------------")
            UpdateHMIInfoCb("\n\n------------- Dossier Tab not found for the reference " + refnum + " please do it manually ------------------")
            result["QIA_PT_input_doc_res"] = -2
            result['ref_number'] = ref_number_arr
            result['raised_slno'] = ""
    except Exception as exp:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        logging.info(f"\nError in downloading the QIA Input Document................. {exp} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
        result["QIA_PT_input_doc_res"] = -1
        result['ref_number'] = ref_number_arr
        result['raised_slno'] = ""

    return result


def getDocNameRefVer(inpDoc, type=None):
    verMod1 = ""
    verMod = ""
    refNum = ""
    DocName = ""
    DocName = re.sub(r"\[V[0-9]{1,2}\.[0-9]{1,2}\]\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]\_", "", inpDoc)
    logging.info(DocName, "findFileName")
    findVer = re.findall(r"\[V[0-9]{1,2}\.[0-9]{1,2}\]", inpDoc)
    logging.info(findVer)
    logging.info(findVer[0].strip("[]"), "findVer[0].strip")
    refNumber = re.findall(r"\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]", inpDoc)
    refNum = refNumber[0].strip("[]") if refNumber else ""
    logging.info("refNum1 ->> ", refNum)
    if findVer:
        verMod = re.sub(r'(v|V)', "", findVer[0].strip("[]"))
        verMod1 = re.sub(r'\.[0]', "", verMod)
    if type == 'ref':
        return refNum
    else:
        return DocName, refNum, verMod1


#####################################
# I/P: testplan book object, dci data
# Desc: it processes the input documents from input folder and check the signal exit or not
# O/P: return numeric value 1 or -1
#####################################
def processInputDocs(tpBook, dciInfo):
    inputDocsList = []
    getDataAsTupleRes = ""
    if os.path.isdir(ICF.getInputFolder()):
        inputFiles = os.listdir(ICF.getInputFolder())
        for ipDoc in inputFiles:
            if (os.path.splitext(ipDoc)[1] == ".docx" or os.path.splitext(ipDoc)[1] == ".doc"):
                inputDocsList.append(ipDoc)

    logging.info("inputDocsList ", inputDocsList)
    # inputDocsList = ['[V4.0][02014_19_00792]SSVS_SSFD_GENx_RSP_AUE_ALERT_UNALLOWED_EVENTS.docx','[V5.0][02014_19_00792]SSVS_SSFD_GENx_RSP_AUE_ALERT_UNALLOWED_EVENTS._Proof_Reading.docx',
    # '[V6.0][02014_19_00792]SSVS_SSFD_GENx_RSP_AUE_ALERT_UNALLOWED_EVENTS.docx']

    # '[V5.0][02014_19_00792]SSVS_SSFD_GENx_RSP_AUE_ALERT_UNALLOWED_EVENTS._Proof_Reading.docx',
    # '[V6.0][02014_19_00792]SSVS_SSFD_GENx_RSP_AUE_ALERT_UNALLOWED_EVENTS.docx'
    logging.info("inputDocsList Len", len(inputDocsList))
    qiaInpDocData = {
        'DocName': '',
        'DocVer': '',
        'DocRefNum': '',
        "signalResponse": ''
    }
    result = {
        "resultResponse": "",
        "reqData": "",
        "qiaInpDocData": "",
        "processedFileList": "",
        'signal_exist_file': ""
    }
    # raisedSlNo_QIA_PT = ""
    signal_exist_file = []
    try:
        isSignalNotMatched = 0
        signalResponse = ""
        refNumList = []
        onlyInTable = 0
        for index, inpDoc in enumerate(inputDocsList):
            if os.path.splitext(inpDoc)[1] == ".docx":
                logging.info("\n\n\ninpDoc>>", inpDoc)
                logging.info("\n\n\ninpDoc index>>", index)
                DocName = re.sub(r"\[V[0-9]{1,2}\.[0-9]{1,2}\]\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]\_", "", inpDoc)
                logging.info(DocName, "findFileName")
                findVer = re.findall(r"\[V[0-9]{1,2}\.[0-9]{1,2}\]", inpDoc)
                logging.info(findVer)
                logging.info(findVer[0].strip("[]"), "findVer[0].strip")
                refNumber = re.findall(r"\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]", inpDoc)
                refNum = refNumber[0].strip("[]") if refNumber else ""
                logging.info("refNum2 ->> ", refNum)
                if findVer:
                    verMod = re.sub(r'(v|V)', "", findVer[0].strip("[]"))
                    verMod1 = re.sub(r'\.[0]', "", verMod)
                    logging.info("ipDocs['ipDocuments'][index], verMod1 >> ", inpDoc, ", ", verMod1)
                else:
                    verMod = ""
                    verMod1 = ""
                # prevDoc = BL.getDocPath(DocName, verMod1)
                prevDoc = BL.getDocPathQIA(DocName, verMod1)
                logging.info("prevDoc - ", prevDoc)
                if (type(prevDoc)) == str:
                    time.sleep(15)
                    prevTableList = WDI.getTables(prevDoc)
                    # logging.info("prevTableList - ", prevTableList)
                    findSignalInDoc, isOnlyInTable = WDI.findTableOfContent(prevTableList, dciInfo['dciSignal'].strip("$"))
                    logging.info("findSignalInDoc - ", findSignalInDoc)
                    if findSignalInDoc != -1 and len(findSignalInDoc) != 0:
                        qiaInpDocData['DocName'] = DocName
                        qiaInpDocData['DocVer'] = findVer[0].strip("[]")
                        qiaInpDocData['DocRefNum'] = refNum
                        logging.info("qiaInpDocData >> ", qiaInpDocData)
                        refNumList.append(refNum)
                        logging.info(f"\nSignal exist in {DocName}")
                        # signalResult, response = processMatchedSignalContent(tpBook, findSignalInDoc, inpDoc, dciInfo)
                        signalResult, getDataAsTupleRes, response = processMatchedSignalContent(tpBook, findSignalInDoc, inpDoc, dciInfo)
                        # if raisedSlNo is not None and raisedSlNo != "":
                        #     raisedSlNo_QIA_PT = raisedSlNo
                        logging.info("\n[[[[[signalResult, response]]]]] ", signalResult, response)
                        logging.info(type(signalResult))
                        if signalResult != 1:
                            isSignalNotMatched = 0
                            break
                        elif signalResult == 1:
                            isSignalNotMatched = 1
                            if response == True:
                                signalResponse = 'P'
                            elif response == False:
                                signalResponse = 'C'
                            else:
                                signalResponse = 'Q'
                            if index != len(inputDocsList) - 1:
                                continue
                            else:
                                logging.info(
                                    "+++++++++++1Signal status is mismatched with DCI and Document raise QIA input docuent+++++++++++++++")
                                qiaInpDocData['signalResponse'] = signalResponse
                                getDataAsTupleRes = getDataAsTuple(dciInfo, inpDoc, 3, 2)
                                if onlyInTable == 1 and len(findSignalInDoc) == 0:
                                    qiaInpDocData['signalResponse'] = 'OT'
                                    signalResponse = 'OT'
                                signal_exist_file.append((inpDoc, dciInfo['dciReq'], dciInfo['dciSignal'], signalResponse, ""))
                                logging.info("signal_exist_file11 ", signal_exist_file)
                                # QIA_Input_DOC = raiseQIA_InputDoc(tpBook, dciInfo, qiaInpDocData, signalResponse)
                                # if QIA_Input_DOC != -1:
                                #     logging.info(
                                #         f"✔Raising QIA PT Input Document done for the signal {dciInfo['dciSignal']} in file {qiaInpDocData['DocName']}...............")
                    else:
                        logging.info(f"\nsignal not present in input document {DocName}")
                        logging.info("indexindex > ", index)
                        if index != len(inputDocsList) - 1:
                            continue
                        else:
                            logging.info("%%%%%%%%%%%%%%%% isSignalNotMatched >> ", isSignalNotMatched, " - ", signalResponse)
                            if isSignalNotMatched != 1 and onlyInTable != 1:
                                logging.info(".........")
                                getDataAsTupleRes = getDataAsTuple(dciInfo, inpDoc, 1, 1)
                                logging.info("getDataAsTupleRes3 >> ", getDataAsTupleRes)
                                # QIA_PT_Response, raisedSlNo = raiseQIA_PT(tpBook, dciInfo, inpDoc, cmtType=3)
                                # if raisedSlNo is not None and raisedSlNo != "":
                                #     raisedSlNo_QIA_PT = raisedSlNo
                                # if QIA_PT_Response != -1:
                                #     logging.info(f"✔Raising QIA PT done for the signal {dciInfo['dciSignal']}...............")
                            else:
                                logging.info(
                                    "+++++++++++2Signal status is mismatched with DCI and Document raise QIA input docuent+++++++++++++++")
                                logging.info("refNumList >> ", refNumList)
                                getDataAsTupleRes = getDataAsTuple(dciInfo, inpDoc, 4, 2)
                                logging.info(f"dciInfo['dciReq'] {dciInfo}")
                                if onlyInTable == 1 and len(findSignalInDoc) == 0:
                                    signalResponse = 'OT'
                                    qiaInpDocData['DocName'] = DocName
                                    qiaInpDocData['DocVer'] = findVer[0].strip("[]")
                                    qiaInpDocData['DocRefNum'] = refNum
                                    qiaInpDocData['signalResponse'] = signalResponse
                                signal_exist_file.append((inpDoc, dciInfo['dciReq'], dciInfo['dciSignal'], signalResponse, ""))
                                logging.info("signal_exist_file22 ", signal_exist_file)
                                # QIA_Input_DOC = raiseQIA_InputDoc(tpBook, dciInfo, qiaInpDocData, signalResponse)
                                # if QIA_Input_DOC != -1:
                                #     logging.info(
                                #         f"✔Raising QIA PT Input Document done for the signal {dciInfo['dciSignal']} in file {qiaInpDocData['DocName']}...............")
        result['resultResponse'] = 1
        result['reqData'] = getDataAsTupleRes
        result['signal_exist_file'] = signal_exist_file
    except Exception as exp:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"Error in processing the signal content:-{exp} {exc_tb.tb_lineno}")
        result['resultResponse'] = -1
    logging.info(f"ffffffffffffffresult {result}")
    return result


def checkIsDocumentExist(reference, version):
    logging.info("\n...............Checking the document is exist in input folder or not..............")
    logging.info("Referece - ", reference)
    logging.info("Version - ", version)
    isFound = 0
    try:
        if os.path.isdir(ICF.getInputFolder()):
            arr = os.listdir(ICF.getInputFolder())
            for inpDoc in arr:
                if os.path.splitext(inpDoc)[1] == ".docx":
                    # or os.path.splitext(inpDoc)[1] == ".doc" or
                    # os.path.splitext(inpDoc)[1] == ".docm" or os.path.splitext(inpDoc)[1] == ".rtf"
                    logging.info("inpDoc ", inpDoc)
                    findVer = re.findall(r"\[V[0-9]{1,2}\.[0-9]{1,2}\]", inpDoc)
                    logging.info("findVer1 ", findVer)
                    logging.info(findVer[0].strip("[]"), "findVer[0].strip")
                    inpDocVer = findVer[0].strip("[]")
                    refNumber = re.findall(r"\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]", inpDoc)
                    inpDocRef = refNumber[0].strip(
                        "[]") if refNumber and refNumber is not None and refNumber != "" else ""
                    logging.info("refNum2 ->> ", inpDocRef)
                    if reference == inpDocRef and version == inpDocVer:
                        isFound = 1
                        break
    except Exception as exp:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"\nError.......... {exp} line no: {exc_tb.tb_lineno}")

    return isFound


#####################################
# I/P: testplan book object, input documents reference and version
# Desc: download all the .doc or docx file
# O/P: return numeric value 1 or -1
#####################################
def downloadIpDocs(tpBook, ipDocs):
    logging.info("++++++++++++++++++++++++++++")
    logging.info("Downloading the previous input documents from summary sheet.....................")
    try:
        logging.info("ipDocs['ipDocuments'][0] ", ipDocs['ipDocuments'][0])
        for index, ipdocs in enumerate(ipDocs['ipDocuments']):
            docRefVer = []
            logging.info(index, "type(index)")
            # exit()
            if ipDocs['ipReferences'][index] is not None and ipDocs['ipVersions'][index] is not None:
                docRefVer.append((ipDocs['ipReferences'][index], str(ipDocs['ipVersions'][index])))
                logging.info("\ndocRefVer -> ", docRefVer)
                isDocFound = checkIsDocumentExist(ipDocs['ipReferences'][index], ipDocs['ipVersions'][index])
                logging.info("isDocFound > ", isDocFound)
                if isDocFound != 1:
                    WI.startDocumentDownload(docRefVer, True)
        return 1
    except Exception as excp:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"\nError in downloading the input documents.......... {excp} line no: {exc_tb.tb_lineno}")
        return -1


#####################################
# I/P: testplan book object
# Desc: get the document details except dci file from summary tab in test plan
# O/P: returns document data as dictionary
#####################################
def getInputDocsExcept_DCI(tpBook):
    result = {
        'ipDocuments': [],
        'ipVersions': [],
        'ipReferences': []
    }
    try:
        sommaireSheet = tpBook.sheets['Sommaire']
        sommaireSheet.activate()
        time.sleep(3)
        maxrow = sommaireSheet.range('A' + str(sommaireSheet.cells.last_cell.row)).end('up').row
        logging.info("\nmaxrow sommaire :- ", maxrow)
        start_row = 5 + 1
        ipDocuments = []
        ipVersions = []
        ipReferences = []
        ref_ver_list = []
        for row in range(start_row, maxrow):
            ref_ver_flag = 0
            if sommaireSheet.range(row, 5).value is not None and sommaireSheet.range(row, 5).value.find("DCI") == -1 and sommaireSheet.range(row, 5).value != '--' and sommaireSheet.range(row,5).value != '-':
                ref = str(sommaireSheet.range(row, 6).value)
                ver = str(sommaireSheet.range(row, 7).value)
                if ref_ver_list:
                    for ind, ref_ver in enumerate(ref_ver_list):
                        if ref in ref_ver and ver in ref_ver:
                            ref_ver_flag = 1
                            break
                        else:
                            if ind == len(ref_ver_list) - 1:
                                ref_ver_list.append([ref, ver])
                                break
                else:
                    ref_ver_list.append([ref, ver])
                if ref_ver_flag != 1:
                    ipDocuments.append(sommaireSheet.range(row, 5).value)
                    ipReferences.append(ref)
                    ipVersions.append(ver)
        result['ipDocuments'] = ipDocuments
        result['ipReferences'] = ipReferences
        result['ipVersions'] = ipVersions
        logging.info(f"result:\n{result}")
    except Exception as exp:
        logging.info(f"\nException in getting the input Documents :- {exp}")
        # writeLog(QIA_PT_LogFile, 'a', f"\nException in getting the input Documents :- {exp}")

    return result


def getReqId(reqID):
    if reqID.find("("):
        splitRequirements = reqID.split("(")
        Requirement = splitRequirements[0]
    else:
        splitRequirements = reqID.split(" ")
        Requirement = splitRequirements[0]
    logging.info("Requirement --> ", Requirement)
    return Requirement


def check_doc_and_req(filep, inpdocList):
    docfound = 0
    istype = 0
    for inp_doc, req, signal, type, signalCont in filep:
        DocName = re.sub(r"\[V[0-9]{1,2}\.[0-9]{1,2}\]\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]", "", inp_doc)
        refNumber = re.findall(r"\[[0-9]{5}\_[0-9]{2}\_[0-9]{5}\]", inp_doc)
        refNum = refNumber[0].strip("[]") if refNumber else ""
        for inpdoc in inpdocList:
            if DocName.strip() in inpdoc['inp_doc']:
                if type == inpdoc['type']:
                    istype = 1
                docfound = 1
    return docfound, istype


def combineQiaPtInpDocData(file_with_signal):
    inpdocList = []
    for filep in file_with_signal:
        i = 0
        for inp_doc, req, signal, type, signalCont in filep:
            signal_strip = signal.strip('$')
            isfound, istype = check_doc_and_req(filep, inpdocList)
            if type == 'P':
                remark = f"Please provide the requirement which consumed this signal"
            elif type == 'C':
                remark = f"Please provide the requirement which produced this signal"
            elif type == 'OT':
                remark = f"please provide the functional requirement of the signal"
            logging.info("\n\n isfound", isfound)
            if isfound != 1:
                inpdocList.append({'inp_doc': inp_doc, 'req': req, 'remark': remark + " " + signal_strip, 'type': type, 'signalCont': signalCont})
            else:
                # inpdocList[i][1] is requirement
                inpdocList[i]['req'] = inpdocList[i]['req'] + ", " + req
                inpdocList[i]['type'] = inpdocList[i]['type'] + ", " + type
                inpdocList[i]['signalCont'] = inpdocList[i]['signalCont'] + ", " + type
                logging.info("\n\n----123", inpdocList[i]['remark'])
                if signal_strip not in inpdocList[i]['remark'] and istype != 1:
                    inpdocList[i]['remark'] = inpdocList[i]['remark'] + ".\n" + remark + " " + str(signal_strip)
                else:
                    if signal_strip not in inpdocList[i]['remark']:
                        inpdocList[i]['remark'] = inpdocList[i]['remark'] + ", " + str(signal_strip)
            i += 1

    logging.info("inpdocList123.... ", inpdocList)
    return inpdocList


def downloadDocsForQIA_PT(tpBookObj):
    try:
        ipDocs = getInputDocsExcept_DCI(tpBookObj)
        logging.info("ipDocs >> ", ipDocs)
        downlod_inp_docs = downloadIpDocs(tpBookObj, ipDocs)
        return downlod_inp_docs
    except Exception as ex:
        logging.info("Something went wrong in downloadDocsForQIA_PT function in downloading the input documents from summary..")
        return -1


# process the interface requirement signal to raise the QIA PT
def processInerfaceReqSignal(tpBookObj, dciInfo):
    BL.displayInformation("\nProcessing the QIA PT.....")
    process_result = ""
    getDataAsTupleRes = ""
    if tpBookObj is None or tpBookObj == "":
        tpBookObj = openTestPlanSheet()
    if tpBookObj != -1 and tpBookObj is not None:
        testSheetList = findSignalInTestSheet(tpBookObj, dciInfo)
        if not testSheetList:
            logging.info(f"\n------------{dciInfo['dciSignal']} not present in Test Plan------------")
            process_result = processInputDocs(tpBookObj, dciInfo)
            logging.info("\n\nprocessDocRes  ", process_result)

    return process_result


def removeReq(data_dict, qia_inp_doc_res):
    logging.info("\n-----------------Removing the requirement------------------")
    if (qia_inp_doc_res['type'].find('P') != -1 or qia_inp_doc_res['type'].find('C') != -1) and qia_inp_doc_res[
        'type'].find('OT') != -1:
        keys = ['SPFNM', 'SPNFR']
    elif qia_inp_doc_res['type'].find('OT') != -1 and qia_inp_doc_res['type'].find('P') == -1 and qia_inp_doc_res[
        'type'].find('C') == -1:
        keys = ['SPNFR']
    else:
        keys = ['SPFNM']

    logging.info(f"keys ==> {keys}")
    for dic_key in keys:
        logging.info("=================================")
        logging.info(f"dic_key11 --> {dic_key}")
        for i, key_value in enumerate(data_dict[dic_key]):
            feps, req_signal = key_value
            logging.info(feps, req_signal, " feps, req_signal\n")
            reqs = qia_inp_doc_res['reqs'].split(',')
            logging.info(f"reqs --> {reqs}")
            find_reqs = re.findall('DCI', req_signal.upper())
            for reqID in reqs:
                logging.info(f"reqID --> {reqID}")
                if req_signal.find(reqID.strip()) != -1 and len(find_reqs) > 1:
                    if re.search(reqID.strip() + r"*\([^()]*\)+\([^()]*\)\,", data_dict[dic_key][i][1]):
                        data_dict[dic_key][i][1] = re.sub(reqID + r"*\([^()]*\)+\([^()]*\)\,", "",
                                                          data_dict[dic_key][i][1]).strip()
                    elif re.search(r"\, " + reqID.strip() + "*\([^()]*\)+\([^()]*\)", data_dict[dic_key][i][1]):
                        data_dict[dic_key][i][1] = re.sub(r"\, " + reqID.strip() + "*\([^()]*\)+\([^()]*\)", "",
                                                          data_dict[dic_key][i][1]).strip()
                elif req_signal.find(reqID.strip()) != -1 and len(find_reqs) <= 1:
                    logging.info("\n\n1data_dict_key ", data_dict[dic_key])
                    data_dict[dic_key].pop(i)

    logging.info("\n\n\n\ndata_dict>> ", data_dict)
    return data_dict


def club_all_qia_data(qia_comment_list):
    newList = []
    combineTxt = ""
    combineType = ""
    indval = 0
    for qiacmt in qia_comment_list:
        logging.info("qiatype ", qiacmt['qia'])
        if qiacmt['qia'] != 'SNP':

            if indval > 0:
                combineTxt += "\n" + qiacmt['qia_cmt']
                combineType += ", " + qiacmt['qia']
            else:
                combineTxt += qiacmt['qia_cmt']
                combineType += qiacmt['qia']
            indval += 1
        else:
            newList.append({'remark_data': qiacmt['qia_cmt'], 'combinedType': qiacmt['qia']})
    newList.append({'remark_data': combineTxt, 'combinedType': combineType})

    logging.info(f"combineTxt --> {combineTxt}")
    logging.info(f"{newList} --> newList")
    return newList

# tpBook = openTestPlanSheet()

# reqDciData = [{'dciSignal': '$ETAT112_SURVEILLANCE112_AS112', 'network': 'LIN_VSM_1', 'pc': 'C', 'framename': 'LIN_VSM_1/NEA_ETAT_MUS_2/ETAT_SURV_AS', 'dciReq': 'DCINT-00069700(0)', 'proj_param': '', 'feps':'FEPS_117614'},{'dciSignal': '$ETAT_SURVEILLANCE_AS', 'network': 'LIN_VSM_1', 'pc': 'C', 'framename': 'LIN_VSM_1/NEA_ETAT_MUS_2/ETAT_SURV_AS', 'dciReq': 'DCINT-00069701(0)', 'proj_param': '', 'feps':'FEPS_117613'},{'dciSignal': '$ETAT_SURVEILLANCE_AS', 'network': 'LIN_VSM_1', 'pc': 'C', 'framename': 'LIN_VSM_1/NEA_ETAT_MUS_2/ETAT_SURV_AS', 'dciReq': 'DCINT-00069699(0)', 'proj_param': '', 'feps':'FEPS_117612'},
# {'dciSignal': '$E.IKHARRAZEN', 'network': 'LIN_VSM_1', 'pc': 'P', 'framename': 'LIN_VSM_1/NEA_ETAT_MUS_2/ETAT_SURV_AS', 'dciReq': 'DCINT-00069698(0)', 'proj_param': '', 'feps':'FEPS_117612'},
# ]

# def sample():
#     QIA_Data_List = []
#     file_with_signal = []
#     for reqData in reqDciData:
#         qia_data_res = processInerfaceReqSignal(tpBook, reqData)
#         logging.info("\n\nQIAPT_Response ", qia_data_res['resultResponse'])
#         logging.info("\nQIA_Data ", qia_data_res)
#         if qia_data_res['resultResponse'] != -1 and qia_data_res['resultResponse'] != "" and qia_data_res['resultResponse'] != None and qia_data_res['reqData']:
#             QIA_Data_List.append(qia_data_res['reqData'])
#         if qia_data_res['signal_exist_file']:
#             file_with_signal.append(qia_data_res['signal_exist_file'])
#         logging.info("file_with_signal3 ", file_with_signal)
#     logging.info("QIA_Data_List ", QIA_Data_List)
#     if QIA_Data_List:
#         result = getDataAsDict(QIA_Data_List)
#         if result["qiaDict"]:
#             logging.info("len(result['qiaDict']['SPNFR']) ", len(result["qiaDict"]['SPNFR']), "----", len(result["qiaDict"]['SPFNM']))
#
#             if file_with_signal:
#                 combine_qia_pt_data = combineQiaPtInpDocData(file_with_signal)
#                 logging.info("combine_qia_pt_data123 ", combine_qia_pt_data)
#                 if combine_qia_pt_data:
#                     for qia_inp_doc_data in combine_qia_pt_data:
#                         qia_pt_inp_doc_res = raiseQIA_InputDoc(tpBook, qia_inp_doc_data)
#                         qia_inp_doc_data['type'] = qia_inp_doc_data['type']
#                         qia_inp_doc_data['reqs'] = qia_inp_doc_data['req']
#                         qia_pt_inp_doc_res['reqs'] = qia_inp_doc_data['req']
#                         logging.info("\n\nqia_pt_inp_doc_res11 ", qia_pt_inp_doc_res)
#                         if qia_pt_inp_doc_res['QIA_PT_input_doc_res'] == -2:
#                             result["qiaDict"] = removeReq(result["qiaDict"], qia_inp_doc_data)
#
#             qia_comment_list = getQiaRemarks(result["qiaDict"])
#             logging.info("qia_comment_list1233 ", qia_comment_list)
#             if qia_comment_list:
#                 final_clubbed_comment = club_all_qia_data(qia_comment_list)
#                 logging.info(f"final_clubbed_comment --> {final_clubbed_comment}")
#                 qia_comment_data = getQiaComment(QIA_Data_List, final_clubbed_comment, qia_pt_inp_doc_res)
#                 logging.info(f"qia_comment_data --> {qia_comment_data}")
#                 qiaResponse = raiseQIA_PT(tpBook, final_clubbed_comment, qia_comment_data)
#                 if qiaResponse['status'] == 1 and qiaResponse['reqData'] != "" and qiaResponse['reqData'] is not None:
#                     logging.info("qiaResponse qia--> ", qiaResponse)
#                     # for req, slno in qiaResponse['reqData']:
#                     #     split_req = req.split('\n')
#                     #     logging.info("split_req ", split_req)
#                     #     for ReqId in split_req:
#                     #         getReqCoords = EI.searchDataInCol(tpBook.sheets['Impact'], 1, ReqId)
#                     #         logging.info("getReqCoords >> ", getReqCoords)
#                     #         if getReqCoords['count'] > 0:
#                     #             tpBook.sheets['Impact'].activate()
#                     #             commentCol = 5
#                     #             for cellPos in getReqCoords['cellPositions']:
#                     #                 row, col = cellPos
#                     #                 EI.setDataFromCell(tpBook.sheets['Impact'], (row, commentCol), f"QIA PT raised for this requirement\nSlNo. {slno}.")
# sample()
if __name__ == "__main__":
    # processInerfaceReqSignal(tpBook, reqDciData)
    ICP.loadConfig()
    logging.info(ICF.getInputFolder())
    logging.info(os.listdir(ICF.getInputFolder()))
    exit()
