import datetime
import os

import nltk

import ExcelInterface as EI
import WordDocInterface as WDI
import InputConfigParser as ICP
from nltk import word_tokenize
import web_interface as WI
import re
import logging
date_time = datetime.datetime.now()

nltk.download('punkt')

# ICP.loadConfig()


vernum = '([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})|([0-9]{1,2}\.[0-9]{1,2})'
refnum = '([0-9]{5})+(_[0-9]{2})+(_[0-9]{5})+'
pattren_ver = "([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})"
ref = r'([0-9]{5}_[0-9]{2}_[0-9]{5})'


def getCellAbsVal(sheet, row, col):
    for i in range(row, 0, -1):
        cellVal = EI.getDataFromCell(sheet, f"{col}{i}")
        if cellVal is not None:
            return cellVal
    return None


def downloadSSD(tpBook):
    global summarySheet
    try:
        summarySheet = tpBook.sheets["Sommaire"]
    except Exception as e:
        logging.info(f"Sommaire sheet not found!")
        exit(1)
    referenceList = []

    # get all reference numbers for ssd files
    nrows = summarySheet.used_range.last_cell.row
    for i in range(6, nrows):
        typeVal = EI.getDataFromCell(summarySheet, f"E{i}")
        if typeVal is not None and re.match("ssd", typeVal.lower()):
            referenceNumber = getCellAbsVal(summarySheet, i, "F")
            refver = getCellAbsVal(summarySheet, i, "G")
            logging.info('refver-->', refver)
            referenceList.append(referenceNumber)
    logging.info("referenceList --> ", referenceList)

    # remove duplicate elements from list
    referenceList = [*(set(referenceList))]

    logging.info("After Duplicate removal referenceList --> ", referenceList)
    # iterate over all reference numbers and download respective documents
    for referenceNum in referenceList:
        WI.startDocumentDownload([[referenceNum, ""]], False, False)


def open_Exact_ssd_file(path):
    SSD = []
    Parameters_Board = []
    dir_list = os.listdir(path)
    logging.info('dir_list---->', dir_list)
    for i in dir_list:
        if (i.find('SSD') != -1) and (i.find('Parameters_Board') == -1):
            SSD.append(i)
            logging.info('li---->', i)
        if (i.find('SSD') != -1) and (i.find('Parameters_Board') != -1):
            Parameters_Board.append(i)
            logging.info('lri---->', i)
    return [SSD, Parameters_Board]


def SSD_Param_Board_Validation(calibration_flow, requirement, SsdParBoardSheet):
    SsdParBoardSheet_value = SsdParBoardSheet.used_range.value
    final = {}
    # result = EI.searchDataInExcel(SsdParBoardSheet, "", calibration_flow)
    result = EI.searchDataInExcelCache(SsdParBoardSheet_value, "", calibration_flow)
    cell_positions = result['cellPositions']
    if calibration_flow!='':
        logging.info('calibration_flow -------->', calibration_flow)
        final.update({'Column_H': calibration_flow})
    for cell_pos in cell_positions:
        logging.info('cell_pos', cell_pos)
        row, col = cell_pos
        Unit_type = SsdParBoardSheet.range(row, 8).value
        logging.info('Unit_type--->', Unit_type)
        rang_min = SsdParBoardSheet.range(row, 9).value
        logging.info('rang_min', rang_min)
        range_max = SsdParBoardSheet.range(row, 10).value
        logging.info('range_max', range_max)
        resolution = SsdParBoardSheet.range(row, 11).value
        logging.info('resolution', resolution)
        Ref_Val = SsdParBoardSheet.range(row, 12).value
        logging.info('Ref_Val--->', Ref_Val)
        if (Unit_type=='TIME') and Ref_Val.find("ms")!=-1:
            found_ms = Ref_Val
            milli_seconds = Ref_Val.replace('ms', '')
            logging.info('replaced', milli_seconds)
            logging.info('found_ms--->', found_ms)
            final.update({'Column_I': milli_seconds})
        elif (Unit_type=='TIME') and Ref_Val.find("s")!=-1:
            replaced2 = Ref_Val.replace('s', '')
            logging.info('replaced2', replaced2)
            milli_seconds = int(replaced2) * 1000
            logging.info('converted_to_milisec --->', milli_seconds)
            final.update({'Column_I': milli_seconds})
        if (Unit_type=='s') or (Unit_type=="ms") or (Unit_type=="TIME"):
            unit = 'ms'
        elif (Unit_type=='N/A') or (Unit_type=='') or (Unit_type is None) or (Unit_type=='None'):
            unit = 'dimensionless'
        else:
            break
        final.update({"Column_L": unit})
        resolution = resolution or ''
        rang_min = rang_min or ''
        range_max = range_max or ''
        if rang_min!='' and range_max!='' and resolution!='':
            Column_M = f'''{requirement}
Range : [{rang_min} : {range_max}]
Resolution : [{resolution}]'''
            final.update({"Column_M": Column_M})
            return final
        elif (rang_min=='' or range_max=='') and resolution!='':
            Column_M = f'''{requirement}
Resolution : [{resolution}]'''
            final.update({"Column_M": Column_M})
            return final
        elif (rang_min!='' and range_max!='') and resolution=='':
            Column_M = f'''{requirement}
Range : [{rang_min} : {range_max}]'''
            final.update({"Column_M": Column_M})
            return final
        elif (rang_min=='' or range_max=='') and resolution=='':
            Column_M = f'{requirement}'
            final.update({"Column_M": Column_M})
            return final


def extract_ssd__param_board_file():
    SSD_File = []
    Param_Board_File = []
    ssd_folder = ICP.getSsdFolder()
    if os.path.isdir(ssd_folder):
        dir_list = os.listdir(ssd_folder)
        logging.info('dir_list --->', dir_list)
        for filename in dir_list:
            if (filename.find('SSD')) != -1 and (filename.find('ALLOC_MATRIX_SSD')) == -1 and (filename.find('Parameters_Board')) == -1:
                SSD_File.append(filename)
            if (filename.find('SSD')) != -1 and (filename.find('ALLOC_MATRIX_SSD')) == -1 and (filename.find('Parameters_Board')) != -1:
                Param_Board_File.append(filename)
    logging.info('SSD_File---->', SSD_File)
    logging.info('Param_Board_File---->', Param_Board_File)
    return [SSD_File, Param_Board_File]


def UpdateQiaParamGlobal(tpBook, valueMap, ReqName, ReqVer, trigram):
    global actual_content
    logging.info('trigram----->', trigram)
    UpdateQiaParam = {}
    list_calibration_flow = []
    logging.info('ICP.getSsdFolder()---->', ICP.getSsdFolder())
    if ICP.getQIAParamCalibration() is not False:
        qiaParamBook = EI.openExcel(ICP.getQIAParamCalibration())
        qiaSheet = qiaParamBook.sheets['NEW_QIA']
        downloadSSD(tpBook)
        files = extract_ssd__param_board_file()
        logging.info('files[0]---->', files[0])
        logging.info('files[1]---->', files[1])
        if files[0] is not None:
            for file in files[0]:
                logging.info('file----->', file)
                joint = (os.path.join(ICP.getSsdFolder(), file))
                logging.info('joint-1-1-1-1-1->', joint)
                requirement1 = (ReqName + ' (' + ReqVer + ')')
                req = (ReqName + ' ' + ReqVer)
                logging.info('req---->', req)
                UpdateQiaParam.update({'Requirement': requirement1})
                logging.info('requirement1 ------>', requirement1)
                actual_content = WDI.getReqContent(joint, ReqName, ReqVer)
                logging.info('actual_content--->', actual_content)
                if actual_content != -1:
                    logging.info('++++++++++++++')
                    # logging.info('actual_content["Content"]-------->', actual_content["Content__"])
                    logging.info('actual_content["Content"]-------->', actual_content["content"])
                    logging.info('++++++++++++++')
                    # words = word_tokenize(actual_content['Content__'])
                    words = word_tokenize(actual_content['content'])
                    UpdateQiaParam.update({'Content': words})
                    logging.info('UpdateQiaParam["Content"]----->', UpdateQiaParam['Content'])
                    param_board_ref = re.search(refnum, file)
                    logging.info('x.group()------>', param_board_ref.group())
                    UpdateQiaParam.update({'Par_calib_ref': param_board_ref.group()})
                    logging.info('"Par_calib_ref": param_board_ref.group() -------->', UpdateQiaParam["Par_calib_ref"])
                    break
                else:
                    logging.info('actual_content------->', actual_content)
        if len(files[1]) > 0 and actual_content != -1:
            for file in files[1]:
                logging.info('file----->', file)
                joint = (os.path.join(ICP.getSsdFolder(), file))
                logging.info('joint-0-0-0-0->', joint)
                Param_Board_doc = EI.openExcel(joint)
                logging.info('Param_Board_doc --->', Param_Board_doc)
                Param_Board_sheet = Param_Board_doc.sheets["Flows Functions"]
                maxrow = Param_Board_sheet.range('E' + str(Param_Board_sheet.cells.last_cell.row)).end('up').row
                logging.info('max_row_A', maxrow)
                for x in range(2, maxrow + 1):
                    list_calibration_flow.append(Param_Board_sheet.range(x, 5).value)
                logging.info('list_calibration_flow---->', list_calibration_flow)
                matches = [x for x in list_calibration_flow if x in UpdateQiaParam["Content"]]
                logging.info("Matches:", matches)
                if len(matches) > 0:
                    final0 = SSD_Param_Board_Validation(matches[0], UpdateQiaParam["Requirement"], Param_Board_sheet)
                    logging.info('final0--->', final0)
                    if final0 is not None:
                        maxrow = qiaSheet.range('B' + str(qiaSheet.cells.last_cell.row)).end('up').row
                        logging.info('max_row_A', maxrow)
                        val_A = qiaSheet.range(f"A{maxrow}").value
                        EI.setDataFromCell(qiaSheet, f"A{maxrow + 1}", int(val_A + 1))
                        logging.info('val_A+1', int(val_A + 1))
                        for key in valueMap:
                            EI.setDataFromCell(qiaSheet, key + str(maxrow + 1), valueMap[key])
                            if key=='I':
                                EI.setDataFromCell(qiaSheet, key + str(maxrow + 1), final0['Column_I'])
                            if key=='L':
                                EI.setDataFromCell(qiaSheet, key + str(maxrow + 1), final0['Column_L'])
                            if key=='M':
                                EI.setDataFromCell(qiaSheet, key + str(maxrow + 1), final0['Column_M'])
                            if key=='H':
                                EI.setDataFromCell(qiaSheet, key + str(maxrow + 1), final0['Column_H'])
                            if key=='U':
                                EI.setDataFromCell(qiaSheet, key + str(maxrow + 1), str(trigram + ' ' + (date_time.strftime("%m/%d/%y")) + " : " + "Information can be found on SSD_Parameter_board Ref : " + UpdateQiaParam['Par_calib_ref']))
                        logging.info(f'QIA of Calibration Updated successfully__for this Requirement____{UpdateQiaParam["Requirement"]}____')
                        break
                    else:
                        logging.info(' ---> None!!! $ None!!! $None!!! <--- ', ' ---> Considering only Time as unit<--- ')
                else:
                    logging.info(f'There is no calibration flow in supporting Document for this requirement--> {UpdateQiaParam["Requirement"]}')
        else:
            logging.info('There is no requirement present in ssd doc......')
        qiaParamBook.save()
        qiaParamBook.close()
    else:
        logging.info('\n*******************************************************\n')
        logging.info('!!!!!!!!!  QIA PARAM file is not present  !!!!!!!!!')
        logging.info('\n*******************************************************\n')


#######################################################################################################################
# valueMap = {
#     "B": 'Testplan_reference[0]',
#     "C": 'functionName',
#     "D": "Creation",
#     "E": "New Calibration for NEA",
#     "F": "NEA",
#     "G": "--",
#     "H": "*********",
#     "I": "*********",
#     "J": "--",
#     "K": 'functionName',
#     "L": "*********",
#     "M": "*********",
#     "N": "--",
#     "O": "--",
#     "P": 'trigram',
#     "Q": 'date_time.strftime("%m/%d/%y")',
#     "R": "Open",
#     "U": 'str(trigram + (date_time.strftime("%m/%d/%y")) + " Information'
# }
#
# input_document = 'fffffffffff'
# ReqName = 'REQ-0770204'
# ReqVer = 'A'
# tpBook = EI.openExcel(r'C:\Users\clakshminarayan\Documents\BSI-VSM(Automation)\QIA_Param_2\QIA_Input_Files' + "\\" + "Tests_20_79_01272_19_01474_FSE_RCTA_V14_VSM.xlsm")
# if __name__=="__main__":
#     UpdateQiaParamGlobal(tpBook, valueMap, ReqName, ReqVer, input_document)
