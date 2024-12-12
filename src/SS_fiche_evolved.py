import ExcelInterface as EI
import re
import json
import time
from datetime import date
import KeyboardMouseSimulator as KMS
import TestPlanMacros as TPM
import InputConfigParser as ICF
import threading
import BusinessLogic as BL
import AnaLyseThematics as AT
import AnalyseTestSheet as ATS
import InputDocLinkPopup as IDLP
import logging
import web_interface as wi


def getArch(taskname):
    arch = "VSM"
    x = re.findall("^F_", taskname)
    if x:
        arch = "BSI"
    return arch

def QIA_ssfiche_dict(req,ver,sf_sheet,comment,req_ver_sf):
    # book=EI.openExcel('C:\\Users\\vgajula\\PycharmProjects\\pythonProject\\python learning\\pythonProject\\F_PICC_06_02_2023\\Output_Files\\Testplan.xlsx')

    #sheet = book.sheets['Impact']
    #values = sheet.used_range.value
   # maxrow = len(values)
  #  sf_sheets = {}
   # req_ver_sf = {'req': [], 'ver': [], 'sf_sheet': [], 'flow': [], 'req_comment': []}
    #if (maxrow > 17):
        # result=EI.searchDataInSpecificRows(sheet,(17,maxrow-17), 4,re.compile("([(BSI)|(VSM)][0-9]{2}_SF_[0-9]{2}_[0-9]{2}_[0-9]{4}[A-Z]?)"))
    #i=values[-1]
          # i=values[-1]
    #try:
     #         tes = re.findall("((?:BSI|VSM)[0-9]{2}_SF_[0-9]{2}_[0-9]{2}_[0-9]{4}[A-Z]?)", i[3])
      #        if (i[4]!=''):
       #         if len(tes) != 0 and ('already present' not in i[4]):
    req_ver_sf['req'].append(req)
    req_ver_sf['ver'].append(ver)
    req_ver_sf['sf_sheet'].append(sf_sheet)
    req_ver_sf['req_comment'].append(comment)
    #except:
     #           pass


    # logging.info(req_ver_sf)
def QIA_ssfiche_update(book, arch, taskname, trigram, disp_info,req_ver_sf):
    sf_sheets = {}
    for i in req_ver_sf['sf_sheet']:
        flow = []
        for j in i:
            sheet = book.sheets[j]
            value = EI.getDataFromCell(sheet, 'C2')
            flow_name = value.split('|')[0]
            flow_value = value.split('|')[1]
            flow.append((flow_name, flow_value))
        req_ver_sf['flow'].append(flow)

    logging.info(req_ver_sf)

    flow_req_com = {}
    for i, j, k in zip(req_ver_sf['req'], req_ver_sf['flow'], req_ver_sf['req_comment']):
        for l in j:
            if l[0] not in flow_req_com:
                flow_req_com[l[0]] = {'req_com': [(k, i)], 'values': [l[1]]}
            else:
                flow_req_com[l[0]]['values'].append(l[1])
                if (k, i) not in flow_req_com[l[0]]['req_com']:
                    flow_req_com[l[0]]['req_com'].append((k, i))
                # if i not in flow_req_com[l[0]]['reqs']:
                #  flow_req_com[l[0]]['reqs'].append(i)
    logging.info(flow_req_com)

    for i, j in flow_req_com.items():
        j['values'] = list(set(j['values']))
        flow_req_com[i]['values'] = j['values']

    logging.info(flow_req_com)
 #   qia_sf_file = EI.findInputFiles()[16]
   # logging.info(f"qia_sf_book {qia_sf_file}")
    while(True):
      qia_sf_file = EI.findInputFiles()[16]
      if (qia_sf_file==''):

            display_info = disp_info
            display_info("Please wait downloading the QIA file for sous fiche")
            arch1 = getArch(ICF.FetchTaskName())
            if arch1 == "BSI":
                wi.startDocumentDownload([["00952_09_01398", ""]])
            elif arch1 == 'VSM':
                wi.startDocumentDownload([['01272_19_00614', ""]])
      else:
          break

    if qia_sf_file:
        qia_sf_book = EI.openExcel(ICF.getInputFolder() + "\\" + qia_sf_file)
        sheet1 = qia_sf_book.sheets['QIA']
        if sheet1.api.AutoFilterMode == True:
            sheet1.api.AutoFilterMode = False
        values = sheet1.used_range.value
        #sheet1.api.AutoFilter.ShowAllData()
        maxrow = sheet1.range('B' + str(sheet1.cells.last_cell.row)).end('up').row
        b = maxrow
        logging.info("b-->", b)
        c = maxrow + 1
        logging.info("maxrow-->", maxrow)
        value = EI.getDataFromCell(sheet1, f'A{maxrow}')
        logging.info("ddd--->", value)
        # maxrow=len(values)
        # logging.info(values)

     #   user_inp=json.load(open('C:\\Users\\vgajula\\PycharmProjects\\pythonProject\\python learning\\pythonProject\\New folder\\user_input\\UserInput.json'))
        Today = date.today()
        curr_date = Today.strftime("%m/%d/%y")
        sheet = book.sheets['Sommaire']
        for i, j in flow_req_com.items():
            # EI.setDataFromCell(sheet1,f'A{maxrow}',value+1)

            value1 = EI.getDataFromCell(sheet, 'C3')[:2]
            value2 = EI.getDataFromCell(sheet, 'C4')[:2]
            taskData = ICF.getTaskDetails()
            value3 = taskname.split('_')[0]
            value4 = taskname.split('_')[1]
            if (arch=='BSI'):
                x=''
            else:
                x='X'
            func = value1 + "_" + value2 + "_" + value3 + "_" + value4
            logging.info(f"func>> {func}")
            final_value_row = [str(value + 1), func, 'General', x, '', i, '\n'.join(j['values']), 'Modification', 'P0',
                               curr_date, ','.join([f"{val[1]} {val[0]}" for val in j['req_com']]),
                               trigram, 'SETTLED', curr_date,
                               '\n'.join([val[0] + val[1] for val in j['req_com']])]
            EI.setDataFromCell(sheet1, f'A{str(c)}:O{str(c)}', final_value_row)
            c = c + 1
            value = value + 1


        for i, j, k in zip(req_ver_sf['req'], req_ver_sf['sf_sheet'], req_ver_sf['req_comment']):
            for l in j:
                if(k!=''):
                      if ('already present' in k)==False:
                        if l not in sf_sheets:
                            sf_sheets[l] = []
                            sf_sheets[l].append((i, k))
                        else:
                            sf_sheets[l].append((i, k))

    return sf_sheets,req_ver_sf


def ss_fiche_update(sf_sheets, ss_fiche):
    macro = EI.getTestPlanAutomationMacro()
    # tpBook = ss_fiche
    sheet2 = ss_fiche.sheets['Sommaire']
    sheets = ['VSM20_N1_99_99_' + i.split('_')[-1] if i[:3] == 'VSM' else 'BSI04_N1_99_99_' + i.split('_')[-1] for i
              in sf_sheets.keys()]

    sheets = [i for i in sheets if i in ss_fiche.sheet_names]

    value = EI.getDataFromCell(sheet2, 'C2')
    Today = date.today()
#    user_inp = json.load(open(
#        'C:\\Users\\vgajula\\PycharmProjects\\pythonProject\\python learning\\pythonProject\\New folder\\user_input\\UserInput.json'))


    # for i,j in sf_sheets.items():
    #    xx=i.split('_')[-1]
    #   sheet_name='VSM20_N1_99_99_'+xx if i[3]=='VSM' else 'BSI04_N1_99_99_'+xx

    # Step 16
    # fillSummary(tpBook, ipDocName, fepsString, referentielList[0], triList[0])
    if (EI.getDataFromCell(sheet2, 'B6') is None):
        # TPM.selectTPInit(macro)
        EI.activateSheet(ss_fiche, ss_fiche.sheets['Sommaire'])
        # time.sleep(1)
        # TPM.unProtectTestSheet(macro)
       # time.sleep(1)
        #TPM.selectTpWritterProfile(macro)
        #time.sleep(1)
        #TPM.selectArch(macro)
        #time.sleep(1)
        #TPM.selectTPInit(macro)
        #time.sleep(1)
        #TPM.selectTestSheetModify(macro)
        #time.sleep(1)
        # TPM.unProtectTestSheet(macro)
        # time.sleep(1)

        # tpBook.sheets['Sommaire'].range('6:7').insert()
        # tpBook.sheets['Sommaire'].range('5:6').insert()
        # for i in range(1,7):
        # tpBook.sheets['Sommaire'].range((6, i), (7, i)).merge()
        try:
            row = [ICF.getTaskDetails()[0]['trigram'], Today.strftime("%d/%m/%Y"),
                   ICF.getTaskDetails()[0]['trigram'] + ' ' + Today.strftime(
                       "%d/%m/%Y") + ' ' + 'Modifying the sheets ' + ','.join(sheets)]

            EI.activateSheet(ss_fiche, ss_fiche.sheets['Sommaire'])
            TPM.unProtectTestSheet(macro)
            time.sleep(1)
            EI.setDataFromCell(sheet2, 'B6:D6', row)
        except:
            pass
    #  fillSummary(tpBook,sheets,user_inp['taskDetails'][0]['trigram'])
    else:

        row = [EI.getDataFromCell(sheet2, 'B6') + '\n' + ICF.getTaskDetails()[0]['trigram'],
               EI.getDataFromCell(sheet2, 'C6') + '\n' + Today.strftime("%d/%m/%y"),
               EI.getDataFromCell(sheet2, 'D6') + '\n' + ICF.getTaskDetails()[0]['trigram'] + ' ' + Today.strftime(
                   "%d/%m/%Y") + ' ' + 'Modifying the sheets ' + ','.join(sheets)]
        EI.activateSheet(ss_fiche, ss_fiche.sheets['Sommaire'])
        TPM.unProtectTestSheet(macro)
        EI.setDataFromCell(sheet2, 'B6:D6', row)
    return sheets


def ts_modify_reqver(sheet, i):
    getString = sheet.range(6, 3).value
    logging.info("ts_modify_reqver Before modify"+ getString)
    logging.info("ts_modify_reqver i = ", i)
  #  for i in keyword:
    res = re.search(i[0] + '\(([A-Z]|[0-9])\)', getString)
    logging.info("res = ", res)
    if res is not None:
        getString = str.replace(getString, res.group(), i[0] + '(' + i[1] + ')')
    else:
        getString = getString + '|' + i[0] + '(' + i[1] + ')'

    logging.info("ts_modify_reqver After modify:", getString)

    return getString
    # code to update requirement version


def fillSheetHistory(sheet, keyword):
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
    EI.setDataFromCell(sheet, (row + 1, col + 1), ICF.getTaskDetails()[0]['trigram'])

    EI.setDataFromCell(sheet, (row + 1, col), getString)

def doEnter(testSheetStatus):
    if testSheetStatus != 'VALIDEE':
        logging.info("vvvvvvvvvvvvvvvvvvvv")
        TPM.doImpactPressEnter()
#UpdateHMIInfoCb = None
pattren_ver = "([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})"


def add_func_them_comment_in_history(tpBook, sf_sheet, themImpactComment, contentFuncImpact, req, ex_them):
    logging.info("themImpactComment-->", themImpactComment)
    logging.info("contentFuncImpact-->", contentFuncImpact)
    logging.info("ATS.newContent-->", ATS.newContent)
    logging.info("ATS.oldContent-->", ATS.oldContent)
    newContent_ev = len(ATS.newContent) if ATS.newContent != -1 else 0
    oldContent_ev = len(ATS.oldContent) if ATS.oldContent != -1 else 0
    # newContent_ev = ""
    # oldContent_ev = ""
    logging.info("newContent_ev-->", newContent_ev, newContent_ev)
    logging.info("oldContent_ev-->", oldContent_ev, oldContent_ev)
    hstryCmt = ""
    if len(themImpactComment) == 0 and len(contentFuncImpact) == 0 and newContent_ev == oldContent_ev:
         hstryCmt = "\nNo functional impact."
    elif (len(themImpactComment) != 0) and (len(contentFuncImpact) == 0):
        # hstryCmt = "Thematic Chnages "+str(', '.join(themImpactComment))+"."
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines(f"\n\nRequirement: {req} \nTest sheet of SF {sf_sheet.name}\nExisting Thematic: {ex_them}\nThematic as per the requirement: {str('| '.join(themImpactComment))}")
        time.sleep(2)
 #   if hstryCmt != "":
    fillSheetHistory(sf_sheet, hstryCmt)


def fillThematicsSF(tpbook,tpsheet,sfbook,sfsheet):

    maxrow = tpbook[tpsheet].range('A' + str(sfsheet.cells.last_cell.row)).end('up').row
    try:
        # cellValue = EI.searchDataInExcel(tpbook[tpsheet], (1, maxrow), "Nature des modifications")
        sheet_value = tpbook[tpsheet].used_range.value
        cellValue = EI.searchDataInExcelCache(sheet_value, (1, maxrow), "Nature des modifications")

    except:
        # cellValue = EI.searchDataInExcel(tpbook[tpsheet], (1, maxrow), "Nature des modifications")
        sheet_value = tpbook[tpsheet].used_range.value
        cellValue = EI.searchDataInExcelCache(sheet_value, (1, maxrow), "Nature des modifications")

    row, col = cellValue["cellPositions"][0]
    logging.info("In History", row, col)
    logging.info("sheet.range(row + 1, col).value", sfsheet.range(row + 1, col).value)
    if tpbook[tpsheet].range(row + 1, col).value is not None and 'Evolved requirement' in tpbook[tpsheet].range(row + 1, col).value :
        getString = sfsheet.range(row + 1, col).value


    thematiclines=''
    r=8
    while(tpbook[tpsheet].range(r,1).value.str.contains('THEMATIQUE')):
        thematiclines= thematiclines+'\n'+thematiclines+tpbook[tpsheet].range(r,1).value + " " + tpbook[tpsheet].range(r,3).value

    foundString=sfbook[sfsheet].range(8,3).value
    with open('../ThematicsSF_Report.txt', 'a') as f:
        f.writelines(
            "\n\n" + f"Found in {sfsheet} thematics {foundString}.The update is {getString}.The thematic lines in testplan sheet {tpsheet} are {thematiclines}")


def treat_SF_evolved_req(tpBook, ssfiche, SF_Sheets,themImpactComment, contentFuncImpact, reqName, reqVer,Req,newreq=""):

    macro = EI.getTestPlanAutomationMacro()
    logging.info(f"SF_Sheets -> {SF_Sheets}")

 #  value = "Updating requirements " + f"Incremented requirement {ReqName} from version {reqVer} to {ReqVer}"
 #    if(newreq !=""):
    if(newreq == Req):
       # value="Updating requirements " + f"Incremented requirement from {Req} to {newreq}"
        Reqt=newreq
    else:
        Reqt=Req
       # value = "Updating requirements " + f"Incremented requirement {ReqName} from version {reqVer} to {ReqVer}"
    if Reqt.find('(') != -1:
        ReqName = Reqt.split("(")[0].split()[0] if len(Reqt.split("(")) > 0 else ""
        ReqVer = Reqt.split("(")[1].split(")")[0] if len(Reqt.split("(")) > 1 else ""
    else:
        ReqName = Reqt.split()[0] if len(Reqt.split()) > 0 else ""
        ReqVer = Reqt.split()[1] if len(Reqt.split()) > 1 else ""
    value = "Evolved requirement." + f"Incremented requirement from {reqName}({reqVer}) to {ReqName}({ReqVer})"
    sheet2 = ssfiche.sheets['Sommaire']
    value1 = EI.getDataFromCell(sheet2, 'C2')
    if (value1 == 'VALIDEE'):
        # TPM.selectTPInit(macro)
        EI.activateSheet(ssfiche, ssfiche.sheets['Sommaire'])
        # time.sleep(1)
        # TPM.unProtectTestSheet(macro)
        time.sleep(1)
        TPM.selectTpWritterProfile(macro)
        time.sleep(1)
        TPM.selectArch(macro)
        time.sleep(1)
        TPM.selectTPInit(macro)
        time.sleep(1)
        TPM.selectTestSheetModify(macro)
        time.sleep(1)

    #if(flag==1):
    #    sf_sheets, req_ver_sf = QIA_ssfiche_update(tpBook)
     #   logging.info(f"sf_sheets -> {sf_sheets}")
     #   sheets = ss_fiche_update(sf_sheets,ssfiche)
      #  logging.info(f"{sheets} - sheets1234")
    sfBook = ssfiche
    macro = EI.getTestPlanAutomationMacro()
    # TPM.selectArch(macro)
    # TPM.selectTpWritterProfile(macro)
    # time.sleep(1)
    # TPM.selectArch(macro)

    for i in SF_Sheets:
        # macro = EI.getTestPlanAutomationMacro()
        sheet = 'VSM20_N1_99_99_' + i.split('_')[-1] if i[:3] == 'VSM' else 'BSI04_N1_99_99_' + i.split('_')[-1]
       # value = "Updating requirements " + ",".join([j[0] + j[1] for j in sf_sheets[i]])
        #value = "Updating requirements " + f"Incremented requirement {ReqName} from version {reqVer} to {ReqVer}"
       # if(newreq !=""):
      #      value="Updating requirements " + f"Incremented requirement from {Req} to {newreq}"
        # logging.info(sf_sheets)
        try:
            if sfBook.sheets[sheet] is not None:
                EI.activateSheet(sfBook, sfBook.sheets[sheet])
                # macro=EI.getTestPlanAutomationMacro()
                time.sleep(1)
                TPM.unProtectTestSheet(macro)
                getString = sfBook.sheets[sheet].range(6, 3).value
                #  for i in keyword:
                res = re.search(ReqName+'\('+ReqVer+'\)', getString)
                if(res is None):
                    testSheetStatus = sfBook.sheets[sheet].range(7, 3).value
                    TPM.selectTestSheetModify(macro)
                    logging.info(f"testSheetStatus {testSheetStatus}")
                   # TPM.unProtectTestSheet(macro)
                    fillSheetHistory(sfBook.sheets[sheet], value)
                    result = ts_modify_reqver(sfBook.sheets[sheet],(ReqName,ReqVer))
                    EI.setDataFromCell(sfBook.sheets[sheet], (6, 3), result)
                    add_func_them_comment_in_history(tpBook, sfBook.sheets[sheet] ,themImpactComment, contentFuncImpact, f"{reqName} ({str(reqVer)})", sfBook.sheets[sheet].range(8,3).value)
        except:
            pass
    logging.info("==============")
    #sfBook.activate()
    #TPM.selectSynthUpdateFor_SF(macro)
    #time.sleep(3)
    #sfBook.save()



    logging.info(
        'ooo999999900000000000000000'
    )



# if __name__ == "__main__":
#     ICF.loadConfig()
#     user_inp = json.load(open(
#         'C:\\Users\\vgajula\\PycharmProjects\\pythonProject\\python learning\\pythonProject\\New folder\\user_input\\UserInput.json'))
#     # book = EI.openExcel(
#     #     'C:\\Users\\vgajula\\PycharmProjects\\pythonProject\\python learning\\pythonProject\\F_PICC_06_02_2023\\Output_Files\\Testplan.xlsx')
#     tpBook = EI.openTestPlan()
#     sf_sheets,req_ver_sf = QIA_ssfiche_update(tpBook)
#     logging.info(sf_sheets)
#     sheets = ss_fiche_update(sf_sheets)
#     logging.info(sheets)
#     sfBook = EI.openExcel(
#         'C:\\Users\\vgajula\\PycharmProjects\\pythonProject\\python learning\\pythonProject\\New folder\\Tests_00952_10_05276_ss_fiches.xlsm')
#     macro = EI.getTestPlanAutomationMacro()
#     # TPM.selectArch(macro)
#     TPM.selectTpWritterProfile(macro)
#     time.sleep(1)
#     TPM.selectArch(macro)
#
#     for i in sf_sheets.keys():
#         # macro = EI.getTestPlanAutomationMacro()
#         sheet = 'VSM20_N1_99_99_' + i.split('_')[-1] if i[:3] == 'VSM' else 'BSI04_N1_99_99_' + i.split('_')[-1]
#         value = "Updating requirements " + ",".join([j[0] + j[1] for j in sf_sheets[i]])
#         # logging.info(sf_sheets)
#         try:
#             if sfBook.sheets[sheet] is not None:
#                 EI.activateSheet(sfBook, sfBook.sheets[sheet])
#                 # macro=EI.getTestPlanAutomationMacro()
#                 time.sleep(1)
#                 testSheetStatus = sfBook.sheets[sheet].range(7,3).value
#                 TPM.selectTestSheetModify(macro)
#                 logging.info(f"testSheetStatus {testSheetStatus}")
#                 # t1 = threading.Thread(target=doEnter(testSheetStatus))
#                 # t1.start()
#                 fillSheetHistory(sfBook.sheets[sheet], value)
#                 # modify req version in test sheet
#                 # if tpBook.sheets[sheet] is not None:
#                 result = ts_modify_reqver(sfBook.sheets[sheet], sf_sheets[i])
#                 # EI.setDataFromCell(tpBook.sheets[sheet],user_inp['taskDetails'][0]['trigram'])
#                 EI.setDataFromCell(sfBook.sheets[sheet], (6, 3), result)
#                 fillThematicsSF(tpBook,i,sfBook,sheet)
#
#                 #verInfo = fillImpactEvolved(tpBook, oldReq, fepsNumber, newReq)
#
#              #   evolReq(sfBook, fepsNumber, macro, Arch, feps, rqIDs, reqname, requirement="")
#         except:
#             pass
#
# #    analyseDeEntrant = EI.openAnalyseDeEntrant(user_inp['taskDetails'][0]['taskName'])
#  #   rqIDs={}
#   #  rqIDs.update(EI.getRequirementIDs(analyseDeEntrant.sheets[user_inp['taskDetails'][0]['taskName']]))
#    # taskArch = user_inp['taskDetails'][0]['taskName'].split("_")[0]
#     #if taskArch == "F":
#
#      #   Arch = "BSI"
#     #else:
#      #   Arch = "VSM"
# #    for feps in rqIDs:
#   #      fepsNumber = (re.search("_[0-9]+", feps)).group()
#    #     logging.info("fepsNumber-------------------->", fepsNumber)
#
#     #    for reqname in dict(rqIDs[feps]):
#       #      if reqname == 'Evolved Requirements':
#      #           rqList = rqIDs[feps][reqname]
#       #          for req in rqList:
# #                    evolReq(tpBook, fepsNumber, macro, Arch, feps, rqIDs, reqname, req)
#      #               time.sleep(2)
#                     #for j,k,l  in req_ver_sf['req'],req_ver_sf['sf_sheet'],req_ver_sf['req_comment']:
#       #              evolReq(tpBook, fepsNumber, macro, Arch, feps, rqIDs, reqname, req)
#
#
#
#
#
#                             # ts_modify_reqver(tpBook.sheets[sheet],ver)
#
#     # if tpBook.sheets[sheet] is not None:
#     # ts_modify(tpBook.sheets[sheet],value)
#     logging.info("==============")
#     TPM.selectSynthUpdate(macro)
#     time.sleep(3)
#     sfBook.save()
#
#     tpBook.activate()
#     TPM.synchronizeSubSheet(macro)
#     time.sleep(2)
#     # tpBook.save()
#  #  time.sleep(3)
#     EI.activateSheet(tpBook, tpBook.sheets['Impact'])
#    # tpBook.activate()
#
#     # TPM.selectTPInit(macro)
#     # time.sleep(1)
#     TPM.selectTPImpact(macro)
#     time.sleep(1)
#     # TPM.selectSynthUpdate(macro)
#
