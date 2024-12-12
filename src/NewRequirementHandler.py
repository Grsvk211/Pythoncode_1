import ExcelInterface as EI
import time
import TestPlanMacros as TPM
import InputConfigParser as ICP
import re
from Backlog_Handler import getCombinedThematicLines
import os, sys
import WordDocInterface as WDI
import AnaLyseThematics as AT
import ParseThematics as parse
import Backlog_Handler as BLH
import BusinessLogic as BL
from lexer import Lexer
from thmParser import Parser
import DTC_Frame_Wired as DFW
import logging




def createCombination(data):
    lexer = Lexer().get_lexer()
    tokens = lexer.lex(data)
    '''
    for token in tokens:
        logging.info(token)'''

    pg = Parser()
    logging.info("parser obj created")
    pg.parse()
    logging.info("Calling parse method")
    parser = pg.get_parser()
    logging.info("parser built - - ")
    combinations = parser.parse(tokens).eval()
    logging.info("combinations", combinations)
    return combinations



def fillImpactNewReq(tpBook, new_req, fepsNum, flag, Arch, rqIDs, fepsForDuplicateReqs):
    logging.info("fillImpactNewReq---fepsForDuplicateReqs-------->",fepsForDuplicateReqs)
    impact_result = {
        'status': 1,
        'testSheetList': []

    }
    logging.info("Fill impact new req function called...................")
    try:
        maxrow = tpBook.sheets['Impact'].range('A' + str(tpBook.sheets['Impact'].cells.last_cell.row)).end('up').row
        sheet = tpBook.sheets['Impact']
        macro = EI.getTestPlanAutomationMacro()
        i = 0
        length = 1
        time.sleep(2)
        EI.activateSheet(tpBook, 'Impact')
        if new_req.find('(') != -1:
            oldReq = "".join(new_req.split())
        logging.info("++++++++Value of ROW = " + str(maxrow))
        logging.info("++++++++Value of length = " + str(length))
        condition = maxrow + 1

        colOfRequirement, colOfVer, colOfFT, colOfComment, colOfFeps = 1, 2, 4, 5, 6

        if new_req.find('(') != -1:
            new_reqName = new_req.split("(")[0].split()[0] if len(new_req.split("(")) > 0 else ""
            new_reqVer = new_req.split("(")[1].split(")")[0] if len(new_req.split("(")) > 1 else ""
            # time.sleep(1)
        else:
            new_reqName = new_req.split()[0] if len(new_req.split()) > 0 else ""
            new_reqVer = new_req.split()[1] if len(new_req.split()) > 1 else ""

        logging.info("new_reqVer = ", new_reqVer)
        impact_result.update({"new_reqVer": new_reqVer})
        logging.info("new_reqName = ", new_reqName)
        impact_result.update({"new_reqName": new_reqName})
        logging.info('impact_result--->', impact_result)
        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfRequirement), new_reqName)
        EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfVer), new_reqVer)
        if fepsForDuplicateReqs:
            BL.addDuplicateFEPS(tpBook, maxrow, colOfFeps, fepsNum, fepsForDuplicateReqs, rqIDs, new_reqName, new_reqVer)
            
        else:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, colOfFeps), fepsNum[1:])

        if flag == -1:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, 5),
                               "The thematic lines of the requirement are NA for " + Arch + ".\nProceed Manually.")
        elif flag == -2:
            EI.setDataFromCell(tpBook.sheets['Impact'], (maxrow + 1, 5),
                               "Input doc is not in the correct format.\n Proceed Manually.")

        # Run impact
        time.sleep(2)
        TPM.selectTPImpact(macro)
        time.sleep(10)
        if sheet.range(maxrow + 1, colOfFT).value is not None and sheet.range(maxrow + 1, colOfFT).value != "":
            testSheetList = sheet.range(maxrow + 1, colOfFT).value.split("\n")
            impact_result['status'] = 2
            impact_result['testSheetList'] = testSheetList
            REQ_COL = 1
            VER_COL = 2
            FEPS_COL = 6
            FT_COL = 5
            reqSearchResult = EI.searchDataInSpecificRows(tpBook.sheets['Impact'], (18, 100), REQ_COL, new_reqName)
            if reqSearchResult['count'] > 0:
                for cellPos in reqSearchResult['cellPositions']:
                    x, y = cellPos
                    EI.setDataFromCell(tpBook.sheets['Impact'], (x, REQ_COL), '')
                    EI.setDataFromCell(tpBook.sheets['Impact'], (x, VER_COL), '')
                    EI.setDataFromCell(tpBook.sheets['Impact'], (x, FEPS_COL), '')
                    EI.setDataFromCell(tpBook.sheets['Impact'], (x, FT_COL), '')
    except Exception as exp:
        logging.info(f"Error Occured Impact: [{exp}]")
        impact_result['status'] = -1

    return impact_result


def getReqVer(req):
    if req.find('(') != -1:
        new_reqName = req.split("(")[0].split()[0] if len(req.split("(")) > 0 else ""
        new_reqVer = req.split("(")[1].split(")")[0] if len(req.split("(")) > 1 else ""
    else:
        new_reqName = req.split()[0] if len(req.split()) > 0 else ""
        new_reqVer = req.split()[1] if len(req.split()) > 1 else ""
    return new_reqName.strip(), new_reqVer.strip()


def combineValues(sheet, combine_str, col, seperator):
    split_value = sheet.range(col).value.split()
    final_res = ""
    if len(split_value) > 0:
        split_value.append(combine_str)
        final_res = seperator.join(split_value)
    return final_res


# already present in BL
def getArch(taskname):
    arch = "VSM"
    x = re.findall("^F_", taskname)
    if x:
        arch = "BSI"
    return arch


def getKPIDocPath(path):
    docList = []
    documents = os.listdir(path)
    logging.info("--", documents)
    for d in documents:
        a = (path + "\\" + d)
        if d.find("") != -1:
            docList.append(a)
    return docList


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


# already present in BL end


def treatBacklog(testSheet):
    refEC = EI.openReferentialEC()
    kpiDocList = getKPIDocPath(ICP.getInputFolder() + "/KPI")
    # rawReqs = analyseThematics.getTsRawReq()
    rawReqs = testSheet.range('C4').value
    ts_reqs = [rawReqs]
    currArch = getArch(ICP.FetchTaskName())
    logging.info("*** kpiDocList ***", kpiDocList)
    logging.info("*** reqs ***", ts_reqs)
    logging.info("*** refEC ***", str(refEC))
    logging.info("*** currArch ***", currArch)
    removeInterfaceReqs = removeInterfaceReq(ts_reqs)
    status, combinedThemLines = getCombinedThematicLines(kpiDocList, removeInterfaceReqs,
                                                         refEC, currArch)
    refEC.close()
    logging.info("***Thematic Combinations = ***", status, combinedThemLines)
    return status, combinedThemLines

def add_initial_and_retour_steps_Normal(sheet, macro):
    keywords = [('---- CONDITIONS INITIALES ----', 1), ('---- RETOUR AUX CONDITIONS INITIALES ----', 3)]
    # initial_cell_pos = EI.searchDataInExcel(sheet, "", '---- CONDITIONS INITIALES ----')
    # retour_cell_pos = EI.searchDataInExcel(sheet, "", ' ---- RETOUR AUX CONDITIONS INITIALES ----')\
    sheet_value = sheet.used_range.value
    for keyword, num in keywords:
        # keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        keyword_cell_pos = EI.searchDataInExcelCache(sheet_value, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        if keyword_cell_pos['count'] > 0:
            for cellPos in keyword_cell_pos['cellPositions']:
                row, col = cellPos
                if num == 1:
                    logging.info("\n\nrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr")
                    TPM.addInitialContionsStep(macro)
                elif num == 3:
                    TPM.addRetourContionsStep(macro)
        # new_val_cel_pos = EI.searchDataInExcel(sheet, "", "PARAMETRE D'ENTREE")
        sheet_value = sheet.used_range.value
        new_val_cel_pos = EI.searchDataInExcelCache(sheet_value, "", "PARAMETRE D'ENTREE")
        logging.info("new_val_cel_pos ", new_val_cel_pos)
        if new_val_cel_pos['count'] > 0:
            for cellPos_CI in new_val_cel_pos['cellPositions']:
                logging.info("cellPos_CI --> ", cellPos_CI)
                row_CI, col_CI = cellPos_CI
                if sheet.range(row_CI + 1, 5).value == "" or sheet.range(row_CI + 1, 5).value is None:
                    if num == 1:
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI), "$ETAT_PRINCIP_SEV")
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI + 1), "CONTACT")
                        EI.setDataFromCell(sheet, (row_CI - 1, 1), "BUT DE L'ETAPE : Put on Contact")
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI - 1), "FONCTION")
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI - 2), 'Put on Contact')
                    elif num == 3:
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI), "$ETAT_PRINCIP_SEV")
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI + 1), "ARRET")
                        EI.setDataFromCell(sheet, (row_CI - 1, 1), "BUT DE L'ETAPE : Put on Arret")
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI - 1), "FONCTION")
                        EI.setDataFromCell(sheet, (row_CI + 1, col_CI - 2), 'Put on Arret')


def add_initial_and_retour_steps(sheet, macro, req, flow, frame, ckt, defectCodeDNFKPI, reqData):
    keywords = [('---- CONDITIONS INITIALES ----', 1),  ('---- CORPS DE TEST ----', 2), ('---- RETOUR AUX CONDITIONS INITIALES ----', 3)]
    flowflag = 0
    frameflag = 0
    wireflag = 0
    # globalDci = EI.openGlobalDCI()
    try:
        flows, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, req_frames, identifier, dtc = DFW.getFlows(req, ckt, defectCodeDNFKPI, reqData)
        globalDci = EI.openGlobalDCI()
        mux_sheet = globalDci.sheets['MUX']
        mux_sheet.activate()
        # if flows:
        #     flow = flows[0]
        if flow:
            # flow = flows[0]
            logging.info("dtc flows--->", dtc, flow)
            # flowPresent = EI.searchDataInCol(mux_sheet, 3, flow)

            sheet_value = mux_sheet.used_range.value
            flowPresent = EI.searchDataInColCache(sheet_value, 3, flow)

            logging.info("flowPresent_keyword_rows-->", flowPresent['cellPositions'])
            if flowPresent['cellPositions']:
                flowflag = 1
        # elif req_frames:
        elif frame:
            # frame = req_frames[0]
            logging.info("dtc frame--->", dtc, frame)
            # framePresent = EI.searchDataInCol(mux_sheet, 9, frame)

            sheet_value = mux_sheet.used_range.value
            framePresent = EI.searchDataInColCache(sheet_value, 9, frame)

            logging.info("framePresent_keyword_rows-->", framePresent['cellPositions'])
            if framePresent['cellPositions']:
                frameflag = 1
        globalDci.close()
        if dtc:
            logging.info("after dtc")
            logging.info("flowflag and frameflag ---->", flowflag, frameflag)
            if flowflag == 1 or frameflag == 1:
                # tpBook.activate()
                sheet.activate()
                if req_frames and flows:
                    flowArr, flowArr_E_Col, flowArr_I_Col = DFW.DTCFlowforTP(flow, defectCode, req_frames, identifier)
                    logging.info("flows-->", flows)
                    logging.info("flowArr happy-->", flowArr)
                    logging.info("rew framesdf----.", req_frames)
                    DFW.rows(keywords, sheet, flow, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, req, req_frames)
            if frameflag == 1:
                # frame_list, identifier = DFW.getFlowFrame(frame, flow)
                identifier = DFW.getIdentifier(frame)
                logging.info("req_frames-->", frame)

                logging.info("hggggg")
                flowArr, flowArr_E_Col, flowArr_I_Col = DFW.DTCFrameforTP(frame, defectCode, identifier)
                # frame--> FD8_DYN_VOL_03F
                DFW.rows(keywords, sheet, flows, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, req, frame)
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"flow or dtc_pattern is not present in the MUX sheet in DCI Global{ex}{exc_tb.tb_lineno}")
        pass
    # flows = getFlow(req)
    # flow = flows[0]
    # wireSignals = ["MANAGE_LED_ZEV"]

    try:
        globalDci = EI.openGlobalDCI()
        logging.info("wire flow--->", flow)
        filare_sheet = globalDci.sheets['FILAIRE']
        filare_sheet.activate()
        # sheet = globalDci.sheets['FILAIRE']
        # wiredFlowsPresent = EI.searchDataInCol(filare_sheet, 3, flow)

        sheet_value = filare_sheet.used_range.value
        wiredFlowsPresent = EI.searchDataInColCache(sheet_value, 3, flow)

        logging.info("wiredFlowsPresent-->", wiredFlowsPresent['cellPositions'])
        if wiredFlowsPresent['cellPositions']:
            wireflag = 1
        globalDci.close()
        # if wiredFlowsPresent['cellPositions']:
        if wireflag == 1:
            # tpBook.activate()
            sheet.activate()
            circuitArr, flowArr_E_Col, WRArr_I_Col, circuit = DFW.getCircuits(req, flow, reqData)
            logging.info("circuitArr, flowArr_E_Col, WRArr_I_Col, circuit--->", circuitArr, flowArr_E_Col, WRArr_I_Col,
                  circuit)
            # circuitArr, flowArr_E_Col, WRArr_I_Col, circuit = DFW.getCircuits(req)
            DFW.WErows(keywords, sheet, WRArr_I_Col, circuitArr, flow, ckt, flow)

    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"wireflow or dtc_pattern is not present in the FILAIRE sheet in DCI Global{ex}")

def getDocVer(inputdoc):
    currVer = re.search("([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})", inputdoc)
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
    return currVer


def grepThematicsCode(rawThematics):
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
    return reducedThm

def getRawThematic(thematic_data):
    logging.info(f"Finding thematic architechture...... {thematic_data}")
    rawThem = ""
    if type(thematic_data) is dict:
    # if type(thematic_data) is not str:
        if thematic_data != -1 and thematic_data != -2:
            if thematic_data['effectivity'] != "":
                rawThem = thematic_data['effectivity']
            elif thematic_data['lcdv'] != "":
                rawThem = thematic_data['lcdv']
            elif thematic_data['diversity'] != "":
                rawThem = thematic_data['diversity']
            elif thematic_data['target'] != "":
                rawThem = thematic_data['target']
    return rawThem

def getReqThematic(ReqName, ReqVer, rqIDs, feps):
    reqThematic = ""
    rqTable = ""
    TableList = ""
    rawThm = ""
    newThemFormat = ""
    thematicLines = []
    for inpFile in rqIDs[feps]['Input_Docs']:
        logging.info("inputdoc --1234 ", inpFile)
        currVer = re.search("([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})", inpFile)
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
            logging.info("No Version found in inputdoc", inpFile)
        logging.info("Current Version = ", currVer)
        currDoc = BL.getDocPath(inpFile, currVer)
        logging.info(f"-------------->currDoc {currDoc} <----------------")
        if type(currDoc) is str:
            logging.info(f"\nDoc Name: {inpFile}")
            reqContent = WDI.getReqContent(currDoc, ReqName, ReqVer)
            logging.info(f"reqContent THM {reqContent}")
            if reqContent != -1:
                rawThm = getRawThematic(reqContent)
            # TableList = WDI.getTables(currDoc)
            # # logging.info("TableList --> ", TableList)
            # rqTable = WDI.threading_findTable(TableList, ReqName)
            # if rqTable == -1:
            #     if (ReqName.find('.') != -1):
            #         ReqName = ReqName.replace('.', '-')
            #     if (ReqName.find('_') != -1):
            #         ReqName = ReqName.replace('_', '-')
            #
            #     # if (reqName.find('.') != -1):
            #     #     reqName = reqName.replace('.', '-')
            #     # if (reqName.find('_') != -1):
            #     #     reqName = reqName.replace('_', '-')
            #
            #     rqTable = WDI.threading_findTable(TableList, ReqName + "(" + ReqVer + ")")
            #     if rqTable != -1:
            #         logging.info("aaaaaaaa")
            #         rawThm = WDI.getThematics(rqTable, ReqName + "(" + ReqVer + ")")
            # else:
            #     rqTable = WDI.threading_findTable(TableList, ReqName + " " + ReqVer)
            #     if rqTable != -1:
            #         logging.info("eeeeeeee")
            #         rawThm = WDI.getThematics(rqTable, ReqName + " " + ReqVer)
                logging.info(f"\nrawThmrawThm {rawThm}")
                if rawThm != -1 and rawThm != "" and rawThm is not None:
                    newThemFormat = grepThematicsCode(rawThm)
                    if newThemFormat.find("( )") != -1:
                        newThemFormat = newThemFormat.replace("( )", "")
                    them_pat = ['AND )', 'AND)', 'OR )', 'OR)']
                    logging.info(f"\nnewThemFormat BF {newThemFormat}")
                    for thmpat in them_pat:
                        if newThemFormat.endswith(thmpat):
                            newThemFormat = newThemFormat.replace(thmpat, " )")
                    logging.info(f"\nnewThemFormat AF {newThemFormat}")
                    rawCombinations = createCombination(newThemFormat)
                    combList = rawCombinations.split('\n')
                    combList = list(set(combList))
                    logging.info("combList ", combList)
                    contraryThem = BLH.removeContrary([combList])
                    logging.info("contraryThem --> ", contraryThem)
                    refEC = EI.openReferentialEC()
                    for themline in contraryThem:
                        for themLine in themline:
                            finalThem = BLH.filterThemForArch(themLine, refEC, BL.getArch(ICP.FetchTaskName()))
                            logging.info("finalThem --> ", finalThem)
                            if finalThem != -1:
                                thematicLines.append(finalThem)
                    logging.info("rqTable --> ", rqTable)
                    logging.info("rawThm --> ", rawThm)
                    logging.info("newThemFormat --> ", newThemFormat)
                    logging.info("rawCombinations --> ", rawCombinations)
                    logging.info("thematicLines Final --> ", thematicLines)
                    refEC.close()
    return thematicLines


def fill_FT(tpBook, sheet_name, req, macro, rqIDs, feps, dtc, flow, frame, ckt, defectCodeDNFKPI, reqData):
    logging.info("Fill FT for new req function called...................")
    # Need to implement function to get the C2 and C3 value
    logging.info("sheet_name --> ", sheet_name)
    sheet_name_new = ""
    thematicLines = []
    for ts in tpBook.sheets:
        # logging.info(ts)
        if re.search("^OLD_", ts.name.upper()):
            ts_name = re.sub("^OLD_", "", ts.name.upper())
            if ts_name == sheet_name:
                logging.info("ts_name --> ", ts_name)
                sheet_name_new = ''.join(('Old_', sheet_name))
                break
    logging.info("sheet_name new : ", sheet_name_new)
    time.sleep(5)
    EI.activateSheet(tpBook, sheet_name.strip())
    # tpBook.sheets[sheet_name_new.strip()].activate()
    shortDESC_col, briefDESC_col, reqs_col, type_col, categorie_col, ponderation_col, trigram, History = 'C2', 'C3', 'C4', 'E5', 'C6', 'E6', 'D17', 'C17'
    shortDESC_VAL = 'Sample short description'
    briefDESC_VAL = 'Sample brief description'
    logging.info("reqreq --> ", req)
    new_reqName, new_reqVer = getReqVer(req)
    logging.info("new_reqVer1 = ", new_reqVer)
    logging.info("new_reqName1 = ", new_reqName)

    req = new_reqName.strip() + "(" + str(new_reqVer) + ")"
    if tpBook.sheets[sheet_name].range(reqs_col).value != "" and tpBook.sheets[sheet_name].range(
            reqs_col).value is not None:
        # combine the new req with existing req in C4 column
        combine_req = combineValues(tpBook.sheets[sheet_name], req, reqs_col, '|')
    else:
        combine_req = req

    thematicLines = getReqThematic(new_reqName, new_reqVer, rqIDs, feps)
    # thematicLines.append(thematics)
    logging.info("thematicLines ", thematicLines)
    logging.info("combine_req ", combine_req)
    # exit()

    try:
        # set value in cell C2
        EI.setDataInCell(tpBook.sheets[sheet_name], shortDESC_col, shortDESC_VAL)
        # set value in cell C3
        EI.setDataInCell(tpBook.sheets[sheet_name], briefDESC_col, briefDESC_VAL)
        # set value in cell C4
        EI.setDataInCell(tpBook.sheets[sheet_name], reqs_col, combine_req)
        # set value in cell E5
        EI.setDataInCell(tpBook.sheets[sheet_name], type_col, 'N1')
        if ckt:
            # set value in cell C6
            EI.setDataInCell(tpBook.sheets[sheet_name], categorie_col, 'MANU')
        else:
            # set value in cell C6
            EI.setDataInCell(tpBook.sheets[sheet_name], categorie_col, 'AUTO')
        # set value in cell E6
        EI.setDataInCell(tpBook.sheets[sheet_name], ponderation_col, 'P2')
        # set value to cell D17
        EI.setDataFromCell(tpBook.sheets[sheet_name], trigram, BL.getTrigram())
        # set value to cell D17
        # EI.setDataFromCell(tpBook.sheets[sheet_name], History, 'Created New Sheet')

        # status, thematicLines = treatBacklog(tpBook.sheets[sheet_name])
        # logging.info("thematicLines --> ", thematicLines)
        if thematicLines:
            for ind, them_line in enumerate(thematicLines):
                logging.info("them_line --> ", them_line)
                if ind > 0:
                    TPM.addThematique(macro)
                    time.sleep(5)
                EI.setDataFromCell(tpBook.sheets[sheet_name], (8, 3), them_line)
        if dtc:
            add_initial_and_retour_steps(tpBook.sheets[sheet_name], macro, req, flow, frame, ckt, defectCodeDNFKPI, reqData)
        else:
            add_initial_and_retour_steps_Normal(tpBook.sheets[sheet_name], macro)
        # exit()
    except Exception as e:
        logging.info(f"+++++++++++++ Error: {e} +++++++++++++")
        return -1

    return 1



if __name__ == "__main__":
    ICP.loadConfig()

    tpBook = EI.openTestPlan()
    thematic_result = EI.searchDataInExcel(tpBook.sheets['VSM20_N1_20_90_0044'], (1, 100), "THEMATIQUE")
    thmLineBefore = tpBook.sheets['VSM20_N1_20_90_0044'].range(8, 1).value
    logging.info(thematic_result)
    exit()
# ICP.loadConfig()
# tpBook = EI.openTestPlan()
# macro = EI.getTestPlanAutomationMacro()
# TPM.selectTestSheetAdd(macro)
# time.sleep(5)
# logging.info("tpBook.sheets.active ", tpBook.sheets.active)
# created_FT = tpBook.sheets.active
# fill_FT(tpBook, "BSI04_N1_02_54_0112", "REQ-0778501 A")


# getReqThematic('REQ-0778501', 'A')
