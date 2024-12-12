import time
from web_interface import startDocumentDownload
import ExcelInterface as EI
import re
import InputConfigParser as ICF
import TestPlanMacros as TPM
import os, sys
from collections import OrderedDict   #  used for to remove duplicates in the circuit
import logging

# These function is used for to get the flows
def getFlow(new_req, reqData):
    flows = ""
    frame = ""
    dtc = ""
    try:
        if reqData['flow'] != None and reqData['flow'] != "":
            flows = reqData['flow']
            logging.info("flows--->", flows)
        else:
            logging.info("Flow is not present in the requirement")
            pass
        if reqData['frame'] != None and reqData['frame'] != "":
            frame = reqData['frame']
            logging.info("frame--->", frame)
    except:
        logging.info("flows for the frame and wired not present in the requirement")

    try:
        celltext = reqData['celltext']
        dtc_pattern = r"record a DTC"
        match = re.search(dtc_pattern, celltext)
        if match:
            dtc = match.group()
            logging.info("Extracted DTC:", dtc)
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"record a DTC' keyword is not present in the requirement{ex}{exc_tb.tb_lineno}")
        pass
    return flows, frame, dtc

def getFrame(new_req, reqData):
    flowframe = ""
    frames = ""
    dtc = ""
    if reqData['frame'] != None and reqData['frame'] != "":
        frames = reqData['frame']
        logging.info("frame--->", frames)
    if reqData['flowframe'] != None and reqData['flowframe'] != "":
        flowframe = reqData['flowframe']
        logging.info("'flowframe--->", flowframe)
    try:
        celltext = reqData['celltext']
        dtc_pattern = r"record a DTC"
        match = re.search(dtc_pattern, celltext)
        if match:
            dtc = match.group()
            logging.info("Extracted DTC:", dtc)
    except:
        logging.info("'record a DTC' keyword is not present in the requirement")
    return frames, dtc, flowframe

# These function is used for to get the flowFrame
def getFlowFrame(flowframe, flows):
    # if Length of Flowframe is less than 3 it is network.
    frame_list = []
    if (flowframe is not None and flowframe != '') and (flows is not None and flows != ''):
        if len(flowframe) <= 3:
            # frame_list = []
            logging.info("less than 3")
            logging.info(1)
            globalDci = EI.openGlobalDCI()
            sheet = globalDci.sheets['MUX']
            for i, flow in enumerate(flows):
                keyword1 = flow
                keyword2 = 'VSM'
                keyword3 = 'CAN_' + flowframe
                # first_keyword_rows = EI.searchDataInCol(sheet, 3, keyword1)
                # logging.info("first_keyword_rows123-->", first_keyword_rows['cellPositions'])
                sheet_value = sheet.used_range.value
                first_keyword_rows = EI.searchDataInColCache(sheet_value, 3, keyword1)

                second_keyword_rows = []
                for row, col in first_keyword_rows['cellPositions']:
                    if sheet.range(f'B{row}').value == keyword2:
                        second_keyword_rows.append(row)
                logging.info(second_keyword_rows)

                third_keyword_rows = []
                for row in second_keyword_rows:
                    if sheet.range(f'J{row}').value == keyword3:
                        third_keyword_rows.append(row)
                logging.info(third_keyword_rows)
                col = 0
                value = EI.getDataFromCell(sheet, (third_keyword_rows[0], col + 9))
                # value---> 1 FD8_DYN_VOL_03F output
                logging.info("value--->", i, value)
                frame_list.append(value)
            # Frame---> FD8_DYN_VOL_03F output
            logging.info("Frame--->", frame_list[0])
            globalDci.close()
        else:
            logging.info("wwwwwxxxxxxx")
            # frame_list = []
            # # if Length of Flowframe is greater than 3 it is Frame.
            logging.info("greater than 3")
            value = flowframe[0]
            framelist = flowframe
            # frame_list.append(value)
            # # frame_list = str(flowframe)
            logging.info("Frame1--->", value)
            logging.info("Frame2  frame_list --->", framelist)
            frame_list.append(framelist)
            logging.info("frame_list------====>",frame_list)
    try:
        MagnetoFrames = EI.findInputFiles()[11]
        path1 = ICF.getInputFolder() + "\\" + MagnetoFrames
        isMagnetoFramesExist = os.path.isfile(path1)
        if isMagnetoFramesExist:
            MagnetoFrame = EI.openExcel(ICF.getInputFolder() + "\\" + MagnetoFrames)
            if MagnetoFrame:
                Magnetosheet = MagnetoFrame.sheets["FRAMES_DEFINITIONS"]
                # value = EI.searchDataInCol(Magnetosheet, 4, frame_list[0])

                sheet_value = Magnetosheet.used_range.value
                value = EI.searchDataInColCache(sheet_value, 4, frame_list[0])

                logging.info("first_keyword_rows-->", value['cellPositions'])
                row, col = value['cellPositions'][0]
                time.sleep(2)
                network = EI.getDataFromCell(Magnetosheet, (row, col - 3))
                ID = EI.getDataFromCell(Magnetosheet, (row, col + 1))
                identifier = network + "_" + "ID" + ID
                logging.info("identifier--->", identifier)
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"magento file is not exists{ex}{exc_tb.tb_lineno}")
    MagnetoFrame.close()
    return [frame_list[0], identifier]


# These function is used for to get the flows, frame, defect code. If Defect code not present we will go to getDefectCode(flows) function
def getFlows(new_req, ckt, defectCodeDNFKPI, reqData):
    flowArr = ""
    flowArr_E_Col = ""
    flowArr_I_Col = ""
    req_frames = ""
    defectCode = ""
    identifier = ""
    try:
        # reqData['comment'] = "@DNF RCTA-179 => VSM-U0131-87"
        logging.info("reqData['comment']--->", reqData['comment'])

        if reqData['flow'] != None and reqData['flow'] != "":
            flows = reqData['flow']
            logging.info("flows--->", flows)
        else:
            logging.info("Flow is not present in the requirement")
            pass

        DID_value = ""
        if re.search("VSM-[A-Z]{2}[0-9]{2}", reqData['comment']):
            extracted_DID = re.findall("VSM-[A-Z]{2}[0-9]{2}", reqData['comment'])
            if extracted_DID:
                logging.info("data1--->", extracted_DID)
                # data1---> ['VSM-U0131-87']
                logging.info("data2=--->", extracted_DID[0])
                # data2=---> VSM-U0131-87
                DID_value = extracted_DID[0]
                logging.info("DID_value1--->", DID_value)
        else:
            # VSM-U0131-81
            extracted_DID = re.findall("VSM-[A-Z]{1,2}[A-Z a-z 0-9]{1,10}-[0-9]{2}", reqData['comment'])
            logging.info("data1--->", extracted_DID)
            if extracted_DID:
                logging.info("data2=--->", extracted_DID[0])

        if extracted_DID and reqData['comment'] != "" and reqData['comment'] is not None:
            # reqData = 'DTC'
            # These loop will go if extracted_DID[0] keyword is present in content.extracted_DID[0] means DTC code
            # example for DTC code is VSM-U0131-87
            if extracted_DID[0] in reqData['comment']:
                defectCode = extracted_DID[0].replace("VSM-", "")
                logging.info("defectCode1--->", defectCode)
        else:
            defectCode = defectCodeDNFKPI
            logging.info("defectCodeDNFKPI--->", defectCode)
    except Exception as ex:
        logging.info(f"Flow is not available for these requirement.please update manually. {ex}")
    pass
    flows = reqData['flow']
    flowframe = reqData['flowframe'] or reqData['frame']

    c = len(flowframe)
    logging.info("flowframe & len of flowframe--->", flowframe, c)
    logging.info("hihihi123")
    logging.info("flowframe and flows ", flowframe[0], flows)
    logging.info("flowframe and flows ", flowframe, flows)
    # if (flowframe is not None and flowframe != '') and (flows is not None and flows != ''):
    try:
        celltext = reqData['celltext']
        dtc_pattern = r"record a DTC"
        match = re.search(dtc_pattern, celltext)
        if match:
            dtc = match.group()
            logging.info("Extracted DTC:", dtc)
        if dtc and (ckt is None or ckt == ''):
            if flowframe and flows is not None and flows != '':
                logging.info("hih")
                logging.info("flowframe--->", flowframe)
                req_frames, identifier = getFlowFrame(flowframe, flows)
            elif flowframe and (flows.strip() is None or flows.strip() == ''):
                logging.info("hihih")
                req_frames, identifier = getFlowFrame(flowframe, "")
    except Exception as ex:
        logging.info(f"DTC not available for these requirement.please update manually. {ex}")
    # req_frames = ["FD8_DYN_VOL_03F"]
    req_frames = req_frames
    logging.info("dtc--->", dtc)
    if dtc:
        flowArr, flowArr_E_Col, flowArr_I_Col = DTCFrameforTP(req_frames, defectCode, identifier)
    logging.info("flows, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, req_frames, identifier, dtc-->", flows, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, req_frames, identifier, dtc)
    return [flows, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, req_frames, identifier, dtc]

def getCircuits(new_req, flow, reqData):
    circuit = ''
    try:
        logging.info("reqData['circuit']--->", reqData['circuit'])
        # reqData['circuit']---> ['short circuit to ground', 'open circuit or short circuit to plus']
        b = reqData['circuit']
        circuit_list = []
        # These for loop will seperate the or, and present in the list of b
        # example b = ['short circuit to ground', 'open circuit', 'short circuit to plus']
        # we will get the output as ['short circuit to ground', 'open circuit',short circuit to plus']
        for item in b:
            if ' or ' in item:
                new_items = item.split(' or ')
                circuit_list.extend(new_items)
            elif ' and ' in item:
                new_items = item.split(' and ')
                circuit_list.extend(new_items)
            else:
                circuit_list.append(item)

        circuit_list = list(OrderedDict.fromkeys(circuit_list))
        circuit = circuit_list
        logging.info("circuit_list----->", circuit)
        # 'flow': ['MANAGE_LED_ZEV']
        # wiredFlows = reqData['flow']

    except Exception as ex:
        logging.info(f"circuit is not available for these requirement.please update manually. {ex}")
        pass
    logging.info("flow12345--->",flow)
    flows = [flow]
    logging.info("flows12345--->", flow)
    circuitArr, flowArr_E_Col, WRArr_I_Col = DTCWIreTP(flows)
    logging.info("12345678--->",circuitArr)
    return [circuitArr, flowArr_E_Col, WRArr_I_Col, circuit]

def getDefectCode(flows,reqData):
    # UI_task_nameEI F_PARAM_06_02_2023
    UI_task_nameEI = ICF.FetchTaskName()
    taskname = UI_task_nameEI.split('_')[1]
    logging.info("taskname------>", taskname)
    defectCode = ""
    dnfflag = 0
    try:
        comment = reqData['comment']

        # Extract the next word after "DNF" in the comment
        next_word_match = re.search(r'DNF\s+(\w+)', comment)
        if next_word_match:
            next_word = next_word_match.group(1)
        else:
            next_word = None

        # Extract the text after ":" in the next word
        # Define the regular expression pattern
        pattern = r"[A-Z]{3,4}-[0-9]{3,4}"

        # Extract the values using regular expressions
        function_list_match = re.findall(pattern, comment)
        if function_list_match:
            function_list = function_list_match[0]
        else:
            function_list = None

        logging.info('new_word_list:', next_word)
        logging.info('Function list:', function_list)
        taskname1 = function_list.split("-")[0]
        logging.info('substring----->',taskname1)
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"comment is not present in the table.{ex}{exc_tb.tb_lineno}")

    # AnalyseDNF_KPI = EI.openExcel(r'C:\Users\vgajula\Documents\08-05-2023\F_PARAM_06_02_2023\Input_Files\Analyse_DNF_kpi.xlsx')
    try:
        AnalyseDNF_KPI = EI.openAnaDNF()
        AnalyseDNF_sheet = AnalyseDNF_KPI.sheets['Synthesis']
        AnalyseDNF_sheet.activate()
        if taskname in comment or next_word in comment:
            if taskname:
                try:
                    # defectCodePresent = EI.searchDataInCol(AnalyseDNF_sheet, 1, taskname)

                    sheet_value = AnalyseDNF_sheet.used_range.value
                    defectCodePresent = EI.searchDataInColCache(sheet_value, 1, taskname)

                    logging.info("defectCodePresent--->",defectCodePresent,taskname)
                    x, y = defectCodePresent['cellPositions'][0]
                    referencenumber = EI.getDataFromCell(AnalyseDNF_sheet, (x, y + 1))
                    logging.info("taskname_referencenumber--->", referencenumber)
                except:
                    logging.info("task name is not present in the Analysis DNF Kpi")
                    pass
            if next_word:
                try:
                    # defectCodePresent = EI.searchDataInCol(AnalyseDNF_sheet, 1, next_word)

                    sheet_value = AnalyseDNF_sheet.used_range.value
                    defectCodePresent = EI.searchDataInColCache(sheet_value, 1, next_word)

                    logging.info("defectCodePresent--->", defectCodePresent, next_word)
                    x, y = defectCodePresent['cellPositions'][0]
                    referencenumber = EI.getDataFromCell(AnalyseDNF_sheet, (x, y + 1))
                    logging.info("next_word_referencenumber--->", referencenumber)
                except:
                    logging.info("next_word is not present in the Analysis DNF Kpi")
                    pass
            AnalyseDNF_KPI.close()
            startDocumentDownload([[referencenumber, ""]])
            DNF_KPI = EI.openDNFKPI()
            try:
                if flows:
                    logging.info("222222222222222222222222222222222222222222222222")
                    DNF_sheet = DNF_KPI.sheets['DIAG NEED']
                    DNF_sheet.activate()
                    logging.info("flows--->", flows)
                    output_list = flows
                    b = '\n'.join(output_list)
                    logging.info(b)
                    logging.info("flows to search in the DNF KPI file--->", b)
                    # G column
                    # defectCodePresent = EI.searchDataInCol(DNF_sheet, 7, b)

                    sheet_value = DNF_sheet.used_range.value
                    defectCodePresent = EI.searchDataInColCache(sheet_value, 7, b)

                    # logging.info("flowPresent_keyword_rows-->", defectCodePresent['cellPositions'])
                    keyword = "VSM"

                    first_keyword_rows = []

                    # F column
                    for row, col in defectCodePresent['cellPositions']:
                        if DNF_sheet.range(f'F{row}').value == keyword:
                            first_keyword_rows.append(row)
                    logging.info(first_keyword_rows)

                    if first_keyword_rows:
                        dnfflag = 1
                    if dnfflag == 1:
                        x = first_keyword_rows[0]
                        logging.info("x--->", x)
                        y = 0
                        # K column
                        defectCode = EI.getDataFromCell(DNF_sheet, (x, y + 11))
                        logging.info("defectcode--->", defectCode)
                    DNF_KPI.close()
                elif function_list:
                    logging.info("111111111111111111111111111111111111111111111111111111111111111111111111")
                    DNF_sheet = DNF_KPI.sheets['DIAG NEED']
                    DNF_sheet.activate()
                    logging.info("flows--->", function_list)
                    # defectCodePresent = EI.searchDataInCol(DNF_sheet, 1, function_list)

                    sheet_value = DNF_sheet.used_range.value
                    defectCodePresent = EI.searchDataInColCache(sheet_value, 1, function_list)

                    keyword = "VSM"

                    first_keyword_rows = []

                    # F column
                    for row, col in defectCodePresent['cellPositions']:
                        if DNF_sheet.range(f'F{row}').value == keyword:
                            first_keyword_rows.append(row)
                    logging.info(first_keyword_rows)
                    if first_keyword_rows:
                        dnfflag = 1
                    if dnfflag == 1:
                        x = first_keyword_rows[0]
                        logging.info("x--->", x)
                        y = 0
                        # K column
                        defectCode = EI.getDataFromCell(DNF_sheet, (x, y + 11))
                        logging.info("defectcode--->", defectCode)
                    DNF_KPI.close()
                else:
                    logging.info("flows or next-word not present so we cannot find the defectcode. Please proceed manually")
                    DNF_KPI.close()
            except Exception as ex:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                logging.info(f"dtc_pattern or frame is not present{ex}{exc_tb.tb_lineno}")
                DNF_KPI.close()
        AnalyseDNF_KPI.close()
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"task name is not present in Analyses DNF KPI{ex}{exc_tb.tb_lineno}")

    return defectCode

# These function is used for to get the Identifier
def getIdentifier(frame):
    identifier = ''
    try:
        MagnetoFrames = EI.findInputFiles()[11]
        path1 = ICF.getInputFolder() + "\\" + MagnetoFrames
        isMagnetoFramesExist = os.path.isfile(path1)
        if isMagnetoFramesExist:
            MagnetoFrame = EI.openExcel(ICF.getInputFolder() + "\\" + MagnetoFrames)
            if MagnetoFrame:
                Magnetosheet = MagnetoFrame.sheets["FRAMES_DEFINITIONS"]
                # value = EI.searchDataInCol(Magnetosheet, 4, frame)

                sheet_value = Magnetosheet.used_range.value
                value = EI.searchDataInColCache(sheet_value, 4, frame)

                logging.info("first_keyword_rows-->", value['cellPositions'])
                row, col = value['cellPositions'][0]
                time.sleep(2)
                network = EI.getDataFromCell(Magnetosheet, (row, col - 3))
                ID = EI.getDataFromCell(Magnetosheet, (row, col + 1))
                identifier = network + "_" + "ID" + ID
                logging.info("identifier--->", identifier)
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"magento file is not exists{ex}{exc_tb.tb_lineno}")
    MagnetoFrame.close()
    return identifier


# These function is used to get the cellPositions in between the (CONDITIONS INITIALES, CORPS DE TEST,RETOUR AUX CONDITIONS INITIALES)
def getRows_FT(sheet, new_val_cel_pos):
    sheet_value = sheet.used_range.value
    # initial_cell_pos = EI.searchDataInExcel(sheet, "", '---- CORPS DE TEST ----')
    initial_cell_pos = EI.searchDataInExcelCache(sheet, "", '---- CORPS DE TEST ----')
    # retour_cell_pos = EI.searchDataInExcel(sheet, "", ' ---- RETOUR AUX CONDITIONS INITIALES ----')
    retour_cell_pos = EI.searchDataInExcelCache(sheet, "", ' ---- RETOUR AUX CONDITIONS INITIALES ----')
    start_position = initial_cell_pos['cellPositions'][0]
    end_position = retour_cell_pos['cellPositions'][0]
    logging.info("start_position & end_position--->", start_position, end_position)

    start_index = None
    end_index = None

    # Find the start and end indices
    for index, pos in enumerate(new_val_cel_pos['cellPositions']):
        if pos[0] >= start_position[0] and pos[0] <= end_position[0]:
            if pos[0] == start_position[0] and pos[1] < start_position[1]:
                continue
            if pos[0] == end_position[0] and pos[1] > end_position[1]:
                break
            if start_index is None:
                start_index = index
            end_index = index + 1

    # Extract the desired cell positions
    if start_index is not None and end_index is not None:
        cell_positions = new_val_cel_pos['cellPositions'][start_index:end_index]
        logging.info("rows-->", cell_positions)
    else:
        logging.info("No cell positions found between the specified positions.")
    return cell_positions


# These function is used to enter the data for the CONDITIONS INITIALES and RETOUR AUX CONDITIONS INITIALES in Excel sheet
def setDatainExcelSheet(sheet,keyword_cell_pos,num):
    new_val_cel_pos = ""
    cellPos_CI = ""
    macro = EI.getTestPlanAutomationMacro()
    try:
        if keyword_cell_pos['count'] > 0:
            for cellPos in keyword_cell_pos['cellPositions']:
                row, col = cellPos
                if num == 1:
                    logging.info("\n\nrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr")
                    TPM.addInitialContionsStep(macro)
                elif num ==2:
                    # UpdateHMIInfoCb(f'Number of flows present for these requirement are {len(flowArray)}')
                    logging.info("12312313245656")
                    for i in range(14):
                        TPM.addCorpDeTestStep(macro)
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
                EI.setDataFromCell(sheet, (row_CI - 1, col_CI - 4), "BUT DE L'ETAPE :Key put into Contact")
                logging.info("rr0--->", (row_CI - 1, col_CI - 4))
                EI.setDataFromCell(sheet, (row_CI + 1, col_CI), "$ETAT_PRINCIP_SEV")
                logging.info("rr1--->", (row_CI + 1, col_CI))
                EI.setDataFromCell(sheet, (row_CI + 1, col_CI + 1), "CONTACT")
                logging.info("rr1--->", (row_CI + 1, col_CI + 1))
                EI.setDataFromCell(sheet, (row_CI + 1, col_CI - 2), "Pass to CONTACT")
            elif num == 3:
                EI.setDataFromCell(sheet, (row_CI - 1, col_CI - 4), "BUT DE L'ETAPE :Key put into ARRET")
                logging.info("rrr0--->", (row_CI - 1, col_CI - 4))
                EI.setDataFromCell(sheet, (row_CI + 1, col_CI), "$ETAT_PRINCIP_SEV")
                logging.info("rrr1--->", (row_CI + 1, col_CI))
                EI.setDataFromCell(sheet, (row_CI + 1, col_CI + 1), "ARRET")
                logging.info("rrr2--->", (row_CI + 1, col_CI + 1))
                EI.setDataFromCell(sheet, (row_CI + 1, col_CI - 2), "Pass to ARRET")
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f"keyword_cell_pos is not present{ex}{exc_tb.tb_lineno}")
    return [new_val_cel_pos,cellPos_CI]


# These function used generate the flows & frames or frames only data to enter in rows
def rows(keywords, sheet, flow, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, req, frame):
    logging.info("rows frame--->", frame)
    flows = flow
    logging.info("flows happppppy--->",flow)
    macro = EI.getTestPlanAutomationMacro()
    sheet_value = sheet.used_range.value
    for keyword, num in keywords:
        # keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        keyword_cell_pos = EI.searchDataInExcelCache(sheet_value, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        new_val_cel_pos, cellPos_CI = setDatainExcelSheet(sheet, keyword_cell_pos, num)
        try:
            cellPosition = getRows_FT(sheet, new_val_cel_pos)
            logging.info("cellPositions34-->", cellPosition)
            if new_val_cel_pos['count'] > 0:
                logging.info("cellPos_CI00--> ", cellPos_CI)
                row_CI, col_CI = cellPos_CI
                # for flow in flows:
                if flow != None and flow != "":
                    if num == 2:
                        create_DTC_FF_FT(sheet, row_CI, num, cellPosition, defectCode, flowArr, flowArr_E_Col, flowArr_I_Col, flow, frame)
                if flows.strip() is None or flows.strip() == '':
                    logging.info("hfsjfdhsjdfh")
                    if num == 2:
                        create_DTC_FF_FT(sheet, row_CI, num, cellPosition, defectCode, flowArr, flowArr_E_Col,flowArr_I_Col, flows, frame)
        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            logging.info(f"\nError in cellpostioins: {ex} line: {exc_tb.tb_lineno}")

def create_DTC_FF_FT(sheet, row_CI, num, cellPositions, defectCode,flowArray,flowArray_E_Col,flowArray_I_Col,flow, frame):
    # req_frame = "FD8_DYN_VOL_03F"
    logging.info("create_DTC_FF_FT frame-->", frame)
    if sheet.range(row_CI + 1, 5).value == "" or sheet.range(row_CI + 1, 5).value is None:
        if num == 2:
            keyword1 = ["DIAG","NO_DIAG","DIAG","NO_DIAG","STOP_TRAME","DIAG","NO_DIAG","RELANCE_TRAME","DIAG","NO_DIAG","DIAG","NO_DIAG","DIAG","NO_DIAG"] # column D
            keyword2 = ["DIAG","","DIAG","","","DIAG","","","DIAG","","DIAG","","DIAG",""] # column I
            keyword3 = ["Clear all the DTC.", "Close the DIAG session","","Close the DIAG session", "Set to Stop the frame "+frame, "", "Close the DIAG session","Set to the Relance Trame", "", "Close the DIAG session","Clear all the DTC.",  "Close the DIAG session","", "Close the DIAG session"] # column C
            keyword4 = ["Clear all the "+defectCode+" using RAZ_JDD.","Close the DIAG Session.","Verify that the "+defectCode+" are absent by sending the read "+defectCode+" request.","Close the DIAG Session.","Stop the "+frame+" to create the fault  "+defectCode,"Verify that the Defect "+defectCode+" is present when we stop the frame "+frame,"Close the DIAG Session.","start the "+frame+" to heal the fault "+defectCode,"Verify that the Defect "+defectCode+" is disappeared when we start the frame "+frame,"Close the DIAG Session.","Clear all the "+defectCode+" using RAZ_JDD.","Close the DIAG Session.","Verify that the "+defectCode+" are absent by sending the read "+defectCode+" request.","Close the DIAG Session."]

            # Indices where the string should be inserted column C
            indices = [2, 5, 8, 12]
            # Iterate over the indices in reverse order and insert the string
            for i in reversed(indices):
                keyword3.pop(i)
                keyword3.insert(i, "Send the request flow to read the DTC")
            logging.info("keyword3-->", keyword3)

            logging.info("flowArrayK-Column--->", flowArray)
            flowArray.insert(0, '$RAZ_JDD_OK')
            flowArray.pop(10)
            flowArray.insert(10, '$RAZ_JDD_OK')
            logging.info("flowarrayK-Column--->", flowArray)

            flowArray_I_Col.pop(0)
            flowArray_I_Col.insert(0, 'All DTC Cleared.')
            flowArray_I_Col.pop(10)
            flowArray_I_Col.insert(10, 'All DTC Cleared.')

            # Indices where the string should be inserted column C
            indices = [1, 3, 4, 6, 7, 9, 11, 13]
            # Iterate over the indices in reverse order and insert the string
            for i in reversed(indices):
                flowArray_I_Col.pop(i)
                flowArray_I_Col.insert(i, "")
            logging.info("flowArray_I_Col1--->", flowArray_I_Col)

            logging.info("num2-->")
            rowslist = cellPositions
            # Increment row values by 1 and decrement column values by 1
            row_values = [(pos[0] + 1, pos[1] - 1) for pos in cellPositions]
            logging.info("row_values-->", row_values)
            # Extract the first element (row) from each tuple
            rows = [row for row, _ in row_values]
            logging.info("rows-->", rows)

            # Increment row values by 1 and decrement column values by 1
            row_values = [(pos[0] - 1, pos[1] - 1) for pos in cellPositions]
            logging.info("row_values-->", row_values)

            # Extract the first element (row) from each tuple
            BUT_DE_rows = [row for row, _ in row_values]
            logging.info("BUT_DE_rows-->", BUT_DE_rows)

            # Append "BUT DE L'ETAPE : " to each element
            Keyword4 = ["BUT DE L'ETAPE : " + keyword for keyword in keyword4]
            logging.info("Keyword4--->",Keyword4)

            logging.info("flowArray_E_Col--->",flowArray_E_Col)
            indices = [1, 3, 6, 9, 11]
            # Iterate over the indices in reverse order and insert the string
            for i in reversed(indices):
                flowArray_E_Col.pop(i)
                flowArray_E_Col.insert(i, "")
            logging.info("flowArray_E_Col before function--->", flowArray_I_Col)

            logging.info("flowArray_I_Col-->", flowArray_I_Col)
            for i, keyword in enumerate(flowArray_I_Col):  # column I
                row = rows[i]
                sheet.range(f"I{row}").value = keyword
                logging.info("I column-->", i, keyword)

            logging.info("flowArray_K_Col---->", flowArray)
            for j, keyword in enumerate(flowArray):  # column K
                row = rows[j]
                sheet.range(f"K{row}").value = keyword
                logging.info("K column-->", j, keyword)

            insertDataInRows(sheet, rows, Keyword4, BUT_DE_rows, keyword3, keyword1, flowArray_E_Col, keyword2)


# These function used generate the wired signal data to enter in rows
def WErows(keywords, sheet, WRArr_I_Col, circuitArr, wireSignals, circuit, flow):
    macro = EI.getTestPlanAutomationMacro()
    sheet_value = sheet.used_range.value
    for keyword, num in keywords:
        # keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        keyword_cell_pos = EI.searchDataInExcelCache(sheet_value, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        new_val_cel_pos, cellPos_CI = setDatainExcelSheet(sheet, keyword_cell_pos, num)
        try:
            cellPosition = getRows_FT(sheet, new_val_cel_pos)
            logging.info("cellPositions34-->", cellPosition)
            if new_val_cel_pos['count'] > 0:
                logging.info("cellPos_CI00--> ", cellPos_CI)
                row_CI, col_CI = cellPos_CI
                if wireSignals[0] != None and wireSignals[0] != "":
                    if num == 2:
                        create_DTC_WR_FT(sheet, row_CI, num, cellPosition, WRArr_I_Col, circuitArr, circuit, flow)
        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            logging.info(f"\nError in cellpostioins: {ex} line: {exc_tb.tb_lineno}")


# These function used when requirement contains both the flow and the frame
def DTCFlowforTP(flow, defectCode, req_frame, identifier):
    flowArr = []
    flowArr_E_Col = []
    flowArr_I_Col = []
    logging.info("happyyy flow---->",flow)
    # req_frame = "FD8_DYN_VOL_03F"
    logging.info("cvxcvxcv--->",req_frame)
    req_flow = flow
    logging.info("find req_flow --->", req_flow)
    try:
        # for req_flow in flow:
        if req_flow != "" and req_flow is not None:
            # if flow exist adding the prefix REQ for request and REP for response
            prefixes = ["", "$REP_DEF_", "", "", "$REP_DEF_", "", "", "$REP_DEF_", "", "", "", "$REP_DEF_", ""]
            suffixes = ["", "_ABSENT", "", "", "_PRESENT", "", "", "_FUGITIF", "", "", "", "_ABSENT", ""]

            # output_values will give the req_flow with pre and suffix and add in the flowArr
            output_values = [f"{prefix}{req_flow}{suffix}" for prefix, suffix in
                             zip(prefixes, suffixes)]
            for i in output_values:
                flowArr.append(i)
            logging.info("flowArrwith --->", flowArr)

            # # Replace the empty string with '$RAZ_JDD_OK'
            # flowArr[index] = req_flow
            # Remove 'FLG_AVOL_ICN' and replace with empty strings
            flowArr = [item if item != req_flow else '' for item in flowArr]
            logging.info("flowArr--->", flowArr)

            prefixes = ["", "$LEC_DEF_", "", "$TRAME_", "$LEC_DEF_", "", "$TRAME_", "$LEC_DEF_", "","", "", "$LEC_DEF_", ""]
            suffixes = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            # output_values will give the req_flow with pre and suffix and add in the flowArr
            output_values = [f"{prefix}{req_flow}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
            for i in output_values:
                flowArr_E_Col.append(i)
            logging.info("flowArr_E_Colwith --->", flowArr_E_Col)

            # Remove 'FLG_AVOL_ICN' and replace with empty strings
            flowArr_E_Col = [item if item != req_flow else '' for item in flowArr_E_Col]
            logging.info("flowArr_E_Col--->", flowArr_E_Col)

            logging.info("flowArray01--->", flowArr_E_Col)
            flowArr_E_Col.insert(0, '$RAZ_JDD')
            flowArr_E_Col.pop(10)
            flowArr_E_Col.insert(10, '$RAZ_JDD')
            logging.info("flowArray_E_Col--->", flowArr_E_Col)

            indices = [4, 7]
            # Iterate over the indices in reverse order and insert the string
            for i in reversed(indices):
                flowArr_E_Col.pop(i)
                flowArr_E_Col.insert(i, "$TRAME_"+identifier+"_"+req_frame)
            logging.info("flowArray_E_Col1--->", flowArr_E_Col)

            prefixes = ["", "", "Check that the DTC ", "", "", "Check that the Defect ", "", "",
                        "Check that the Defect ", "", "", "", "Check that the DTC ", ""]
            suffixes = ["", "", " are absent after clearing all the DTC.", "", "",
                        " are Present after creating the fault", "", "", " are Disappeared after removing the fault.",
                        "", "", "", " are absent after clearing all the DTC.", ""]

            # output_values will give the req_flow with pre and suffix and add in the flowArr
            output_values = [f"{prefix}'{defectCode}'{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
            for i in output_values:
                flowArr_I_Col.append(i)
            logging.info("flowArr_I_Colwith --->", flowArr_I_Col)

    except Exception as ex:
        logging.info(f"Flow is not available for these requirement.please update manually. {ex}")
    pass
    return [flowArr, flowArr_E_Col, flowArr_I_Col]


# These function used when requirement contains frame only present these logic used
def DTCFrameforTP(req_frame, defectCode, identifier):
    logging.info("framessssssss")
    flowArr = []
    flowArr_E_Col = []
    flowArr_I_Col = []
    try:
        if req_frame != "" and req_frame is not None:
            # if flow exist adding the prefix REQ for request and REP for response
            prefixes = ["","$REP_DEF_TRAME_", "","", "$REP_DEF_TRAME_", "","", "$REP_DEF_TRAME_", "","", "","$REP_DEF_TRAME_", ""]
            suffixes = ["","_ABSENT","", "", "_PRESENT","", "", "_FUGITIF", "","", "","_ABSENT", ""]
            # output_values will give the req_flow with pre and suffix and add in the flowArr
            output_values = [f"{prefix}{identifier}{'_'}{req_frame}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
            for i in output_values:
                flowArr.append(i)
            logging.info("frameArrwith --->", flowArr)

            # # Replace the empty string with '$RAZ_JDD_OK'
            # flowArr[index] = req_flow
            # Remove 'FLG_AVOL_ICN' and replace with empty strings
            flowArr = [item if item != req_frame else '' for item in flowArr]
            logging.info("frameArrkcolumn--->", flowArr)

            prefixes = ["","$LEC_DEF_TRAME_", "","$TRAME_", "$LEC_DEF_TRAME_", "","$TRAME_", "$LEC_DEF_TRAME_","", "","", "$LEC_DEF_TRAME_", ""]
            suffixes = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            # output_values will give the req_flow with pre and suffix and add in the flowArr
            output_values = [f"{prefix}{identifier}{'_'}{req_frame}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
            for i in output_values:
                flowArr_E_Col.append(i)
            logging.info("frameArr_E_Colwith --->", flowArr_E_Col)

            # Remove 'FLG_AVOL_ICN' and replace with empty strings
            flowArr_E_Col = [item if item != req_frame else '' for item in flowArr_E_Col]
            logging.info("frameArr_EColumn--->", flowArr_E_Col)

            logging.info("frameArr_EColumn--->", flowArr_E_Col)
            flowArr_E_Col.insert(0, '$RAZ_JDD')
            flowArr_E_Col.pop(10)
            flowArr_E_Col.insert(10, '$RAZ_JDD')
            flowArr_E_Col.pop(13)
            flowArr_E_Col.insert(13, '')
            logging.info("flowArray_E_Col--->", flowArr_E_Col)


            prefixes = ["","", "Check that the DTC ", "","", "Check that the Defect ","", "", "Check that the Defect ", "","","", "Check that the DTC ", ""]
            suffixes = ["", ""," are absent after clearing all the DTC.","", "", " are Present after creating the fault","", "", " are Disappeared after removing the fault.","", "","", " are absent after clearing all the DTC.", ""]
            # output_values will give the req_flow with pre and suffix and add in the flowArr
            output_values = [f"{prefix}'{defectCode}'{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
            for i in output_values:
                flowArr_I_Col.append(i)
            logging.info("flowArr_I_Colwith --->", flowArr_I_Col)
    except Exception as ex:
        logging.info(f"Frame is not available for these requirement.please update manually. {ex}")
        pass
    return [flowArr, flowArr_E_Col, flowArr_I_Col]


# These function used when requirement contains wired siganls
def DTCWIreTP(wireSignal):
    circuitArr = []
    flowArr_E_Col = []
    WRArr_I_Col = []
    try:
        for req_flow in wireSignal:
            if req_flow != "" and req_flow is not None:
                # if flow exist adding the prefix REQ for request and REP for response

                prefixes = ["", "$REP_DEF_", "", "", "$REP_DEF_", "", "", "$REP_DEF_", "", "", "", "$REP_DEF_", ""]
                suffixes = ["",   "_ABSENT", "", "",  "_PRESENT",  "", "", "_FUGITIF",  "", "", "",  "_ABSENT", ""]
                # output_values will give the req_flow with pre and suffix and add in the flowArr
                output_values = [f"{prefix}{req_flow}{suffix}" for prefix, suffix in
                                 zip(prefixes, suffixes)]
                for i in output_values:
                    circuitArr.append(i)
                logging.info("flowArrwith --->", circuitArr)

                circuitArr = [item if item != req_flow else '' for item in circuitArr]
                logging.info("flowArr--->", circuitArr)

                prefixes = ["", "$LEC_", "", "$", "$LEC_", "", "$", "$LEC_", "", "", "", "$LEC_", ""]
                suffixes = ["",      "", "",  "",      "", "",  "",      "", "", "", "",      "", ""]
                # output_values will give the req_flow with pre and suffix and add in the flowArr
                output_values = [f"{prefix}{req_flow}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
                for i in output_values:
                    flowArr_E_Col.append(i)
                logging.info("flowArr_E_Colwith --->", flowArr_E_Col)

                # Remove 'FLG_AVOL_ICN' and replace with empty strings
                flowArr_E_Col = [item if item != req_flow else '' for item in flowArr_E_Col]
                logging.info("flowArr_E_Col--->", flowArr_E_Col)

                prefixes = ["All DTC Cleared", "",                     "Check that the DTC ", "", "",                "Check that the Defect ", "", "",                     "Check that the Defect ", "", "All DTC Cleared.", "",                     "Check that the DTC ", "Check that the DTC"]
                suffixes = [               "", "", " are absent after clearing all the DTC.", "", "", " are Present after creating the fault", "", "", " are Disappeared after removing the fault.", "",                 "", "", " are absent after clearing all the DTC.",                   ""]
                # output_values will give the req_flow with pre and suffix and add in the flowArr
                output_values = [f"{prefix}{''}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
                for i in output_values:
                    WRArr_I_Col.append(i)
                logging.info("WRArr_I_Col --->", WRArr_I_Col)
    except Exception as ex:
        logging.info(f"Flow is not available for these requirement.please update manually. {ex}")
    pass
    return [circuitArr, flowArr_E_Col, WRArr_I_Col]


# These function used when requirement having wired Signals (cc,cc+,ccm) in the content
def create_DTC_WR_FT(sheet, row_CI, num, cellPosition ,WRArr_I_Col, circuitArr, circuit, WRFrame):
    # WRFrame = "MANAGE_LED_ZEV"
    b = circuit
    logging.info("circuitb-->", b)
    if sheet.range(row_CI + 1, 5).value == "" or sheet.range(row_CI + 1, 5).value is None:
        if num == 2:
            keyword2 = [                            "DIAG",                 "NO_DIAG",                                                           "DIAG",                        "", "", "DIAG",                        "",                                                  "",                                                                        "DIAG",                        "",                             "DIAG",                        "",                                                           "DIAG", ""] # column I
            keyword4 = ["Clear all the DTC using RAZ_JDD.", "Close the DIAG Session.","Verify that the DTC are absent by sending the read DTC request.", "Close the DIAG Session.", "",     "", "Close the DIAG Session.", "Make correct connections to heal the fault '+DTC'", "Verify that the Defect 'DTC' is disappeared when we correct the connections", "Close the DIAG Session.", "Clear all the DTC using RAZ_JDD.", "Close the DIAG Session.", "Verify that the DTC are absent by sending the read DTC request.", "Close the DIAG Session."]
            # WRArr_I_Col = getFlows(req)[4]
            indices = [1, 3, 4, 6, 7, 9, 11, 13]
            # Iterate over the indices in reverse order and insert the string
            for i in reversed(indices):
                WRArr_I_Col.pop(i)
                WRArr_I_Col.insert(i, "")
            logging.info("flowArray_I_Col0--->", WRArr_I_Col)

            WRArr_I_Col.pop(0)
            WRArr_I_Col.insert(0, 'All DTC Cleared.')
            WRArr_I_Col.pop(10)
            WRArr_I_Col.insert(10, 'All DTC Cleared.')
            logging.info("flowArray_I_Col1--->", WRArr_I_Col)

            logging.info("num2-->")
            rowslist = cellPosition
            # Increment row values by 1 and decrement column values by 1
            row_values = [(pos[0] + 1, pos[1] - 1) for pos in cellPosition]
            logging.info("row_values-->", row_values)
            # Extract the first element (row) from each tuple
            rows = [row for row, _ in row_values]
            logging.info("rows-->", rows)

            # Increment row values by 1 and decrement column values by 1
            row_values = [(pos[0] - 1, pos[1] - 1) for pos in cellPosition]
            logging.info("row_values-->", row_values)

            # Extract the first element (row) from each tuple
            BUT_DE_rows = [row for row, _ in row_values]
            logging.info("BUT_DE_rows-->", BUT_DE_rows)
            # Append "BUT DE L'ETAPE : " to each element
            keyword4 = ["BUT DE L'ETAPE : " + keyword for keyword in keyword4]
            logging.info("Keyword4--->", keyword4)

            if 'short circuit to plus' in b:
                keyword1, keyword3, keyword4, flowArray_E_Col = DTCWR_CCPlus(keyword4,  WRFrame)
            if 'short circuit to ground' in b:
                keyword1, keyword3, keyword4, flowArray_E_Col = DTCWR_CCM(keyword4, WRFrame)
            if 'open circuit' in b:
                keyword1, keyword3, keyword4, flowArray_E_Col = DTCWR_CO(keyword4,  WRFrame)

            logging.info("flowArray01--->", flowArray_E_Col)
            flowArray_E_Col.insert(0, '$RAZ_JDD')
            flowArray_E_Col.pop(10)
            flowArray_E_Col.insert(10, '$RAZ_JDD')
            logging.info("flowArray_E_Col--->", flowArray_E_Col)

            flowArray_E_Col.pop(4)
            flowArray_E_Col.insert(4, '$' + WRFrame)
            flowArray_E_Col.pop(7)
            flowArray_E_Col.insert(7, '$' + WRFrame)

            logging.info("CircuitArray00--->", circuitArr)
            circuitArr.insert(0, '$RAZ_JDD_OK')
            circuitArr.pop(10)
            circuitArr.insert(10, '$RAZ_JDD_OK')
            logging.info("Circuitarray--->", circuitArr)

            for i, keyword in enumerate(WRArr_I_Col):  # column I
                row = rows[i]
                sheet.range(f"I{row}").value = keyword
                logging.info("I column-->", i, keyword)

            logging.info("flowArray_K_Col before function---->", circuitArr)

            logging.info("flowArray_K_Col---->", circuitArr)
            for j, keyword in enumerate(circuitArr):  # column K
                row = rows[j]
                sheet.range(f"K{row}").value = keyword
                logging.info("K column-->", j, keyword)

            insertDataInRows(sheet, rows, keyword4, BUT_DE_rows, keyword3, keyword1, flowArray_E_Col, keyword2)


# These function is used when we have 'open circuit or short circuit to plus'.
def DTCWR_CCPlus(keyword4,WRFrame):
    WRCCPlusArr_E_Col = []
    keyword1 = [              "DIAG",                "NO_DIAG", "DIAG",                "NO_DIAG",           "COURT_CIRCUIT(CC12)", "DIAG",                "NO_DIAG",     "COURT_CIRCUIT(OFF)","DIAG",                "NO_DIAG",               "DIAG",                "NO_DIAG", "DIAG",                "NO_DIAG"] # column D
    keyword3 = ["Clear all the DTC.", "Close the DIAG session",     "", "Close the DIAG session", "Short circuit to the "+WRFrame,     "", "Close the DIAG session", "Correct the connections",   "", "Close the DIAG session", "Clear all the DTC.", "Close the DIAG session",     "", "Close the DIAG session"] # column C
    # Indices where the string should be inserted column C
    indices = [2, 5, 8, 12]
    # Iterate over the indices in reverse order and insert the string
    for i in reversed(indices):
        keyword3.pop(i)
        keyword3.insert(i, "Send the request to read the " + WRFrame)
    logging.info("keyword3-->", keyword3)
    # CC+
    keyword4.pop(4)
    keyword4.insert(4, "BUT DE L'ETAPE :Short circuit the " + WRFrame + " to create the fault DTC")
    keyword4.pop(5)
    keyword4.insert(5, "BUT DE L'ETAPE :Verify that the Defect DTC is present when we Short circuit the "+WRFrame)

    prefixes = ["", "$LEC_", "", "", "$LEC_", "", "", "$LEC_", "", "", "", "$LEC_", ""]
    suffixes = ["", "_CC12", "", "", "_CC12", "", "", "_CC12", "", "", "", "_CC12", ""]
    # output_values will give the req_flow with pre and suffix and add in the flowArr
    output_values = [f"{prefix}{WRFrame}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
    for i in output_values:
        WRCCPlusArr_E_Col.append(i)
    logging.info("flowArr_E_Colwith --->", WRCCPlusArr_E_Col)
    # Remove 'FLG_AVOL_ICN' and replace with empty strings
    flowArr_E_Col = [item if item != WRFrame else '' for item in WRCCPlusArr_E_Col]
    logging.info("flowArr_E_Col--->", flowArr_E_Col)

    return [keyword1, keyword3, keyword4, flowArr_E_Col]


# These function is used when we have 'short circuit to ground'.
def DTCWR_CCM(keyword4, WRFrame):
    WRCCPlusArr_E_Col = []
    # CCM
    keyword1 = ["DIAG", "NO_DIAG", "DIAG", "NO_DIAG", "COURT_CIRCUIT(CCM)", "DIAG", "NO_DIAG", "COURT_CIRCUIT(OFF)",
                "DIAG", "NO_DIAG", "DIAG", "NO_DIAG", "DIAG", "NO_DIAG"]  # column D
    keyword3 = ["Clear all the DTC.", "Close the DIAG session", "", "Close the DIAG session",
                "Short circuit to Ground the " + WRFrame, "", "Close the DIAG session", "Correct the connections", "",
                "Close the DIAG session", "Clear all the DTC.", "Close the DIAG session", "",
                "Close the DIAG session"]  # column C

    # Indices where the string should be inserted column C
    indices = [2, 5, 8, 12]
    # Iterate over the indices in reverse order and insert the string
    for i in reversed(indices):
        keyword3.pop(i)
        keyword3.insert(i, "Send the request to read the " + WRFrame)
    logging.info("keyword3-->", keyword3)
    keyword4.pop(4)
    keyword4.insert(4, "BUT DE L'ETAPE :Short circuit to ground for the "+WRFrame+" to create the fault DTC")
    keyword4.pop(5)
    keyword4.insert(5, "BUT DE L'ETAPE :Verify that the Defect DTC is present when we Short circuit to ground the flow "+WRFrame)

    prefixes = ["", "$LEC_", "", "", "$LEC_", "", "", "$LEC_", "", "", "", "$LEC_", ""]
    suffixes = ["", "_CCM", "", "", "_CCM", "", "", "_CCM", "", "", "", "_CCM", ""]
    # output_values will give the req_flow with pre and suffix and add in the flowArr
    output_values = [f"{prefix}{WRFrame}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
    for i in output_values:
        WRCCPlusArr_E_Col.append(i)
    logging.info("flowArr_E_Colwith --->", WRCCPlusArr_E_Col)
    # Remove 'FLG_AVOL_ICN' and replace with empty strings
    flowArr_E_Col = [item if item != WRFrame else '' for item in WRCCPlusArr_E_Col]
    logging.info("flowArr_E_Col--->", flowArr_E_Col)
    return [keyword1, keyword3, keyword4, flowArr_E_Col]


# These function is used when we have open circuit.
def DTCWR_CO(keyword4, WRFrame):
    WRCCPlusArr_E_Col = []
    # # CO

    keyword1 = ["DIAG", "NO_DIAG", "DIAG", "NO_DIAG", "COURT_CIRCUIT(CO)", "DIAG", "NO_DIAG", "COURT_CIRCUIT(OFF)",
                "DIAG", "NO_DIAG", "DIAG", "NO_DIAG", "DIAG", "NO_DIAG"]  # column D
    keyword3 = ["Clear all the DTC.", "Close the DIAG session", "", "Close the DIAG session",
                "Open circuit to the " + WRFrame, "", "Close the DIAG session", "Correct the connections", "",
                "Close the DIAG session", "Clear all the DTC.", "Close the DIAG session", "",
                "Close the DIAG session"]  # column C
    # Indices where the string should be inserted column C
    indices = [2, 5, 8, 12]
    # Iterate over the indices in reverse order and insert the string
    for i in reversed(indices):
        keyword3.pop(i)
        keyword3.insert(i, "Send the request to read the " + WRFrame)
    logging.info("keyword3-->", keyword3)
    keyword4.pop(4)
    keyword4.insert(4, "BUT DE L'ETAPE :Open circuit for the " + WRFrame + " to create the fault DTC")
    keyword4.pop(5)
    keyword4.insert(5,"BUT DE L'ETAPE :Verify that the Defect DTC is present when we Open circuit to ground the flow " + WRFrame)

    prefixes = ["", "$LEC_", "", "", "$LEC_", "", "", "$LEC_", "", "", "", "$LEC_", ""]
    suffixes = ["", "_CO", "", "", "_CO", "", "", "_CO", "", "", "", "_CO", ""]
    # output_values will give the req_flow with pre and suffix and add in the flowArr
    output_values = [f"{prefix}{WRFrame}{suffix}" for prefix, suffix in zip(prefixes, suffixes)]
    for i in output_values:
        WRCCPlusArr_E_Col.append(i)
    logging.info("flowArr_E_Colwith --->", WRCCPlusArr_E_Col)
    # Remove 'FLG_AVOL_ICN' and replace with empty strings
    flowArr_E_Col = [item if item != WRFrame else '' for item in WRCCPlusArr_E_Col]
    logging.info("flowArr_E_Col--->", flowArr_E_Col)
    return [keyword1, keyword3, keyword4, flowArr_E_Col]


# These function is used to insert the data in the Excel sheet by giving input as arguments
def insertDataInRows(sheet, rows, Keyword4, BUT_DE_rows, keyword3, keyword1, flowArray_E_Col, keyword2):
    for i, keyword in enumerate(Keyword4):  # column BUT_DE
        row = BUT_DE_rows[i]
        sheet.range(f"A{row}").value = keyword
        logging.info("BUT_DE_column-->", i, keyword)

    for i, keyword in enumerate(keyword3):  # column c
        row = rows[i]
        sheet.range(f"C{row}").value = keyword
        logging.info("C column-->", i, keyword)

    #  to find out the values in A column.
    for i in rows:
        b = sheet.range(f"A{i}").value
        logging.info("A column-->", b)

    for i, keyword in enumerate(keyword1):  # column D
        row = rows[i]
        sheet.range(f"D{row}").value = keyword
        logging.info("D column-->", i, keyword)

    logging.info("flowArray_E_Col---->",flowArray_E_Col)
    for i, keyword in enumerate(flowArray_E_Col):  # column E
        row = rows[i]
        sheet.range(f"E{row}").value = keyword
        logging.info("E column-->", i, keyword)

    for j, keyword in enumerate(keyword2):  # column J
        row = rows[j]
        sheet.range(f"J{row}").value = keyword
        logging.info("j column-->", j, keyword)