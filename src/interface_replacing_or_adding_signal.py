import os
import TestPlanMacros as TPM
import InputConfigParser as ICF
import ExcelInterface as EI
import web_interface as WB
import re
# import KeyboardMouseSimulator as KMS
import logging

vernum = '([vV]{1}[0-9]{1,2}\.[0-9]{1,2})|([vV]{1}[0-9]{1,2})'
refnum = '([0-9]{5})+(_[0-9]{2})+(_[0-9]{5})+'
keywords_DCI = ['DCI', 'dci', 'DCINT', 'DCIINT', 'dciint', 'dcint']
pattern = r'^(v|V)'
remove_ver_from_req = r'\(.*\)$'


def remove_duplicates(my_list):
    unique_list = list(set(my_list))
    return unique_list


# ===>> Download the Previous Documents in Sommaire sheet <<=== #
def Download_Dci_Docs_In_Sommaire(tpBook):
    list_dci_files = []
    EI.activateSheet(tpBook, 'Sommaire')
    sheet = tpBook.sheets['Sommaire']
    max_row = tpBook.sheets['Sommaire'].range('E' + str(tpBook.sheets['Sommaire'].cells.last_cell.row)).end('up').row
    startingRow = 6
    max = max_row + 1
    for x in range(startingRow, max):
        DCI_FileName = sheet.range(x, 5).value
        for key in keywords_DCI:
            if DCI_FileName is not None:
                if DCI_FileName.find(key)!=-1:
                    reference = sheet.range(x, 6).value
                    version = sheet.range(x, 7).value
                    logging.info('version-if-none-->', version, 'reference-if-none-->', reference)
                    if (reference or version) is not None:
                        logging.info('version---->', version, 'reference---->', reference)
                        match = bool(re.search(pattern, str(version)))
                        logging.info('match---->', match)
                        if match:
                            list_dci_files.append((DCI_FileName, reference, str(version)))
                        else:
                            version_with_prefix = "V" + str(version)
                            list_dci_files.append((DCI_FileName, reference, version_with_prefix))
    list_dci_files = remove_duplicates(list_dci_files)
    logging.info('list_dci_files', list_dci_files)
    for doc_name, ref, ver in list_dci_files:
        doc_ref_ver = [(ref, ver)]
        logging.info(doc_name, ref, ver)
        logging.info('doc_ref_ver ----> ', doc_ref_ver)
        WB.startDocumentDownload(doc_ref_ver, False)


def verify_req_Downloaded_dci(reqirements):
    keyword_dci_sheet_names = ['MUX', 'FILAIRE', 'DCi_NT']
    signal_names = []
    functional_Dci_files = []
    lis = os.listdir(ICF.getInputFolder())
    for l in lis:
        for key in keywords_DCI:
            if ((l.find(key))!=-1) and (l.find('DCI_GLOBAL'))==-1 and (l.find('~$'))==-1:
                functional_Dci_files.append(l)
    logging.info('functional_Dci_files--->', functional_Dci_files)
    for reqs in reqirements:
        logging.info('reqs--->', reqs)
        for dci_file in functional_Dci_files:
            dci_doc = EI.openExcel(ICF.getInputFolder() + "\\" + dci_file)
            sheet_names = [sheet.name for sheet in dci_doc.sheets]
            logging.info('dci_doc---->', dci_doc)
            for keys in keyword_dci_sheet_names:
                try:
                    logging.info('keys--->', keys)
                    EI.activateSheet(dci_doc, keys)
                    MUX_Sig = EI.searchDataInCol(dci_doc.sheets[keys], 1, reqs)
                    # logging.info('DCI_REQ searched in MUX ===>', MUX_Sig)
                    if MUX_Sig['cellPositions']:
                        for cell in MUX_Sig['cellPositions']:
                            row, col = cell
                            arc = str(dci_doc.sheets[keys].range(row, 2).value)
                            logging.info('arc-->', arc)
                            if arc=='VSM':
                                signal_ = str(dci_doc.sheets[keys].range(row, 3).value)
                                logging.info('signal_', signal_, 'dci_file', dci_file)
                                signal_names.append(signal_)
                except:
                    # logging.info('!!! ["MUX", "FILAIRE", "DCi_NT"] only these keywords are handled !!!')
                    # pass
                    # if any(item in keyword_dci_sheet_names for item in sheet_names):
                    #     pass
                    # else:
                    #     logging.info('in keys present', sheet_names)
                    #     raise Exception(f'Sorry, !!! ["MUX", "FILAIRE", "DCi_NT"] only these sheet names are handled please verify with the particular dci document {dci_file}!!!')
                    pass
            dci_doc.close()
    logging.info('signal_names--->', signal_names)
    return signal_names


def verify_flow(tpBook, param_data, dci_info, sheet_name):
    EI.activateSheet(tpBook, tpBook.sheets[sheet_name])
    searchResultForSignal = EI.searchDataInExcel(tpBook.sheets[sheet_name], "", dci_info['dciSignal'])
    logging.info('searchResultForSignal---->', searchResultForSignal)
    val = [item.replace('$', '').strip() for item in searchResultForSignal['cellValue']]
    logging.info('val--->', val)
    parm_sig = param_data[0]['PARAM_SIGNAL']
    logging.info('param_sig --->', parm_sig)
    all_match = all(item==parm_sig for item in val)
    logging.info('all_match--->', all_match)
    return all_match


def replace_signal_in_sheet(tpBook, sheet_name, dci_signal, new_signal):
    results = EI.searchDataInExcel(tpBook.sheets[sheet_name], "", dci_signal)
    logging.info('new_signal---->', new_signal)
    logging.info('sheet_name---->', sheet_name)
    logging.info('results-->', results)
    logging.info('param_signal--->', new_signal)
    EI.activateSheet(tpBook, sheet_name)
    for cell in results['cellPositions']:
        tpBook.sheets[sheet_name].range(cell).value = str(new_signal)


def update_dynamically(tpBook, sheet_name, searchResultForSignal, new_name):
    count = 0
    macro = EI.getTestPlanAutomationMacro()
    popupribbon = macro.macro("TPMacros.TriggerPopup")
    logging.info('sheet ==>', sheet_name)
    popupribbon(sheet_name)
    for coor in searchResultForSignal['cellPositions']:
        row, col = coor
        # KMS.showWindow(tpBook.name.split('.')[0])
        tpBook.sheets[sheet_name].api.Unprotect()
        testsheet = tpBook.sheets[sheet_name]
        is_empty = row_is_empty(tpBook, testsheet, coor)
        if is_empty==1:
            logging.info('There is an empty row in test sheet')
            add_new_line(tpBook, tpBook.sheets[sheet_name], (row, col),
                         new_name)
        elif is_empty==0:
            logging.info('There is no empty row so Adding New line--------> here')
            Actual_Value = str(testsheet.range(row, col).value)
            cell_Value1 = str(testsheet.range(row + 1, 1).value)
            logging.info('Actual_Value--->', Actual_Value)
            logging.info('cell_Value1------>', cell_Value1)
            status = bool(re.search("BUT DE L'ETAPE :", cell_Value1))
            value_____ = str(testsheet.range(75, 5).value)
            logging.info('value_____', value_____)
            logging.info('status------>', status)
            TPM.addNewLine(testsheet, ((row + count) + 1), col)
            logging.info('count---->', count)
            logging.info('(row + count, col)--->', (row + count, col))
            add_new_line(tpBook, tpBook.sheets[sheet_name], (row + count, col),
                         new_name)
            count += 1


def dci_value_update(param_data, dci_info, data_dic):
    # PARAM_SIGNAL
    # dciSignal
    logging.info('dci_info------>', dci_info)
    logging.info('param_data------>', param_data)
    if (data_dic!=1) and (len(data_dic['new_signal']) > 0) and (len(data_dic['test_sheet_name']) > 0) and (
            data_dic['new_signal']!='') and (data_dic['test_sheet_name']!='') and (
            data_dic['new_signal'] is not None) and (data_dic['test_sheet_name'] is not None):
        sheet_name = data_dic['test_sheet_name']
        new_name = data_dic['new_signal']
        tpBook = EI.openTestPlan()
        Dci_reqs = []
        status = verify_flow(tpBook, param_data, dci_info, sheet_name)
        if status is False:
            logging.info('Not matched !!!')
            c4_req = (tpBook.sheets[sheet_name].range('C4').value.split('|'))
            logging.info('c4_req--->', c4_req)
            for key in keywords_DCI:
                for c4 in c4_req:
                    if c4.find(key)!=-1:
                        req_c4_ = re.sub(remove_ver_from_req, '', c4.strip())
                        Dci_reqs.append(req_c4_.strip())
            unix = remove_duplicates(Dci_reqs)
            current_dci_req = re.sub(remove_ver_from_req, '', dci_info['dciReq'].strip())
            logging.info('current_dci_req---->', current_dci_req)
            if str(current_dci_req).strip() in unix:
                unix.remove(str(current_dci_req.strip()))
            logging.info('unix--->', unix)

            if not unix:
                logging.info('No requirement is present in the c4 colm directly replacing the signal')
                replace_signal_in_sheet(tpBook, sheet_name, dci_info['dciSignal'].strip(),
                                        param_data[0]['PARAM_SIGNAL'].strip())
            else:
                logging.info('proceed with comparing process')
                Download_Dci_Docs_In_Sommaire(tpBook)
                signal_names = verify_req_Downloaded_dci(unix)
                # signal_names.append(dci_info['dciSignal'].strip()) ################3
                dci_signal = dci_info['dciSignal'].replace('$', '').strip()
                if dci_signal.strip() in signal_names:
                    logging.info('matched please continue the process adding new line')
                    searchResultForSignal = EI.searchDataInExcel(tpBook.sheets[sheet_name], "", dci_info['dciSignal'])
                    logging.info("searchResultForSignal------>> ", searchResultForSignal)
                    if searchResultForSignal['count']!=-1:
                        if searchResultForSignal['count'] > 0:
                            signal_cell_position = creating_Group_flow(tpBook, sheet_name,
                                                                       searchResultForSignal['cellPositions'])
                            logging.info('----', signal_cell_position, '----')
                            if len(searchResultForSignal['cellPositions']) > 0:
                                logging.info('searchResultForSignal["cellPositions"]--->',
                                      searchResultForSignal['cellPositions'])
                                update_dynamically(tpBook, sheet_name, searchResultForSignal, new_name)
                else:
                    logging.info('continue with replacing the signal')
                    replace_signal_in_sheet(tpBook, sheet_name, dci_info['dciSignal'].strip(), new_name.strip())
        elif status is True:
            logging.info(' |***| matched |***| no need to do any thing |***| ')


# ===>> Grouping the Signal coordinates as per Input and output signal <<=== #
def creating_Group_flow(tpBook, sheet, cellpos):
    # input
    global input_final_result
    Input_Buffer = []
    start_input_ranges = []
    input_cell_pos = []
    input_final = []

    # output
    global output_final_result
    output_Buffer = []
    start_output_ranges = []
    output_cell_pos = []
    output_final = []

    testsheet = tpBook.sheets(sheet)
    EI.activateSheet(tpBook, sheet)

    for cell in cellpos:
        row, col = cell
        logging.info('row', row)
        logging.info('col', col)
        if col==5:
            input_cell_pos.append((row, col))
        elif col==11:
            output_cell_pos.append((row, col))
    logging.info('input_cell_pos', input_cell_pos)
    logging.info('output_cell_pos', output_cell_pos)

    # Input signal
    if len(input_cell_pos) > 0:
        group_input_signal = EI.searchDataInExcel(testsheet, '', "PARAMETRE D'ENTREE")
        logging.info('group_input_signal----->', group_input_signal)
        logging.info('input_cell_pos----->', input_cell_pos)

        for a_tuple in group_input_signal["cellPositions"]:
            Input_Buffer.append(a_tuple[0])
        logging.info('Buffer0', Input_Buffer)
        logging.info('len(Buffer0)', len(Input_Buffer))
        length = len(Input_Buffer) - 1  # test length of buffer#######
        try:
            for i in range(0, length, 1):
                logging.info((Input_Buffer[i], Input_Buffer[i + 1]))
                start_input_ranges.append((Input_Buffer[i], Input_Buffer[i + 1]))
        except IndexError:
            logging.info("The end of list index")
        logging.info('start---->', start_input_ranges)
        logging.info('input_signal_cell_pos----->', input_cell_pos)

        last_cell_row = start_input_ranges[-1][1]
        last_cell_col = input_cell_pos[-1][1]
        logging.info('Last cell position', last_cell_row)
        logging.info('Last cell position', last_cell_col)
        last_row = testsheet.range(last_cell_row, last_cell_col).end('down').row
        logging.info('lst_row', last_row)
        var_val = str(testsheet.range(last_cell_row + 1, last_cell_col).value)
        logging.info('var_val', var_val)
        if var_val=='' or var_val=='None':
            logging.info('hhhh')
            last_empty_row = last_cell_row + 2
            logging.info('Last row of empty cell', last_empty_row)
            logging.info('Last cell position --->', (last_cell_row, last_empty_row))
            start_input_ranges.append((last_cell_row, last_empty_row))
        elif last_row > 50000:
            logging.info('llll')
            last_empty_row = last_cell_row + 2
            logging.info('Last row of empty cell', last_empty_row)
            logging.info('Last cell position --->', (last_cell_row, last_empty_row))
            start_input_ranges.append((last_cell_row, last_empty_row))
        else:
            logging.info('gggg')
            last_empty_row = last_row + 1
            logging.info('Last row of empty cell', last_empty_row)
            logging.info('Last cell position --->', (last_cell_row, last_empty_row))

            start_input_ranges.append((last_cell_row, last_empty_row))

        logging.info('start---->final----->', start_input_ranges)

        range_groups = [[] for _ in range(len(start_input_ranges))]
        lower_bounds = [0] * len(start_input_ranges)
        for sig_pos in input_cell_pos:
            for i, start in enumerate(start_input_ranges):
                logging.info('ssssssssss', start)
                if start[0] < sig_pos[0] < start[1] and sig_pos[0] > lower_bounds[i]:
                    range_groups[i].append(sig_pos)
                    lower_bounds[i] = sig_pos[0]

        for i, start in enumerate(start_input_ranges):
            logging.info(f"Signal group for input ranges {start}: {range_groups[i]}")
            input_final.append(range_groups[i])
        logging.info('final', input_final)
        input_final_result = list(filter(None, input_final))
        logging.info('\ninput_final_result', input_final_result)
        logging.info('innnnnn''''''''')
        if input_cell_pos!=[] and output_cell_pos==[]:
            logging.info('outttttttttttt''''''''')
            return input_final_result

    # Output signal
    if len(output_cell_pos) > 0:
        group_output_signal = EI.searchDataInExcel(testsheet, '', "PARAMETRE DE SORTIE")
        logging.info('group_output_signal----->', group_output_signal)
        logging.info('output_cell_pos----->', output_cell_pos)

        for a_tuple in group_output_signal["cellPositions"]:
            output_Buffer.append(a_tuple[0])
        logging.info('Buffer', output_Buffer)
        logging.info('len(Buffer)', len(output_Buffer))
        length = len(output_Buffer) - 1
        try:
            for i in range(0, length, 1):
                logging.info((output_Buffer[i], output_Buffer[i + 1]))
                start_output_ranges.append((output_Buffer[i], output_Buffer[i + 1]))
        except IndexError:
            logging.info("The end of list index")
        logging.info('start---->', start_output_ranges)
        logging.info('output_signal_cell_pos----->', output_cell_pos)

        last_cell_row = start_output_ranges[-1][1]
        last_cell_col = output_cell_pos[-1][1]
        logging.info('Last cell position', last_cell_row)
        logging.info('Last cell position', last_cell_col)
        last_row = testsheet.range(last_cell_row, last_cell_col).end('down').row
        logging.info('lst_row', last_row)
        var_val = str(testsheet.range(last_cell_row + 1, last_cell_col).value)
        logging.info('var_val', var_val)
        if var_val=='' or var_val=='None':
            logging.info('hhhh')
            last_empty_row = last_cell_row + 2
            logging.info('Last row of empty cell', last_empty_row)
            logging.info('Last cell position --->', (last_cell_row, last_empty_row))
            start_output_ranges.append((last_cell_row, last_empty_row))
        elif last_row > 50000:
            logging.info('llll')
            last_empty_row = last_cell_row + 2
            logging.info('Last row of empty cell', last_empty_row)
            logging.info('Last cell position --->', (last_cell_row, last_empty_row))
            start_output_ranges.append((last_cell_row, last_empty_row))
        else:
            logging.info('gggg')
            last_empty_row = last_row + 1
            logging.info('Last row of empty cell', last_empty_row)
            logging.info('Last cell position --->', (last_cell_row, last_empty_row))
            start_output_ranges.append((last_cell_row, last_empty_row))
        logging.info('start---->final----->', start_output_ranges)

        range_groups = [[] for _ in range(len(start_output_ranges))]
        lower_bounds = [0] * len(start_output_ranges)
        for sig_pos in output_cell_pos:
            for i, start in enumerate(start_output_ranges):
                if start[0] < sig_pos[0] < start[1] and sig_pos[0] > lower_bounds[i]:
                    range_groups[i].append(sig_pos)
                    lower_bounds[i] = sig_pos[0]

        for i, start in enumerate(start_output_ranges):
            logging.info(f"Signal group for output ranges {start}: {range_groups[i]}")
            output_final.append(range_groups[i])
        logging.info('final', output_final)
        output_final_result = list(filter(None, output_final))
        logging.info('\noutput_final_result', output_final_result)
        if input_cell_pos==[] and output_cell_pos!=[]:
            return output_final_result
        elif input_cell_pos!=[] and output_cell_pos!=[]:
            final_list = (input_final_result + output_final_result)
            logging.info('final listttttttt-------->', final_list)
            return final_list


# ===>> Adding the new line in the each test sheets <<=== #RR
def add_new_line(tpBook, testsheet, coord, signal):
    # testsheet = tpBook.sheets[sheet_name]
    # KMS.showWindow(tpBook.name.split('.')[0])
    logging.info('coordin', coord)
    row, col = coord
    if row and col:
        logging.info('rowe', row, 'col', col)
        Default_Value = ''
        EI.activateSheet(tpBook, testsheet)
        cellValue = str(testsheet.range(coord).value)
        logging.info('cellValue ==>', cellValue)
        logging.info('Signal ==>', signal)
        # Input signal
        if 0 <= col <= 6:
            logging.info('++++++++++INPUT SIGNAL++++++++++For ADDING NEW ROW++++++++++')
            logging.info('rowinitial', row)
            val1 = str(testsheet.range(row, 2).value)
            testsheet.range(row + 1, 2).value = val1
            logging.info('rowfinal', row)
            logging.info('val1 ==>', val1)
            if val1=='None':
                testsheet.range(row + 1, 2).value = Default_Value
            val2 = str(testsheet.range(row, 3).value)
            testsheet.range(row + 1, 3).value = val2
            if val2=='None':
                testsheet.range(row + 1, 3).value = Default_Value
            val3 = str(testsheet.range(row, 4).value)
            testsheet.range(row + 1, 4).value = val3
            if val3=='None':
                testsheet.range(row + 1, 4).value = Default_Value
            testsheet.range(row + 1, 5).value = signal
            val5 = str(testsheet.range(row, 6).value)
            testsheet.range(row + 1, 6).value = val5
            if val5=='None':
                testsheet.range(row + 1, 6).value = Default_Value
            return 1

        # Output signal
        elif 7 <= col <= 13:
            logging.info('++++++++++OUTPUT SIGNAL++++++++++For ADDING NEW ROW++++++++++')
            val6 = str(testsheet.range(row, 7).value)
            testsheet.range(row + 1, 7).value = val6
            if val6=='None':
                testsheet.range(row + 1, 7).value = Default_Value
            val7 = str(testsheet.range(row, 8).value)
            testsheet.range(row + 1, 8).value = val7
            if val7=='None':
                testsheet.range(row + 1, 8).value = Default_Value
            val8 = str(testsheet.range(row, 9).value)
            testsheet.range(row + 1, 9).value = val8
            if val8=='None':
                testsheet.range(row + 1, 9).value = Default_Value
            val9 = str(testsheet.range(row, 10).value)
            testsheet.range(row + 1, 10).value = val9
            if val9=='None':
                testsheet.range(row + 1, 10).value = Default_Value
            testsheet.range(row + 1, 11).value = signal
            val11 = str(testsheet.range(row, 12).value)
            testsheet.range(row + 1, 12).value = val11
            if val11=='None':
                testsheet.range(row + 1, 12).value = Default_Value
            val12 = str(testsheet.range(row, 13).value)
            testsheet.range(row + 1, 13).value = val12
            if val12=='None':
                testsheet.range(row + 1, 13).value = Default_Value
            return 1
        else:
            logging.info('Failed to add or copy the new row for signal')


# ===>> For checking row is Empty row in the each test sheets <<=== #
def row_is_empty(tpBook, sheet, coord):
    testsheet = tpBook.sheets(sheet)
    logging.info('coordi33', coord)
    row, col = coord
    if row and col:
        logging.info('row33', row)
        EI.activateSheet(tpBook, sheet)
        cellValue = str(testsheet.range(coord).value)
        logging.info('cellValue ==>', cellValue)
        # Input signal
        if 0 <= col <= 6:
            logging.info('++++++++++INPUT SIGNAL++++++++++For checking row is Empty++++++++++')
            val1 = str(testsheet.range(row + 1, 1).value)
            logging.info('val1--->', val1)
            val2 = str(testsheet.range(row + 1, 2).value)
            logging.info('val2--->', val2)
            val3 = str(testsheet.range(row + 1, 3).value)
            logging.info('val3--->', val3)
            val4 = str(testsheet.range(row + 1, 4).value)
            logging.info('val4--->', val4)
            val5 = str(testsheet.range(row + 1, 5).value)
            logging.info('val5--->', val5)
            val6 = str(testsheet.range(row + 1, 6).value)
            logging.info('val6--->', val6)
            if ((val1 is None) and (val2 is None) and (val3 is None) and (val4 is None) and (val5 is None) and (
                    val6 is None)) or (
                    (val1=='None') and (val2=='None') and (val3=='None') and (val4=='None') and (val5=='None') and (
                    val6=='None')):
                logging.info('An Empty input row is there')
                return 1
            elif ((val1 is not None) and (val2 is not None) and (val3 is not None) and (val4 is not None) and (
                    val5 is not None) and (val6 is not None)) or (
                    (val1!='None') and (val2!='None') and (val3!='None') and (val4!='None') and (val5!='None') and (
                    val6!='None')):
                logging.info('No Empty input row is there')
                return 0
        elif 7 <= col <= 13:
            logging.info('++++++++++OUTPUT SIGNAL++++++++++For checking row is Empty++++++++++')
            val1 = str(testsheet.range(row + 1, 1).value)
            logging.info('val1--->', val1)
            val7 = str(testsheet.range(row + 1, 7).value)
            logging.info('val7--->', val7)
            val8 = str(testsheet.range(row + 1, 8).value)
            logging.info('val8--->', val8)
            val9 = str(testsheet.range(row + 1, 9).value)
            logging.info('val9--->', val9)
            val10 = str(testsheet.range(row + 1, 10).value)
            logging.info('val10--->', val10)
            val11 = str(testsheet.range(row + 1, 11).value)
            logging.info('val11--->', val11)
            val12 = str(testsheet.range(row + 1, 12).value)
            logging.info('val12--->', val12)
            val13 = str(testsheet.range(row + 1, 13).value)
            logging.info('val13--->', val13)
            if ((val1 is None) and (val7 is None) and (val8 is None) and (val9 is None) and (val10 is None) and (
                    val11 is None) and (
                        val12 is None) and (val13 is None)) or (
                    (val1=='None') and (val7=='None') and (val8=='None') and (val9=='None') and (val10=='None') and (
                    val11=='None') and (
                            val12=='None') and (val13=='None')):
                logging.info('An Empty output row is there')
                return 1
            elif ((val1 is not None) and (val7 is not None) and (val8 is not None) and (val9 is not None) and (
                    val10 is not None) and (
                          val11 is not None) and (val12 is not None) and (val13 is not None)) or (
                    (val1!='None') and (val7!='None') and (val8!='None') and (val9!='None') and (val10!='None') and (
                    val11!='None') and (
                            val12!='None') and (val13!='None')):
                logging.info('No Empty output row is there')
                return 0
    else:
        logging.info('No Coordinates present!!!!!')

# if __name__=="__main__":
# req = ['GEN-VHL-DCINT-SSY-IHV-COHG.1654(0)', 'GEN-VHL-DCINT-SSY-IHV-COHG.1641(0)',
#        'GEN-VHL-DCINT-SSY-IHV-COHG.1629(0)', 'GEN-VHL-DCINT-NT-MASCOM-121(0)']
# verify_req_Downloaded_dci(req)
# dci_value_update(ParamData, dciInfo, sheet_flow_name)
# coord = [(13, 5), (27, 5), (30, 5), (47, 5)]
# coord = [(20, 5), (33, 5), (39, 5), (44, 5), (47, 5)]
# # coord = [(14, 11), (20, 11), (33, 11), (39, 11), (44, 11)]
# logging.info('coord--->', coord)
# flag = 0
# count = 0
# sheet_name = 'VSM20_GC_02_54_0038'
# macro = EI.getTestPlanAutomationMacro()
# popupribbon = macro.macro("TPMacros.TriggerPopup")
# logging.info('sheet ==>', sheet_name)
# popupribbon(sheet_name)
# for coor in coord:
#     row, col = coor
#     # KMS.showWindow(tpBook.name.split('.')[0])
#     tpBook.sheets[sheet_name].api.Unprotect()
#     testsheet = tpBook.sheets[sheet_name]
#     is_empty = row_is_empty(tpBook, testsheet, coor)
#     if is_empty == 1:
#         logging.info('There is an empty row in test sheet')
#         add_new_line(tpBook, tpBook.sheets[sheet_name], (row, col), ParamData['paramSignal'])
#     elif is_empty == 0:
#         logging.info('There is no empty row so Adding New line--------> here')
#         Actual_Value = str(testsheet.range(row, col).value)
#         cell_Value1 = str(testsheet.range(row + 1, 1).value)
#         logging.info('Actual_Value--->', Actual_Value)
#         logging.info('cell_Value1------>', cell_Value1)
#         status = bool(re.search("BUT DE L'ETAPE :", cell_Value1))
#         value_____ = str(testsheet.range(75, 5).value)
#         logging.info('value_____', value_____)
#         logging.info('status------>', status)
#         TPM.addNewLine(testsheet, ((row+count) + 1), col)
#         logging.info('count---->', count)
#         logging.info('(row + count, col)--->', (row + count, col))
#         add_new_line(tpBook, tpBook.sheets[sheet_name], (row + count, col), ParamData['paramSignal'])
#         count += 1
#

# if flag == 1:
#     row, col = coor
#     coor = (row + count, col)
#     add_new_line(tpBook, tpBook.sheets[sheet_name], coor, ParamData['paramSignal'])
#     flag = 0
# else:
#     row, col = coor
#     coor_ = (row + count, col)
#     var = add_new_line(tpBook, tpBook.sheets[sheet_name], coor_,
#                        ParamData['paramSignal'])
#     logging.info('var----->', var)
#     if var == 1:
#         flag = 1
#         count += 1
#     elif var == 2:
#         logging.info('empty row+++')
# logging.info('count--->', count, 'row--->', (coor[0]+count))


# ICF.loadConfig()
# tpBook = EI.openTestPlan()
# sheets = tpBook.sheets

# dciInfo = {
#     "dciSignal": "ETAT_PRINCIP_SEV",
#     "network": "hs_8",
#     "pc": "",
#     "thm": "",
#     "framename": "",
#     "dciReq": "DCINT-00027946(2)",
#     "proj_param": ""
# }
# ParamData = {
#     "paramSignal": "ETAT_PRINCIP_SEV_hs10",
#     "network": "hs_7",
#     "pc": "",
#     "framename": "",
#     "dciReq": "DCINT-00027946(2)"
# }
# sheet_flow_name = 'VSM20_GC_02_54_0038'

# DCIdoc = ['DCI_SUBSYST_PARK_HMI_22Q4 -01991_19_01856 V11.0']