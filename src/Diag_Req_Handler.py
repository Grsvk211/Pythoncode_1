import os
import re
import sys
import time
# from time import sleep
import WordDocInterface as WDI
import logging
import TestPlanMacros as TPM
import ExcelInterface as EI
import InputConfigParser as ICF
import BusinessLogic as BL
import extract_life_cycle as lyf
import QIA_Param as QP
import NewRequirementHandler as NRH
import DIAG_other_value_key_words as DOVKW
import DocumentSearch as DS

import json
# from nltk import word_tokenize
# import KeyboardMouseSimulator as KMS
# from collections import OrderedDict

# ICF.loadConfig()

# ReqName0 = 'REQ-0742878'
# ReqVer0 = 'A'
# file_name = 'SSD_HMIF_LONGITUDINAL_MOBILITY_MOBY_HMI_23Q1.docx'
# pat = r'C:\Users\clakshminarayan\Documents\BSI-VSM(Automation)\Integarted, DTC, DIAG, CALIBRATION\Input_Files'
DID_pattren = r'([A-Za-z]+[-_]{1}[A-Za-z]+[0-9]+)'
Pat_tren = r'[A-Za-z]+[-_]'
patt = r'[A-Za-z]+'


def extract_data_DCI_Global_G_column(var):
    # logging.basicConfig(level=logging.DEBUG)
    data = {}
    lines = var.split('\n')
    for line in lines:
        if line.strip():
            if '=' in line:
                key, value = line.split('=')
                data[key.strip().lower().capitalize()] = value.strip()
                logging.debug(f"Key: {key.strip()}, Value: {value.strip()}")
            elif ':' in line:
                key, value = line.split(':')
                data[key.strip().lower().capitalize()] = value.strip()
                logging.debug(f"Key: {key.strip()}, Value: {value.strip()}")
    logging.debug(f'data-extract_diag_Content-->{data}')
    return data


# variable = 'Resolution : 1.0\nInitValue=0000\nOffset=0.0\nEst signé=false\nBit start=1\nTaille=4.0\nValeur_Invalide_S=Not Applicable\nValeur_Interdite_S=Not Applicable\nNORMAL:0000=0000\nSPORT:0001=0001\nCONFORT:0010=0010\nECO:0011=0011\nSABLE:0100=0100\nBOUE:0101=0101\nNEIGE:0110=0110\nZEV:0111=0111\nEAWD:1000=1000\nHYBRID:1001=1001\nZEV_ECO:1010=1010\nHYBRID_ECO:1011=1011\nECO_PLUS:1100=1100\nREINFORCED LOAD:1101=1101\nRESERVE:1110=1110\nRESERVE:1111=1111\nE_evt=Yes\nColumnEvtValue=E(trig)+P1000'


def Init_Value_extraction(var):
    data = extract_data_DCI_Global_G_column(var)
    logging.info('data', data)
    dic = {}
    init_value_keys = ['Initvalue', 'InitValue', 'initvalue', 'initValue']
    for key in init_value_keys:
        if key in data:
            dic['InitValue'] = data[key]
            return dic
    logging.info('  ################################  ')
    logging.info('!!!!! InitValue is not present !!!!!')
    logging.info('  ################################  ')
    dic['InitValue'] = ''
    return dic


def is_hexadecimal(value):
    try:
        int(float(value))
        logging.info('Given string is Decimal --->', value)
        return False
    except ValueError:
        logging.info('Given string is Hexadecimal --->', value)
        return True


def is_binary(Value):
    b = '10'
    count = 0
    for char in Value:
        if char not in b:
            count = 1
            break
        else:
            pass
    if count:
        logging.info("StringA is not a binary string", Value)
        return False
    else:
        logging.info("StringA is a binary string", Value)
        return True


def convert_to_decimal(value):
    logging.info('convert_to_decimal(value)--->', value)
    if value!='':
        binary = is_binary(value)
        hex_decimal = is_hexadecimal(value)
        if hex_decimal is True:
            return int(value, 16)
        elif binary is True:
            return int(value, 2)
        elif value:
            return int(float(value))
        else:
            return None
    else:
        return ''


def extract_G_col(data):
    keyVal = []
    data = data.split('\n')
    logging.info('data--->', data)
    # data = [item.split('=')[0] if '=' in item else item for item in data]
    # data = [item.split('=')[0].split(':')[0] if ('=' in item or ':' in item) else item for item in data]
    # logging.info('\ndata--->', data)
    try:
        # splitting the G colum value as tuple key and value
        for dt in data:
            if re.search(r'[a-zA-Z0-9]+\s=\s[a-zA-Z0-9]+:\s[a-zA-Z]|[a-zA-Z0-9]+=[a-zA-Z0-9]+:[a-zA-Z0-9]+', dt):
                dt = re.sub(r'(:\w+_\w+|:\s[a-zA-Z0-9]+|:[a-zA-Z0-9]+)', "", dt)
                logging.info(f"dt==={dt}")
                keyVal.append((dt.split('=')[0], dt.split('=')[1]))
            elif re.search(r'[a-zA-Z0-9]+\s:\s[a-zA-Z0-9]+=\s[a-zA-Z]+|[a-zA-Z0-9]+:[a-zA-Z0-9]+=[a-zA-Z0-9]+', dt):
                dt = re.sub(r'(=\w+_\w+|=\s[a-zA-Z0-9]+|=[a-zA-Z0-9]+)', "", dt)
                logging.info(f"dt2==={dt}")
                keyVal.append((dt.split(':')[0], dt.split(':')[1]))
            elif re.search(r'[a-zA-Z0-9]+\s:\s[a-zA-Z0-9]+|[a-zA-Z0-9]+:[a-zA-Z0-9]+', dt):
                keyVal.append((dt.split(':')[0], dt.split(':')[1]))
            elif re.search(r'[a-zA-Z0-9]+\s=\s[a-zA-Z0-9]+|[a-zA-Z0-9]+=[a-zA-Z0-9]+', dt):
                keyVal.append((dt.split('=')[0], dt.split('=')[1]))
        logging.info(f"\nkeyVal -> {keyVal}")
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        logging.info(f"\nError in splitKeyValues: {ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
    return keyVal


def remove_elements(var, value):
    return [x for x in var if x!=value]


def dci_extraction(var):
    init_value_keys = ['InitValue=', 'initValue=', 'initvalue=', 'Initvalue=', 'InitValue =', 'initValue =',
                       'initvalue =', 'Initvalue =']
    elements = var.split()
    # Remove the first element if it starts with 'InitValue='
    for init_value in init_value_keys:
        if elements[0].startswith(init_value):
            elements = elements[1:]
    extracted_elements = list(elements)
    logging.info('extracted_elements---->', extracted_elements)
    return extracted_elements


def Init_Value_SS_fiches_Other_value_extraction(variable_E_, variable_G_, ss_fiche_init):
    dic = {
        "InitValue": "",
        "ss_fiches_Init_Value": "",
        "ss_fiches_Other_Value": "",
    }
    logging.info('ss_fiche_init ---->', ss_fiche_init)
    ss_fiche_init = []
    logging.info('variable_G_ ---->', variable_G_)
    logging.info('variable_E_ ---->', variable_E_)
    E_column = dci_extraction(variable_E_)
    logging.info('E_column---->', E_column)
    logging.info('ss_fiche_init ---->', ss_fiche_init)
    buffer_storage = ''
    var = extract_G_col(variable_G_)
    logging.info('var --->', var)
    data = dict(var)
    logging.info('data', data)
    init_value_keys = ['InitValue', 'initValue', 'initvalue', 'Initvalue']
    for key in init_value_keys:
        if key in data:
            logging.info('data --->', data)
            logging.info('key--->', key)
            logging.info('data[key]', data[key])
            buffer_storage = data[key]
            logging.info('buffer_storage', buffer_storage)
            dic['InitValue'] = data[key]  # if needed concatenate--> .split('=')[0].split(':')[0]
    for key, value in data.items():
        for keys_s in init_value_keys:
            if key!=keys_s and value==buffer_storage:
                logging.info('key', key)
                logging.info("Matching value:", value)
                logging.info("Matching value:", (key, value))
                dic['Init_Actual_Value'] = key
    if (dic["Init_Actual_Value"])!='' and len(ss_fiche_init) > 0:
        logging.info('Flow found in Ss_fiches and DCI also process Begin !!!!')
        exact_other_value = remove_elements(ss_fiche_init, dic['Init_Actual_Value'])
        logging.info('exact_other_value-->', exact_other_value)
        dic["Other_Value"] = exact_other_value[0]
        logging.info('dic["Other_Value"]----->', dic["Other_Value"])
    elif (ss_fiche_init==[]) and (dic["Init_Actual_Value"])!='':
        logging.info('Flow found in Dci process Begin !!!!')
        exact_other_value_dci = remove_elements(E_column, dic['Init_Actual_Value'])
        logging.info('exact_other_value_dci-->', exact_other_value_dci)
        dic["Other_Value"] = exact_other_value_dci[0]
        logging.info('dic["Other_Value"]----->', dic["Other_Value"])
    return dic


def search_DCI_Global_P_C(signal):
    DCI_Data = {
        "DCI_global_signal": "",
        "Produced|Consumed": "",
        "DCI_global_E_column": "",
        "DCI_global_G_column": "",
        "Status_Dci_Global": False
    }
    P_C__ = []
    count = int()
    listdir = os.listdir(ICF.getInputFolder())
    if (len(listdir) > 0) or (listdir is not None) or (listdir!=""):
        for lis in listdir:
            if lis.lower().find('dci_global')!=-1 and lis.lower().find('~$')==-1:
                dci_doc = EI.openExcel(ICF.getInputFolder() + "\\" + lis)
                logging.info('signal----->', signal)
                EI.activateSheet(dci_doc, 'MUX')
                MUX_Sig = EI.searchDataInCol(dci_doc.sheets['MUX'], 3, signal)
                if len(MUX_Sig['cellPositions']) > 0:
                    for coord in MUX_Sig['cellPositions']:
                        row, col = coord
                        coordinate_VSM = row, 2
                        coordinated_C_P = row, 11
                        coordinated_E_column = row, 5
                        coordinated_G_column = row, 7
                        dci_signal = str(dci_doc.sheets['MUX'].range(coord).value)
                        DCI_Data["DCI_global_signal"] = dci_signal
                        P_C = str(dci_doc.sheets['MUX'].range(coordinated_C_P).value)
                        VSM = str(dci_doc.sheets['MUX'].range(coordinate_VSM).value)
                        logging.info('VSM---->', VSM)
                        Init_Value = str(dci_doc.sheets['MUX'].range(coordinated_G_column).value)
                        logging.info('Init_Value---->', str(Init_Value))
                        Default_Values = str(dci_doc.sheets['MUX'].range(coordinated_E_column).value)
                        logging.info('Init_Value---->', str(Default_Values))
                        DCI_Data["DCI_global_E_column"] = Default_Values
                        DCI_Data["DCI_global_G_column"] = Init_Value
                        DCI_Data["Status_Dci_Global"] = True
                        if VSM=='VSM':
                            P_C__.append(P_C)
                            count += 1
                    logging.info('P_C__ --->', P_C__)
                    if 'P' in P_C__ and 'C' in P_C__:
                        logging.info('Both P and C are there')
                        default_case1 = 'C'
                        logging.info('@@@@@@@ returning default--> C @@@@@@')
                        DCI_Data["Produced|Consumed"] = default_case1
                        dci_doc.close()
                        return DCI_Data
                    elif 'P' in P_C__:
                        logging.info('Only P is there')
                        logging.info('###### returning P ######')
                        DCI_Data["Produced|Consumed"] = 'P'
                        logging.info('DCI_Data["Produced|Consumed"] = P_C---->', 'P')
                        dci_doc.close()
                        return DCI_Data
                    elif 'C' in P_C__:
                        logging.info('Only C is there')
                        logging.info('###### returning C ######')
                        DCI_Data["Produced|Consumed"] = 'C'
                        logging.info('DCI_Data["Produced|Consumed"] = P_C---->', 'C')
                        dci_doc.close()
                        return DCI_Data
                    else:
                        logging.info('Neither P nor C is there')
                        dci_doc.close()
                        return DCI_Data
                dci_doc.close()
                return DCI_Data


def compare_list(list1, keyword):
    set1 = set(list1)
    set2 = set(keyword)
    if set1.issubset(set2):
        logging.info(f"All elements in {list1} are present in {keyword} ----> {True}")
        return True
    else:
        logging.info(f"Not all elements in {list1} are present in {keyword} ----> {False}")
        return False


def vehicle_mode_extraction(var):
    keyword = ["EMPTY", "USINE", "CONTROLE", "STOCKAGE_TRANSPORT", "APV", "SHOWROOM"]
    key_word = ['ALL', 'all', 'All', 'aLL', 'AlL']
    split_var = []
    if len(var) > 0:
        lis1 = compare_list(var, keyword)
        logging.info('lis1--->', lis1)
        for element in var:
            if (lis1 is False) and (element in key_word):
                logging.info('in')
                logging.info('returning---->keyword--->', keyword)
                return keyword
            elif (lis1 is False) and (element not in key_word):
                logging.info('inn')
                split_element = element.split(',')
                split_var.extend(split_element)
                logging.info('splitted successfully')
            elif (lis1 is True) and (element not in key_word):
                logging.info('innn')
                logging.info('element', element)
                split_var.append(element)
                logging.info('split_var.append(element)---else--->')
        return split_var


def search_ss_fiche(signal):
    global Ss_fiches_doc
    ss_fiche_Data = {
        "SS_fiche_signal": "",
        "D_Colum_Other_Value": [],
        "Status_ss_fiche": False
    }
    listdir = os.listdir(ICF.getInputFolder())
    if (len(listdir) > 0) or (listdir is not None) or (listdir!=""):
        for lis in listdir:
            if lis.find('ss_fiches')!=-1 and lis.find('~$') == -1:
                Ss_fiches_doc = EI.openExcel(ICF.getInputFolder() + "\\" + lis)
                logging.info('signal----->', signal)
                EI.activateSheet(Ss_fiches_doc, 'Matrice de tests')
                Matrice_de_tests = EI.searchDataInCol(Ss_fiches_doc.sheets['Matrice de tests'], 2, signal)
                logging.info('Matrice_de_tests---->', Matrice_de_tests)
                if len(Matrice_de_tests['cellPositions']) > 0:
                    for cellpos in Matrice_de_tests['cellPositions']:
                        row, col = cellpos
                        other_value_coord = row, 4
                        ss_fiche_signal = str(
                            Ss_fiches_doc.sheets['Matrice de tests'].range(Matrice_de_tests['cellPositions'][0]).value)
                        D_column_value = str(Ss_fiches_doc.sheets['Matrice de tests'].range(other_value_coord).value)
                        ss_fiche_Data["SS_fiche_signal"] = ss_fiche_signal
                        ss_fiche_Data["D_Colum_Other_Value"].append(D_column_value)
                        ss_fiche_Data["Status_ss_fiche"] = True
                    logging.info('+++++++ signal present in Sous fiche +++++++')
    Ss_fiches_doc.close()
    return ss_fiche_Data
    # else:
    #     return ss_fiche_Data


def Cases_logic(dci_glob_pc, ss_fiches):
    if (dci_glob_pc["Produced|Consumed"]=='C') and (dci_glob_pc["Produced|Consumed"]!='') and (
            dci_glob_pc["Produced|Consumed"] is not None) and (dci_glob_pc["Status_Dci_Global"]==True) and (
            ss_fiches["Status_ss_fiche"]==False) and (dci_glob_pc is not None):
        logging.info('Case---> 1')
        return 'Case1'
    elif (dci_glob_pc["Produced|Consumed"]=='P') and (dci_glob_pc["Produced|Consumed"]!='') and (
            dci_glob_pc["Produced|Consumed"] is not None) and (ss_fiches["Status_ss_fiche"]==True) and (
            dci_glob_pc["Status_Dci_Global"]==True) and (dci_glob_pc is not None):
        logging.info('Case---> 2')
        return 'Case2'
    elif (dci_glob_pc["Produced|Consumed"]=='P') and (dci_glob_pc["Produced|Consumed"]!='') and (
            dci_glob_pc["Produced|Consumed"] is not None) and (dci_glob_pc["Status_Dci_Global"]==True) and (
            ss_fiches["Status_ss_fiche"]==False) and (
            ss_fiches["Status_ss_fiche"]!=True) and (dci_glob_pc is not None):
        logging.info('Case---> 3')
        return 'Case3'
    elif (dci_glob_pc["Produced|Consumed"]=='') and (dci_glob_pc["Produced|Consumed"] is None) and (
            dci_glob_pc["Status_Dci_Global"]==False) and (
            ss_fiches["Status_ss_fiche"]==True) and (ss_fiches["Status_ss_fiche"]!=False) and (
            dci_glob_pc is None) and (ss_fiches is not None):
        logging.info('Case---> 4')
        return 'Case4'
    elif ((dci_glob_pc["Produced|Consumed"]=='') and (ss_fiches["Status_ss_fiche"]==False) and (
            dci_glob_pc["Status_Dci_Global"]==False)) or ((dci_glob_pc is None) and (ss_fiches is None)):
        logging.info('Case---> 5')
        return 'Case5'
    else:
        logging.info('None of the above')
        return -1


def remove_common(a, b):
    listt = []
    for i in a[:]:
        if i in b:
            a.remove(i)
            listt.append(i)
            b.remove(i)
    logging.info("list1 : ", a)
    logging.info("list2 : ", b)
    mylist = b
    sorted_list = sorted(mylist, key=lambda x: x[0])
    logging.info('sorted_list---->', sorted_list)
    logging.info('listt--->', listt)
    return listt, sorted_list


def find_the_ranges(tuple_list):
    # condition_initials
    final_range = []
    start = tuple_list[0][0]
    end = tuple_list[1][0]
    numbers_range = range(start, end + 1)
    rangenum = list(numbers_range)
    logging.info('rangenum--->', rangenum)
    for range0 in rangenum:
        final_range.append((range0, 5))
    logging.info(final_range)
    return final_range


def extract_common_elements(lis1, lis2):
    common_elements = set(lis1).intersection(lis2)
    return common_elements


def create_groups(keywords, sheet):
    cell_pos_dic = {
        "condition_initials": [],
        "CORPS_DE_TEST": [],
        "RETOUR_AUX": []
    }

    # Get the cell positions of keywords
    parameter_de_entreee = EI.searchDataInExcel(sheet, "", "PARAMETRE D'ENTREE")
    CONDITIONS_INITIALES_pos = EI.searchDataInExcel(sheet, "", keywords[0])
    CORPS_DE_TEST_pos = EI.searchDataInExcel(sheet, "", keywords[1])
    RETOUR_AUX_pos = EI.searchDataInExcel(sheet, "", keywords[2])
    last_cell_pos = EI.searchDataInExcel(sheet, "", "Historique&")

    # Group the headings
    condition_initials = sum([CONDITIONS_INITIALES_pos["cellPositions"], CORPS_DE_TEST_pos['cellPositions']], [])
    condition_initials = find_the_ranges(condition_initials)
    condition_initials = list(extract_common_elements(parameter_de_entreee["cellPositions"], condition_initials))
    cell_pos_dic["condition_initials"] = sorted(condition_initials, key=lambda x: x[0])

    CORPS_DE = sum([CORPS_DE_TEST_pos["cellPositions"], RETOUR_AUX_pos['cellPositions']], [])
    CORPS_DE = find_the_ranges(CORPS_DE)
    CORPS_DE = list(extract_common_elements(parameter_de_entreee["cellPositions"], CORPS_DE))
    cell_pos_dic["CORPS_DE_TEST"] = sorted(CORPS_DE, key=lambda x: x[0])

    RETOUR_AUX = sum([RETOUR_AUX_pos["cellPositions"], last_cell_pos['cellPositions']], [])
    RETOUR_AUX = find_the_ranges(RETOUR_AUX)
    RETOUR_AUX = list(extract_common_elements(parameter_de_entreee["cellPositions"], RETOUR_AUX))
    cell_pos_dic["RETOUR_AUX"] = sorted(RETOUR_AUX, key=lambda x: x[0])

    logging.info("cell_pos_dic['condition_initials']---->", cell_pos_dic["condition_initials"])
    logging.info("cell_pos_dic['CORPS_DE_TEST']----->", cell_pos_dic["CORPS_DE_TEST"])
    logging.info("cell_pos_dic['RETOUR_AUX']----->", cell_pos_dic["RETOUR_AUX"])
    return cell_pos_dic


def remove_common_element_list_tuples(a, b):
    for i in a[:]:
        if i in b:
            a.remove(i)
            b.remove(i)
    logging.info("list1 : ", a)
    logging.info("list2 : ", b)
    return a


def updated_cell_positions(present, final):
    logging.info('present---->', present)
    logging.info('final---->', final)
    CORPS_DE = list(remove_common_element_list_tuples(present, final))
    CORPS_DE = sorted(CORPS_DE, key=lambda x: x[0])
    logging.info('CORPS_DE-->', CORPS_DE)
    if len(CORPS_DE)==0:
        logging.info('len(CORPS_DE) == 0 --->', final)
        return
    else:
        logging.info('len(CORPS_DE) == 0 -else-->', CORPS_DE)
        return CORPS_DE


# #########################################################################################################################################################################################################
# Case 1 supporting functions


def case1_mode_config(cellpos, sheet, vehic_mode):
    row, col = cellpos[0]
    logging.info('cellpos[0]-innnn-->', cellpos[0])
    EI.setDataFromCell(sheet, (row - 1, 1),
                       f"BUT DE L'ETAPE : Set vehicle mode to {vehic_mode}")
    EI.setDataFromCell(sheet, (row + 1, col), "$MODE_CONFIG_VHL")
    EI.setDataFromCell(sheet, (row + 1, col - 1), "FONCTION")
    EI.setDataFromCell(sheet, (row + 1, col - 2), f'Set the vehicle mode to {vehic_mode}')
    EI.setDataFromCell(sheet, (row + 1, col + 1), vehic_mode)
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


def case1_life_cycle(cellpos, sheet, signal, init_value):
    row, col = cellpos[0]
    EI.setDataFromCell(sheet, (row - 1, 1),
                       f"BUT DE L'ETAPE : Set the signal {signal} at {init_value} before verifing the defect")
    EI.setDataFromCell(sheet, (row + 1, col - 2), f'Set the signal {signal} at value {init_value}')
    EI.setDataFromCell(sheet, (row + 1, col), f"${signal}")
    EI.setDataFromCell(sheet, (row + 1, col + 1), init_value)
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


def case1_involved_flow(cellpos, sheet, involved_flow, other_value, other_name):
    row, col = cellpos[0]
    EI.setDataFromCell(sheet, (row - 1, 1), f"BUT DE L'ETAPE : Put the flow {involved_flow} in {other_name}")
    EI.setDataFromCell(sheet, (row + 1, col), f"${involved_flow}")
    EI.setDataFromCell(sheet, (row + 1, col - 2), f"{involved_flow} take the value {other_name}")
    EI.setDataFromCell(sheet, (row + 1, col + 1), other_value)
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


########################################################################################################
# we need to confirm that flow_name should be same from step QIA of PARAM --> step III. --> Column H.
# For Diag ----> case1_Diag_flow
########################################################################################################


def case1_diag_flow(cellpos, sheet, involved_flow, input_did_flow, output_did_flow):
    row, col = cellpos[0]
    EI.setDataFromCell(sheet, (row - 1, 1),
                       "BUT DE L'ETAPE : Send the request for the diagnostic device and verify the response")
    EI.setDataFromCell(sheet, (row + 1, col - 2), f"Use the DID {input_did_flow} to request the DIAG")
    EI.setDataFromCell(sheet, (row + 1, col - 1), f"DIAG")
    EI.setDataFromCell(sheet, (row + 1, col), f"$REQ_{involved_flow}")
    EI.setDataFromCell(sheet, (row + 1, col + 4), f"Use the DID {output_did_flow} as response of the DIAG")
    EI.setDataFromCell(sheet, (row + 1, col + 5), f"DIAG")
    EI.setDataFromCell(sheet, (row + 1, col + 6), f"$REP_{involved_flow}")
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


def case1_no_diag(cellpos, sheet):
    row, col = cellpos[0]
    EI.setDataFromCell(sheet, (row - 1, 1), f"BUT DE L'ETAPE : Close the DIAG Session")
    EI.setDataFromCell(sheet, (row + 1, col - 2), f'Close the DIAG Session')
    EI.setDataFromCell(sheet, (row + 1, col - 1), f'NO_DIAG')
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


def case1_retour_de_tests_Set_the_signal(cellpos, sheet, signal, init_value):
    row, col = cellpos[0]
    EI.setDataFromCell(sheet, (row - 1, 1), f"BUT DE L'ETAPE : Set the signal ${signal} at {init_value}")
    EI.setDataFromCell(sheet, (row + 1, col), f'Set the signal ${signal} at {init_value}')
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


def case1_retour_de_tests_ARRET(cellpos, sheet):
    row, col = cellpos[0]
    EI.setDataFromCell(sheet, (row - 1, 1), "BUT DE L'ETAPE : Put on arret")
    EI.setDataFromCell(sheet, (row + 1, col), "$ETAT_PRINCIP_SEV")
    EI.setDataFromCell(sheet, (row + 1, col + 1), "ARRET")
    EI.setDataFromCell(sheet, (row + 1, col - 1), "FONCTION")
    EI.setDataFromCell(sheet, (row + 1, col - 2), 'Put on arret')
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


def condition_initials_default_step(sheet, row, col):
    EI.setDataFromCell(sheet, (row - 1, 1), "BUT DE L'ETAPE : Put the CONTACT")
    EI.setDataFromCell(sheet, (row + 1, col - 2), "Put on contact")
    EI.setDataFromCell(sheet, (row + 1, col - 1), "FONCTION")
    EI.setDataFromCell(sheet, (row + 1, col), "$ETAT_PRINCIP_SEV")
    EI.setDataFromCell(sheet, (row + 1, col + 1), "CONTACT")


# #########################################################################################################################################################################################################
# case 1 Driver function

def case1_External_Input_Signal(sheet, macro, vcal_mod, life_cyc, involved_flow, init_value, DCI_other_value, DCI_other_name, DCI_init_value, DCI_init_name, input_did_flow, output_did_flow):
    keywords = [('---- CONDITIONS INITIALES ----', 1), ('---- CORPS DE TEST ----', 2),
                ('---- RETOUR AUX CONDITIONS INITIALES ----', 3)]
    Keywords = ['---- CONDITIONS INITIALES ----', '---- CORPS DE TEST ----',
                '---- RETOUR AUX CONDITIONS INITIALES ----']
    for keyword, num in keywords:
        keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        if keyword_cell_pos['count'] > 0:
            for cellPos in keyword_cell_pos['cellPositions']:
                row, col = cellPos
                if num==1:
                    logging.info("\n\nrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr")
                    TPM.addInitialContionsStep(macro)
                elif num==2:
                    if vcal_mod!="EMPTY":
                        for i in range(7):
                            TPM.addCorpDeTestStep(macro)
                    else:
                        for i in range(6):
                            TPM.addCorpDeTestStep(macro)
                elif num==3:
                    if vcal_mod!="EMPTY":
                        for i in range(2):
                            TPM.addRetourContionsStep(macro)
                    else:
                        TPM.addRetourContionsStep(macro)
        new_val_cel_pos = EI.searchDataInExcel(sheet, "", "PARAMETRE D'ENTREE")
        condition_initials_result = create_groups(Keywords, sheet)
        logging.info("new_val_cel_pos ", new_val_cel_pos)
        if new_val_cel_pos['count'] > 0:
            if num==1:
                condition_initials_result = create_groups(Keywords, sheet)
                logging.info('condition_initials_result---->', condition_initials_result)
                for cell_pos in condition_initials_result["condition_initials"]:
                    row, col = cell_pos
                    condition_initials_default_step(sheet, row, col)
            elif num==2:
                Corp_De_Test_result = create_groups(Keywords, sheet)
                logging.info('Corp_De_Test_result---->', Corp_De_Test_result)
                if len(condition_initials_result["CORPS_DE_TEST"]) >= 0:
                    if vcal_mod!="EMPTY":
                        condition_initials_result["CORPS_DE_TEST"] = case1_mode_config(
                            condition_initials_result["CORPS_DE_TEST"], sheet, vcal_mod)
                        # for sig in life_cyc:
                        #     for tup in sig:
                        #         sig1, sig2 = tup
                        #         logging.info('sig1 --->', sig1)
                        #         logging.info('sig2 --->', sig2)
                        #         condition_initials_result["CORPS_DE_TEST"] = case1_life_cycle(
                        #             condition_initials_result["CORPS_DE_TEST"], sheet, sig1, sig2)
                    condition_initials_result["CORPS_DE_TEST"] = case1_involved_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, DCI_other_value, DCI_other_name)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
                    condition_initials_result["CORPS_DE_TEST"] = case1_involved_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, DCI_init_value, DCI_init_name)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
            elif num==3:
                Retour_Contions_result = create_groups(Keywords, sheet)
                logging.info('Retour_Contions_result---->', Retour_Contions_result)
                if len(condition_initials_result["RETOUR_AUX"]) >= 0:
                    # condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_Set_the_signal(
                    #     condition_initials_result["RETOUR_AUX"], sheet, involved_flow, init_value)
                    if vcal_mod!="EMPTY":
                        Default_value = 'CLIENT'
                        condition_initials_result["RETOUR_AUX"] = case1_mode_config(
                            condition_initials_result["RETOUR_AUX"], sheet, Default_value)
                    condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_ARRET(
                        condition_initials_result["RETOUR_AUX"], sheet)


# #########################################################################################################################################################################################################
# Case 2 supporting functions

def case2_involved_flow_sos_fiches(cellpos, sheet, involved_flow, other_value):
    row, col = cellpos[0]
    EI.setDataFromCell(sheet, (row - 1, 1), f"BUT DE L'ETAPE : Put the flow {involved_flow} in {other_value}")
    EI.setDataFromCell(sheet, (row + 1, col), f"${involved_flow}")
    EI.setDataFromCell(sheet, (row + 1, col - 1), "FONCTION")
    EI.setDataFromCell(sheet, (row + 1, col - 2), f"{involved_flow} take the value {other_value}")
    EI.setDataFromCell(sheet, (row + 1, col + 1), other_value)
    final_cell__pos = updated_cell_positions(cellpos, [cellpos[0]])
    return final_cell__pos


# #########################################################################################################################################################################################################
# case 2 Driver function


def case2_External_Output_Signal(sheet, macro, vcal_mod, life_cyc, involved_flow, other_value, actual_init_value,
                                 input_did_flow, output_did_flow):
    output = []
    keywords = [('---- CONDITIONS INITIALES ----', 1), ('---- CORPS DE TEST ----', 2),
                ('---- RETOUR AUX CONDITIONS INITIALES ----', 3)]
    Keywords = ['---- CONDITIONS INITIALES ----', '---- CORPS DE TEST ----',
                '---- RETOUR AUX CONDITIONS INITIALES ----']
    for keyword, num in keywords:
        keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        if keyword_cell_pos['count'] > 0:
            for cellPos in keyword_cell_pos['cellPositions']:
                row, col = cellPos
                if num==1:
                    logging.info("\n\nrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr")
                    TPM.addInitialContionsStep(macro)
                elif num==2:
                    if vcal_mod!="EMPTY":
                        for i in range(7):
                            TPM.addCorpDeTestStep(macro)
                    else:
                        for i in range(6):
                            TPM.addCorpDeTestStep(macro)
                elif num==3:
                    if vcal_mod!="EMPTY":
                        for i in range(2):
                            TPM.addRetourContionsStep(macro)
                    else:
                        TPM.addRetourContionsStep(macro)
        new_val_cel_pos = EI.searchDataInExcel(sheet, "", "PARAMETRE D'ENTREE")
        condition_initials_result = create_groups(Keywords, sheet)
        logging.info("new_val_cel_pos ", new_val_cel_pos)
        if new_val_cel_pos['count'] > 0:
            if num==1:
                condition_initials_result = create_groups(Keywords, sheet)
                logging.info('condition_initials_result---->', condition_initials_result)
                for cell_pos in condition_initials_result["condition_initials"]:
                    row, col = cell_pos
                    condition_initials_default_step(sheet, row, col)
            elif num==2:
                Corp_De_Test_result = create_groups(Keywords, sheet)
                logging.info('Corp_De_Test_result---->', Corp_De_Test_result)
                if len(condition_initials_result["CORPS_DE_TEST"]) >= 0:
                    # for sig in life_cyc:
                    #     for tup in sig:
                    #         sig1, sig2 = tup
                    #         logging.info('sig1 --->', sig1)
                    #         logging.info('sig2 --->', sig2)
                    #         condition_initials_result["CORPS_DE_TEST"] = case1_life_cycle(
                    #             condition_initials_result["CORPS_DE_TEST"], sheet, sig1, sig2)
                    if vcal_mod!="EMPTY":
                        condition_initials_result["CORPS_DE_TEST"] = case1_mode_config(
                            condition_initials_result["CORPS_DE_TEST"], sheet, vcal_mod)
                    condition_initials_result["CORPS_DE_TEST"] = case2_involved_flow_sos_fiches(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, other_value)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
                    condition_initials_result["CORPS_DE_TEST"] = case2_involved_flow_sos_fiches(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, actual_init_value)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
            elif num==3:
                Retour_Contions_result = create_groups(Keywords, sheet)
                logging.info('Retour_Contions_result---->', Retour_Contions_result)
                if len(condition_initials_result["RETOUR_AUX"]) >= 0:
                    # condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_Set_the_signal(
                    #     condition_initials_result["RETOUR_AUX"], sheet, involved_flow, init_value)
                    if vcal_mod!="EMPTY":
                        Default_value = 'CLIENT'
                        condition_initials_result["RETOUR_AUX"] = case1_mode_config(
                            condition_initials_result["RETOUR_AUX"], sheet, Default_value)
                    condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_ARRET(
                        condition_initials_result["RETOUR_AUX"], sheet)


# #########################################################################################################################################################################################################
# Case 3 supporting functions

# #########################################################################################################################################################################################################
# case 3 Driver function
#                     case1_External_Input_Signal(tpBook.sheets[sheet_name], macro, vcal_mod, life_cycle, inv_flow, init_value, DCI_other_value, DCI_other_name, DCI_init_value, DCI_init_name, input_did_flow, output_did_flow)

def case3_External_Input_Signal(sheet, macro, vcal_mod, life_cyc, involved_flow, init_value, DCI_other_value, DCI_other_name, DCI_init_value, DCI_init_name, input_did_flow, output_did_flow):
    output = []
    keywords = [('---- CONDITIONS INITIALES ----', 1), ('---- CORPS DE TEST ----', 2),
                ('---- RETOUR AUX CONDITIONS INITIALES ----', 3)]
    Keywords = ['---- CONDITIONS INITIALES ----', '---- CORPS DE TEST ----',
                '---- RETOUR AUX CONDITIONS INITIALES ----']
    for keyword, num in keywords:
        keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        if keyword_cell_pos['count'] > 0:
            for cellPos in keyword_cell_pos['cellPositions']:
                row, col = cellPos
                if num==1:
                    logging.info("\n\nrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr")
                    TPM.addInitialContionsStep(macro)
                elif num==2:
                    if vcal_mod!="EMPTY":
                        for i in range(7):
                            TPM.addCorpDeTestStep(macro)
                    else:
                        for i in range(6):
                            TPM.addCorpDeTestStep(macro)
                elif num==3:
                    if vcal_mod!="EMPTY":
                        for i in range(2):
                            TPM.addRetourContionsStep(macro)
                    else:
                        TPM.addRetourContionsStep(macro)
        new_val_cel_pos = EI.searchDataInExcel(sheet, "", "PARAMETRE D'ENTREE")
        condition_initials_result = create_groups(Keywords, sheet)
        logging.info("new_val_cel_pos ", new_val_cel_pos)
        if new_val_cel_pos['count'] > 0:
            if num==1:
                condition_initials_result = create_groups(Keywords, sheet)
                logging.info('condition_initials_result---->', condition_initials_result)
                for cell_pos in condition_initials_result["condition_initials"]:
                    row, col = cell_pos
                    condition_initials_default_step(sheet, row, col)
            elif num==2:
                Corp_De_Test_result = create_groups(Keywords, sheet)
                logging.info('Corp_De_Test_result---->', Corp_De_Test_result)
                if len(condition_initials_result["CORPS_DE_TEST"]) >= 0:
                    # for sig in life_cyc:
                    #     for tup in sig:
                    #         sig1, sig2 = tup
                    #         logging.info('sig1 --->', sig1)
                    #         logging.info('sig2 --->', sig2)
                    #         condition_initials_result["CORPS_DE_TEST"] = case1_life_cycle(
                    #             condition_initials_result["CORPS_DE_TEST"], sheet, sig1, sig2)
                    if vcal_mod!="EMPTY":
                        condition_initials_result["CORPS_DE_TEST"] = case1_mode_config(
                            condition_initials_result["CORPS_DE_TEST"], sheet, vcal_mod)
                    condition_initials_result["CORPS_DE_TEST"] = case1_involved_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, DCI_other_value, DCI_other_name)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
                    condition_initials_result["CORPS_DE_TEST"] = case1_involved_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, DCI_init_value, DCI_init_name)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
            elif num==3:
                Retour_Contions_result = create_groups(Keywords, sheet)
                logging.info('Retour_Contions_result---->', Retour_Contions_result)
                if len(condition_initials_result["RETOUR_AUX"]) >= 0:
                    # condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_Set_the_signal(
                    #     condition_initials_result["RETOUR_AUX"], sheet, involved_flow, init_value)
                    if vcal_mod!="EMPTY":
                        Default_value = 'CLIENT'
                        condition_initials_result["RETOUR_AUX"] = case1_mode_config(
                            condition_initials_result["RETOUR_AUX"], sheet, Default_value)
                    condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_ARRET(
                        condition_initials_result["RETOUR_AUX"], sheet)


# #########################################################################################################################################################################################################
# Case 4 supporting functions

# #########################################################################################################################################################################################################
# case 4 Driver function


def case4_Internal_Signal(sheet, macro, vcal_mod, life_cyc, involved_flow, init_value, other_value, actual_init_value,
                          input_did_flow,
                          output_did_flow):
    output = []
    keywords = [('---- CONDITIONS INITIALES ----', 1), ('---- CORPS DE TEST ----', 2),
                ('---- RETOUR AUX CONDITIONS INITIALES ----', 3)]
    Keywords = ['---- CONDITIONS INITIALES ----', '---- CORPS DE TEST ----',
                '---- RETOUR AUX CONDITIONS INITIALES ----']
    for keyword, num in keywords:
        keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        if keyword_cell_pos['count'] > 0:
            for cellPos in keyword_cell_pos['cellPositions']:
                row, col = cellPos
                if num==1:
                    logging.info("\n\nrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr")
                    TPM.addInitialContionsStep(macro)
                elif num==2:
                    if vcal_mod!="EMPTY":
                        for i in range(7):
                            TPM.addCorpDeTestStep(macro)
                    else:
                        for i in range(6):
                            TPM.addCorpDeTestStep(macro)
                elif num==3:
                    if vcal_mod!="EMPTY":
                        for i in range(2):
                            TPM.addRetourContionsStep(macro)
                    else:
                        TPM.addRetourContionsStep(macro)
        new_val_cel_pos = EI.searchDataInExcel(sheet, "", "PARAMETRE D'ENTREE")
        condition_initials_result = create_groups(Keywords, sheet)
        logging.info("new_val_cel_pos ", new_val_cel_pos)
        if new_val_cel_pos['count'] > 0:
            if num==1:
                condition_initials_result = create_groups(Keywords, sheet)
                logging.info('condition_initials_result---->', condition_initials_result)
                for cell_pos in condition_initials_result["condition_initials"]:
                    row, col = cell_pos
                    condition_initials_default_step(sheet, row, col)
            elif num==2:
                Corp_De_Test_result = create_groups(Keywords, sheet)
                logging.info('Corp_De_Test_result---->', Corp_De_Test_result)
                if len(condition_initials_result["CORPS_DE_TEST"]) >= 0:
                    # for sig in life_cyc:
                    #     for tup in sig:
                    #         sig1, sig2 = tup
                    #         logging.info('sig1 --->', sig1)
                    #         logging.info('sig2 --->', sig2)
                    #         condition_initials_result["CORPS_DE_TEST"] = case1_life_cycle(
                    #             condition_initials_result["CORPS_DE_TEST"], sheet, sig1, sig2)
                    if vcal_mod!="EMPTY":
                        condition_initials_result["CORPS_DE_TEST"] = case1_mode_config(
                            condition_initials_result["CORPS_DE_TEST"], sheet, vcal_mod)
                    condition_initials_result["CORPS_DE_TEST"] = case2_involved_flow_sos_fiches(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, other_value)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
                    condition_initials_result["CORPS_DE_TEST"] = case2_involved_flow_sos_fiches(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, actual_init_value)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
            elif num==3:
                Retour_Contions_result = create_groups(Keywords, sheet)
                logging.info('Retour_Contions_result---->', Retour_Contions_result)
                if len(condition_initials_result["RETOUR_AUX"]) >= 0:
                    # condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_Set_the_signal(
                    #     condition_initials_result["RETOUR_AUX"], sheet, involved_flow, init_value)
                    if vcal_mod!="EMPTY":
                        Default_value = 'CLIENT'
                        condition_initials_result["RETOUR_AUX"] = case1_mode_config(
                            condition_initials_result["RETOUR_AUX"], sheet, Default_value)
                    condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_ARRET(
                        condition_initials_result["RETOUR_AUX"], sheet)


# #########################################################################################################################################################################################################
# Case 5 supporting functions

# #########################################################################################################################################################################################################
# case 5 Driver function


def case5_Internal_Signal(sheet, macro, vcal_mod, life_cyc, involved_flow, init_value, other_value, input_did_flow,
                          output_did_flow):
    output = []
    keywords = [('---- CONDITIONS INITIALES ----', 1), ('---- CORPS DE TEST ----', 2),
                ('---- RETOUR AUX CONDITIONS INITIALES ----', 3)]
    Keywords = ['---- CONDITIONS INITIALES ----', '---- CORPS DE TEST ----',
                '---- RETOUR AUX CONDITIONS INITIALES ----']
    for keyword, num in keywords:
        keyword_cell_pos = EI.searchDataInExcel(sheet, "", keyword)
        logging.info("initial_cell_pos - ", keyword_cell_pos)
        if keyword_cell_pos['count'] > 0:
            for cellPos in keyword_cell_pos['cellPositions']:
                row, col = cellPos
                if num==1:
                    logging.info("\n\nrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr")
                    TPM.addInitialContionsStep(macro)
                elif num==2:
                    if vcal_mod!="EMPTY":
                        for i in range(3):
                            TPM.addCorpDeTestStep(macro)
                    else:
                        for i in range(2):
                            TPM.addCorpDeTestStep(macro)
                elif num==3:
                    if vcal_mod!="EMPTY":
                        for i in range(2):
                            TPM.addRetourContionsStep(macro)
                    else:
                        TPM.addRetourContionsStep(macro)
        new_val_cel_pos = EI.searchDataInExcel(sheet, "", "PARAMETRE D'ENTREE")
        condition_initials_result = create_groups(Keywords, sheet)
        logging.info("new_val_cel_pos ", new_val_cel_pos)
        if new_val_cel_pos['count'] > 0:
            if num==1:
                condition_initials_result = create_groups(Keywords, sheet)
                logging.info('condition_initials_result---->', condition_initials_result)
                for cell_pos in condition_initials_result["condition_initials"]:
                    row, col = cell_pos
                    condition_initials_default_step(sheet, row, col)
            elif num==2:
                Corp_De_Test_result = create_groups(Keywords, sheet)
                logging.info('Corp_De_Test_result---->', Corp_De_Test_result)
                if len(condition_initials_result["CORPS_DE_TEST"]) >= 0:
                    # for sig in life_cyc:
                    #     for tup in sig:
                    #         sig1, sig2 = tup
                    #         logging.info('sig1 --->', sig1)
                    #         logging.info('sig2 --->', sig2)
                    #         condition_initials_result["CORPS_DE_TEST"] = case1_life_cycle(
                    #             condition_initials_result["CORPS_DE_TEST"], sheet, sig1, sig2)
                    if vcal_mod!="EMPTY":
                        condition_initials_result["CORPS_DE_TEST"] = case1_mode_config(
                            condition_initials_result["CORPS_DE_TEST"], sheet, vcal_mod)
                    condition_initials_result["CORPS_DE_TEST"] = case1_diag_flow(
                        condition_initials_result["CORPS_DE_TEST"], sheet, involved_flow, input_did_flow,
                        output_did_flow)
                    condition_initials_result["CORPS_DE_TEST"] = case1_no_diag(
                        condition_initials_result["CORPS_DE_TEST"], sheet)
            elif num==3:
                logging.info('in side loop Retour_Contions')
                Retour_Contions_result = create_groups(Keywords, sheet)
                logging.info('Retour_Contions_result---->', Retour_Contions_result)
                if len(condition_initials_result["RETOUR_AUX"]) >= 0:
                    # condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_Set_the_signal(
                    #     condition_initials_result["RETOUR_AUX"], sheet, involved_flow, init_value)
                    if vcal_mod!="EMPTY":
                        Default_value = 'CLIENT'
                        condition_initials_result["RETOUR_AUX"] = case1_mode_config(
                            condition_initials_result["RETOUR_AUX"], sheet, Default_value)
                    condition_initials_result["RETOUR_AUX"] = case1_retour_de_tests_ARRET(
                        condition_initials_result["RETOUR_AUX"], sheet)


# #########################################################################################################################################################################################################
# data_dic---> {'Involved flow': 'GPE_PUSH_MODE', 'Parameter description': 'Reading on the push MODE',
# 'Life cycle': '', 'Vehicle mode': 'All'}

def diag_FT_creation(tpBook, ReqName, ReqVer, update_sheet, rqIDs, feps):
    # global sheet_name, macro, vcal_mod
    Case = update_sheet["CASES"]
    inv_flow = update_sheet["Involved_Flow"]
    veh_mode = update_sheet["Vehicle_Mode"]
    life_cycle = update_sheet["Life_Cycle"]
    init_value = update_sheet["INIT_Value"]
    ss_fiches_init_value = update_sheet["Ss_fiche_init_value"]
    logging.info('ss_fiches_init_value------>', ss_fiches_init_value)
    ss_fiches_other_value = update_sheet["ss_fiches_Other_Value"]
    logging.info('ss_fiches_other_value------>', ss_fiches_other_value)
    DCI_other_value = update_sheet["DCI_Other_Value"]
    logging.info('DCI_other_value------>', DCI_other_value)
    DCI_other_name = update_sheet["DCI_Other_Name"]
    logging.info('DCI_other_name------>', DCI_other_name)
    DCI_init_value = update_sheet["DCI_init_value"]
    logging.info('DCI_init_value------>', DCI_init_value)
    DCI_init_name = update_sheet["DCI_init_Name"]
    logging.info('DCI_init_name------>', DCI_init_name)
    input_did_flow = update_sheet["DID_CODE"]["Input_DID_FLOW"]
    output_did_flow = update_sheet["DID_CODE"]["Output_DID_FLOW"]
    logging.info('veh_mode---->', veh_mode)
    if len(veh_mode) > 0:
        for vcal_mod in veh_mode:
            logging.info('veh_mode---->', veh_mode)
            logging.info('vcal_mod---->', vcal_mod)
            logging.info("fffffffffffffffffffffffffffffffffffff")
            # tpBook = EI.openTestPlan()
            # tpBook.activate()
            # KMS.showWindow(tpBook.name.split('.')[0])
            macro = EI.getTestPlanAutomationMacro()
            TPM.selectTpWritterProfile(macro)
            TPM.selectTestSheetAdd(macro)
            logging.info('tpBook', tpBook)
            logging.info("tpBook.sheets.active ", tpBook.sheets.active)
            logging.info("tpBook.sheets.active.name ", tpBook.sheets.active.name)
            sheet_name = tpBook.sheets.active.name
            logging.info("Fill FT for Diag req function called...................")
            # Need to implement function to get the C2 and C3 value
            logging.info("sheet_name -->", sheet_name)
            sheet_name_new = ""
            thematicLines = []
            for ts in tpBook.sheets:
                # logging.info(ts)
                if re.search("^OLD_", ts.name.upper()):
                    ts_name = re.sub("^OLD_", "", ts.name.upper())
                    if ts_name==sheet_name:
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
            # new_reqName, new_reqVer = NRH.getReqVer(req)
            logging.info("ReqName = ", ReqName)
            logging.info("ReqVer = ", ReqVer)
            # if req.find("(") == -1:
            req = ReqName + "(" + str(ReqVer) + ")"
            if tpBook.sheets[sheet_name].range(reqs_col).value!="" and tpBook.sheets[sheet_name].range(
                    reqs_col).value is not None:
                # combine the new req with existing req in C4 column
                combine_req = NRH.combineValues(tpBook.sheets[sheet_name], req, reqs_col, '|')
            else:
                combine_req = req

            thematicLines = NRH.getReqThematic(ReqName, ReqVer, rqIDs, feps)
            # thematicLines.append(thematics)
            logging.info("thematicLines ", thematicLines)
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
                if Case=='Case1':
                    logging.info('Case1--->', Case)
                    case1_External_Input_Signal(tpBook.sheets[sheet_name], macro, vcal_mod, life_cycle,
                                                inv_flow, init_value, DCI_other_value, DCI_other_name, DCI_init_value, DCI_init_name, input_did_flow,
                                                output_did_flow)
                elif Case=='Case2':
                    logging.info('Case2--->', Case)
                    case2_External_Output_Signal(tpBook.sheets[sheet_name], macro, vcal_mod, life_cycle,
                                                 inv_flow, ss_fiches_other_value, ss_fiches_init_value, input_did_flow,
                                                 output_did_flow)
                elif Case=='Case3':
                    logging.info('Case3--->', Case)
                    case3_External_Input_Signal(tpBook.sheets[sheet_name], macro, vcal_mod, life_cycle,
                                                inv_flow, init_value, DCI_other_value, DCI_other_name, DCI_init_value, DCI_init_name, input_did_flow,
                                                output_did_flow)
                elif Case=='Case4':
                    logging.info('Case4--->', Case)
                    case4_Internal_Signal(tpBook.sheets[sheet_name], macro, vcal_mod, life_cycle, inv_flow, init_value,
                                          ss_fiches_other_value, ss_fiches_init_value, input_did_flow, output_did_flow)
                elif Case=='Case5':
                    logging.info('Case5--->', Case)
                    case5_Internal_Signal(tpBook.sheets[sheet_name], macro, vcal_mod, life_cycle, inv_flow, init_value,
                                          DCI_other_value, input_did_flow, output_did_flow)
                if (Case=='Case1') or (Case=='Case2') or (Case=='Case3') or (Case=='Case4') or (Case=='Case5'):
                    logging.info('innnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn')
                    try:
                        BL.fillHistoryAndTrigram(tpBook.sheets[sheet_name], "Created new sheet")
                        logging.info(f"reqreq111 {req}")
                        if req.find("("):
                            splitRequirements = req.split("(")
                            Requirement = splitRequirements[0]
                        else:
                            splitRequirements = req.split(" ")
                            Requirement = splitRequirements[0]
                        new_req_pos = EI.searchDataInCol(tpBook.sheets['Impact'], 1, str(Requirement).strip())
                        logging.info("new_req_pos ", new_req_pos)
                        if new_req_pos['count'] > 0:
                            ts_col = 4
                            comment_col = 5
                            for cell_pos in new_req_pos['cellPositions']:
                                x, y = cell_pos
                                logging.info(f"cell_pos {cell_pos}")
                                EI.setDataFromCell(tpBook.sheets['Impact'], (x, comment_col),
                                                   "New Requirement.")
                    except Exception as e:
                        logging.info(f"\nError in filling history for new requirement.. {e}")
                # exit()
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                logging.info(f'''+++++++++++++ Error: {e} 
                    line number : {exc_tb.tb_lineno}+++++++++++++''')


def did_code(reqData):
    DIDVal = ''
    final_set = {
        "Input_DID_FLOW": '',
        "Output_DID_FLOW": ''
    }
    if (reqData != '') or (reqData is not None):
        if re.search("VSM-[A-Z]{1,2}[A-Z a-z0-9]{1,10}-[0-9]{2}", reqData):
            # VSM-U0131-81
            extracted_DID = re.findall("VSM-[A-Z]{1,2}[A-Z a-z0-9]{1,10}-[0-9]{2}", reqData)
            DID_value = QP.convertDID(extracted_DID[0])
            if DID_value!="" and DID_value is not None:
                DIDVal = QP.split_did_with_dot(DID_value)
            final_set["Input_DID_FLOW"] = f"22.{DIDVal}"
            final_set["Output_DID_FLOW"] = f"62.{DIDVal}"
        elif re.search("VSM-[A-Z a-z0-9]{1,20}", reqData):
            DID_value = did_code_comment_formatting(reqData)
            final_set["Input_DID_FLOW"] = f"22.{DID_value}"
            final_set["Output_DID_FLOW"] = f"62.{DID_value}"
    else:
        logging.info('!!!!!!!!!! DID VALUE IS EMPTY !!!!!!!!!!')
        final_set["Input_DID_FLOW"] = f"22.Read"
        final_set["Output_DID_FLOW"] = f"62.Write"
    return final_set


def did_code_comment_formatting(comment):
    formatted_did = ''
    lines = comment.split('\n')
    for line in lines:
        if line.strip():
            code = re.search(DID_pattren, line)
            DID_code = re.sub(Pat_tren, "", code.group())
            formatted_did = '.'.join(DID_code[i:i + 2] for i in range(0, len(DID_code), 2))
            break
    return formatted_did


def extract_diag_req(ReqName, ReqVer, rqIDs, feps):
    # global actual_content
    update_sheet = {
        "Involved_Flow": "",
        "INIT_Value": "",
        "Life_Cycle": [],
        "Vehicle_Mode": [],
        "CASES": "",
        "DID_CODE": {},
        "ss_fiches_Other_Value": "",
        "Ss_fiche_init_value": "",
        "DCI_Other_Value": "",
        "DCI_init_value": "",
        "DCI_Other_Name": "",
        "DCI_init_Name": ""
    }
    requirement = (ReqName + ' (' + ReqVer + ')')
    logging.info('requirement ------>', requirement)
    y = EI.findInputFiles()[17]
    logging.info('y = EI.findInputFiles()[18]---->', y)
    if len(y) > 0:
        for i in y:
            path = ICF.getInputFolder() + "\\" + i
            logging.info("path----000-->", path)
            actual_content = WDI.getReqContent(path, ReqName, ReqVer)
            # actual_content = DS.find_requirement_content(path, ReqName + "(" + str(ReqVer) + ")")
            # logging.info(f"DIAG actual_content {actual_content}")
            # if actual_content == -1 or not actual_content:
            #     actual_content = DS.find_requirement_content(path, ReqName + " " + str(ReqVer))
            #     logging.info(f"DIAG oldRqTable2 {actual_content}")
            #     if actual_content == -1 or not actual_content:
            #         actual_content = DS.find_requirement_content(path, ReqName + "  " + str(ReqVer))
            #         if actual_content == -1 or not actual_content:
            #             actual_content = DS.find_requirement_content(path, ReqName + " (" + str(ReqVer) + ")")
            logging.info('actual_content--->', actual_content)
            if actual_content != -1 and actual_content:
                content = actual_content['content']
                logging.info('content--->', content)
                diag_pattern = r"diagnostic device"
                match = re.search(diag_pattern, content)
                if match:
                    diag = match.group()
                    logging.info("Extracted DIAG:", diag)
                    if actual_content['content']:
                        dat_dic = extract_diag_content(str(actual_content['content']))
                        logging.info('dat_dic--->', dat_dic)
                        if ((len(dat_dic['Involved flow']) > 0) or (len(dat_dic['Involved flows']) > 0)) and (
                                (dat_dic['Involved flow']!=[]) or (dat_dic['Involved flows']!=[])):
                            logging.info('dat_dic["Involved flow"]---->', dat_dic['Involved flow'])
                            for inv_flow in dat_dic['Involved flow']:
                                logging.info('inv_flow----->', inv_flow)
                                update_sheet["Involved_Flow"] = inv_flow
                                if (len(dat_dic['Life cycle']) > 0) and (dat_dic['Life cycle']!='') and (
                                        dat_dic['Life cycle']!=-1):
                                    life_cycle = lyf.extract_condn(dat_dic['Life cycle'])
                                    update_sheet["Life_Cycle"] = life_cycle
                                    logging.info('life_cycle---->', life_cycle)
                                    update_sheet["Vehicle_Mode"] = vehicle_mode_extraction(dat_dic['Vehicle mode'])
                                    logging.info('update_sheet["Vehicle_Mode"] --->', update_sheet["Vehicle_Mode"])
                                    dci_glob_pc = search_DCI_Global_P_C(inv_flow)
                                    logging.info('dci_glob_pc --->', dci_glob_pc)
                                    ss_fiches = search_ss_fiche(inv_flow)
                                    logging.info('ss_fiches --->', ss_fiches)
                                    initial_value = Init_Value_extraction(dci_glob_pc["DCI_global_G_column"])
                                    logging.info('initial_value--------->', initial_value)
                                    init_value = convert_to_decimal(initial_value['InitValue'])
                                    logging.info('init_value---->', init_value)
                                    update_sheet["INIT_Value"] = init_value
                                    if ss_fiches['Status_ss_fiche'] is True:
                                        other_value = Init_Value_SS_fiches_Other_value_extraction(dci_glob_pc["DCI_global_E_column"],
                                                                                        dci_glob_pc["DCI_global_G_column"],
                                                                                        ss_fiches['D_Colum_Other_Value'])
                                        update_sheet["ss_fiches_Other_Value"] = other_value['ss_fiches_Other_Value']
                                        update_sheet["Ss_fiche_init_value"] = other_value['ss_fiches_Init_Value']
                                        logging.info('dci_glob_pc["Produced|Consumed"] --->', dci_glob_pc["Produced|Consumed"])
                                    if dci_glob_pc['Status_Dci_Global'] is True:
                                        DCI_init_other_values = extract_DCI_int_value_other_Value(dci_glob_pc["DCI_global_G_column"])
                                        logging.info('DCI_init_other_values----->', DCI_init_other_values)
                                        name0, value0 = DCI_init_other_values['Other_Value']
                                        update_sheet["DCI_Other_Value"] = value0
                                        update_sheet["DCI_Other_Name"] = name0
                                        name1, value1 = DCI_init_other_values['Init_Value']
                                        update_sheet["DCI_init_value"] = value1
                                        update_sheet["DCI_init_Name"] = name1
                                    # if actual_content['comment'] != '':
                                    DID_CODE = did_code(str(actual_content['comment']))
                                    logging.info('DID_CODE = did_code(reqData)---->', DID_CODE)
                                    update_sheet["DID_CODE"] = DID_CODE
                                    logging.info('update_sheet["DID_CODE"] = DID_CODE---->', update_sheet["DID_CODE"])
                                    input_did_flow = update_sheet["DID_CODE"]["Input_DID_FLOW"]
                                    output_did_flow = update_sheet["DID_CODE"]["Output_DID_FLOW"]
                                    logging.info('input_did_flow--->', input_did_flow)
                                    logging.info('output_did_flow--->', output_did_flow)
                                    logic = Cases_logic(dci_glob_pc, ss_fiches)
                                    logging.info('logic---->', logic)
                                    update_sheet["CASES"] = logic
                                    tpBook = EI.openTestPlan()
                                    tpBook.activate()
                                    logging.info("tpBook.sheets.active ", tpBook.sheets.active)
                                    diag_FT_creation(tpBook, ReqName, ReqVer, update_sheet, rqIDs, feps)
                                    # logging.info('fill_history---->', fill_history)
                                    # if fill_history != -1:
                                    #     try:
                                    #         BL.fillHistoryAndTrigram(tpBook, "Created new sheet")
                                    #         new_req_pos = EI.searchDataInCol(tpBook.sheets['Impact'], 1, ReqName)
                                    #         logging.info("new_req_pos ", new_req_pos)
                                    #         if new_req_pos['count'] > 0:
                                    #             ts_col = 4
                                    #             comment_col = 5
                                    #             for cell_pos in new_req_pos['cellPositions']:
                                    #                 x, y = cell_pos
                                    #                 logging.info(f"cell_pos {cell_pos}")
                                    #                 EI.setDataFromCell(tpBook.sheets['Impact'], (x, comment_col),
                                    #                                    "New Requirement.")
                                    #     except Exception as e:
                                    #         logging.info(f"\nError in filling history for new requirement.. {e}")
                                else:
                                    logging.info(':::: Life cycle is Empty please debug ::::')
                            logging.info('DIAGNOSTICS Requirement is Treated')
                        else:
                            logging.info(':::: Involved flow is Empty please debug ::::')
                    else:
                        logging.info(':::: Content is empty ::::')
                else:
                    logging.info('::::::::::DIAG REQ NOT FOUND::::::::::')
                    return -1


def convert_to_decimal__(value):
    # logging.info('convert_to_decimal(value)--->', value)
    if value!='':
        value = value.strip()
        try:
            # Try converting from binary to decimal
            decimal_value = int(value, 2)
            logging.info('Try converting from binary to decimal', decimal_value)
            return decimal_value
        except ValueError:
            try:
                # Try converting from hexadecimal to decimal
                decimal_value = int(value, 16)
                logging.info('Try converting from hexadecimal to decimal', decimal_value)
                return decimal_value
            except ValueError:
                try:
                    # Try converting to decimal directly
                    decimal_value = int(float(value))
                    logging.info('converting to decimal directly', decimal_value)
                    return decimal_value
                except ValueError:
                    logging.info("Returning same value")
                    return value


def find_matching_elements(element, full_element):
    matching_elements = []
    for i in element:
        for j in full_element:
            if i == j:
                if (i[0].strip() == j[0].strip()) and (i[1].strip() == j[1].strip()):
                    vat = i
                    matching_elements.append(vat)
    final = list(set(matching_elements))
    return final


def extract_DCI_int_value_other_Value(G_column):
    DCI_Values_dic = {
        "Init_Value": "",
        "Other_Value": ""
    }
    keyword = ['initvalue', 'init value', 'Init Value']
    buff_other_value = []

    modified_var_ = extract_G_col(G_column)
    logging.info('\n\nmodified_var--->', modified_var_)
    final_other_value = find_matching_elements(modified_var_, DOVKW.Other_value_keyword_with_values)
    logging.info('\n\nfinal_other_value--0->', final_other_value)
    for key in keyword:
        for mod in modified_var_:
            if mod[0].strip().lower()==key:
                init_value = convert_to_decimal__(mod[1])
                for other_value in final_other_value:
                    oth_value = convert_to_decimal__(other_value[1])
                    if init_value==oth_value:
                        DCI_Values_dic['Init_Value'] = (str(other_value[0]), str(init_value))
                        buff_other_value.append(other_value)
                        break
    for element in final_other_value:
        if element in buff_other_value:
            final_other_value.remove(element)
    o_v = convert_to_decimal__(final_other_value[0][1])
    f_o_v = str(final_other_value[0][0]), str(o_v)
    DCI_Values_dic['Other_Value'] = f_o_v
    return DCI_Values_dic


def extract_content(sentence):
    text = sentence.replace(":", ": l")
    # Split the string using the regex pattern and include the splitter
    involvedFlowIndex = text.lower().index('involved flow')
    text = text[involvedFlowIndex:]
    categories = re.findall(r'[A-Za-z0-9_ \'\"]+:', text)
    contents = re.split(r'[A-Za-z0-9_ \'\"]+:', text)

    # Remove any empty strings from the resulting list
    categories = [item.replace(":", "").strip() for item in categories]
    contents = [item.strip() for item in contents if item.strip()]

    cat2content = dict()
    for category, content in zip(categories, contents):
        if (category.lower().strip() == "involved flow") or (category.lower().strip() == "involved flows"):
            cat2content["Involved flow"] = content.strip().split("\n")
        elif category.lower().strip()=="vehicle mode":
            cat2content["Vehicle mode"] = content.strip().split(",")
        elif category.lower().strip()=="Life cycle":
            cat2content["Life cycle"] = content.strip()
        else:
            cat2content[category.strip()] = content.strip()
    # logging.info(json.dumps(cat2content, indent=2))
    return cat2content

def extract_diag_content(text):
    json_object = extract_content(text)
    for key in json_object:
        if isinstance(json_object[key], list):
            json_object[key] = [item[2:].strip() if item.startswith('l ') else item.strip() for item in
                                json_object[key]]
            json_object[key] = [item for item in json_object[key] if item!='']
        elif isinstance(json_object[key], str):
            json_object[key] = json_object[key][2:].strip()
    # logging.info the modified JSON object
    logging.info("\n", json_object)
    return json_object


# if __name__=='__main__':
#     final_lis = []
#     # ReqName = 'REQ-0785220'
#     # ReqVer = 'A'
#     lis = [('REQ-0742874', 'A'), ('REQ-0742875', 'A'), ('REQ-0742876', 'A'), ('REQ-0742879', 'B'), ('REQ-0742880', 'B'), ('REQ-0742882', 'A'), ('REQ-0742881', 'A'), ('REQ-0742883', 'B'), ('REQ-0742884', 'B')]
#     # actual_content = {
#     # 'content' : 'The VSM shall allow the reading, on a diagnostic device, of the following flows values with respect of given conditions:\n\nInvolved flow : PILOTAGE_PLAF1_FONC \nParameter description : control the lighting of rooflamp row 2\t \nLife cycle : ECL_ACCUEIL_ACTIVABLE=Activable Or ECL_LOC_FORC_PLAF = activable\n\nVehicle mode : All'
#     # }
#     # file_name = r'C:\Users\clakshminarayan\Documents\BSI-VSM(Automation)\Integarted, DTC, DIAG, CALIBRATION\test_Extraction_Involved_flows\CABIN_SSD_LIC_GEN2_00998_16_02483_v23_23Q1.docx'
#     # file_name = r'C:\Users\clakshminarayan\Documents\BSI-VSM(Automation)\Integarted, DTC, DIAG, CALIBRATION\test_Extraction_Involved_flows\SSD_HMIF_ENERGY_HMI.docx'
#     file_name = r'C:\Users\clakshminarayan\Documents\BSI-VSM(Automation)\Integarted, DTC, DIAG, CALIBRATION\test_Extraction_Involved_flows\SSD_HMIF_LONGITUDINAL_MOBILITY_MOBY_HMI_23Q1.docx'
#
#     for i in lis:
#         ReqName, ReqVer = i
#         actual_content = WDI.getReqContent(file_name, ReqName, ReqVer)
#         logging.info('actual_content--->', actual_content)
#         final_lis.append(str(actual_content['content']))
#
#     logging.info('final_lis---->', final_lis)
    # for foo in final_lis:
    #     data = extract_diag_content(foo)
    #     logging.info('\n\ndata--->', data)
    #     logging.info('\n\n')



#     if actual_content != -1:
#         content = actual_content['content']
#         data = extract_diag_content(content)
#         # data = extract(content)
#         logging.info('data--->', data)
#     # val = search_ss_fiche('P_INFO_CMDM_MODE_VHL')
#     # logging.info('val--->', val)
#     # extract_diag_req(ReqName, ReqVer, file_name)
