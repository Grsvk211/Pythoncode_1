import re
import WordDocInterface as WDI
import ExcelInterface as EI
import InputConfigParser as ICP
import Thematic as THM
import logging

def optimize_input_string(input_string):
    input_string = input_string.replace("{", " ( ")
    input_string = input_string.replace("}", " ) ")
    # input_string = input_string.replace(",", " AND ")
    input_string = input_string.replace(",", " OR ")
    input_string = input_string.replace("(", " ( ")
    input_string = input_string.replace(")", " ) ")

    output_string = ""
    foundNewWord = False
    for index, char in enumerate(input_string):
        # logging.info(char)
        logging.info(f"input_string[index + 1:] {input_string[index + 1:]}")
        if char in "()":
            output_string += char
        elif input_string[index + 1:].startswith(" OR "):
            output_string += " OR "
        elif input_string[index + 1:].startswith("AND "):
            output_string += " AND "
        elif char == ' ':
            # New Word
            foundNewWord = True
        else:
            if (index + 5 < len(input_string)) and foundNewWord:
                temp = input_string[index:index + 6]
                pattern = r"[A-Za-z0-9]{3}_\d{2}"
                res = re.search(pattern, temp)
                if res is not None:
                    output_string += temp
                else:
                    # goto next word
                    foundNewWord = False

    return output_string


def getTestSheetReqVer(tpBook, test_sheet, reqName):
    getReqList = tpBook.sheets[test_sheet].range('C4').value.split("|")
    # logging.info("req list = ", getReqList)
    logging.info("before req list = ", getReqList)
    # Use list comprehension to remove empty elements
    getReqList = [item for item in getReqList if item]

    # Print the filtered list
    logging.info("after req list = ", getReqList)
    testReqName = ""
    testReqVer = ""
    for req in getReqList:
        if req.find("(") != -1:
            tempReqName = req.split("(")[0]
            tempReqVer = req.split("(")[1].split(")")[0]
            if tempReqName == reqName.strip():
                testReqName = tempReqName
                testReqVer = tempReqVer
        else:
            logging.info("req = ", req)
            tempReqName = req.split()[0]
            try:
                tempReqVer = req.split()[1]
            except:
                tempReqVer = ""
            if tempReqName == reqName.strip():
                testReqName = tempReqName
                testReqVer = tempReqVer

    return testReqName, testReqVer


def getThematic(thematic_data):
    # logging.info(f"Getting the thematic ...... {thematic_data}")
    if thematic_data != -1 and thematic_data != -2:
        if thematic_data['effectivity'] != "":
            return thematic_data['effectivity']
        elif thematic_data['lcdv'] != "":
            return thematic_data['lcdv']
        elif thematic_data['diversity'] != "":
            return thematic_data['diversity']
        elif thematic_data['target'] != "":
            return thematic_data['target']


def find_difference(string1, string2):
  string1 = string1.split()
  string2 = string2.split()

  str1 = {re.findall('[A-Z]{3}_[0-9]{2}', x)[0] for x in string1 if re.search('[A-Z]{3}_[0-9]{2}', x)}
  str2 = {re.findall('[A-Z]{3}_[0-9]{2}', y)[0] for y in string2 if re.search('[A-Z]{3}_[0-9]{2}', y)}

  oldThemDiff = list(str1-str2)
  newThemDiff = list(str2-str1)

  logging.info(f"A-B {oldThemDiff}")
  logging.info(f"\n\nB-A {newThemDiff}")

  return  oldThemDiff, newThemDiff


def getRawThematic(tpBook, refEC, currDoc, prevDoc,listOfTestSheets, reqName, reqVer, Arch, newReq=""):
    logging.info("Get Raw Theatic....")
    req_name = ""
    req_ver = ""
    if len(newReq) == 0:
        req_name = reqName.strip()
        req_ver = reqVer.strip()
    else:
        if newReq.find("(") != -1:
            req_name = newReq.split("(")[0]
            req_ver = newReq.split("(")[1].split(")")[0]
        else:
            req_name = newReq.split(" ")[0]
            req_ver = newReq.split(" ")[1]
            req_name = reqName.strip()
            req_ver = reqVer.strip()
    logging.info(f"req_name {req_name} req_ver {req_ver}")
    OldRawThematic = ""
    NewRawThematic = ""

    # getting the previous requirement and version from testplan sheet
    test_sheet_req, test_sheet_ver = getTestSheetReqVer(tpBook, listOfTestSheets[0], req_name)
    logging.info(f"test_sheet_req, test_sheet_ver {test_sheet_req, test_sheet_ver}")

    # getting the thematic content from document
    oldRawThem = WDI.getReqContent(prevDoc, test_sheet_req, test_sheet_ver)
    newRawThem = WDI.getReqContent(currDoc, req_name, req_ver)

    if oldRawThem != -1 and oldRawThem is not None and oldRawThem != "":
        thematicOld = getThematic(oldRawThem)
        logging.info(f"\n\nthematicOld {thematicOld}")

        # optimizing the raw thematic in proper format
        OldRawThematic = optimize_input_string(thematicOld)

    if newRawThem != -1 and newRawThem is not None and newRawThem != "":
        thematicNew = getThematic(newRawThem)
        logging.info(f"\n\nthematicNew {thematicNew}")

        # optimizing the raw thematic in proper format
        NewRawThematic = optimize_input_string(thematicNew)

    logging.info(f"\n\noldRawThem -> {OldRawThematic}")
    logging.info(f"NewRawThemactic -> {NewRawThematic}")

    return OldRawThematic, NewRawThematic


if __name__ == "__main__":
    # input_string = "HP-0000873[B-∞]AND(HP-0000873[DFH{DFH_05}]ORHP-0000873[LYQ{LYQ_01}])"
    input_string = "HP-0000873 [ DICO MULTIGAMME Q3_2020 - ∞ ] AND ( DFH AUXILIARY BRAKE (DFH_05 ELECTRIC)  AND LYQ TYPE_DIVERSITY (LYQ_01 BEFORE_FUNCT_CODIF)  ) "
    # input_string = "HP-0000873[B-∞] AND ( HP-0000873[DFH{DFH_05}] AND HP-0000873[LYQ{LYQ_01}] )"
    # input_string = "HP-0000873 [ DICO MULTIGAMME Q3_2020 - ∞ ] AND ( DFH AUXILIARY BRAKE (DFH_05 ELECTRIC)  AND LYQ TYPE_DIVERSITY (LYQ_01 BEFORE_FUNCT_CODIF)  ) "
    # input_string = "HP-0000873[A-∞] AND ( HP-0000873[DFH{DFH_05}] AND HP-0000873[ALO{ALO_01}] ) OR ( HP-0000873[DFH{DFH_05}] AND HP-0000873[ALO{ALO_02}] )"
    # input_string = "HP-0000873 [ DICO MULTIGAMME T2_2016 - ∞ ] OR ( ALO TYPE_FSE (ALO_01 EMOVE_2004 , ALO_02 EMOVE_2010)  AND DFH AUXILIARY BRAKE (DFH_05 ELECTRIC)  ) "
    # input_string = "HP-0000873 [ DICO MULTIGAMME Q2_2021 - ∞ ] AND EEK OPTION_ROW2_PASS_PRES_DETECTIO(EEK_00 WITHOUT) AND IZB INHIBITION_FCT_TNB_AR(IZB_00 WITHOUT) AND LUG TYPE_SBR_HMI(LUG_01 3_STATES) AND LWK TYPE_SBR_ALERT_VARIANTS(LWK_02 AUDIO_VISUAL_OUT_NORTH_AMERICA) AND (EEN OPTION_ROW3_PASS_PRES_DETECTIO(EEN_00 WITHOUT) AND IZE NB_CONTACT_CNB(IZE_08 7) OR IZE NB_CONTACT_CNB(IZE_05 4,IZE_06 5,IZE_07 6)) AND ((LME DISPO_INFO_PRESENCE_PASS_AV (LME_00 ABSENT) AND INHIBITION_FCT_TNB_PASS_AV (IZA_00 Non inhibée)) OR (OPTION_ROW1_CENTER_PRES_DETECT (JXN_00 WITHOUT) AND ROW1_SEATBELT_WITH_SENSOR (AWM_03 3)))"
    #
    logging.info(f"Output String -> {optimize_input_string(input_string)}")
    exit()
    tpBook = EI.openTestPlan()
    prevDoc = ICP.getInputFolder()+"\\[V1.0][01843_21_00466]TechnicalNote_HMIF_BSI_VSM_ANBC_SBR_ISS_0103550__DDM_10704905.docx"
    currDoc = ICP.getInputFolder()+"\\[V2.0][01843_22_00768]NT_HMIF_BSI_VSM_SBR_ALTIS-10951193_ISS-0171583_EFFECTIVITY_CORRECTION.docx"
    refEC = ""

    oldThematic, newThematic = getRawThematic(tpBook, refEC, currDoc, prevDoc, listOfTestSheets=['VSM20_N1_02_58_0304'], reqName='REQ-0615092', reqVer='C', Arch='VSM', newReq="")
    thematicLines = THM.convertRawToThematicLines(newThematic)
    for thematicLine in thematicLines:
        if thematicLine != -1 or thematicLine != -2:
            logging.info(thematicLine)




