import sys
import re
import WordDocInterface as WDI
import ExcelInterface as EI
import DocumentSearch as DS
import Thematic_Convert as TC
import Backlog_Handler as BH
import NewRequirementHandler as NRH
import logging

def getRawThematicReq(currDoc,reqName, reqVer, newReq=""):
    print("currDoc---------------->", currDoc)
    NewRawThematic = " "
    try:
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
        NewRawThematic = ""

        # getting the thematic content from document
        newRawThem = WDI.getReqContent(currDoc, req_name, req_ver)

        if newRawThem != -1 and newRawThem is not None and newRawThem != "":
            thematicNew = TC.getThematic(newRawThem)
            logging.info(f"\n\nthematicNew {thematicNew}")

            # optimizing the raw thematic in proper format
            NewRawThematic = TC.optimize_input_string(thematicNew)

        logging.info(f"NewRawThemactic -> {NewRawThematic}")
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info(f'NewRawThemactic error .{ex}{exc_tb.tb_lineno}')
    return NewRawThematic


def format_expression(expression):
    matches = re.findall(r'\{([\d,]+)\}', expression)
    if matches:
        numbers = matches[0].split(',')
        formatted_numbers = [f"{expression.replace(matches[0], num)}" for num in numbers]
        return ' '.join(formatted_numbers)
    return expression


# def createApplicableCombination(thematiqueList, refEC, ARCH):
def createApplicableCombination(thematiqueList, ARCH):
    # Loop through each element
    tempflagR0 = 0
    tempflagR1 = 0
    tempflagR2 = 0
    ListOfThematics = []
    for i in thematiqueList:
        # Open the Excel file
        # wb = xw.Book(refEC)
        wb = EI.openReferentialEC()
        try:
            # Activate the 'Liste EC' sheet
            sheet = wb.sheets['Liste EC']
            sheet.activate()
            maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            sheet_value = sheet.used_range.value
            logging.info("Maxrows  and Listof thematics----->>>", maxrow, i)
            if i:
                logging.info("In filterThemForArch (maxrow, i, sheet)", maxrow, i, sheet)
                try:
                    searchResults = EI.searchDataInColCache(sheet_value, 7, i)
                except:
                    searchResults = EI.searchDataInColCache(sheet_value, 7, i)
                if searchResults["count"] != 0:
                    x, y = searchResults["cellPositions"][0]
                    applicableBSI = sheet.range(x, y + 38).value
                    applicableR1 = sheet.range(x, y + 39).value
                    applicableR2 = sheet.range(x, y + 40).value
                    logging.info("Thematique = ", i, "Aplicable to = ", applicableBSI, applicableR1, applicableR2)
                    # for BSi Arch
                    if ARCH == "BSI":
                        if applicableBSI == "Y":
                            flagR0 = 1
                            tempflagR0 = flagR0
                            pass
                        else:
                            logging.info("not applicable for BSI but its present in req\n")
                    # for VSM Arch
                    elif ARCH == "VSM":
                        applicableR1_1 = " "
                        applicableR2_list = []
                        final_value_R2 = ''
                        if (applicableR1 != "Y" and applicableR1 != "N") or (applicableR2 != "Y" and applicableR2 != "N"):
                            logging.info("applicableR2----->",applicableR2)
                            duplicateResult = sheet.range(x, y + 1).value
                            replace_Code = [duplicateResult]
                            # Replace_Code-------> ['DXD{03,05} AND DXD{06}']
                            logging.info("Replace_Code------->", replace_Code)
                            output_string = [format_expression(input_str) for input_str in replace_Code]
                            logging.info("output_strings------->", output_string)
                            # output_strings-------> ['DXD{03} DXD{05} AND DXD{06}']
                            output_strings = [string.replace(' AND ', ' ') for string in output_string]
                            # output_strings-------> ['DXD{03} DXD{05} DXD{06}']
                            for k in output_strings:
                                duplicateRes = k.replace('{', '_').replace('}', '')
                                logging.info("searchResult--------->",duplicateRes)
                                # searchResult---------> DXD_03 DXD_05 DXD_06
                                duplicateResl = duplicateRes.split()
                                # duplicateResl------->['DXD_03', 'DXD_05', 'DXD_06']
                                for m in duplicateResl:
                                    if m:
                                        logging.info("m--->", m)
                                        logging.info("In filterThemForArch (maxrow, i, sheet)", maxrow, m, sheet)
                                        try:
                                            searchResults = EI.searchDataInExcelCache(sheet_value, (maxrow, 7), m)
                                        except:
                                            searchResults = EI.searchDataInExcelCache(sheet_value, (maxrow, 7), m)
                                        if searchResults["count"] != 0:
                                            x, y = searchResults["cellPositions"][0]
                                            applicableBSI_1 = sheet.range(x, y + 38).value
                                            applicableR1_1 = sheet.range(x, y + 39).value
                                            applicableR2_1 = sheet.range(x, y + 40).value
                                            logging.info("Thematique2 = ", i, "Aplicable to = ", applicableBSI_1, applicableR1_1, applicableR2_1)
                                            test_cases = [(applicableBSI, applicableR1, m), (applicableBSI_1, applicableR1_1, applicableR2_1)]
                                            logging.info("test_cases-------->",test_cases)
                                            logging.info(f'{[(applicableBSI, applicableR1, m),(applicableBSI_1, applicableR1_1, applicableR2_1)]}={(applicableBSI, applicableR1, applicableR2_1)}')
                                            applicableR2 = applicableR2_1
                                            logging.info("applicableR2->", applicableR2)
                                            applicableR2_list.append(applicableR2)
                            logging.info("applicableR2_list->", applicableR2_list)
                            final_value_R2 = "N"  # Default value if the pattern is not found
                            if "Y" in applicableR2_list and "N" in applicableR2_list:
                                final_value_R2 = "Y"
                            logging.info(f'after changing the replacing code Applicability combination-R2----->{(applicableBSI, applicableR2, final_value_R2)}')
                        flagR1 = 1
                        flagR2 = 1
                        if ((applicableR1 == "Y") and (applicableR2 == "Y")) or ((applicableR1_1 == "Y") and (final_value_R2 == "Y")):
                            logging.info("Aplicable to R1 & R2")
                            # tempflagR0 = flagR0
                            tempflagR1 = flagR1
                            tempflagR2 = flagR2
                            pass
                        elif (applicableR1 == "Y") or (applicableR2 == "Y"):
                            if (applicableR1 == "Y"):
                                logging.info("NEA R1 applicable")
                                tempflagR1 = flagR1
                            elif (applicableR2 == "Y"):
                                flagR2 = 1
                                logging.info("NEA R2 applicable")
                                tempflagR2 = flagR2
                        else:
                            logging.info(f'Not applicable for VSM but its present in req\n')
                    else:
                        logging.info("arch not found\n")
                else:
                    print(f"Thematique {i} not found in referential EC")
                    return -1,-1,-1
            logging.info("TempFlag = ", tempflagR0, tempflagR1, tempflagR2)

        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            logging.info(f"An error occurred:{ex}{exc_tb.tb_lineno}")
        wb.close()
        return tempflagR0, tempflagR1, tempflagR2


# def checkReq(currDoc, requirement_id, refEC, ARCH):
def checkReq(currDoc, requirement_id, ARCH, newReq = ''):
    print("currDoc------------>", currDoc)
    flag = 0
    thematic_Lines = ''
    unique_list = []
    if newReq != '':
        requirement_id = newReq
    logging.info("currDoc------>", currDoc)
    if currDoc != -1:
        reqName, reqVer = NRH.getReqVer(requirement_id)
        if requirement_id == reqName + "(" + reqVer + ")":
            con = DS.find_requirement_content(currDoc, requirement_id)
            logging.info("con0---->", con)
        elif requirement_id == reqName + " (" + reqVer + ")":
            con = DS.find_requirement_content(currDoc, requirement_id)
            logging.info("con1---->", con)
        elif requirement_id == reqName + " " + reqVer:
            con = DS.find_requirement_content(currDoc, requirement_id)
            logging.info("con2---->", con)
        elif requirement_id == reqName + "  " + reqVer:
            con = DS.find_requirement_content(currDoc, requirement_id)
            logging.info("con3---->", con)
        # reqName, reqVer = NRH.getReqVer(requirement_id)

        try:
            rawThematics = getRawThematicReq(currDoc, reqName, reqVer, newReq="")
            if rawThematics:
                logging.info("rawThematics----->",rawThematics)
                thematic_Combination_Line = BH.grepThematicsCode(rawThematics)
                thematic_Lines = BH.createCombination(thematic_Combination_Line)
                logging.info("thematic_Lines----->",thematic_Lines)
                output_list = []
                list_consider = []
                lines = thematic_Lines.split('\n')
                for line in lines:
                    elements = line.split('|')
                    output_list.append(tuple(elements))
                logging.info(output_list)
                for i, n in enumerate(output_list):
                    first_elements = [t for t in n]
                    logging.info(f'Elements{i+1}--:{first_elements}')
                    result_list = []
                    for item in first_elements:
                        result_list.append([item])
                    tempflagR0_list = []
                    tempflagR1_list = []
                    tempflagR2_list = []
                    for i, l in enumerate(result_list):
                        # refEC = EI.openReferentialEC()
                        # tempflagR0, tempflagR1, tempflagR2 = createApplicableCombination(l, refEC, ARCH)
                        tempflagR0, tempflagR1, tempflagR2 = createApplicableCombination(l,  ARCH)
                        logging.info("tempflagR0, tempflagR1, tempflagR2---->",tempflagR0, tempflagR1, tempflagR2)
                        tempflagR0_list.append(tempflagR0)
                        tempflagR1_list.append(tempflagR1)
                        tempflagR2_list.append(tempflagR2)

                    logging.info("TempFlag list for tempflagR0(BSI):", tempflagR0_list)
                    logging.info("TempFlag list for tempflagR1(VSM):", tempflagR1_list)
                    logging.info("TempFlag list for tempflagR2(VSM):", tempflagR2_list)

                    if all(x == tempflagR0_list[0] for x in tempflagR0_list):
                        # if any(flag != 0 for flag in tempflagR0_list[0]):
                        if tempflagR0_list[0] != 0:
                            logging.info("All elements in tempflagR0_list are the same:", tempflagR0_list[0])
                            output_string = '|'.join(first_elements)
                            logging.info(output_string)
                            list_consider.append(output_string)
                    else:
                        logging.info("Elements in tempflagR0_list are not the same")

                    if all(x == tempflagR1_list[0] for x in tempflagR1_list):
                        # if any(flag != 0 for flag in tempflagR1_list):
                        if tempflagR1_list[0] != 0:
                            logging.info("All elements in tempflagR1_list :", tempflagR1_list[0])
                            output_string = '|'.join(first_elements)
                            logging.info(output_string)
                            list_consider.append(output_string)
                    else:
                        logging.info("Elements in tempflagR1_list are not the same")

                    if all(x == tempflagR2_list[0] for x in tempflagR2_list):
                        # if any(flag != 0 for flag in tempflagR2_list):
                        if tempflagR2_list[0] != 0:
                            logging.info("All elements in tempflagR2_list are the same:", tempflagR2_list[0])
                            output_string = '|'.join(first_elements)
                            logging.info(output_string)
                            list_consider.append(output_string)
                    else:
                        logging.info("Elements in tempflagR2_list are not the same")
                logging.info("list_consider------>", list_consider)
                unique_list = list(dict.fromkeys(list_consider))
                logging.info(f'ThematicsLines associated with given Requirement{requirement_id}---> {unique_list}')

                if ARCH == 'VSM':
                    if unique_list:
                        flag = 1
                        logging.info(f'These Requirement "{requirement_id}" is applicable for VSM')

                    else:
                        flag = -1
                        logging.info(f'"{requirement_id}" not applicable for {ARCH}')
                        # displayInformation(f'"{requirement_id}" not being applicable for {ARCH}, Proceed Manually.')

                else:

                    if unique_list:
                        flag = 1
                        logging.info(f'These Requirement "{requirement_id}" is applicable for BSI')

                    else:
                        flag = -1
                        logging.info(f'"{requirement_id}" not applicable for {ARCH}')
            else:
                flag = 1
                logging.info("No Thematics for these requirement in Document")

        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            logging.info(f'"{requirement_id}" Input document is not in correct format.{ex}{exc_tb.tb_lineno}')
            print(f'\n"{requirement_id}" Thematics in the input doc are not in the correct format. Example check for the open/close parenthesis or Req & Ver may have two times in the document.')
            flag = -2
    else:
        flag = -3
    return flag


if __name__ == "__main__":
    currDoc = r"C:/Users/vgajula/Downloads/table.docx"
    reqName = 'REQ-0580241 '
    reqVer = 'E'
    b = getRawThematicReq(currDoc, reqName, reqVer, newReq="")
    print("b-------------->", b)
