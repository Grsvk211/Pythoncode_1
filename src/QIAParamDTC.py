import QIA_Updater as QIAU
from datetime import date
import os
import ExcelInterface as EI
import InputConfigParser as ICF
import re
import NewRequirementHandler as NRH
import WordDocInterface as WDI
import QIA_Param as QP
import logging
ICF.loadConfig()


# these function will work when we have flow in the requirements.
def creatQIAParamDTC(testPlanReference, functionName, new_req, trigram):
    logging.info("\nProcessing the QIA Sheet to treat the DTC!!!\n")
    valueMap00 = {"A": "", "B": testPlanReference, "C": functionName, "D": 'Creation', "E": 'New DIAG for NEA',
                  "F": "", "G": '--', "H": '', "I": '', "J": '--', "K": '', "L": functionName, "M": "", "N": '--',
                  "O": '--', "P": '' + '(EXP)', "Q": '', "R": '',
                  "U": '', }
    isQAFileExist = os.path.isfile(ICF.getInputFolder() + "\\" + EI.findInputFiles()[10])
    logging.info("isQIAParamExist- ", isQAFileExist)
    if isQAFileExist:
        QIABook = EI.openExcel(ICF.getInputFolder() + "\\" + EI.findInputFiles()[10])
        if QIABook:
            QIAsheet = QIABook.sheets["NEW_QIA"]
            QIAsheet.activate()
            maxrow = QIAsheet.range('B' + str(QIAsheet.cells.last_cell.row)).end('up').row
            valueMap00['A'] = int(QIAsheet.range(maxrow, 1).value) + 1
            flowArr = []
            QiaFile = EI.findInputFiles()[10]
            path = ICF.getInputFolder() + "\\" + QiaFile
            isQAFileExist = os.path.isfile(path)
            logging.info("QA File Exist- ", isQAFileExist)

            # it used for to remove the gap between the reqno and version
            newRequir = re.sub(r'(\w-\d+)\s+(\()', r'\1\2', new_req)
            logging.info("newRequir--->", newRequir)
            new_reqName, new_reqVer = NRH.getReqVer(newRequir)
            logging.info("new_reqVer1=", new_reqVer)
            logging.info("new_reqName1=", new_reqName)
            req = new_reqName + "(" + str(new_reqVer) + ")"

            # Split the trigram at the first space character
            parts = trigram.split(' ', 1)
            # Extract the first part of the trigram
            tri = parts[0]

            Today = date.today()
            curr_date = Today.strftime("%m/%d/%y")

            valueMap00['I'] = req
            valueMap00['P'] = str(tri + '(EXP)')
            valueMap00['Q'] = curr_date

            valueMap00['F'] = ""
            valueMap00['H'] = ""

            EEAD = EI.findInputFiles()[12]
            path = ICF.getInputFolder() + "\\" + EEAD
            # docName = "EEAD_SUBSYST_PARK_HMI_22Q4.docx"
            # doc_path = ICF.getInputFolder() + "\\" + docName
            # doc = ICF.getInputFolder() + "\\" + "(new)EEAD_SUBSYST_PARK_HMI_22Q4.docx"
            reqData = WDI.getReqContent(path, new_reqName, new_reqVer)
            logging.info(f"reqData 1--> {reqData}")

            if type(reqData) is not dict:
                reqDataCond = "reqData.strip() != -1 and reqData.strip() != -2 and reqData.strip() != "" and reqData.strip() is not None"
            else:
                reqDataCond = True

            if reqDataCond:
                try:
                    if re.search("VSM-[A-Z]{2}[0-9]{2}", reqData['comment']):
                        extracted_DID = re.findall("VSM-[A-Z]{2}[0-9]{2}", reqData['comment'])
                        DID_value = extracted_DID[0]
                    else:
                        # VSM-U0131-81
                        extracted_DID = re.findall("VSM-[A-Z]{1,2}[A-Z a-z 0-9]{1,10}-[0-9]{2}", reqData['comment'])
                        DID_value = QP.convertDID(extracted_DID[0])
                except Exception as ex:
                    logging.info(f"Comment is empty in requirement table.please update manually. {ex}")
                    pass

                them_arch = QP.getThemArchi(reqData)
                valueMap00['F'] = them_arch

                try:
                    # reqData['comment'] = "@DNF RCTA-179 => VSM-U0131-87"
                    logging.info("reqData['comment']--->", reqData['comment'])

                    if reqData['flow']!=None and reqData['flow']!="":
                        flows = reqData['flow']
                        logging.info("flows--->", flows)
                    else:
                        logging.info("Flow is not present in the requirement")
                        pass

                    for req_flow in flows:
                        if req_flow!="" and req_flow is not None:
                            # if flow exist adding the prefix REQ for request and REP for response
                            prefixes = ["LEC_", "REP_", "REP_", "REP_"]
                            suffixes = ["", "_ABSENT", "_FUGITIF", "_PRESENT"]
                            # output_values will give the req_flow with pre and suffix and add in the flowArr
                            output_values = [f"{prefix}{req_flow}{suffix}" for prefix, suffix in
                                             zip(prefixes, suffixes)]
                            for i in output_values:
                                flowArr.append(i)
                            logging.info("flowArr--->", flowArr)

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
                            DID_value = QP.convertDID(extracted_DID[0])
                            logging.info("DID_value2--->", DID_value)

                    if DID_value!="" and DID_value is not None:
                        DIDVal = QP.split_did_with_dot(DID_value)
                        logging.info(f"DIDVal >> {DIDVal}")

                    if extracted_DID and reqData['comment']!="" and reqData['comment'] is not None:
                        # reqData = 'DTC'
                        # These loop will go if extracted_DID[0] keyword is present in content.extracted_DID[0] means DTC code
                        # example for DTC code is VSM-U0131-87
                        if extracted_DID[0] in reqData['comment']:
                            defectCode = extracted_DID[0].replace("VSM-", "")
                            logging.info("defectCode--->", defectCode)
                            valueMap00['M'] = str('Defect Code :' + defectCode)
                            valueMap00['R'] = "Open"
                            valueMap00['U'] = str(
                                tri + '(EXP)' + " " + curr_date + " Information can be found in Cerebro.")
                            for flow_req_res in flowArr:
                                valueMap00['H'] = f"{flow_req_res}"
                                if 'LEC_' in flow_req_res:
                                    valueMap00['K'] = f"19.04.{DIDVal}.FF"
                                elif 'REQ_' and '_ABSENT' in flow_req_res:
                                    valueMap00['K'] = f"7F.19.31"
                                elif 'REQ_' and '_PRESENT' in flow_req_res:
                                    valueMap00['K'] = f"59.04.{DIDVal}.09#"
                                elif 'REQ_' and '_FUGITIF' in flow_req_res:
                                    valueMap00['K'] = f"59.04.{DIDVal}.08#"
                                QIAU.UpdateQiaParamGlobal(QIAsheet, valueMap00)

                        # These loop will go if DTC keyword is not present in content.
                    else:
                        logging.info("hiiii")
                        valueMap00['M'] = str('Defect Code : TBD')
                        valueMap00['R'] = "QIA point opened"
                        valueMap00['U'] = str(
                            tri + '(EXP)' + " " + curr_date + "Raised QIA of Spec. QIA no. ##### and QIA of spec reference : '#####_##_#####'")
                        for flow_req_res in flowArr:
                            valueMap00['H'] = f"{flow_req_res}"
                            if 'LEC_' in flow_req_res:
                                valueMap00['K'] = "--"
                            elif 'REQ_' and '_ABSENT' in flow_req_res:
                                valueMap00['K'] = "--"
                            elif 'REQ_' and '_PRESENT' in flow_req_res:
                                valueMap00['K'] = "--"
                            elif 'REQ_' and '_FUGITIF' in flow_req_res:
                                valueMap00['K'] = "--"
                            QIAU.UpdateQiaParamGlobal(QIAsheet, valueMap00)
                            # UpdateQiaParamGlobal(QIAsheet, valueMap00)
                except Exception as ex:
                    logging.info(f"Flow or DTC code is not available for these requirement.please update manually. {ex}")
                    pass

        QIABook.save()
        QIABook.close()

# if __name__ == "__main__":
#      ICF.loadConfig()
#      testPlanReference= "12344_45_45678"
#      functionName = "FSEE_VVXSA"
#      # new_req = "REQ-0743278 (B)"
#      new_req = "REQ-0743275(B)"
#      # new_req = "REQ-0778501 (A)"
#      trigram = "VKG"
#
#      creatQIAParamDTC(testPlanReference, functionName, new_req, trigram)
#     logging.info("__name__ "+__name__)
#     # exit()
#     extracted_DID = "VSM-U1FD6-87"
#     # extracted_DID = "VSM-B1812-12"
#     b = QP.convertDID(extracted_DID)
#     logging.info("b-->", b)
#     defectCode = b.replace("VSM-","")
#     logging.info("defectCode-->", defectCode)
#     xx = QP.split_did_with_dot(b)
#     logging.info("xx--->", xx)
# # output == VSM-DFD6-87 for the input "VSM-U1FD6-87"
