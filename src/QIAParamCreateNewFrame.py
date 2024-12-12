import QIA_Updater as QIAU
from datetime import date
import os
import ExcelInterface as EI
import InputConfigParser as ICF
import re
import NewRequirementHandler as NRH
import WordDocInterface as WDI
import logging
# i is the inputDocumentReference
def creatNewFrameQIA(testPlanReference,functionName,new_req,trigram,i):
    logging.info("\nProcessing the QIA Sheet!!!\n")
    QiaFile = EI.findInputFiles()[10]
    MagnetoFrames = EI.findInputFiles()[11]
    path= ICF.getInputFolder() + "\\" + QiaFile
    isQAFileExist = os.path.isfile(path)
    logging.info("QA File Exist- ", isQAFileExist)
    Today = date.today()
    curr_date = Today.strftime("%m/%d/%Y")
    # it used for to remove the gap between the reqno and version
    newRequir = re.sub(r'(\w-\d+)\s+(\()', r'\1\2', new_req)
    logging.info("newRequir--->",newRequir)
    new_reqName, new_reqVer = NRH.getReqVer(newRequir)
    logging.info("new_reqVer1=",new_reqVer)
    logging.info("new_reqName1=",new_reqName)
    req = new_reqName + "(" + str(new_reqVer) + ")"
    pattern = "([A-Z0-9]{4,5})+(_[0-9]{2})+(_[A-Z0-9]{4,5})+"
    inputDocumentReference = re.search(pattern,i)
    logging.info("inputDocumentReference--->",inputDocumentReference)
    # Split the trigram at the first space character
    parts = trigram.split(' ', 1)
    # Extract the first part of the trigram
    tri = parts[0]

    EEAD = EI.findInputFiles()[12]
    path = ICF.getInputFolder() + "\\" + EEAD
    reqData = WDI.getReqContent(path, new_reqName, new_reqVer)
    logging.info(f"reqData 2--> {reqData}")

    if isQAFileExist:
        QIABook = EI.openExcel(ICF.getInputFolder() + "\\" + QiaFile)
        try:
            if QIABook:
                QIAsheet = QIABook.sheets["NEW_QIA"]
                QIAsheet.activate()
                maxrow = QIAsheet.range('B' + str(QIAsheet.cells.last_cell.row)).end('up').row
                path1 = ICF.getInputFolder() + "\\" + MagnetoFrames
                isMagnetoFramesExist = os.path.isfile(path1)
                if isMagnetoFramesExist:
                    MagnetoFrame = EI.openExcel(ICF.getInputFolder() + "\\" + MagnetoFrames)
                    if MagnetoFrame:
                        Magnetosheet = MagnetoFrame.sheets["FRAMES_DEFINITIONS"]

                        frames = reqData['frame']
                        logging.info("Frame--->", frames)
                        frame = frames[0]

                        ntew = frame[:3]
                        # ID0 = EI.searchDataInCol(Magnetosheet, 1, ntew)

                        sheet_value = Magnetosheet.used_range.value
                        ID0 = EI.searchDataInColCache(sheet_value, 1, ntew)

                        logging.info("ID0['cellPositions'][0]--->", ID0['cellPositions'][0])
                        x, y = ID0['cellPositions'][0]
                        logging.info("x, y--->", x, y)
                        ntew = EI.getDataFromCell(Magnetosheet, (x, y))
                        Pro = EI.getDataFromCell(Magnetosheet, (x, y + 1))
                        consumer = EI.getDataFromCell(Magnetosheet, (x, y + 2))
                        if Pro == 'PASS_VSM':
                            Proc = 'P'
                        elif consumer == 'PASS_VSM':
                            Proc = 'C'
                        else:
                            Proc = '--'
                        if frame:
                            logging.info("ntew--->", ntew)
                            if ntew == 'FD7':
                                parameter = 'NEA_R1_2'
                            elif ntew == 'FD8':
                                parameter = 'NEA_R1_X'
                            elif ntew == 'HS7':
                                parameter = 'NEA_R1|NEA_R1_1'
                            else:
                                parameter = 'NEA'
                        ID = EI.getDataFromCell(Magnetosheet, (x, y + 4))
            valueMap00 = {"A": int(QIAsheet.range(maxrow, 1).value) + 1,
                "B": testPlanReference, "C": functionName, "D": 'Creation', "E": 'New Flow for NEA',
                "F": parameter, "G": '--', "H": str("TRAME_" + ntew + "_" + ID + "_" + frame),
                "I": req, "J": '--', "K": '--', "L": '--', "M": '--',
                "N": str("CAN_" + ntew + "/" + frame), "O": Proc, "P": tri+'(EXP)',"Q": curr_date,
                "R": 'Open',
                "U": str(
                    curr_date + " " + tri+'(EXP)' + " : Information can be found on DCI Ref :" + inputDocumentReference.group() +"."),
            }
            QIAU.UpdateQiaParamGlobal(QIAsheet, valueMap00)

        except:
            logging.info("frame not present in requirement")
            pass
    QIABook.save()
    MagnetoFrame.close()
    QIABook.close()


def getdocContent(table):
    data = []
    # these pattern will fetch the frame from the column
    framePattern = r'^([A-Z]{2})+([0-9]{1})+_'
    # these pattern will fetch the 'FD8' or 'FD8_DYN_VOL_03F' from the label column
    flowframePattern = r'^([A-Z]{2})+([0-9]{1})_(.*_[A-Za-z0-9]{3})|([A-Z]{2})+([0-9]{1})'
    result = {"frame": "", "flows": "", "flowframe": "", "circuit": ""}
    # for row in table.rows:  # Iterate through each row in the table
    #     # Iterate through each cell in the row
    #     for cell in row.cells:
    cell = table
    # Check if the cell contains a table
    if cell.tables:
        # Iterate through each nested table in the cell
        for nested_table in cell.tables:
            # Iterate through each row in the nested table
            for nested_row in nested_table.rows:
                # Iterate through each cell in the nested row
                for nested_cell in nested_row.cells:
                    # logging.info the content of the cell
                    logging.info("content inside the table--->", nested_cell.text)
                    if nested_cell.text is not None:
                        data.append(nested_cell.text)
                        logging.info("data = ", data)
                        frame = [string for string in data if re.match(framePattern, string)]
                        result["frame"] = frame
                        if "Flow" in nested_cell.text or "Functional flows" in nested_cell.text:
                            logging.info("hi")
                            if 'Flow' in nested_cell.text or "Functional flows" in nested_cell.text:
                                target_table = table
                                break
                            break
                        if "Label" in nested_cell.text:
                            if "Label" in nested_cell.text:
                                target_table = table
                                break
                            break

    if "Flow" in data[0] or "Functional flows" in data[0]:
        lst = data[1].split('\n')
        lst = [x.strip() for x in lst if x.strip()]
        logging.info("lst--->",lst)
        result["flows"] = lst

        try:
            for string in data:
                match = re.search(flowframePattern, string)
                if match:
                    flowframe = match.group()
            result["flowframe"] = flowframe

        except:
            logging.info("flowframe is not present in requirement table")
            pass
    try:
        if "Flow" in data[0] or "Functional flows" in data[0]:
            lst = data[2].split('\n')
            lst = [x.strip() for x in lst if x.strip()]
            logging.info("lst--->",lst)
            circuit_list = []
            if 'short circuit to ground' in lst:
                circuit_list.append('short circuit to ground')
            if 'open circuit or short circuit to plus' in lst:
                circuit_list.append('open circuit or short circuit to plus')
            if 'open circuit' in lst:
                circuit_list.append('open circuit')
            if 'short circuit to plus' in lst:
                circuit_list.append('short circuit to plus')
            result["circuit"] = circuit_list
            logging.info("Circuitlist--->", circuit_list)
    except:
        logging.info("Label is not present in requirement table")
        pass

    return result

# if __name__ == "__main__":
#      ICF.loadConfig()
#      testPlanReference= "12344_45_45678"
#      functionName = "FSEE_VVXSA"
#      new_req = "REQ-0743278 (B)"
#      # new_req = "REQ-0743275(B)"
#      trigram = "VKG"
#      i = "23456_46_45678"
#
#      creatNewFrameQIA(testPlanReference, functionName, new_req, trigram, i)