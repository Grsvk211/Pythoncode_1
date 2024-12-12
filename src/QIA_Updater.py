import ExcelInterface as EI


def UpdateQiaParamGlobal(qiaSheet, valueMap):
    maxrow = qiaSheet.range('B' + str(qiaSheet.cells.last_cell.row)).end('up').row
    for key in valueMap:
        EI.setDataFromCell(qiaSheet, key + str(maxrow + 1), valueMap[key])

# 6 --> Column F
# if frame is on FD7 == column H need to place NEA_R1_2 or
# if frame is on FD8 == column H need to place NEA_R1_X or
# if frame is on HS7 == column H need to place NEA_R1|NEA_R1_1 or
# if frame is on others == column H need to place NEA
# archi_values = [{'FD7': 'NEA_R1_2', 'FD8': 'NEA_R1_X', 'HS7': 'NEA_R1|NEA_R1_'}]
# df = []
# for i, row in df.iterrows():
#     if row['FD7'] == 'yes':
#         df.loc[i, 'NEA'] = 'NEA_R1_2'
#     elif row['FD8'] == 'yes':
#         df.loc[i, 'NEA'] = 'NEA_R1_X'
#     elif row['HS7'] == 'yes':
#         df.loc[i, 'NEA'] = 'NEA_R1|NEA_R1_1'
#     else:
#         df.loc[i, 'NEA'] = 'NEA'

# def creatNewFrameQIA(testPlanReference,taskName,new_req,trigram):
#     logging.info("\nProcessing the QIA Sheet!!!\n")
#     QiaFile = findInputFiles()[10]
#     path= ICF.getInputFolder() + "\\" + QiaFile
#     isQAFileExist = os.path.isfile(path)
#     logging.info("QA File Exist- ", isQAFileExist)
#     Today = date.today()
#     curr_date = Today.strftime("%m/%d/%Y")
#     newRequir = re.sub(r'(\w-\d+)\s+(\()', r'\1\2', newreq) it used for to remove the gap between the reqno and version
#     if isQAFileExist:
#         QIABook = EI.openExcel(ICF.getInputFolder() + "\\" + QiaFile)
#         if QIABook:
#             QIAsheet = QIABook.sheets["NEW_QIA"]
#             QIAsheet.activate()
#             ntew = "FD7"
#             ID = "ID03F"
#             canid_frame_name = "FD8_DYN_VOL_03F"
#             valueMap00 = {
#                     "B": testPlanReference,
#                     "C": 'taskName',
#                     "D": 'Creation',
#                     "E": 'New Flow for NEA',
#                     "F": '',
#                     "G": '--',
#                     "H": str("TRAME_"+ntew+"_"+ID+"_"+canid_frame_name),
#                     "I": newRequir,
#                     "J": '--',
#                     "K": '--',
#                     "L": '--',
#                     "M": '--',
#                     "N": str("CAN_"+ntew+"/"+canid_frame_name),
#                     "O": '',
#                     "P": trigram,
#                     "Q": curr_date,
#                     "R": 'Open',
#                     "U": str(curr_date + " " + tirgram + " : Information can be found on DCI Ref : " + Reference_of_DCI + "."),
#                 }
#             UpdateQiaParamGlobal(QIAsheet, valueMap00)


# if __name__=="__main__":
