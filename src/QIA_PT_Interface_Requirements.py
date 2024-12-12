import ExcelInterface as xi
import re
import WebInterface_For_QIA_PT as wi
import web_interface as wif
import os
from datetime import date
import datetime as dt
import xlwings as xw
import InputConfigParser as ICF
import logging

cols = {-1: "XXXXXXX", 0: "A", 1: "B", 2: "C", 3: "D", 4: "E", 5: "F", 6: "G",
        7: "H", 8: "I", 9: "J", 10: "K", 11: "L", 12: "M", 13: "N", 14: "O",
        15: "P", 16: "Q", 17: "R", 18: "S", 19: "T", 20: "U", 21: "V", 22: "W",
        23: "X", 24: "Y", 25: "Z"}

pointpattern = re.compile(
    r'\bpoint no[\.:]\s*\d+\s*\(ref\.?:?\s*\d{5}_\d{2}_\d{5}\)' + "|" +
    r'\bpoint no[\.:]?\s*\b\d+\b\s*in\s*\d{5}_\d{2}_\d{5}' + "|" +
    r"(?:QIA point)\s*\d+\s* of input document \(?\d{5}_\d{2}_\d{5}\)?" + "|" +
    r"point\s*\(?\d{5}_\d{2}_\d{5}\)?\s*.*\b\d{1,3}\b" + "|" +
    r'\d{5}_\d{2}_\d{5}?.{1, 20}(?:point) (?:no.)\s*\d{1,3}' + "|" +
    r'\(?\d{5}_\d{2}_\d{5}\)?\s*(?:point)?\s*(?:no.)?\s*\d{1,3}' + "|" +
    r'(?:point)\s*(?:no.)?\s*\d{1,3}\s*\d{1,3}.{1,25}\d{5}_\d{2}_\d{5}', re.IGNORECASE)
ppointNumberPattern = re.compile(r'\b\d+\b')

functionalReqPattern = r'REQ-\d{7}\s*(?:\(\w+\)|[A-Z]\b)' + "|" + \
                       r'REQ_\w{4}_\w{3}_\w{3}_\w{3}\s*(?:\(\w+\)|[A-Z]\b).?' + \
                       "|" + r'GEN-(?!.*(?:dci|DCI))'

# DCINT-00001438(2)
dciRequirementPatterns = r'GEN-VHL-DCINT-[A-Za-z0-9_.]*\(\d\)' + \
                         "|" + r'DCINT-\d{8}\(?\d{0,2}\)?'

dciReferenceNumberPattern = re.compile(r'Ref.[\b:]?\d{5}_\d{2}_\d{5}' + "|" +
                                       r'ref \d{5}_\d{2}_\d{5}' + "|" +
                                       r'(?:Ref\. )?\d{5}_\d{2}_\d{5} V\d\.*\d*' + "|" +
                                       r'(?:Ref\. )?\d{5}_\d{2}_\d{5} V\d\.*\d*' + "|" +
                                       r'Ref.[\b:]?\d{5}_\d{2}_\d{5}' + "|" +
                                       r'ref \d{5}_\d{2}_\d{5}' + "|" +
                                       r'ref.[\b:]?\d{5}_\d{2}_\d{5}', re.IGNORECASE)

listOfAcceptedWords = ["Treated", "Completed",
                       "Accepted", "Accepté", "Traitée", "Soldé", "Closed"]
listOfRejectedWords = ["Rejected", "Refusée", "Canceled"]
listOfOpenWords = ["Open", "", None]

listOfAcceptedColumns = ["Etat", "Status", ]
listOfDescriptiveColumns = ["Remarks", "Remark", "Remarque", ]
listOfCommentsColumns = ["Comments", "Answers",
                         "Answers / Comments", "Réponse / Commentaire", ]
listOfLocationColumns = ["Localisation /  Date"]

tdate = dt.datetime.today().strftime("%d_%m_%Y_%H_%M")


# datetime.today().strftime("%d_%m_%Y_%H_%M")


def createOutputFilename():
    outputFolder = os.path.normpath(ICF.getOutputFiles())
    outputFolder = os.path.abspath(outputFolder)
    outputFilename = os.path.join(outputFolder, f"QIA_Results{tdate}.txt")
    commentFileName = os.path.join(outputFolder, f"QIA_Comments{tdate}.txt")
    return outputFilename, commentFileName


def oneOrMany(lis):
    return " are" if len(lis) > 1 else "is"


def compareFunctionalRequirement(str1, str2):
    # Remove parentheses and spaces from both strings
    str1 = str1.replace("(", "").replace(")", "").replace(" ", "")
    str2 = str2.replace("(", "").replace(")", "").replace(" ", "")
    # Compare the modified strings
    return str1 == str2


def getDCIFilesFromFolder(folder_path):
    """
        Gives all dci files form folder

        Parameters:
            folder_path: path in which dci files are stored

        Return:
            list of paths of all dci files present in folder

        Author:
            Yogesh Jagtap
    """
    dciFiles = []
    folder_path = os.path.abspath(folder_path)
    # logging.info("absolute path of folder", folder_path)
    for filename in os.listdir(folder_path):
        # logging.info(f"filename: {filename}")
        file_path = os.path.join(folder_path, filename)

        if os.path.isfile(file_path) and "dci" in filename.lower() and ".xlsx" in filename or "xlsm" in filename:
            # logging.info(f"File path: {file_path}")
            dciFiles.append(file_path)
    return dciFiles


def getDownloadedFileNamesWithReferenceNo(folder_path, ref_no, type_of_file=".xls"):
    """
        Gives list of downloaded files with reference number from given folder
    """
    files = []
    folder_path = os.path.abspath(folder_path)
    logging.info("absolute path of folder", folder_path)
    for filename in os.listdir(folder_path):
        # logging.info(f"filename: {filename}")
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(
                file_path) and ref_no in filename.lower() and type_of_file in filename.lower() and "v0.0" in filename.lower():
            files.append(file_path)
    return files


def getDCIFileinfo(filename):
    """
    Gives version, reference number and name of file from filename
    """
    ref = re.search(r"\d{5}_\d{2}_\d{5}", filename)
    version = re.search(r"V\d+\.\d", filename)
    name = re.search(r"[A-Za-z0-9_\.]*.xlsx$", filename)
    return {"ref": ref.group(0) if ref else "", "version": version.group(0) if version else "",
            "name": name.group(0) if name else ""}


def write(line, opFile):
    """
    The function writes a given line to an output file and handles any exceptions that may occur.
    Parameters
        line (str) : The string that needs to be written to the output file
        opFile (object) : opFile is a file object that represents the output file to which the function will
    write the given line

    Author:
        Saurav Kokane
    """
    try:
        opFile.write(f"{line} \n")
    except Exception as e:
        logging.info(f"Error in write in output file {str(e)}")


def findSheetNameMatch(wbook, sheetName):
    """
        Finds sheet in workbook

        Parameters:
            wbook (object) : excel Interface workbook
            sheetName (str) : name of sheet to be search in workbooks

        Returns:
            returns object of excel sheet if sheet with sheetName found
            else returns None

        Author: Prabhav Pandya
    """
    try:
        for sheet in wbook.sheets:
            if sheet.name.lower().strip() == sheetName.lower().strip():
                return sheet
        return None
    except:
        return None


# @Modifier: Saurav Kokane, last modification 25/07/2023


def findReqInAnalyzeDeEntrant(requirement, analyse_entrant_data):
    """

    Searches functional requirement in analyse de entrant document

    Parameters:
        analyse_de_entrant_doc (str) : analyze de entrant file path(absolute path)
        functionality (str) : sheet or functionality name in which we have to search functional requirement
        requirement (str) : requirement to be searched in analyze de entrant

    Returns:
        data where requirement is found (in new or evolved requirement)
        None when data not found
    Author:
        Saurav Kokane
    """
    reqNamePattern = re.compile(
        r'REQ-\d{7}|GEN-(?!.*(?:dci|DCI))|REQ_\w{4}_\w{3}_\w{3}_\w{3}', re.IGNORECASE)

    reqName = re.search(reqNamePattern, requirement).group()
    data = {"requirement": requirement, "found": False,
            "location": None, "req type": None}

    all_req_types = ['Evolved Requirements', 'New Requirements']
    for feps, fepsData in analyse_entrant_data.items():
        logging.info(feps, fepsData)
        for req_type, reqs in fepsData.items():
            if req_type.strip() not in all_req_types:
                logging.info(f"not a req type, {req_type}")
                continue
            logging.info("-----")
            for req in reqs:
                logging.info(f"req in list {req}")
                if reqName.upper() == req.upper().split(" ")[0]:
                    logging.info(f"req found{req}")
                    data = {"requirement": requirement, "found": True,
                            "location": None, "req type": "new" if "new" in req_type.lower().strip() else "evolved"}
                    return data
    return data


def findDCIRequirementInDCISheets(references=[], dciFiles=[], search_values=[]):
    """
        Finds Interface requirement in dci files

        Parameters:
            file (object) : file object of output file
            references (list(list(str))) : references of files in which we have to search dci requirements
            dciFiles (list(str)) : all dci files
            search values (list(str)) : list of all Interface requirements to search

        Returns:
            Requirements data if requirements found

        Author:
            Yogesh Jagtap
    """

    # logging.info(f"IN DCI Sheets reference: {references}, dcifiles: {dciFiles}, searchvalues: {search_values} ")

    dciReq_data = []
    for ref in references:
        for dciFile in dciFiles:
            if ref[0] not in dciFile:
                # logging.info(f"{ref[0]} not in {dciFile}")
                skip = True
                continue
            try:
                downloaded_book = xi.openExcel(dciFile)
                mux_sheet = downloaded_book.sheets["MUX"]
                for i in range(1, mux_sheet.used_range.last_cell.row + 1):
                    try:
                        aval = xi.getDataFromCell(
                            mux_sheet, f'A{i}')
                        match_found = False
                        for value in search_values:
                            if value and aval and value.strip() in str(aval).strip():
                                # file.write(
                                #     f"{value} matched at row {i} in file: {dciFile}\n")
                                info = {"Req": value, "Flux": xi.getDataFromCell(mux_sheet, f"C{i}"),
                                        "CorP": xi.getDataFromCell(
                                            mux_sheet, f"K{i}"), "DCISheet": getDCIFileinfo(dciFile)}
                                dciReq_data.append(
                                    info)
                                # logging.info(f"info {info}")

                                # pass
                    except Exception as e:
                        logging.info(
                            f'''Error in {dciFile} in MUX at {i}th row: {str(e)}, {e.__cause__}''')
            except Exception as e:
                logging.info(
                    f"Error in findDCIRequirementInDCISheets {e} {e.__cause__}")
            finally:
                try:
                    downloaded_book.close()
                except:
                    pass
    return dciReq_data


#  Error is ocurring of visual basic.
def getFunctionalRequirementsFromQIAInputDocument(point_no, QIADocumentPath, colsLocatedAt=7):
    '''
        Finds functional requirement in qia input document
        Parameters:
            point_no (int, str): interger point number to find in  qia input
            QIADocumentPath (str): filepath of qia input document
            colsLocatedAt (int): row number at which column names are located
        Returns:
            tuple with status of point and functional requirements
    '''
    logging.info(f"handling point no: {point_no}, {QIADocumentPath}")
    Requirement_info = ("", [])

    try:
        QIA_Book = None
        QIA_Book = xi.openExcel(QIADocumentPath)

        QIA_sheet = findSheetNameMatch(QIA_Book, "Remarks follow up")
        # logging.info("Opened", QIA_sheet)

        QIA_sheet = findSheetNameMatch(
            QIA_Book, "Suivi des Remarques") if QIA_sheet == None else QIA_sheet

        logging.info(QIA_sheet)
        if QIA_sheet == None:
            raise Exception(
                "Could not find \"Suivi des Remarques\" or \"Remarks follow up\" sheet")
        logging.info("Opened")
        statusColLoc = -1
        descriptionColLoc = -1
        answerColLoc = -1
        localisationColLoc = -1
        for j in range((len(cols)) - 1):
            colname = xi.getDataFromCell(
                QIA_sheet, f"{cols[j]}{colsLocatedAt}")
            # logging.info(j, colname)
            if colname in listOfAcceptedColumns:
                # logging.info("listOfAcceptedColumns", j, colname)
                statusColLoc = j
            elif colname in listOfDescriptiveColumns:
                # logging.info("listOfDescriptiveColumns", j, colname)
                descriptionColLoc = j
            elif colname in listOfCommentsColumns:
                # logging.info("listOfCommentsColumns", j, colname)
                answerColLoc = j
            elif colname in listOfLocationColumns:
                # logging.info("listOfLocationColumns", j, colname)
                localisationColLoc = j
            else:
                # logging.info("continue")
                pass

        # logging.info(statusColLoc, descriptionColLoc, answerColLoc, localisationColLoc)
        logging.info("mid of the function", statusColLoc,
                     descriptionColLoc, answerColLoc, localisationColLoc)
        for i in range(colsLocatedAt + 1, QIA_sheet.used_range.last_cell.row + 1):
            # logging.info("start", QIA_sheet)
            cellDataA = xi.getDataFromCell(QIA_sheet, f"A{i}")
            # logging.info(type(cellDataA), cellDataA, point_no)
            try:
                point = float(cellDataA)
            except:
                pass
            point = str(int(cellDataA)) if str(type(cellDataA)) in [
                "<class 'float'>", "<class 'int'>"] else ""

            if point == str(point_no).strip():

                cellDataJ = xi.getDataFromCell(
                    QIA_sheet, f"{cols[statusColLoc]}{i}")
                cellDataJ = cellDataJ.strip() if cellDataJ else ""
                cellDataE = xi.getDataFromCell(
                    QIA_sheet, f"{cols[descriptionColLoc]}{i}")
                cellDataE = cellDataE if cellDataE else ""
                cellDataM = xi.getDataFromCell(
                    QIA_sheet, f"{cols[answerColLoc]}{i}")
                cellDataM = cellDataM.split("\n") if cellDataM else ""

                functional_requirements = []

                functional_requirements.extend(re.findall(
                    functionalReqPattern, cellDataE))
                for line in cellDataM:
                    functional_requirements.extend(re.findall(
                        functionalReqPattern, line))
                logging.info("Functional Requirement", functional_requirements)
                if cellDataJ in listOfAcceptedWords:

                    # logging.info(f"{cellDataA}point is accepted", cellDataJ)
                    pass
                elif cellDataJ in listOfRejectedWords:
                    Requirement_info = ("Rejected", functional_requirements)
                    # logging.info(f"{cellDataA}point is rejected", cellDataJ)
                    try:
                        if QIA_Book.fullname in [x.fullname for x in xw.books]:
                            QIA_Book.close()
                    except:
                        pass
                    return ("Rejected", functional_requirements)
                elif not cellDataJ or cellDataJ in listOfOpenWords:
                    # logging.info(f"{cellDataA}point is open {cellDataJ}")
                    Requirement_info = ("Open", functional_requirements)
                    try:
                        if QIA_Book.fullname in [x.fullname for x in xw.books]:
                            QIA_Book.close()
                    except:
                        pass
                    return ("Open", functional_requirements)
                else:
                    Requirement_info = (cellDataJ, functional_requirements)
                    # logging.info(f"{cellDataA}point is {cellDataJ}")
                    try:
                        if QIA_Book.fullname in [x.fullname for x in xw.books]:
                            QIA_Book.close()
                    except:
                        pass
                    return (cellDataJ, functional_requirements)

                if not len(functional_requirements) or not functional_requirements:
                    Requirement_info = (
                        "Accepted", functional_requirements if functional_requirements else [])
                    # logging.info(f"{cellDataA}point is accepted2, has  no requirement")
                    try:
                        if QIA_Book.fullname in [x.fullname for x in xw.books]:
                            QIA_Book.close()
                    except:
                        pass
                    return ("Accepted", [])
                else:
                    Requirement_info = ("Accepted", functional_requirements)
                    # logging.info(f"{cellDataA}point is accepted  and has reqs {functional_requirements}")
                    try:
                        if QIA_Book.fullname in [x.fullname for x in xw.books]:
                            QIA_Book.close()
                    except:
                        pass
                    return ("Accepted", functional_requirements)
    except Exception as e:
        logging.info("_" * 20, e.__str__(), e.__cause__, e,
                     " in getFunctionalRequirementsfromQIA inputDocuments ", "_" * 20)

    logging.info("before last")
    try:
        if QIA_Book.fullname in [x.fullname for x in xw.books]:
            QIA_Book.close()
    except:
        pass
    # logging.info("at last", cellDataJ)
    return Requirement_info


def dcifileTreatment(qiabookPath, analyse_entrant_data, user=""):
    dt = date.today().strftime("%d/%m/%Y")
    # file = open(outputFileName, "w")

    # This code is searching for all files in the directory "./docs" that have "dci" in their filename
    outputFileName = createOutputFilename()[0]
    folder_path = wi.destinationFolder
    dciFiles = []
    opFile = open(outputFileName, "w")
    file1 = open(createOutputFilename()[1], "w")
    write("Output: ", file1)
    write("Output:", opFile)
    null_counter = 0
    try:
        qiaBook = xi.openExcel(qiabookPath)
        # added by saurav from line 441 to 452
        qiaSheet = findSheetNameMatch(qiaBook, "QIA")
        if qiaSheet is None:
            qiaSheet = findSheetNameMatch(qiaBook, "New QIA")
        if qiaSheet is None:
            qiaSheet = findSheetNameMatch(qiaBook, "Old QIA")

        if qiaSheet is None:
            logging.info("There is no qia sheet present in qia workbook")
            return
        logging.info("Files opened")
        final_search_values = []
        final_search_ref = []
        final_all_values = []
        logging.info("No of row in sheet", qiaSheet.used_range.last_cell.row)
        for i in range(2, qiaSheet.used_range.last_cell.row + 1):
            search_values = []
            search_ref = []
            try:
                cellDataA = xi.getDataFromCell(qiaSheet, f'A{i}')
                cellDataC = xi.getDataFromCell(qiaSheet, f'C{i}')
                cellDataJ = xi.getDataFromCell(qiaSheet, f'J{i}')
                cellDataE = xi.getDataFromCell(qiaSheet, f'E{i}')
                cellDataM = xi.getDataFromCell(qiaSheet, f"M{i}")

                cellDataA = cellDataA if cellDataA else ""
                cellDataC = cellDataC if cellDataC else ""
                cellDataJ = cellDataJ if cellDataJ else ""
                cellDataE = cellDataE if cellDataE else ""
                cellDataM = cellDataM if cellDataM else ""

                # added by Saurav, line 462 to 472

                if null_counter >= 5:
                    break
                if cellDataA != "" and cellDataC != "" and cellDataJ != "" and cellDataE != "" and cellDataM != "":
                    null_counter = 0
                else:
                    null_counter += 1
                    if null_counter >= 5:
                        break

                if cellDataC and cellDataE and cellDataJ and cellDataJ == "Opened":
                    logging.info("**" * 30)
                    search_values.extend(re.findall(
                        dciRequirementPatterns, cellDataC))
                    search_values.extend(re.findall(
                        dciRequirementPatterns, cellDataE))
                    search_ref.extend(re.findall(
                        dciReferenceNumberPattern, cellDataE))
                    search_ref.extend(re.findall(
                        dciReferenceNumberPattern, cellDataE))
                    search_ref.extend(re.findall(
                        dciReferenceNumberPattern, cellDataE))

                    search_ref.extend(re.findall(
                        dciReferenceNumberPattern, cellDataM))
                    search_ref.extend(re.findall(
                        dciReferenceNumberPattern, cellDataM))
                    search_ref.extend(re.findall(
                        dciReferenceNumberPattern, cellDataM))

                    search_values = list(set(search_values))
                    search_ref = list(set(search_ref))

                    final_search_values.extend(search_values)
                    final_search_ref.extend(search_ref)
                    logging.info(f"search values {search_values}")
                    if len(search_values):
                        allFunctionalReqs = []
                        pointNo = re.findall(pointpattern, cellDataE)
                        # logging.info(f"Point numbers found:-------------------{cellDataA}{pointNo} {cellDataE} {cellDataM}")
                        for line in cellDataM.split("\n"):
                            reqs = re.findall(functionalReqPattern, line)
                            if not len(reqs):
                                reqs = re.findall(r'REQ-\d{7}', line)
                            pointNo.extend(re.findall(pointpattern, line))
                            # logging.info(f"Point numbers found:-------------------{pointNo} {line}")
                            if len(reqs):
                                allFunctionalReqs.append(reqs)
                        allFunctionalReqs = allFunctionalReqs[-1] if len(
                            allFunctionalReqs) else None

                        logging.info(
                            f"{cellDataA} Reference numbers {search_ref}, {pointNo} {allFunctionalReqs}")
                        searches = []
                        references = []
                        for ref in search_ref:
                            refg = re.search(
                                r"\d{5}_\d{2}_\d{5}", str(ref))
                            ref_no = refg.group(0) if refg else ""
                            verg = re.search(re.compile(
                                r"V\d+\.?\d*", re.IGNORECASE), str(ref))
                            version = verg.group(0).split(
                                "V")[1] if verg else ""
                            references.append((ref_no, version))
                            # logging.info(ref_no, version)
                        references = list(map(list, references))
                        if ICF.getAutoDownloadStatusInputDocument():
                            wi.startDocumentDownload(references)

                        # downloadRefs = [ref.split()[1:]
                        #                 for ref in search_ref]
                        # logging.info(f"downloadRefs{int(cellDataA)} {search_ref}")
                        write(
                            f"{i} point No. {int(cellDataA)}----, functional reqs:{search_values} {search_ref}------- pointNo in comment: {pointNo} -----func req: {allFunctionalReqs} ",
                            opFile)
                        logging.info(
                            f"{i} point No. {int(cellDataA)}----, functional reqs:{search_values} {search_ref}------- pointNo in comment: {pointNo} -----func req: {allFunctionalReqs} ")

                        # ********************************************************************************************

                        dciFiles = getDCIFilesFromFolder(
                            folder_path=folder_path)

                        dciReq_data = []
                        try:
                            dciReq_data = findDCIRequirementInDCISheets(
                                references, dciFiles, search_values)
                        except Exception as e:
                            logging.info(e, e.__cause__)
                        pt = ""

                        logging.info("Checking for functional requirement or point number")
                        # cas1 1
                        if allFunctionalReqs:
                            logging.info(
                                f"functional requirement at {cellDataA}, {allFunctionalReqs}")
                            try:  # try:
                                data = []
                                msg = ""
                                notFoundInAnalyse_de_entrant = []
                                foundInAnalyse_de_entrant = []
                                for functional_requirement in allFunctionalReqs:
                                    logging.info(functional_requirement)
                                    available = findReqInAnalyzeDeEntrant(
                                        functional_requirement, analyse_entrant_data)
                                    logging.info(functional_requirement, available)
                                    if not available["found"]:
                                        notFoundInAnalyse_de_entrant.append(
                                            available['requirement'])
                                    else:
                                        foundInAnalyse_de_entrant.append(
                                            available)

                                if len(notFoundInAnalyse_de_entrant):
                                    msg += " and ".join(notFoundInAnalyse_de_entrant) + " " + oneOrMany(
                                        notFoundInAnalyse_de_entrant) + " not comes to update and it will be treated once the requirement comes to update. "
                                if len(foundInAnalyse_de_entrant):
                                    newRequirements = []
                                    evolvedRequirements = []
                                    for req in foundInAnalyse_de_entrant:
                                        if req['req type'] == "new":
                                            newRequirements.append(
                                                req['requirement'])
                                        elif req['req type'] == "evolved":
                                            evolvedRequirements.append(
                                                req['requirement'])
                                    if len(newRequirements):
                                        msg += " and ".join(newRequirements) + \
                                               " will be treated as new requirements"
                                    if len(evolvedRequirements):
                                        msg += " and ".join(evolvedRequirements) + \
                                               " will be treated as evolved requirements"
                                msg = f"{user} {dt}: " + \
                                      msg if msg != "" else ""
                                logging.info(msg)

                                write(f"{cellDataA} {msg}", file1)
                                qiaSheet.range(
                                    f"M{i}").value = cellDataM + "\n" + msg
                            except Exception as e:
                                logging.info(
                                    f"Error in dci file treatement {e}, {e.__cause__}")

                        elif len(pointNo):
                            logging.info("*" * 10, pointNo, "*" * 10)
                            try:
                                logging.info(cellDataA, pointNo)
                                msg = ""
                                #                     logging.info("<"* 20, i, cellDataA, pointNo, ">"*20)
                                if str(
                                        type(pointNo)) == "<class 'list'>":
                                    pointNo = pointNo[-1]
                                elif str(type(pointNo)) == "<class 'str'>":
                                    pass
                                else:
                                    pointNo = ""
                                    continue
                                logging.info(
                                    f"{cellDataA} point no present in it {pointNo}")
                                refr = re.search(
                                    r"\d{5}_\d{2}_\d{5}", pointNo)
                                refr = refr.group() if refr else ""

                                pt = re.search(
                                    ppointNumberPattern, pointNo)
                                pt = pt.group() if pt else ""
                                logging.info("point Number and ref no", pt, refr)

                                wi.startDocumentDownload([[refr, "latest"]])
                                # if ICF.getAutoDownloadStatusInputDocument():
                                #     wi.startDocumentDownload(
                                #         [[refr, "latest"]])
                                files = getDownloadedFileNamesWithReferenceNo(
                                    wi.destinationFolder, refr)
                                functionalRequirements = []
                                for QIAInputfile in files:
                                    functionalRequirement = getFunctionalRequirementsFromQIAInputDocument(
                                        pt, QIAInputfile)
                                    functionalRequirements.append(
                                        functionalRequirement)

                                notFoundInAnalyse_de_entrant = []
                                foundInAnalyse_de_entrant = []
                                for status, requirements in functionalRequirements:
                                    logging.info("req in input doc ",
                                                 status, requirements)
                                    if status == "Rejected":
                                        msg += f"  is cancelled state, need to treat manually"
                                    elif status == "Open":
                                        msg += f" The QIA of input document (point no. {pt} {refr}) is still in OPEN, it will be treated in next update."
                                    elif status == "Accecpted":
                                        available = None
                                        if len(requirements):

                                            for requirement in requirements:
                                                available = findReqInAnalyzeDeEntrant(
                                                    requirement, analyse_entrant_data)
                                                if not available["found"]:
                                                    notFoundInAnalyse_de_entrant.append(
                                                        requirement['requirement'])
                                                else:
                                                    foundInAnalyse_de_entrant.append(
                                                        requirement['requirement'])
                                        else:
                                            msg += " Requirements not found in input document, point need to be treat manually"
                                    else:
                                        msg += f" is in {status if status else ''} state, with {' and '.join(requirements) if len(requirements) else 'no'} requirements."
                                if len(notFoundInAnalyse_de_entrant):
                                    msg += f" " + " and ".join(notFoundInAnalyse_de_entrant) + " " + oneOrMany(
                                        notFoundInAnalyse_de_entrant) + " not comes to update and it will be treated once the requirement comes to update. "
                                if len(foundInAnalyse_de_entrant):
                                    newRequirements = []
                                    evolvedRequirements = []
                                    for req in foundInAnalyse_de_entrant:
                                        if req['req type'] == "new":
                                            newRequirements.append(
                                                req['requirement'])
                                        elif req['req type'] == "evolved":
                                            evolvedRequirements.append(
                                                req['requirement'])
                                    if len(newRequirements):
                                        msg += " and ".join(newRequirements) + \
                                               " will be treated as new requirements"
                                    if len(evolvedRequirements):
                                        msg += " and ".join(evolvedRequirements) + \
                                               " will be treated as evolved requirements"

                                msg = f"{user} {dt}: At point no. {pt} in ref. {refr} " + \
                                      msg if msg != "" else ""
                                logging.info(msg)
                                write(f"{cellDataA} {msg}", file1)
                                qiaSheet.range(
                                    f"M{i}").value = cellDataM + "\n" + msg
                            except Exception as e:
                                logging.info(
                                    f"Error in case 2 of dci file treatement, {e}, {e.__cause__}")

                        final_value = {"line no": i, "point No": int(cellDataA), "DCI REQS": search_values,
                                       "refs": references, "comment data": {
                                "pointNo": pt, "funcReq": allFunctionalReqs}, "reqsData": dciReq_data}

                        final_all_values.append(final_value)
            except Exception as e:
                logging.info(
                    f"Error in dcifileTreatement at row {i} {e}, {e.__cause__}")
            finally:
                pass
        # logging.info("checkPoint 2")
        for value in final_all_values:
            logging.info(value)
            write(f"{value}", opFile)
    except Exception as e:
        logging.info(
            f"Error in qia file opening in dcifiletreatement {e}, {e.__cause__}")
    finally:
        # closing qia excel file
        if qiaBook.fullname in [x.fullname for x in xw.books]:
            qiaBook.save()
            qiaBook.close()
    if opFile:
        opFile.close()
    logging.info("QIA all values completed")


def findQiaSheet(inputFolder, qia_sheet_pattern, qia_reference):
    qia_sheet = None
    for filename in os.listdir(inputFolder):
        file_path = os.path.join(inputFolder, filename)
        if len(re.findall(qia_sheet_pattern, file_path)) and qia_reference in file_path:
            qia_sheet = file_path
            break
    return qia_sheet

# @Modifier: Saurav Kokane, last modification 25/07/2023


def downloadQIA_PTDocument(taskName, tpBook):
    # analyse_de_entrants = []
    # ade_pattern = re.compile(r'(?:VSM|BSI)_ANALYSE_DES_ENTRANT', re.IGNORECASE)
    # test_plan_pattern = re.compile(
    #     r'[^_a-zA-Z0-9]Tests_\d{2}_\d{2}_\d{5}_\d{2}_\d{5}_[a-zA-Z0-9_]+(?:VSM|BSI)', re.IGNORECASE)
    qia_sheet_pattern = re.compile(
        r'[^_a-zA-Z0-9]QIA_\d{2}_\d{2}_\d{5}_\d{2}_\d{5}_[a-zA-Z0-9_]+(?:VSM|BSI)', re.IGNORECASE)
    inputFolder = ICF.getInputFolder()
    inputFolder = os.path.abspath(inputFolder)
    logging.info(inputFolder)
    # qia_sheet = None #commented on AUG 15 modifier: Priya

    testsheet = findSheetNameMatch(tpBook, "Sommaire")
    testsheet = testsheet if testsheet else findSheetNameMatch(
        tpBook, "Summary")
    qia_reference = xi.getDataFromCell(testsheet, "F1")
    if qia_reference == None:
        return
    logging.info(qia_reference)
    # wif.startDocumentDownload([[qia_reference, ""]]) #commented on AUG 15 modifier: Priya
    # if ICF.getAutoDownloadStatusInputDocument():
    #     wif.startDocumentDownload([[qia_reference, ""]])

    # start commented on AUG 15 modifier: Priya
    # for filename in os.listdir(inputFolder):
    #     file_path = os.path.join(inputFolder, filename)
    #     if len(re.findall(qia_sheet_pattern, file_path)) and qia_reference in file_path:
    #         qia_sheet = file_path
    #         break
    # end commented on AUG 15 modifier: Priya

    #Added on AUG 15 Line 784-786 modifier: Priya
    qia_sheet = findQiaSheet(inputFolder, qia_sheet_pattern, qia_reference)
    if qia_sheet is None:
        wif.startDocumentDownload([[qia_reference, ""]])
        qia_sheet = findQiaSheet(inputFolder, qia_sheet_pattern, qia_reference)

    if not qia_sheet:
        logging.info(f"There is no QIA of PT file in input folder: {inputFolder}")
    return qia_sheet


# @Modifier: Saurav Kokane, last modification 25/07/2023


def execute_interface_requirement_treatment(task_name, tpBook, analyse_entrant_data):
    qia_sheet = downloadQIA_PTDocument(task_name, tpBook)
    logging.info(qia_sheet)
    dcifileTreatment(qia_sheet, analyse_entrant_data=analyse_entrant_data, user=ICF.gettrigram())


if __name__ == "__main__":
    # qiaDocument = input("Enter path of QIA of PT: ")
    # qiaDocument = os.path.abspath(qiaDocument)
    # analyse_de_entrant = input("Enter path of analyse de entrant: ")
    # analyse_de_entrant = os.path.abspath(analyse_de_entrant)
    # taskName = input(
    #     "Enter name of taskName or sheet name in analyse de entrant: ")
    # dcifileTreatment(qiaDocument,
    #                  analyse_de_entrant, taskName)

    ICF.loadConfig()
    # logging.info(ICF.getTaskName())
    # execute_interface_requirement_treatment(ICF.getTaskName())
    xi.openTestPlan()
