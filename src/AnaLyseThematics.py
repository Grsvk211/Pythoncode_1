import re
import ParseThematics as parse
import ExcelInterface as EI
import TestPlanMacros as TPM
import KeyboardMouseSimulator as KMS
import InputConfigParser as ICF
import WordDocInterface as WDI
import time
import difflib
import DocumentSearch as DS
import logging
from Backlog_Handler import remove_trailing_and_or


class AnalyseThematics:
    # funcImpactComment = 0
    def __init__(self, tpBook, refEC, currDoc, prevDoc, listOfTestSheets, reqName, reqVer, Arch, newReq=""):
        self.ARCH = Arch
        self.tpBook = tpBook
        self.refEC = refEC
        self.listOfSheets = listOfTestSheets
        self.testSheet = listOfTestSheets[0]
        self.oldRawThm = ""
        self.newRawThm = ""
        self.reqName = reqName
        self.reqVer = str(reqVer)
        self.testReqName = ""
        self.testReqVer = ""
        self.currDoc = currDoc
        self.prevDoc = prevDoc
        self.funcImpact = ""
        self.newReq = newReq
        self.comment = ""
        # self.themImpactComment = []
        # self.funcImpactComment = 0

    def getTsRawReq(self):
        return self.testSheet.range('C4').value

    def getTestSheetRequirement(self):
        getReqList = self.testSheet.range('C4').value.split("|")
        logging.info("before req list = ", getReqList)
        # Use list comprehension to remove empty elements
        getReqList = [item for item in getReqList if item]

        # Print the filtered list
        logging.info("after req list = ", getReqList)
        for req in getReqList:
            if req.find("(") != -1:
                tempReqName = req.split("(")[0]
                tempReqVer = req.split("(")[1].split(")")[0]
                if tempReqName == self.reqName.strip():
                    self.testReqName = tempReqName
                    self.testReqVer = tempReqVer
            else:
                logging.info("req = ", req)
                tempReqName = req.split()[0]
                try:
                    tempReqVer = req.split()[1]
                except:
                    tempReqVer = ""
                if tempReqName == self.reqName.strip():
                    self.testReqName = tempReqName
                    self.testReqVer = tempReqVer
        return getReqList

    def createThmCombinations(self, thm):
        logging.info("Creating Thematic Combinations ")
        rawCombinations = parse.createCombination(thm)
        return rawCombinations

    def grepThematicsCode(self, rawThematics):
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
            # logging.info("a = ", a)
            # logging.info(a.find("{"))
            # logging.info(re.search("[(]", a))
            if re.search("[a-zA-Z0-9]{3}[(][0-9]{2}[)]", a) is not None:
                a = a.replace("(", "_")
                a = a.replace(")", "")
                a = a.strip()
                thematics_code.append(a)
                logging.info("thematics_code = ", thematics_code)
            else:
                if a.find("AND")==0:
                    if (thematics_code[-1]!="AND"):
                        thematics_code.append(a)
                if a.find("OR")==0:
                    if (thematics_code[-1]!="OR"):
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
        if len(re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", thematics_code[0]))==0:
            if thematics_code[0]!="(":
                logging.info("removing first element", thematics_code[0],
                      re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", thematics_code[0]))
                thematics_code.remove(thematics_code[0])
        # elif thematics_code[0] != "(":
        #     thematics_code.remove(thematics_code[0])
        logging.info("Thematic = ", thematics_code)
        logging.info("Thematic code final(1) = ", ''.join(thematics_code))

        openBracket = []
        closeBracket = []
        for i in range(len(thematics_code)):
            if thematics_code[i]=="(":
                openBracket.append(i)
            if thematics_code[i]==")":
                closeBracket.append(i)
        logging.info("Indices = ", openBracket, closeBracket)
        logging.info("Thematic code = ", thematics_code)
        for n, i in enumerate(thematics_code):
            # logging.info("N & i", n, i)
            if i.find('_')!=-1:
                # logging.info("thm code",thematics_code[n+1])
                if n < (len(thematics_code) - 1):
                    if thematics_code[n + 1].find('_')!=-1:
                        thematics_code[n + 1] = ',' + thematics_code[n + 1]

        reducedThm = ' '.join(thematics_code)
        logging.info("Thematic code final(2) = ", reducedThm)
        return reducedThm

    def filterThemForArch(self, thematicLine):
        # refBook = xw.Book(r"C:/Users/6451/Desktop/bsi_auto/09-11-2021/Modified/Aptest/Input/Referentiel_EC.xlsm")
        logging.info("In filterThemForArch function - ", thematicLine)
        ListOfThematics = thematicLine.split("|")
        time.sleep(1)
        sheet = self.refEC.sheets['Liste EC']
        logging.info("sheet = ", sheet)
        maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        sheet_value = sheet.used_range.value
        logging.info("Maxrows  and Listof thematics----->>>", maxrow, ListOfThematics)
        # final_VSMr2_thm = []
        tempflagR1 = 0
        tempflagR2 = 0
        ListOfThematicsCopy = ListOfThematics.copy()
        for i in ListOfThematicsCopy:
            flagR1 = 0
            flagR2 = 0
            logging.info("In filterThemForArch (maxrow, i, sheet)", maxrow, i, sheet)
            try:
                # searchResults = EI.searchDataInExcel(sheet, (maxrow, 7), i)
                searchResults = EI.searchDataInExcelCache(sheet_value, (maxrow, 7), i)
            except:
                # searchResults = EI.searchDataInExcel(sheet, (maxrow, 7), i)
                searchResults = EI.searchDataInExcelCache(sheet_value, (maxrow, 7), i)
            # logging.info("searchresult------->a5:", searchResults)
            if searchResults["count"]!=0:
                x, y = searchResults["cellPositions"][0]
                applicableBSI = sheet.range(x, y + 38).value
                applicableR1 = sheet.range(x, y + 39).value
                applicableR2 = sheet.range(x, y + 40).value
                logging.info("Thematique = ", i, "Aplicable to = ", applicableBSI, applicableR1, applicableR2)
                # for BSi Arch
                if self.ARCH=="BSI":
                    if applicableBSI=="Y":
                        pass
                    else:
                        ListOfThematics.remove(i)
                        logging.info("not applicable for BSI but its present in req\n")
                # for VSM Arch
                elif self.ARCH=="VSM":
                    if (applicableR1=="Y") and (applicableR2=="Y"):
                        logging.info("Aplicable to R1 & R2")
                        pass
                    elif (applicableR1=="Y") or (applicableR2=="Y"):
                        if (applicableR1=="Y"):
                            flagR1 = 1
                            logging.info("NEA R1 applicable")
                            tempflagR1 = flagR1
                        elif (applicableR2=="Y"):
                            flagR2 = 1
                            logging.info("NEA R2 applicable")
                            tempflagR2 = flagR2
                    else:
                        ListOfThematics.remove(i)
                        logging.info("not applicable for VSM but its present in req\n")
                else:
                    logging.info("arch not found\n")
            else:
                logging.info("Thematique not found in referntial EC")
                return -1
        logging.info("TempFlag = ", tempflagR1, tempflagR2)
        if (tempflagR1==1) and (tempflagR2==1):
            logging.info("CONFLICT")
            return -1
        else:
            logging.info("Returning thematics ", ListOfThematics)
            i = 1
            thematicLine = ""
            for l in ListOfThematics:
                thematicLine = thematicLine + l
                if i < len(ListOfThematics):
                    thematicLine = thematicLine + "|"
                    i = i + 1
            logging.info("Returning thematic line", thematicLine)
            return thematicLine

    def getThematicsFromPT(self):
        EI.activateSheet(self.tpBook, self.testSheet)
        self.testSheet_value = self.testSheet.used_range.value
        # minrowCellValue = EI.searchDataInExcel(self.testSheet, (1, 100), "THEMATIQUE")
        minrowCellValue = EI.searchDataInExcelCache(self.testSheet_value, (1, 100), "THEMATIQUE")
        a, b = minrowCellValue["cellPositions"][0]
        minrow = a, b + 2
        # logging.info("minrow = ", minrow)
        # maxrowCellValue = EI.searchDataInExcel(self.testSheet, (26, 100), "CONDITIONS INITIALES")
        maxrowCellValue = EI.searchDataInExcelCache(self.testSheet_value, (26, 100), "CONDITIONS INITIALES")
        x, y = maxrowCellValue["cellPositions"][0]
        # logging.info("x, y = ", x, y)
        maxrow = (x - 2, y + 2)
        # logging.info(maxrow)
        condition = range(a, x - 2)
        # logging.info("condition", condition)
        them = []
        for row in range(a, x - 1):
            col = 3
            # logging.info("row = ", row)
            thematicLine = EI.getDataFromCell(self.testSheet, (row, col))
            # logging.info(thematicLine)
            them.append(thematicLine)
        # logging.info(them)
        return them

    def checkThematicsInPT(self, thematics, newThem, themImpactComment):
        macro = EI.getTestPlanAutomationMacro()
        logging.info("In checkThematicsInPT function ", thematics, newThem, themImpactComment)
        global funcImpactComment
        # funcImpactComment = 0
        thmLineBefore = EI.getDataFromCell(self.testSheet, (8, 1))
        thmLineBefore = thmLineBefore.strip()
        try:
            thmLineBefore = thmLineBefore.split()[1]
        except:
            thmLineBefore = "1"
        logging.info("thmLineBefore = ", thmLineBefore, type(thmLineBefore))
        logging.info("In checkThematicsInPT = ", thematics, newThem)
        if (''.join(thematics))=="--":
            logging.info("No thematic present")
            EI.setDataFromCell(self.testSheet, (8, 3), newThem)
            logging.info("Thematiques11 " + str(', '.join(newThem)) + " added at thematique line 1")
            funcImpactComment = 1
            for i in newThem.split("|"):
                themImpactComment.append(i)
            # self.comment = self.comment + "Thematiques " + str(
            #     ', '.join(newThem.split("|"))) + " added at thematique line 1."
            self.comment = self.comment +"."
            logging.info("checkThematicsInPT parametres(1) - ", funcImpactComment, themImpactComment, self.comment)
        else:
            flag = 0
            splittedNewThm = newThem.split("|")
            logging.info("splittedNewThm = ", splittedNewThm, len(splittedNewThm))
            if len(splittedNewThm)==1:
                for n, thematic in enumerate(thematics):
                    logging.info("thematic", thematic, n, ''.join(splittedNewThm), thematic.find(''.join(splittedNewThm)))
                    if thematic.find(''.join(splittedNewThm))!=-1:
                        logging.info("Thematic already present(1)")
                        break
                    else:
                        if n==(len(thematics) - 1):
                            logging.info("Thems22")
                            funcImpactComment = 1
                            # minrowCellValue = EI.searchDataInExcel(self.testSheet, (1, 100), "THEMATIQUE")
                            minrowCellValue = EI.searchDataInExcelCache(self.testSheet_value, (1, 100), "THEMATIQUE")
                            a, b = minrowCellValue["cellPositions"][0]
                            # maxrowCellValue = EI.searchDataInExcel(self.testSheet, (26, 100), "CONDITIONS INITIALES")
                            maxrowCellValue = EI.searchDataInExcelCache(self.testSheet_value, (26, 100), "CONDITIONS INITIALES")
                            x, y = maxrowCellValue["cellPositions"][0]
                            for row in range(a, x - 1):
                                col = 3
                                thematicLine = EI.getDataFromCell(self.testSheet, (row, col))
                                splittedThem = thematicLine.split("|")
                                splittedThem.sort()
                                if thematicLine.find(''.join(splittedNewThm).split("_")[0])==-1:
                                    thematicLine = thematicLine + "|" + ''.join(splittedNewThm)
                                    EI.setDataFromCell(self.testSheet, (row, col), thematicLine)
                                    themImpactComment = themImpactComment + splittedNewThm
                                    specificLine = EI.getDataFromCell(self.testSheet, (row, col - 2))
                                    specificLine = specificLine.strip()
                                    try:
                                        specificLine = specificLine.split()[1]
                                    except:
                                        specificLine = "1"
                                    logging.info("specificLine = ", specificLine)
                                    # self.comment = self.comment + "Thematiques " + str(
                                    #     ', '.join(splittedNewThm)) + " added at thematique line " + str(
                                    #     specificLine) + ". "

                                    self.comment = self.comment+ ". "
                                    logging.info("checkThematicsInPT parametres(2) - ", funcImpactComment, themImpactComment,
                                          self.comment)
                                    funcImpactComment = 1
                                else:
                                    logging.info("themss33")
                                    if row==(x - 2):
                                        # commented on 24/04/2023 start

                                        # TPM.addThematique(macro)
                                        # time.sleep(5)
                                        # EI.setDataFromCell(self.testSheet, (8, 3), newThem)

                                        # commented on 24/04/2023 end

                                        logging.info("themss44")
                                        themImpactComment.append(newThem)
                                        funcImpactComment = 1
                            # logging.info("Add new line", newThem)
                            # TPM.addThematique()
                            # time.sleep(5)
                            # EI.setDataFromCell(self.testSheet, (8, 3), newThem)
                            # for i in newThem.split("|"):
                            #     themImpactComment.append(i)
            else:
                splittedNewThm.sort()
                lengthAdded = []
                lengthDeleted = []
                addedList = []
                deletedList = []
                for thematic in thematics:
                    thematicIndex = thematics.index(thematic)
                    splittedThem = thematic.split("|")
                    splittedThem.sort()
                    logging.info("sorted thematics = ", ''.join(splittedNewThm), ''.join(splittedThem))
                    if (''.join(splittedNewThm)) in (''.join(splittedThem)):
                        logging.info("Thematic already present(2)")
                        flag = 1
                        break
                    else:
                        logging.info("Thematic in testplan = ", splittedThem)
                        logging.info("Thematic in input = ", splittedNewThm)
                        thmDiff = difflib.ndiff(splittedThem, splittedNewThm)
                        addedThematics = []
                        deletedThematics = []
                        for i in thmDiff:
                            if i.startswith('-'):
                                deleted = i.split()[1]
                                deletedThematics.append(deleted)
                            if i.startswith('+'):
                                addedThematics.append(i.split()[1])
                            if i.startswith(' '):
                                pass
                        if (len(addedThematics)!=0):
                            for deleted in deletedThematics:
                                for added in addedThematics:
                                    logging.info("--------1", added, deleted)
                                    if added.find(deleted.split("_")[0])!=-1:
                                        if added!=deleted:
                                            logging.info("--------2", added, deleted)
                                            addedThematics = []
                                            deletedThematics = []
                                            logging.info(
                                                "Contrary thematic found so cannot add in this line so emptying the list")
                                    elif len(deletedThematics)==len(splittedThem):
                                        addedThematics = []
                                        deletedThematics = []
                                        logging.info(
                                            "Full thematique line getting deleted so cannot add in this line so emptying the list")
                            lengthAdded.append(len(addedThematics))
                            lengthDeleted.append(len(deletedThematics))
                            addedList.append(addedThematics)
                            deletedList.append(deletedThematics)
                            logging.info("Added Thematics = ", addedThematics)
                            logging.info("Deleted Thematics = ", deletedThematics)
                        else:
                            logging.info("Thematic already present(3)")
                            flag = 1
                if flag==0:
                    logging.info("lengthAdded & addedList = ", lengthAdded, addedList)
                    logging.info("lengthDeleted & deletedList = ", lengthDeleted, deletedList)
                    val = 0
                    copyOfLengthAdded = lengthAdded.copy()
                    try:
                        while True:
                            copyOfLengthAdded.remove(val)
                    except ValueError:
                        pass
                    logging.info("copyOfLengthAdded  = ", copyOfLengthAdded, len(copyOfLengthAdded))
                    if len(copyOfLengthAdded)==0:
                        logging.info("Add new line12", newThem)
                        funcImpactComment = 1
                        # commented on 24/04/2023 start

                        # TPM.addThematique(macro)
                        # time.sleep(5)
                        # EI.setDataFromCell(self.testSheet, (8, 3), newThem)

                        # commented on 24/04/2023 end

                        for i in newThem.split("|"):
                            themImpactComment.append(i)
                        logging.info("checkThematicsInPT parametres(3) - ", funcImpactComment, themImpactComment, self.comment)
                    else:
                        count = 0
                        smallest = max(lengthAdded)
                        for i in lengthAdded:
                            if i==0:
                                pass
                            else:
                                if i <= smallest:
                                    smallest = i
                                    position = count
                            count = count + 1
                        logging.info("smallest, position = ", smallest, position)
                        logging.info("thematics[position] = ", thematics[position], "addedList[position] = ",
                              addedList[position])
                        lineToAddThematique = EI.searchDataInSpecificRows(self.testSheet, (6, 100), 3,
                                                                          thematics[position])
                        x, y = lineToAddThematique["cellPositions"][0]
                        specificLine = EI.getDataFromCell(self.testSheet, (x, 1))
                        logging.info('line - ', lineToAddThematique, (lineToAddThematique["cellPositions"][0]), specificLine)
                        specificLine = specificLine.strip()
                        try:
                            specificLine = specificLine.split()[1]
                        except:
                            specificLine = "1"
                        logging.info("specificLine = ", specificLine)
                        EI.setDataFromCell(self.testSheet, (lineToAddThematique["cellPositions"][0]),
                                           (str(thematics[position] + "|" + str('|'.join(addedList[position])))))
                        logging.info("Thematiques56 " + str(
                            ''.join(addedList[position])) + " added at thematique line " + specificLine)
                        funcImpactComment = 1
                        themImpactComment = themImpactComment + addedList[position]
                        # self.comment = self.comment + "Thematiques " + str(
                        #     ', '.join(addedList[position])) + " added at thematique line " + str(specificLine) + ". "
                        self.comment = self.comment+ ". "
                        logging.info("checkThematicsInPT parametres(4) - ", funcImpactComment, themImpactComment, self.comment)
        thmLineAfter = EI.getDataFromCell(self.testSheet, (8, 1))
        thmLineAfter = thmLineAfter.strip()
        try:
            thmLineAfter = thmLineAfter.split()[1]
        except:
            thmLineAfter = "1"
        logging.info("thmLineAfter = ", thmLineAfter, type(thmLineAfter))
        logging.info("thmLineBefore = ", thmLineBefore, type(thmLineBefore))
        logging.info(f"funcImpactComment {funcImpactComment}")
        if thmLineBefore!=thmLineAfter:
            thmLineBefore = str(int(thmLineBefore) + 1)
            if thmLineBefore==thmLineAfter:
                logging.info("Added thematique line " + str(thmLineAfter))
                funcImpactComment = 1
                # self.comment = self.comment + "Added thematique line " + str(thmLineAfter) + ". "
                self.comment = self.comment+". "
            else:
                logging.info("Added thematique line " + str(thmLineBefore) + " to line " + str(thmLineAfter))
                funcImpactComment = 1
                # self.comment = self.comment + "Added thematique line " + str(thmLineBefore) + " to line " + str(
                #     thmLineAfter) + ". "
                self.comment = self.comment + ". "
        logging.info("self.comment = ", self.comment, themImpactComment, funcImpactComment)
        return themImpactComment
        # themImpactComment = list(dict.fromkeys(themImpactComment))
        # logging.info("funcImpactComment_final = ", themImpactComment)

    def oldToNewThm(self, listOfContents):
        # keys=['=','AND','OR']
        brackets = ['(', ')', '[', ']', '{', '}']
        newThm = []
        found = 0
        last_index = -1  # to get index of duplicate elements#
        # referntialEC = EI.openReferentialEC()
        logging.info("referential EC opened in oldToNewThm", listOfContents)
        for w in listOfContents:
            if w == '=':
                if re.search("_", listOfContents[listOfContents.index(w) - 1]).group():
                    thm = listOfContents[listOfContents.index(w, last_index + 1) - 1]
                    val = listOfContents[listOfContents.index(w, last_index + 1) + 1]
                    logging.info("thm & val before = ", thm, val)
                    if last_index < (len(listOfContents) - 1):
                        last_index = listOfContents.index(w, last_index + 1)
                    else:
                        last_index = -1
                    thmCopy = thm
                    valCopy = val
                    for c in thm:
                        if c in brackets:
                            found = 1
                            bIndex = thm.index(c)
                            b = c
                            thmCopy = thmCopy.replace(c, '')
                    for c in val:
                        if c in brackets:
                            found = 1
                            bIndex = val.index(c)
                            b = c
                            valCopy = valCopy.replace(c, '')
                    logging.info("thm & val after = ", thmCopy, valCopy)
                    newThmName = EI.getNewThematics(thmCopy, valCopy, self.refEC)
                    logging.info("newThmName = ", newThmName)
                    if newThmName!=-1:
                        if found==1:
                            found = 0
                            if bIndex==0:
                                newThmName = b + newThmName
                            else:
                                newThmName = newThmName + b
                        newThm.append(newThmName)
                    else:
                        logging.info("returning -1 from oldToNewThm")
                        # referntialEC.close()
                        logging.info("referential EC closed in oldToNewThm and returning -1")
                        return -1
            if w=='AND':
                newThm.append(w)
            if w=='OR':
                newThm.append(w)
        logging.info("returning newThm from oldToNewThm")
        # referntialEC.close()
        logging.info("referential EC closed in oldToNewThm", newThm)
        if len(newThm)!=0:
            return newThm
        else:
            return -1

    def getRawThm_old(self, Doc, reqName, reqVer):
        TableList = WDI.getTables(Doc)
        # RqTable=threading_findTable(TableList, ReqName+"("+ReqVer+")")
        rawThm = ""
        rqTable = WDI.threading_findTable(TableList, reqName)
        if rqTable==-1:
            if (reqName.find('.')!=-1):
                reqName = reqName.replace('.', '-')
            if (reqName.find('_')!=-1):
                reqName = reqName.replace('_', '-')

            rqTable = WDI.threading_findTable(TableList, reqName + "(" + reqVer + ")")
        else:
            rqTable = WDI.threading_findTable(TableList, reqName + "(" + reqVer + ")")
        if rqTable!=-1:
            chkOldFormat = WDI.checkFormat(rqTable, reqName + "(" + reqVer + ")")
            logging.info("check1", chkOldFormat)
            if (chkOldFormat==0):
                oldVerThm = WDI.getOldThematics(rqTable, reqName + "(" + reqVer + ")")
                logging.info("old ver ", oldVerThm)
                for i in oldVerThm:
                    for j in i:
                        if '=' in j:
                            j = j.replace(j, " = ")
                            oldVerThm = oldVerThm.replace(i, j)
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nOld raw Thematics " + str(oldVerThm))
                try:
                    rawThm = rawThm.join(self.oldToNewThm(oldVerThm.split()))
                except:
                    rawThm = -1
            elif (chkOldFormat==2):
                rawThm = WDI.getThematicsGateway(rqTable, reqName + "(" + reqVer + ")")
            else:
                rawThm = WDI.getThematics(rqTable, reqName + "(" + reqVer + ")")
        else:
            rqTable = WDI.threading_findTable(TableList, reqName + " " + reqVer)
            if rqTable!=-1:
                chkOldFormat = WDI.checkFormat(rqTable, reqName + " " + reqVer)
                logging.info("check2", chkOldFormat)
                if chkOldFormat==0:
                    oldVerThm = WDI.getOldThematics(rqTable, reqName + " " + reqVer)
                    for i in oldVerThm:
                        for j in i:
                            if '=' in j:
                                j = j.replace(j, " = ")
                                oldVerThm = oldVerThm.replace(i, j)
                    try:
                        rawThm = rawThm.join(self.oldToNewThm(oldVerThm.split()))
                    except:
                        rawThm = -1
                elif (chkOldFormat==2):
                    rawThm = WDI.getThematicsGateway(rqTable, reqName + " " + reqVer)
                else:
                    rawThm = WDI.getThematics(rqTable, reqName + " " + reqVer)
            else:
                rqTable = WDI.threading_findTable(TableList, reqName + "  " + reqVer)
                if rqTable!=-1:
                    chkOldFormat = WDI.checkFormat(rqTable, reqName + "  " + reqVer)
                    logging.info("check3", chkOldFormat)

                    if chkOldFormat==0:
                        oldVerThm = WDI.getOldThematics(rqTable, reqName + "  " + reqVer)
                        for i in oldVerThm:
                            for j in i:
                                if '=' in j:
                                    j = j.replace(j, " = ")
                                    oldVerThm = oldVerThm.replace(i, j)
                        try:
                            rawThm = rawThm.join(self.oldToNewThm(oldVerThm.split()))
                        except:
                            rawThm = -1
                    elif (chkOldFormat==2):
                        rawThm = WDI.getThematicsGateway(rqTable, reqName + "  " + reqVer)
                    else:
                        rawThm = WDI.getThematics(rqTable, self.reqName + "  " + self.reqVer)
                else:
                    rqTable = WDI.threading_findTable(TableList, reqName + " (" + reqVer + ")")
                    if rqTable!=-1:
                        chkOldFormat = WDI.checkFormat(rqTable, reqName + " (" + reqVer + ")")
                        logging.info("check3", chkOldFormat)

                        if chkOldFormat==0:
                            oldVerThm = WDI.getOldThematics(rqTable, reqName + " (" + reqVer + ")")
                            for i in oldVerThm:
                                for j in i:
                                    if '=' in j:
                                        j = j.replace(j, " = ")
                                        oldVerThm = oldVerThm.replace(i, j)
                            try:
                                rawThm = rawThm.join(self.oldToNewThm(oldVerThm.split()))
                            except:
                                rawThm = -1
                        elif (chkOldFormat==2):
                            rawThm = WDI.getThematicsGateway(rqTable, reqName + " (" + reqVer + ")")
                        else:
                            rawThm = WDI.getThematics(rqTable, self.reqName + " (" + self.reqVer + ")")
                    else:
                        rawThm = -1
        return rawThm

    def compare_thematiques(self, thematique_version_a, thematique_version_b):
        present_thematiques, modified_thematiques, added_thematiques, deleted_thematiques = set(), set(), set(), set()
        thematique_codes_a = set(
            thematiqueCode for thematiqueLine in thematique_version_a for thematiqueCode in thematiqueLine.split('|'))
        thematique_codes_b = set(
            thematiqueCode for thematiqueLine in thematique_version_b for thematiqueCode in thematiqueLine.split('|'))

        for a_thematique_code in thematique_codes_a:
            if a_thematique_code in thematique_codes_b:
                present_thematiques.add(a_thematique_code)
            else:
                for b_thematique_code in thematique_codes_b:
                    if b_thematique_code.startswith(
                            a_thematique_code[:3]) and b_thematique_code not in present_thematiques:
                        modified_thematiques.add(b_thematique_code)

        temp_set = modified_thematiques | present_thematiques
        added_thematiques = thematique_codes_b.difference(temp_set)
        possible_deleted_thematiques = thematique_codes_a.difference(temp_set)

        deleted_thematiques = {thematique_code for thematique_code in possible_deleted_thematiques
                               if not any(code.startswith(thematique_code[:3]) for code in temp_set)}

        result = {
            'present': present_thematiques,
            'added': added_thematiques,
            'modified': modified_thematiques,
            'deleted': deleted_thematiques
        }

        return result

    def getThem(self, req_data):
        logging.info(f"Finding thematic from doc data...... {req_data}")
        reqThm = -1
        if req_data != -1 and req_data != -2:
            if 'LCDV' in req_data:
                if req_data['LCDV'] != "":
                    return req_data['LCDV']
            if 'effectivity' in req_data:
                if req_data['effectivity'] != "":
                    return req_data['effectivity']
            if 'diversity' in req_data:
                if req_data['diversity'] != "":
                    return req_data['diversity']
            if 'target' in req_data:
                if req_data['target'] != "":
                    return req_data['target']

        return reqThm

    def getRawThm(self, Doc, reqName, reqVer):
        logging.info(f"Doc {Doc}")
        logging.info(f"\nFinding the thematic New..... {reqName, reqVer}")
        reqThm = ''
        reqData = DS.find_requirement_content(Doc, reqName + "(" + reqVer + ")")
        logging.info(f"res for {reqName}({reqVer}) ==> {reqData}")
        if reqData == -1 or not reqData:
            reqData = DS.find_requirement_content(Doc, reqName + " (" + reqVer + ")")
            logging.info(f"res for {reqName} ({reqVer}) ==> {reqData}")
        if reqData == -1 or not reqData:
            reqData = DS.find_requirement_content(Doc, reqName + " " + reqVer)
            logging.info(f"res for {reqName} {reqVer} ==> {reqData}")
        if reqData == -1 or not reqData:
            reqData = DS.find_requirement_content(Doc, reqName + "  " + reqVer)
            logging.info(f"res for {reqName}  {reqVer} ==> {reqData}")

        reqThm = self.getThem(reqData)
        return reqThm


    def compareThematics(self):
        logging.info("In compareThematics function")
        self.testReqName = self.testReqName.strip()
        self.testReqVer = self.testReqVer.strip()
        # oldReq = oldReq.strip()
        self.newReq = self.newReq.strip()
        logging.info("compareThematics parametres = ", self.testReqName, self.testReqVer, "oldReq =*" + self.reqName + "*",
              "newReq =*" + self.newReq + "*")
        if len(self.newReq)==0:
            self.reqName = self.reqName.strip()
            self.reqVer = self.reqVer.strip()
        else:
            if self.newReq.find("(")!=-1:
                self.reqName = self.newReq.split("(")[0]
                self.reqVer = self.newReq.split("(")[1].split(")")[0]
            else:
                self.reqName = self.newReq.split(" ")[0]
                self.reqVer = self.newReq.split(" ")[1]
                self.reqName = self.reqName.strip()
                self.reqVer = self.reqVer.strip()
        logging.info("In compareThematics function", self.testReqName, self.testReqVer, "*" + self.reqName + "*",
              "*" + self.reqVer + "*")
        oldRawThm = ""
        newRawThm = ""
        oldThm = ""
        newThm = ""
        oldRawThm = self.getRawThm(self.prevDoc, self.testReqName, self.testReqVer)
        logging.info("old Raw thematics = ", type(oldRawThm), oldRawThm)
        newRawThm = self.getRawThm(self.currDoc, self.reqName, self.reqVer)
        logging.info("New Raw thematics = ", type(newRawThm), newRawThm, )
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nNew Raw thematics = " + str(newRawThm))
        # logging.info("old Raw thematics ", oldRawThm)
        # logging.info("New Raw thematics ", newRawThm)

        if ((type(oldRawThm)==str) and (type(newRawThm)==str)):
            logging.info(f"old thematics#############")
            if len(oldRawThm)!=0:
                oldThm = self.grepThematicsCode(oldRawThm)
                logging.info("old thematics ", oldThm)
            if len(newRawThm)!=0:
                newThm = self.grepThematicsCode(newRawThm)
                logging.info("New thematics ", type(newThm))

            if oldThm.split()==newThm.split():
                logging.info("No functional Imapct")
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nNo functional Impact in Thematics")
                return -1
            else:
                logging.info("Functional Imapct In thematics")
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nFunctional Imapct In thematics")
                return newThm
        elif (oldRawThm==-1) and (newRawThm!=-1):
            logging.info("Functional Impact as Old Thematique not present in SSD")
            newThm = self.grepThematicsCode(newRawThm)
            return newThm
        elif ((oldRawThm!=-1) and (newRawThm==-1)):  # check from where -2 is cmng
            logging.info("Issue in referentiel EC")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines(
                    "\n\nfor Requirement " + self.reqName + " " + self.reqVer + " thematique not found in referential EC. Please check the files manually")
            time.sleep(2)
            return -3
        else:
            logging.info("Unable to get both current and previous thematiques from SSD or referentiel EC ")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines("\n\nFor requirement " + str(
                    self.reqName + " " + self.reqVer) + " unable to get both current and previous thematiques from SSD or referentiel EC. Please check the files manually")
            time.sleep(2)
            return -2

    def removeChars(self):
        logging.info("removing unwanted characters...")
        final_raw_comb = self.funcImpact
        if self.funcImpact.find("( )") != -1:
            final_raw_comb = self.funcImpact.replace("( )", "")
        if final_raw_comb != '':
            if final_raw_comb.strip().endswith("OR'") or final_raw_comb.strip().endswith("AND'"):
                final_raw_comb = remove_trailing_and_or(final_raw_comb)

        return final_raw_comb

    def Analyse(self):
        macro = EI.getTestPlanAutomationMacro()
        global themImpactComment
        themImpactComment = []
        global funcImpactComment
        funcImpactComment = 0
        logging.info("themImpactComment(1)", themImpactComment)
        themFlagDiffArch = 0
        thmConflictFlag = 0
        self.getTestSheetRequirement()
        logging.info("Test sheet Req Name ", self.testReqName, self.testReqVer)
        # compare thm
        # Check functional impact
        # create attribute which will be used in main()
        self.funcImpact = self.compareThematics()
        logging.info("self.funcImpact BF = ", self.funcImpact)
        if type(self.funcImpact)==str:
            self.funcImpact = self.removeChars()
            logging.info("self.funcImpact AF= ", self.funcImpact)
            rawCombinations = self.createThmCombinations(self.funcImpact)
            combList = rawCombinations.split('\n')
            combList = list(set(combList))
            logging.info("combList ", combList)
            for self.testSheet in self.listOfSheets:
                logging.info("-------self.testSheet in Analyse fucntion of AnaLyseThematiques module - ", self.testSheet)
                funcImpactComment = 0
                self.comment = ""
                KMS.showWindow((self.tpBook.name).split(".")[0])
                time.sleep(2)
                EI.activateSheet(self.tpBook,  self.testSheet)
                time.sleep(1)
                if self.testSheet.range('C7').value=='VALIDEE':
                    logging.info('Executing the selectTestSheetModify')
                    TPM.selectTestSheetModify(macro)
                applicableThemLine = ""
                for newComb in combList:
                    logging.info("new line newComb", newComb)
                    applicableThemLine = self.filterThemForArch(newComb)
                    them = self.getThematicsFromPT()
                    prevThem = them.copy()
                    if applicableThemLine!=-1:
                        if len(applicableThemLine)!=len(newComb):
                            logging.info("Thematic not applicable to architecture")
                            themFlagDiffArch = 1
                        if len(applicableThemLine)!=0:
                            themImpactComment = self.checkThematicsInPT(them, applicableThemLine, themImpactComment)
                    elif applicableThemLine==-1:
                        logging.info("Conflict in Thematic")
                        thmConflictFlag = 1
                logging.info("Functional Impact", themImpactComment)
                if self.newReq!="":
                    logging.info("In AnaLyseThematics self.newReq != """)
                    if len(self.comment)==0:
                        logging.info("In AnaLyseThematics len(self.comment) == 0:", self.testReqName, str(self.testReqVer),
                              str(self.newReq))
                        EI.fillSheetHistory(self.testSheet,
                                            "Evolved requirement. Changed requirement name from " + self.testReqName + "(" + str(
                                                self.testReqVer) + ")" + " to " + str(self.newReq) + ". ")
                    else:
                        logging.info("In AnaLyseThematics in ELSE(1)", self.testReqName, str(self.testReqVer),
                              str(self.newReq), str(self.comment))
                        EI.fillSheetHistory(self.testSheet,
                                            "Evolved requirement. Changed requirement name from " + self.testReqName + "(" + str(
                                                self.testReqVer) + ")" + " to " + str(self.newReq) + ". " + str(
                                                self.comment))
                else:
                    logging.info("In AnaLyseThematics self.newReq != "" ELSE part")
                    if len(self.comment)==0:
                        logging.info("In AnaLyseThematics len(self.comment) == 0:", self.reqName, str(self.testReqVer),
                              str(self.reqVer))
                        EI.fillSheetHistory(self.testSheet,
                                            "Evolved requirement. Incremented version of requirement " + self.reqName + " from " + str(
                                                self.testReqVer) + " to " + str(self.reqVer) + ". ")
                    else:
                        logging.info("In AnaLyseThematics in ELSE(1)", self.reqName, str(self.testReqVer), str(self.reqVer),
                              str(self.comment))
                        EI.fillSheetHistory(self.testSheet,
                                            "Evolved requirement. Incremented version of requirement " + self.reqName + " from " + str(
                                                self.testReqVer) + " to " + str(self.reqVer) + ". " + str(self.comment))
                # logging.info("Modified Text ", modContent)
                if themFlagDiffArch==1:
                    with open('../Aptest_Tool_Report.txt', 'a') as f:
                        f.writelines(
                            "\n\nFor " + self.reqName + "(" + str(
                                self.reqVer) + ") some thematics are not applicable to architecture. Please check manually to raise any QIA if needed.")
                    time.sleep(2)
                if thmConflictFlag==1:
                    with open('../Aptest_Tool_Report.txt', 'a') as f:
                        f.writelines(
                            "\n\nFor " + self.reqName + "(" + str(
                                self.reqVer) + ") Unable to analyse Thematiques because there is conflict in thematics. Please proceed manually")
                    time.sleep(2)
                modThem = self.getThematicsFromPT()
                addedThms = []
                for mod in modThem:
                    thms = mod.split('|')
                    for old in prevThem:
                        oldThms = old.split('|')
                        for t in thms:
                            if t not in oldThms:
                                if t not in addedThms:
                                    addedThms.append((t, modThem.index(mod)))
            logging.info("themImpactComment(2)", themImpactComment)
        elif self.funcImpact in (-1, -2):
            if self.funcImpact==-2:
                logging.info("Thematiques not found in both the documents")
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines(
                        "\n\nThematiques not found in both the documents for " + self.testReqName + ". Please proceed manually")
                time.sleep(2)
            logging.info("No functional Impact")
            for self.testSheet in self.listOfSheets:
                KMS.showWindow((self.tpBook.name).split(".")[0])
                time.sleep(2)
                EI.activateSheet(self.tpBook, self.testSheet)
                time.sleep(1)
                if self.testSheet.range('C7').value=='VALIDEE':
                    TPM.selectTestSheetModify(macro)
                if self.newReq!="":
                    logging.info("In AnaLyseThematics self.newReq != "" No fucntional impact", self.testReqName,
                          str(self.testReqVer), str(self.newReq))
                    EI.fillSheetHistory(self.testSheet,
                                        "Evolved requirement. Changed requirement name from " + self.testReqName + "(" + str(
                                            self.testReqVer) + ")" + " to " + str(self.newReq) + ". ")
                else:
                    logging.info("In AnaLyseThematics self.newReq != "" No fucntional impact ELSE part", self.reqName,
                          str(self.testReqVer), str(self.reqVer))
                    EI.fillSheetHistory(self.testSheet,
                                        "Evolved requirement. Incremented version of requirement " + self.reqName + " from " + str(
                                            self.testReqVer) + " to " + str(self.reqVer) + ". ")
        # elif self.funcImpact==-2:
        #     logging.info("Thematiques not found in both the documents")
        else:
            logging.info("Thematiques not found in the documents or in refrentiel EC document")

# tpBook = EI.openExcel(r"C:/Users/10388/Downloads/BSI AUtomation/Input/Tests_20_64_01272_19_00870_FSEE_TURNKEY_CPK4_V5_VSM.xlsm")
# data="IF P_INFO_ACPK_MENU=MANEUVER AND P_INFO_ACPK_MANEUVER_STATE=IN_PROGRESS THEN P_SON_ACPK_ACTIVATION=DEMANDE ELSE P_SON_ACPK_ACTIVATION=PAS_DE_DEMANDE"
# testData="IF P_INFO_ACPK_MENU=10 AND P_INFO_ACPK_MANEUVER_STATE=1 THEN P_SON_ACPK_ACTIVATION=DEMANDE ELSE P_SON_ACPK_ACTIVATION=PAS_DE_DEMANDE"
# thm='AJO_01 AND (LYQ_01 AND (AJA_02 , AJA_03) AND ABC_00 OR LYQ_02 AND (DUB_23 , DUB_24))'
# thm1='AJO_01 AND (DEF_01 AND (LYQ_01 , LYQ_02))'
# testSheet = 'VSM20_GC_20_64_0025'
# thematic = r"HP-0000873 [ DICO MULTIGAMME Q1_2021 - ∞ ] AND AJO_01 WITH AND {LYQ_01 BEFORE_FUNCT_CODIF AND AJA(AJA_02 DOP_AV_ET_AR,AJA_03 DOP_AV_AR_LAT) AND ABC_00 OR LYQ_02 FUNCT_CODIF AND DUB(DUB_23 FRONT+REAR,DUB_24 FRONT+REAR+TRAJECTORY+LATERAL)}"
# analyseThematics(thematic,tpBook,testSheet)
# thm = r"HP-0000873 [ DICO MULTIGAMME Q2_2021 - ∞ ] AND AUU OPTION_DYNAMIC_EPS_ALERT(AUU_00 WITHOUT) AND INL TYPE_DA(INL_02 DAV) OR INL TYPE_DA(INL_01 DAEH,INL_03 DAE)"
# grepThematicsCode(thm)
# refBook = EI.openExcel("C:\\Users\\9346\\OneDrive - Expleo France\\Desktop\\Team Code\\TeamTesting Documents\\VSM_General_Documents\\Referentiel_EC.xlsm")
# ARCH = "BSI"
# thematicLine = "IDK_01|CHB_01|IOC_02|LSP_00|AZX_00"
# filterThemForArch(refBook, ARCH, thematicLine)

#
# refEC=EI.openReferentialEC()
# tpBook=EI.openTestPlan()
# EI.openExcel(ICF.getTestPlanMacro())
# time.sleep(1)
# prevDoc = input("Enter prev Doc path ")
# logging.info("\n")
# currDoc = input("Enter curr Doc path ")
# logging.info("\n")
# lst = []
#
# # number of elements as input
# n = int(input("Enter number of testSheets : "))
# logging.info("\n")
#
# # iterating till the range
# for i in range(0, n):
#     logging.info("\n")
#     ele = str(input("Enter testsheet name :"))
#     sheet = tpBook.sheets[ele]
#     lst.append(sheet)  # adding the element
# newReq=""
# logging.info(lst)
# reqName = input("Enter old requirement name to be checked in current doc ")
# logging.info("\n")
# reqVer = input("Enter old requirement version to be checked in current doc ")
# logging.info("\n")
# newReq=input("Enter new requirement version to be checked in current doc if any else press enter ")
# time.sleep(2)
# KMS.showWindow(tpBook.name.split('.')[0])
# time.sleep(1)
# TPM.selectArch()
# time.sleep(1)
# TPM.selectTpWritterProfile()
# time.sleep(1)
# TPM.selectToolbox()
# time.sleep(1)
# testSheetName='VSM20_GC_20_64_0025'
# testSheet=tpBook.sheets[testSheetName]
# Arch='VSM'


# oldRawThm=r"HP-0000873 [ DICO MULTIGAMME Q1_2021 - ∞ ] AND AJO_01 WITH AND {LYQ_01 BEFORE_FUNCT_CODIF AND AJA(AJA_02 DOP_AV_ET_AR,AJA_03 DOP_AV_AR_LAT) AND ABC_00 OR LYQ_02 FUNCT_CODIF AND DUB(DUB_23 FRONT+REAR,DUB_24 FRONT+REAR+TRAJECTORY+LATERAL)}"
# newRawThm=r"HP-0000873 [ DICO MULTIGAMME Q1_2021 - ∞ ] AND AJO_01 WITH AND {LYQ_01 BEFORE_FUNCT_CODIF AND AJA(AJA_02 DOP_AV_ET_AR,AJA_03 DOP_AV_AR_LAT) AND ABC_00 OR LYQ_02 FUNCT_CODIF AND DUB(DUB_23 FRONT+REAR,DUB_24 FRONT+REAR+TRAJECTORY+LATERAL)}"
# data=AnalyseThematics(tpBook, refEC, currDoc, prevDoc, lst, reqName, reqVer, Arch,newReq)
# data.Analyse()
# def getTestSheetRequirement():
