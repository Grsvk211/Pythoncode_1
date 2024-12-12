import time
import xlwings as xw
import ExcelInterface as EI
import ParseContents as parser
import TestPlanMacros as TPM
import KeyboardMouseSimulator as KMS
import InputConfigParser as ICF
import AnaLyseThematics as AT
import WordDocInterface as WDI
import json
import difflib
import logging
import re
# import numpy as np
import DocumentSearch as DS
import logging


class AnalyseTestSheet:
    def __init__(self, dciBook, tpBook, alertDoc, ssFiches, currDoc, prevDoc, listOfTestSheets, reqName, reqVer,
                 newReq=""):
        self.dciBook = dciBook
        self.tpBook = tpBook
        self.alertDoc = alertDoc
        self.ssFiches = ssFiches
        self.currDoc = currDoc
        self.prevDoc = prevDoc
        self.reqName = reqName
        self.reqVer = reqVer
        self.testReqName = ""
        self.testReqVer = ""
        self.oldContent = ""
        self.newContent = ""
        self.funcImpact = ""
        self.newReq = newReq
        self.listOfSheets = listOfTestSheets
        self.testSheet = listOfTestSheets[0]
        self.ssFlag = 0
        self.ssSteps = []
        self.ssSignals = []
        self.operators = [r"=", r"<", r">", r"+", r"-", r">=", r"<=", r"(", r")", r"[", r"]", r"{", r"}"]
        self.keys = [r"IF", r"THEN", r"ELSE", r"AND", r"OR", r"=", r"<", r">", r"+", r"-", r">=", r"<=", r"(", r")",
                     r"[", r"]", r"{", r"}", r"FOR", r"NOT"]
        logging.basicConfig(filename="analysePT.log",
                            format='%(asctime)s %(message)s',
                            filemode='w')
        self.logger = logging.getLogger()

    def joinWords(self, data):  # still more testing required
        # use this only after removing noise and replacing grammer with operators
        # Adding '_' between words (if any)
        global keys  # make keys as class attribute
        copyofData = data.split()
        logging.info("copy ", copyofData, len(copyofData))
        for n, d in enumerate(copyofData):
            if d in self.operators:
                i = n + 1
                if i < len(copyofData) - 1:
                    while copyofData[i + 1] not in self.keys:
                        try:
                            wrd = copyofData[i] + "_" + copyofData[i + 1]
                            data = data.replace(copyofData[i] + " " + copyofData[i + 1], wrd)
                            # data=data.replace(copyofData[i+1],'')
                        except:
                            break
                        if i < len(copyofData) - 2:
                            i = i + 1
                        else:
                            break
                        # i=i+1
        return data

    def loadKeywords(self):  # tested ok
        f_keys = open('../user_input/Keywords.json', "r")
        # f_keys = open(r'C:\Users\vgajula\Downloads\New folder (2)\FSEE_ACTIVE_BRAKE_14_10_2022\user_input\Keywords.json', "r")
        keywords = json.load(f_keys)
        return keywords

    def convertData(self, keywords, data):  # tested
        for key in keywords["keywords"]:
            for text in keywords["keywords"][key]:
                if re.findall(text.lower(), data.lower()):
                    data = data.replace(text.upper(), key.upper() + " ")
        return data

    def AddSpaces(self, data):  # tested ok
        for d in data:
            for i in d:
                # logging.info(i)
                if i in self.operators:
                    # logging.info("yes")
                    i = i.replace(i, " " + i + " ")
            data = data.replace(d, i)
        return data

    def txtPreProcessing(self, data):
        keywords = self.loadKeywords()
        data = self.AddSpaces(data)
        data = self.convertData(keywords, data)
        data = self.joinWords(data)
        return data

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
        logging.info("test", self.testReqName, self.testReqVer)

    def getContent(self, Doc, ReqName, ReqVer):

        logging.info("in get content ", ReqName, ReqVer)
        TableList = WDI.getTables(Doc)
        # RqTable=threading_findTable(TableList, ReqName+"("+ReqVer+")")
        RqTable = WDI.threading_findTable(TableList, ReqName)
        logging.info("table ", RqTable)
        if RqTable == -1:
            if ((ReqName.find('.') != -1) | (ReqName.find('_') != -1)):
                ReqName = ReqName.replace('.', '-')
            RqTable = WDI.threading_findTable(TableList, ReqName + "(" + ReqVer + ")")
        else:
            RqTable = WDI.threading_findTable(TableList, ReqName + "(" + ReqVer + ")")
            logging.info("table2 ", RqTable)
        if RqTable != -1:
            chkOldFormat = WDI.checkFormat(RqTable, ReqName + "(" + ReqVer + ")")  # Identify gateway req here
            if chkOldFormat == 0:
                Content = WDI.getOldContents(RqTable, ReqName + "(" + ReqVer + ")")
            elif chkOldFormat == 2:
                Content = WDI.getGatewayContent(RqTable, ReqName + "(" + ReqVer + ")")
            else:
                Content = WDI.getNewContents(RqTable, ReqName + "(" + ReqVer + ")")
        else:
            RqTable = WDI.threading_findTable(TableList, ReqName + " " + ReqVer)
            if RqTable != -1:
                chkOldFormat = WDI.checkFormat(RqTable, ReqName + " " + ReqVer)
                if chkOldFormat == 0:
                    Content = WDI.getOldContents(RqTable, ReqName + " " + ReqVer)
                elif chkOldFormat == 2:
                    Content = WDI.getGatewayContent(RqTable, ReqName + " " + ReqVer)
                else:
                    Content = WDI.getNewContents(RqTable, ReqName + " " + ReqVer)
            else:
                RqTable = WDI.threading_findTable(TableList, ReqName + "  " + ReqVer)
                if RqTable != -1:
                    chkOldFormat = WDI.checkFormat(RqTable, ReqName + "  " + ReqVer)
                    if chkOldFormat == 0:
                        Content = WDI.getOldContents(RqTable, ReqName + "  " + ReqVer)
                    elif chkOldFormat == 2:
                        Content = WDI.getGatewayContent(RqTable, ReqName + " " + ReqVer)
                    else:
                        Content = WDI.getNewContents(RqTable, ReqName + "  " + ReqVer)
                else:
                    RqTable = WDI.threading_findTable(TableList, ReqName + " (" + ReqVer + ")")
                    if RqTable != -1:
                        chkOldFormat = WDI.checkFormat(RqTable, ReqName + " (" + ReqVer + ")")
                        if chkOldFormat == 0:
                            Content = WDI.getOldContents(RqTable, ReqName + " (" + ReqVer + ")")
                        elif chkOldFormat == 2:
                            Content = WDI.getGatewayContent(RqTable, ReqName + " (" + ReqVer + ")")
                        else:
                            Content = WDI.getNewContents(RqTable, ReqName + " (" + ReqVer + ")")
                    else:
                        Content = -1
        return Content

    def getRequirementContent(self, Doc, reqName, reqVer):
        reqData = DS.find_requirement_content(Doc, reqName + "(" + reqVer + ")")
        logging.info(f"AT reqData1 {reqData}")
        if reqData == -1 or not reqData:
            reqData = DS.find_requirement_content(Doc, reqName + " (" + reqVer + ")")
            logging.info(f"AT reqData2 {reqData}")
        if reqData == -1 or not reqData:
            reqData = DS.find_requirement_content(Doc, reqName + " " + reqVer)
            logging.info(f"AT reqData3 {reqData}")
        if reqData == -1 or not reqData:
            reqData = DS.find_requirement_content(Doc, reqName + "  " + reqVer)
            logging.info(f"AT reqData4 {reqData}")
        if reqData and reqData is not None:
            return reqData['content']
        else:
            -1

    def compareDocs(self):
        global newContentSplit
        newContentSplit = []

        global oldContentSplit
        oldContentSplit = []

        self.testReqName = self.testReqName.strip()
        self.testReqVer = self.testReqVer.strip()
        # oldReq = oldReq.strip()
        self.newReq = self.newReq.strip()
        logging.info("compareDocs parametres = ", self.testReqName, self.testReqVer, "oldReq =*" + self.reqName + "*",
              "newReq =*" + self.newReq + "*")
        if len(self.newReq) == 0:
            self.reqName = self.reqName
            self.reqVer = self.reqVer.strip()
        else:
            if self.newReq.find("(") != -1:
                self.reqName = self.newReq.split("(")[0]
                self.reqVer = self.newReq.split("(")[1].split(")")[0]
            else:
                self.reqName = self.newReq.split(" ")[0]
                self.reqVer = self.newReq.split(" ")[1]
                self.reqName = self.reqName
                self.reqVer = self.reqVer
        logging.info("In compareDocs function", self.testReqName, self.testReqVer, "*" + self.reqName + "*",
              "*" + self.reqVer + "*")
        # logging.info("In compareDocs function", testReqName, testReqVer, reqName, reqVer)
        # oldContent = self.getContent(self.prevDoc, self.testReqName, self.testReqVer)
        logging.info(f"self.prevDoc {self.prevDoc}")
        logging.info(f"self.currDoc {self.currDoc}")
        oldContent = self.getRequirementContent(self.prevDoc, self.testReqName, self.testReqVer)
        logging.info("Old content found", oldContent)
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nOld content found " + str(oldContent))
        # newContent = self.getContent(self.currDoc, self.reqName, self.reqVer)
        newContent = self.getRequirementContent(self.currDoc, self.reqName, self.reqVer)
        logging.info("New content found", newContent)
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\nNew content found " + str(newContent))
        # logging.info(oldRqTable())
        if ((type(oldContent) == str) and (type(newContent) == str)):

            logging.info("new Content ", newContent.split())
            logging.info("old Content ", oldContent.split())
            '''for i in oldContent:
                for j in i:
                    if '=' in j:
                        j=j.replace(j," = ")
                oldContent=oldContent.replace(i,j)
            for i in newContent:
                for j in i:
                    if '=' in j:
                        j=j.replace(j," = ")
                newContent=newContent.replace(i,j)'''
            for i in oldContent:
                for j in i:
                    # logging.info(i)
                    if j in self.operators:
                        # logging.info("yes")
                        j = j.replace(j, " " + j + " ")
                oldContent = oldContent.replace(i, j)
            for i in newContent:
                for j in i:
                    # logging.info(i)
                    if j in self.operators:
                        # logging.info("yes")
                        j = j.replace(j, " " + j + " ")
                newContent = newContent.replace(i, j)
            newContentSplit = newContent.split()
            oldContentSplit = oldContent.split()
            logging.info("new Content ", newContentSplit)
            logging.info("old Content ", oldContentSplit)

            for n, i in enumerate(newContentSplit):
                # logging.info("I = ", i)
                if i == "IF":
                    indexIF = n
                    logging.info(indexIF)
                    del newContentSplit[:indexIF]
                    break
            logging.info("new Content Modified = ", newContentSplit)
            for n, i in enumerate(oldContentSplit):
                if i == "IF":
                    indexIF = n
                    del oldContentSplit[:indexIF]
                    break
            logging.info("old Content Modified = ", oldContentSplit)
            newModContent = " ".join(newContentSplit)
            oldModContent = " ".join(oldContentSplit)
            newModContent = newModContent.upper()
            oldModContent = oldModContent.upper()

            '''contentDiff = difflib.ndiff(oldContentSplit, newContentSplit)
            diffDict = {"equal": [], "del": [], "add": []}
            
            #create dictionary
            for i in contentDiff:
                if i.startswith('-'):
                    #logging.info("Deleted Word",i.split()[1])
                    diffDict["del"].append(i.split()[1])
                if i.startswith('+'):
                    #logging.info("Added Word",i.split()[1])
                    diffDict["add"].append(i.split()[1])
                if i.startswith(' '):
                    #logging.info("No change",i.split()[0])
                    diffDict["equal"].append(i.split()[0])'''

            if oldContentSplit == newContentSplit:
                logging.info("No functional Imapct")
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nNo Functional Impact in Contents")
                return -1
            else:
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nFunctional Impact in Contents")
                # logging.info(diffDict)
                return newModContent
        elif ((type(oldContent) == dict) and (type(newContent) == dict)):
            logging.info("GateWay requirement ")
            if oldContent == newContent:
                return -1
            else:
                return newContent
        else:
            return -2

    def getRawSteps(self, content):
        logging.info("Generating raw Steps ")
        return parser.createSteps(content).split()

    def getNewStep(self):
        logging.info("Identifying new Steps ")
        maxrow = self.testSheet.range('A' + str(self.testSheet.cells.last_cell.row)).end('up').row
        self.testSheet_value = self.testSheet.used_range.value
        # searchCondInit = EI.searchDataInExcel(self.testSheet, (26, maxrow), "CONDITIONS INITIALES")
        searchCondInit = EI.searchDataInExcelCache(self.testSheet_value , (26, maxrow), "CONDITIONS INITIALES")
        # searchCorpsTest = EI.searchDataInExcel(self.testSheet, (26, maxrow), "CORPS DE TEST")
        searchCorpsTest = EI.searchDataInExcelCache(self.testSheet_value , (26, maxrow), "CORPS DE TEST")
        # searchRetour = EI.searchDataInExcel(self.testSheet, (26, maxrow), 'RETOUR AUX CONDITIONS INITIALES')
        searchRetour = EI.searchDataInExcelCache(self.testSheet_value , (26, maxrow), 'RETOUR AUX CONDITIONS INITIALES')
        testBodyStart, col = searchCorpsTest["cellPositions"][0]
        testBodyEnd, col = searchRetour["cellPositions"][0]
        logging.info(testBodyStart, testBodyEnd)

        steps = []
        i = testBodyStart
        z = testBodyEnd

        rowRange = 0
        while (i < z):
            if self.testSheet.range(i, 1).merge_cells is False:
                if self.testSheet.range(i, 1).value != 'ETAPE':
                    # logging.info("1",i)
                    # logging.info(testSheet.range(i,1).value)
                    j = i
                    # logging.info("11",j)
                    while (self.testSheet.range(j, 5).merge_cells is False):
                        logging.info(self.testSheet.range(j, 5).value, j)
                        if (self.testSheet.range(j, 5).value is None) and (self.testSheet.range(j, 11).value is None):
                            # logging.info(j)
                            rowRange = j
                        j = j + 1
                    # logging.info("after",j)
                    if rowRange != 0:
                        if rowRange not in steps:
                            steps.append(rowRange)
                    i = j + 1
            i = i + 1

            # logging.info("incr ",i)
        logging.info("Steps", steps)
        return steps

    def getRows(self):
        logging.info("identifying row value of steps in FT ")
        maxrow = self.testSheet.range('A' + str(self.testSheet.cells.last_cell.row)).end('up').row
        self.testSheet_value = self.testSheet.used_range.value
        # searchCondInit = EI.searchDataInExcel(self.testSheet, (26, maxrow), "CONDITIONS INITIALES")
        searchCondInit = EI.searchDataInExcelCache(self.testSheet_value, (26, maxrow), "CONDITIONS INITIALES")
        # searchCorpsTest = EI.searchDataInExcel(self.testSheet, (26, maxrow), "CORPS DE TEST")
        searchCorpsTest = EI.searchDataInExcelCache(self.testSheet_value, (26, maxrow), "CORPS DE TEST")
        # searchRetour = EI.searchDataInExcel(self.testSheet, (26, maxrow), 'RETOUR AUX CONDITIONS INITIALES')
        searchRetour = EI.searchDataInExcelCache(self.testSheet_value, (26, maxrow), 'RETOUR AUX CONDITIONS INITIALES')
        testBodyStart, col = searchCorpsTest["cellPositions"][0]
        testBodyEnd, col = searchRetour["cellPositions"][0]
        logging.info(testBodyStart, testBodyEnd)

        steps = []
        i = testBodyStart
        z = testBodyEnd

        while (i < z):
            if self.testSheet.range(i, 1).merge_cells is False:
                if self.testSheet.range(i, 1).value != 'ETAPE':
                    # logging.info("1",i)
                    # logging.info(testSheet.range(i,1).value)
                    j = i
                    rowRange = []
                    # logging.info("11",j)
                    while self.testSheet.range(j, 1).merge_cells is False:
                        # logging.info(j)
                        rowRange.append(j)
                        j = j + 1
                    # logging.info("after",j)
                    steps.append(rowRange)
                    i = j + 1
            i = i + 1

            # logging.info("incr ",i)
        logging.info("Steps", steps)
        return steps

    def list_duplicates_of(self, seq, item):
        start_at = -1
        locs = []
        while True:
            try:
                loc = seq.index(item, start_at + 1)
            except ValueError:
                break
            else:
                locs.append(loc)
                start_at = loc
                return locs

    def checkIpSignals(self, rowSteps, ip):
        logging.info("Checking input signals in FT ", ip)
        listOfIp = ip
        ipFound = []
        match = []
        result = []

        for step in rowSteps:
            found = 0
            for ipsignal in listOfIp:
                if ipsignal not in (-1, -2):
                    signal, valueName, value, tempo, typeStimuli = ipsignal
                    logging.info("Value from dci ", value)
                    for r in step:
                        if self.testSheet.range(r, 5).value is not None:
                            if signal in self.testSheet.range(r, 5).value:
                                logging.info("1st", self.testSheet.range(r, 2).value)
                                if (self.testSheet.range(r, 2).value == tempo):
                                    logging.info("2nd", self.testSheet.range(r, 3).value)
                                    if self.testSheet.range(r, 3).value is not None:
                                        if valueName in self.testSheet.range(r, 3).value:
                                            logging.info("3rd", self.testSheet.range(r, 4).value)
                                            if self.testSheet.range(r, 4).value == typeStimuli:
                                                logging.info("4th", self.testSheet.range(r, 6).value)
                                                if self.testSheet.range(r, 6).value is not None:
                                                    if type(self.testSheet.range(r, 6).value) == float:
                                                        logging.info("5th", )
                                                        if str(int(float(self.testSheet.range(r, 6).value))) == str(
                                                                value):
                                                            logging.info("hh")
                                                            found = found + 1
                                                            match.append(ipsignal)
                                                    else:
                                                        if str(self.testSheet.range(r, 6).value) == str(value):
                                                            logging.info("hh2")
                                                            found = found + 1
                                                            match.append(ipsignal)

            # if found!=0:
            ipFound.append(found)
        if ((len(ipFound) != 0) & (max(ipFound) > 0)):
            matchID = self.list_duplicates_of(ipFound, max(ipFound))[0]
            logging.info("mID", matchID)
            copyOfIp = []
            result = rowSteps[matchID]
            logging.info("Match found", ipFound, max(ipFound), len(listOfIp), match)
            if max(ipFound) < len(listOfIp):
                copyOfIp = listOfIp.copy()
                for ip in listOfIp:
                    if ip in match:
                        copyOfIp.remove(ip)
            return result, copyOfIp
        else:
            logging.info("Match not found")
            return -1, -1

    def checkOpSignals(self, step, listOfOp):
        logging.info("Checking Output Signals in FT ")
        opFound = 0
        match = []
        for r in step:
            for op in listOfOp:
                outSignal, outValueName, outValue, tempMin, tempMax, timePeriod = op
                logging.info("op", op)
                if (outSignal.find('SON') != -1):
                    if self.testSheet.range(r, 10).value is not None:
                        if "SON" in self.testSheet.range(r, 10).value:
                            logging.info("h")
                            if self.testSheet.range(r, 9).value is not None:
                                if ((outSignal in self.testSheet.range(r, 9).value) or (
                                        outValueName in self.testSheet.range(r, 9).value)):
                                    logging.info("hh")
                                    if self.testSheet.range(r, 11).value is not None:
                                        if "$DMD_EM_SON_NUM" in self.testSheet.range(r, 11).value:
                                            logging.info("hh")
                                            if self.testSheet.range(r, 7).value == tempMin and self.testSheet.range(r,
                                                                                                                    8).value == tempMax:
                                                if outValueName == 'DEMANDE':
                                                    logging.info("hd")
                                                    if "$" + outValue in self.testSheet.range(r, 12).value:
                                                        logging.info("op found")
                                                        opFound = 1
                                                        match.append(op)
                                                if ((outValueName == 'PAS_DE_DEMANDE') or (outValueName == 'PAS')):
                                                    logging.info("hpas")
                                                    if "<>$" + outValue in self.testSheet.range(r, 12).value:
                                                        logging.info("op found")
                                                        opFound = 1
                                                        match.append(op)
                elif (outSignal.find('TEM') != -1):
                    if self.testSheet.range(r, 10).value is not None:
                        if self.testSheet.range(r, 9).value is not None:
                            if ((outSignal in self.testSheet.range(r, 9).value) or (
                                    outValueName in self.testSheet.range(r, 9).value)):
                                if "$" + outSignal in self.testSheet.range(r, 11).value:
                                    if self.testSheet.range(r, 7).value == tempMin and self.testSheet.range(r,
                                                                                                            8).value == tempMax:
                                        if outValueName == 'DEMANDE':
                                            if "$" + outValue in self.testSheet.range(r, 12).value:
                                                opFound = 1
                                                match.append(op)
                                        if ((outValueName == 'PAS_DE_DEMANDE') or (outValueName == 'PAS')):
                                            if "$<>" + outValue in self.testSheet.range(r, 12).value:
                                                opFound = 1
                                                match.append(op)
                elif (outSignal.find('MSG') != -1):
                    if self.testSheet.range(r, 10).value is not None:
                        if (('PULSE' in self.testSheet.range(r, 10).value) & (
                                str(int(re.search(r"\d+", timePeriod).group()) * 1000) in self.testSheet.range(r,
                                                                                                               10).value)):
                            if self.testSheet.range(r, 9).value is not None:
                                if ((outSignal in self.testSheet.range(r, 9).value) or (
                                        outValueName in self.testSheet.range(r, 9).value)):
                                    if "$DMD_AFF_MSG_NUM_MSG" in self.testSheet.range(r, 11).value:
                                        if self.testSheet.range(r, 7).value == tempMin and self.testSheet.range(r,
                                                                                                                8).value == tempMax:
                                            if outValueName == 'DEMANDE':
                                                if "$" + outValue in self.testSheet.range(r, 12).value:
                                                    logging.info("op found")
                                                    opFound = 1
                                                    match.append(op)
                                                    # flag=1
                                            if ((outValueName == 'PAS_DE_DEMANDE') or (outValueName == 'PAS')):
                                                if "$<>" + outValue in self.testSheet.range(r, 12).value:
                                                    logging.info("op found")
                                                    opFound = 1
                                                    match.append(op)
                else:
                    if self.testSheet.range(r, 7).value == tempMin and self.testSheet.range(r, 8).value == tempMax:
                        if self.testSheet.range(r, 10).value is None:
                            if self.testSheet.range(r, 9).value is not None:
                                if ((outSignal in self.testSheet.range(r, 9).value) or (
                                        outValueName in self.testSheet.range(r, 9).value)):
                                    if "$" + outSignal in self.testSheet.range(r, 11).value:
                                        if outValue in self.testSheet.range(r, 12).value:
                                            opFound = 1
                                            match.append(op)
        if opFound == 1:
            logging.info("op Match found")
            irow = step[-1]
            copyOfOp = listOfOp.copy()
            for o in listOfOp:
                if o in match:
                    copyOfOp.remove(o)
            return irow, copyOfOp
        else:
            logging.info("invalid step")
            return -1, -1

    def addIpData(self, rows, ipList):
        logging.info("add Ip data in rows", rows, ipList)
        self.logger.info("add Ip data in rows")
        if len(ipList) != 0:
            for n, row in enumerate(rows):
                # signal,valueName,value,tempo,typeStimuli
                try:
                    ipSignal, ipValName, ipValue, tempo, typeStimuli = ipList[n]
                    self.testSheet.range(row, 2).value = tempo
                    self.testSheet.range(row, 4).value = typeStimuli
                    self.testSheet.range(row, 5).value = "$" + ipSignal
                    self.testSheet.range(row, 6).value = ipValue
                    self.testSheet.range(row, 3).value = "Set the " + ipSignal + " on value " + ipValName
                except:
                    break
        else:
            logging.info("No input signals to add for this step")

    def addOpData(self, rows, opList):
        logging.info("add Op data in rows", rows, opList)
        z = 0
        if len(opList) != 0:
            for n, dRow in enumerate(rows):
                logging.info("z", z)

                if self.testSheet.range(dRow, 11).value is None:
                    # signal,valueName,value,tempMin,tempMax,timePeriod
                    logging.info(n)
                    try:
                        if -1 not in opList[n]:
                            opSignal, opValName, opValue, tempMin, tempMax, timePeriod = opList[n]

                            self.testSheet.range(dRow, 7).value = tempMin
                            self.testSheet.range(dRow, 8).value = tempMax
                            self.testSheet.range(dRow, 9).value = opSignal + " = " + opValName
                            if opSignal.find('SON') != -1:
                                self.testSheet.range(dRow, 10).value = "SON"
                                self.testSheet.range(dRow, 11).value = "$DMD_EM_SON_NUM"
                            elif opSignal.find('TEM') != -1:
                                self.testSheet.range(dRow, 10).value = ""
                                self.testSheet.range(dRow, 11).value = "$" + opSignal
                            elif opSignal.find('MSG') != -1:
                                self.testSheet.range(dRow, 10).value = "PULSE(" + str(
                                    int(re.search(r"\d+", timePeriod).group()) * 1000) + ")"
                                self.testSheet.range(dRow, 11).value = "$DMD_AFF_MSG_NUM_MSG"
                            else:
                                if opValue != -1:
                                    self.testSheet.range(dRow,
                                                         9).value = "Set the " + opSignal + " on value " + opValName
                                    self.testSheet.range(dRow, 11).value = "$" + opSignal
                                    self.testSheet.range(dRow, 12).value = opValue
                                else:
                                    logging.info("Signal not valid to be added in test sheet")
                            if opValName == 'DEMANDE':
                                self.testSheet.range(dRow, 12).value = "$" + str(opValue)
                            elif opValName == 'PAS_DE_DEMANDE':
                                self.testSheet.range(dRow, 12).value = "<>$" + str(opValue)
                            else:
                                self.testSheet.range(dRow, 12).value = str(opValue)
                        else:
                            logging.info("No output signals to add -1")
                    except Exception as e:
                        logging.info(e)
                        break
        else:
            logging.info("No output signals to add")

    def addDataInSheet(self, rowData, stepData):
        logging.info("Adding data in sheet")
        dataRow = []
        corp = 0
        retour = 0
        logging.info("len of rowData ", len(rowData))
        addedRow = 0
        KMS.showWindow((self.tpBook.name).split('.')[0])
        EI.activateSheet(self.tpBook, self.testSheet.name)
        if len(rowData) != 0:
            col = 1
            for row, ipList, opList in rowData:
                if len(ipList) != 0:
                    inRow = row
                    dataRow = []
                    row = row + addedRow
                    logging.info("inserting row at ", row)
                    logging.info("Adding ip ", ipList)
                    logging.info("Adding op ", opList)
                    addlines = max([len(ipList), len(opList)])
                    logging.info("Add lines ", addlines)
                    # KMS.showWindow((self.tpBook.name).split('.')[0])
                    time.sleep(1)

                    time.sleep(1)
                    r = 1
                    # dataRow.append(row)
                    while r == addlines:
                        if (self.testSheet.range(row + r, col).value) != 0:
                            TPM.addLineInStep(self.testSheet, row, col)
                            time.sleep(2)
                            row = row + r
                            dataRow.append(row)
                        r = r + 1
                    addedRow = dataRow[-1] - inRow
                    logging.info("data rows ", dataRow, addedRow)
                    opDataRow = [i for i in range(dataRow[-1] - addedRow, dataRow[-1] + 1)]
                    self.addIpData(dataRow, ipList)
                    self.addOpData(opDataRow, opList)
        if len(stepData) != 0:
            for step in stepData:
                logging.info("inserting step ", step)
                if len(step) != 0:
                    ipList, opList = step
                    # KMS.showWindow((self.tpBook.name).split('.')[0])
                    # time.sleep(1)
                    # EI.activateSheet(self.tpBook, self.testSheet.name)
                    time.sleep(1)
                    TPM.addStepInCORPS(self.testSheet)
                    time.sleep(1)
                    emptyRows = self.getNewStep()
                    logging.info("empty row", emptyRows)
                    newRowRange = []
                    try:
                        row = emptyRows[-1]
                        newRowRange.append(row)
                        ptLines = 1
                        addlines = max([len(ipList), len(opList)])
                        logging.info("addlines ", addlines, ptLines)
                        while ptLines < addlines:
                            TPM.addLineInStep(self.testSheet, row, 1)
                            time.sleep(2)
                            ptLines = ptLines + 1
                            row = row + 1
                            newRowRange.append(row)

                        logging.info("Row Range ", newRowRange)
                        self.addIpData(newRowRange, ipList)
                        self.addOpData(newRowRange, opList)
                    except:
                        logging.info("Unable to add steps as new step is not added")
                        break


                else:
                    logging.info("No steps to add in the testSheet no value")

    def parseOutputLine(self, op, opSignal):
        tempMin = None
        tempMax = None
        timePeriod = None
        signal = opSignal.split(op)[0]
        valueName = opSignal.split(op)[1]
        if (valueName.find('$') != -1):
            tempMin = valueName.split('$')[1]
            valueName = valueName.split('$')[0]

        if (valueName.find('#') != -1):
            tempMax = valueName.split('#')[1]
            valueName = valueName.split('#')[0]

        if op == '=':
            if (valueName.find('+') != -1) or (valueName.find('-') != -1):
                value = valueName
            else:
                if ((signal.find('SON') != -1) or (signal.find('TEM') != -1) or (signal.find('MSG') != -1)):
                    value, timePeriod = EI.alertSignalValue(self.alertDoc, signal)
                else:
                    value = EI.getSignalValueOutput(self.dciBook, signal, valueName)
        if op == '<':
            value = valueName + '-1'
        if op == '>':
            value = valueName + '-1'
        if value not in (-1, -2):
            logging.info("value of op ", value)
            return (signal, valueName, value, tempMin, tempMax, timePeriod)
        else:
            logging.info("output signal value not found ")
            return -1

    def parseInputLine(self, op, ipSignals):
        tempo = None
        typeStimuli = None
        funcName = self.tpBook.sheets['Sommaire'].range(4, 3).value
        functionName = funcName.split("-")[1].strip()
        logging.info("Function Name ", functionName, "operator ", op)
        signal = ipSignals.split(op)[0]
        valueName = ipSignals.split(op)[1]
        if (valueName.find('$') != -1) or (valueName.find('#') != -1):
            try:
                tempo = valueName.split('$')[1]
                valueName = valueName.split('$')[0]
            except:
                tempo = valueName.split('#')[1]
                valueName = valueName.split('#')[0]

        # logging.info("Signal and Value ",signal,valueName,value)
        ssFlag = EI.check_ss_fiches(self.ssFiches, signal, valueName, functionName)
        if ssFlag == valueName:
            logging.info("SIgnal is ssFiches")
            typeStimuli = 'FONCTION'
            value = valueName

        if op == '=':
            if (valueName.find('+') != -1) or (valueName.find('-') != -1):
                value = valueName
            elif type(ssFlag) != str:
                value = EI.getSignalValueInput(self.dciBook, self.ssFiches, signal, valueName, functionName)

            else:
                value = valueName
        if op == '!':
            if (valueName.find('+') != -1) or (valueName.find('-') != -1):
                value = valueName
            elif type(ssFlag) != str:
                value = EI.getSignalInitValue(self.dciBook, self.tpBook, self.ssFiches, signal, valueName, functionName)

            else:
                value = EI.getSousFichesInitialValue(self.tpBook, signal)
                logging.info("ss Init value ", value)
        if op == '<':
            value = valueName + '-1'
        if op == '>':
            value = valueName + '+1'

        if type(value) == str:
            return (signal, valueName, value, tempo, typeStimuli)
        elif value == -1:
            logging.info("value of input signal not found")
            return -1
        else:
            logging.info("internal signal")
            return -2

    def getFTSignals(self):
        lstOfSignals = []
        maxrow = self.testSheet.range('A' + str(self.testSheet.cells.last_cell.row)).end('up').row
        for r in range(1, maxrow):
            if self.testSheet.range(r, 5).merge_cells is False:
                if self.testSheet.range(r, 5).value is not None:
                    if type(self.testSheet.range(r, 5).value) == str:
                        if self.testSheet.range(r, 5).value.find('_') != -1 and self.testSheet.range(r, 5).value.find(
                                '$') != -1:
                            lstOfSignals.append(self.testSheet.range(r, 5).value)
        return lstOfSignals

    def checkFuncImpact(self, old, contentFuncImpact):
        macro = EI.getTestPlanAutomationMacro()
        # global funcImpactComment
        logging.info("funcImpactComment = ", AT.funcImpactComment)
        new = self.getFTSignals()
        diff = difflib.ndiff(old, new)
        diffDict = {"equal": [], "del": [], "add": []}

        for i in diff:
            # logging.info("i",i)
            if i.startswith('-'):
                # logging.info("Deleted Word",i.split()[1])
                diffDict["del"].append(i.split()[1])
            if i.startswith('+'):
                # logging.info("Added Word",i.split()[1])
                diffDict["add"].append(i.split()[1])
            if i.startswith(' '):
                # logging.info("No change",i.split()[0])
                diffDict["equal"].append(i.split()[0])

        if self.testSheet.range('C7').value == 'VALIDEE':
            TPM.selectTestSheetModify(macro)
        if len(diffDict["add"]) == 0 and len(diffDict["del"]) == 0:
            logging.info("No Functional impact after analysing FT.....")
            if self.newReq != "":
                if AT.funcImpactComment == 0:
                    logging.info("No functional impact in thematics & content")
                    EI.fillSheetHistory(self.testSheet, "No functional impact")
            else:
                if AT.funcImpactComment == 0:
                    logging.info("No functional impact in thematics & content")
                    EI.fillSheetHistory(self.testSheet, "No functional impact")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines(
                    "\n\nFor " + self.reqName + "(" + self.reqVer + ") contents with functional impact are sucessfully analysed")
        elif len(diffDict["add"]) != 0:
            logging.info("Functional impact after analysing FT Signal added.....")
            if self.newReq != "":
                EI.fillSheetHistory(self.testSheet, " Signals" + " ".join(diffDict["add"]) + " are added.")
            else:
                EI.fillSheetHistory(self.testSheet, " Signals" + " ".join(diffDict["add"]) + " are added.")

            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines(
                    "\n\nFor " + self.reqName + "(" + self.reqVer + ") contents with functional impact are sucessfully analysed")
        elif len(diffDict["del"]) != 0:
            logging.info("Functional impact after analysing FT signal deleted.....")
            if self.newReq != "":
                EI.fillSheetHistory(self.testSheet, " Signals" + " ".join(diffDict["del"]) + " are deleted.")
            else:
                EI.fillSheetHistory(self.testSheet, " Signals" + " ".join(diffDict["del"]) + " are deleted.")

            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines(
                    "\n\nFor " + self.reqName + "(" + self.reqVer + ") contents with functional impact are sucessfully analysed")
        # contentFuncImpact = contentFuncImpact + (diffDict["add"])
        # logging.info("Functional impact Checked.....", (diffDict["add"]))
        contentFuncImpact.extend(diffDict["add"])
        logging.info("contentFuncImpact0000000------->",contentFuncImpact)
        logging.info("Functional impact Checked.....", (diffDict["add"]))

    def TestAnalyse(self, content):
        content = self.txtPreProcessing(content)
        logging.info("func impact", content)
        newContentSteps = self.getRawSteps(content)
        logging.info("new content Steps ", newContentSteps)
        for self.testSheet in self.listOfSheets:
            oldSignals = self.getFTSignals()
            rowSteps = self.getRows()
            self.ssFlag = 0
            insertrow = []
            newStepData = []
            sigDict = {}
            logging.info("row Steps ", rowSteps)
            for contentStep in newContentSteps:
                ipSignals = []
                opSignals = []
                parsedIp = -1
                if contentStep.find('==') != -1:
                    ipSignalList = contentStep.split('==')[0].split('|')
                    opSignalList = contentStep.split('==')[1].split('|')
                    for ip in ipSignalList:  # to find operators and values for signal
                        logging.info("ip signals ", ip)
                        if ip in sigDict:
                            if sigDict[ip] not in (-1, -2):

                                ipSignals.append(sigDict[ip])
                            else:
                                parsedIp = sigDict[ip]
                        else:
                            if ip.find('=') != -1:
                                parsedIp = self.parseInputLine('=', ip)
                                sigDict[ip] = parsedIp
                            elif ip.find('!') != -1:
                                parsedIp = self.parseInputLine('!', ip)
                                sigDict[ip] = parsedIp
                            elif ip.find('<') != -1:
                                parsedIp = self.parseInputLine('<', ip)
                                sigDict[ip] = parsedIp
                            elif ip.find('>') != -1:
                                parsedIp = self.parseInputLine('>', ip)
                                sigDict[ip] = parsedIp
                            else:
                                logging.info("this is not possible in parse input line ")
                        if parsedIp != -1 and parsedIp != -2:

                            ipSignals.append(parsedIp)
                        else:
                            logging.info("Value of input signal " + ip + " not found")
                            break

                    if parsedIp not in (-1, -2):

                        for op in opSignalList:
                            if op in sigDict:
                                if sigDict[op] != -1:
                                    opSignals.append(sigDict[op])
                                else:
                                    parsedOp = -1
                            else:

                                if op.find('=') != -1:
                                    parsedOp = self.parseOutputLine('=', op)
                                    sigDict[op] = parsedOp
                                elif op.find('<') != -1:
                                    parsedOp = self.parseOutputLine('<', op)
                                    sigDict[op] = parsedOp
                                elif op.find('>') != -1:
                                    parsedOp = self.parseOutputLine('>', op)
                                    sigDict[op] = parsedOp
                                else:
                                    logging.info("this is not possible in parse output line ")
                                if (parsedOp != -1):
                                    opSignals.append(parsedOp)
                                    # sigDict[op]=parsedOp
                                else:
                                    break
                                    logging.info("Value of output signal " + op + " not found")

                        self.ssSteps.append((self.ssSignals, opSignals))
                        stepMatch, insertData = self.checkIpSignals(rowSteps, ipSignals)
                        logging.info("insert data ", insertData)
                        logging.info("insert row at ", stepMatch)
                        # logging.info("insert ss data ",self.ssSteps)
                        if stepMatch != -1:
                            irow, idata = self.checkOpSignals(stepMatch, opSignals)
                            if irow != -1:
                                ssSteps = []
                                idCopy = insertData.copy()
                                for i in idCopy:
                                    logging.info("i", i, insertData.index(i))
                                    iSignal, iValName, iVal, tempo, typeStimuli = i

                                    if typeStimuli == 'FONCTION':
                                        ssSteps.append(i)
                                        newStepData.append((ssSteps, opSignals))
                                        if i in insertData:
                                            insertData.remove(i)

                                insertrow.append((irow, insertData, idata))
                            else:
                                newStepData.append((ipSignals, opSignals))

                        else:

                            newStepData.append((ipSignals, opSignals))

            logging.info("Signal Dictionary", sigDict)
            logging.info("insert row and data ", insertrow)
            logging.info("New Step", newStepData)
            self.addDataInSheet(insertrow, newStepData)
            logging.info("Data added in test Sheet ")
            self.checkFuncImpact(oldSignals, contentFuncImpact)
        # EI.addEvovledReq(self.tpBook, self.testSheet, self.reqName+"("+self.reqVer+")",self.newReq)

        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines(
                "\n\nFor " + self.reqName + "(" + self.reqVer + ") Contents with functional impact are sucessfully analysed")
        time.sleep(2)

    def Analyse(self):
        logging.info("Analysing Test Sheet....")
        global contentFuncImpact
        contentFuncImpact = []
        self.logger.info("Analysing Test Sheet....")
        self.getTestSheetRequirement()
        self.funcImpact = self.compareDocs()
        with open('../Aptest_Tool_Report.txt', 'a') as f:
            f.writelines("\n\ncreating Steps")
        if type(self.funcImpact) == str:
            # oldContentSteps=self.getRawSteps(self.oldcontent)
            self.funcImpact = self.txtPreProcessing(self.funcImpact)
            logging.info("func impact", self.funcImpact)
            newContentSteps = self.getRawSteps(self.funcImpact)
            # logging.info("old content Steps ",oldContentSteps)
            logging.info("new content Steps ", newContentSteps)

            # logging.info("new content Steps ",newContentSteps)
            for self.testSheet in self.listOfSheets:
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nchecking steps in " + self.testSheet.name)
                EI.activateSheet(self.tpBook, self.testSheet)
                oldSignals = self.getFTSignals()
                rowSteps = self.getRows()
                self.ssFlag = 0
                insertrow = []
                newStepData = []
                sigDict = {}
                logging.info("row Steps ", rowSteps)
                for contentStep in newContentSteps:
                    ipSignals = []
                    opSignals = []
                    parsedIp = -1
                    if contentStep.find('==') != -1:
                        ipSignalList = contentStep.split('==')[0].split('|')
                        opSignalList = contentStep.split('==')[1].split('|')
                        for ip in ipSignalList:  # to find operators and values for signal
                            logging.info("ip signals ", ip)
                            if ip in sigDict:
                                if sigDict[ip] not in (-1, -2):

                                    ipSignals.append(sigDict[ip])
                                else:
                                    parsedIp = sigDict[ip]
                            else:
                                if ip.find('=') != -1:
                                    parsedIp = self.parseInputLine('=', ip)
                                    sigDict[ip] = parsedIp
                                elif ip.find('!') != -1:
                                    parsedIp = self.parseInputLine('!', ip)
                                    sigDict[ip] = parsedIp
                                elif ip.find('<') != -1:
                                    parsedIp = self.parseInputLine('<', ip)
                                    sigDict[ip] = parsedIp
                                elif ip.find('>') != -1:
                                    parsedIp = self.parseInputLine('>', ip)
                                    sigDict[ip] = parsedIp
                                else:
                                    logging.info("this is not possible in parse input line ")
                            if parsedIp != -1 and parsedIp != -2:

                                ipSignals.append(parsedIp)
                            else:
                                logging.info("Value of input signal " + ip + " not found")
                                break

                        if parsedIp not in (-1, -2):

                            for op in opSignalList:
                                if op in sigDict:
                                    if sigDict[op] != -1:
                                        opSignals.append(sigDict[op])
                                    else:
                                        parsedOp = -1
                                else:
                                    if op.find('=') != -1:
                                        parsedOp = self.parseOutputLine('=', op)
                                        sigDict[op] = parsedOp
                                    elif op.find('<') != -1:
                                        parsedOp = self.parseOutputLine('<', op)
                                        sigDict[op] = parsedOp
                                    elif op.find('>') != -1:
                                        parsedOp = self.parseOutputLine('>', op)
                                        sigDict[op] = parsedOp
                                    else:
                                        logging.info("this is not possible in parse output line ")
                                    if (parsedOp != -1):
                                        opSignals.append(parsedOp)
                                        # sigDict[op]=parsedOp
                                    else:
                                        break
                                        logging.info("Value of output signal " + op + " not found")

                            self.ssSteps.append((self.ssSignals, opSignals))
                            stepMatch, insertData = self.checkIpSignals(rowSteps, ipSignals)
                            logging.info("insert data ", insertData)
                            logging.info("insert row at ", stepMatch)
                            # logging.info("insert ss data ",self.ssSteps)
                            if stepMatch != -1:
                                irow, idata = self.checkOpSignals(stepMatch, opSignals)
                                if irow != -1:
                                    ssSteps = []
                                    idCopy = insertData.copy()
                                    for i in idCopy:
                                        logging.info("i", i, insertData.index(i))
                                        iSignal, iValName, iVal, tempo, typeStimuli = i

                                        if typeStimuli == 'FONCTION':
                                            ssSteps.append(i)
                                            newStepData.append((ssSteps, opSignals))
                                            if i in insertData:
                                                insertData.remove(i)

                                    insertrow.append((irow, insertData, idata))
                                else:
                                    newStepData.append((ipSignals, opSignals))

                            else:

                                newStepData.append((ipSignals, opSignals))

                logging.info("Signal Dictionary", sigDict)
                logging.info("insert row and data ", insertrow)
                logging.info("New Step", newStepData)
                self.addDataInSheet(insertrow, newStepData)
                logging.info("Data added in test Sheet ")
                with open('../Aptest_Tool_Report.txt', 'a') as f:
                    f.writelines("\n\nData added in test Sheet")
                self.checkFuncImpact(oldSignals, contentFuncImpact)
            # EI.addEvovledReq(self.tpBook, self.testSheet, self.reqName+"("+self.reqVer+")",self.newReq)

            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines(
                    "\n\nFor " + self.reqName + "(" + self.reqVer + ") Contents with functional impact are sucessfully analysed")
            time.sleep(2)
        elif type(self.funcImpact) == dict:
            logging.info("Functional Impact in GateWay Contents")
            with open('../Aptest_Tool_Report.txt', 'a') as f:
                f.writelines(
                    "\n\nFor " + self.reqName + "(" + self.reqVer + ") Functional impact in gateway requirment. Please proceed manually.")
        elif self.funcImpact == -1:
            logging.info("self.listOfSheets = ", self.listOfSheets)
            for self.testSheet in self.listOfSheets:
                EI.activateSheet(self.tpBook, self.testSheet)
                logging.info("no functional impact in content")
                logging.info("self.testSheet = ", self.testSheet)
                self.checkFuncImpact(self.getFTSignals(), contentFuncImpact)
            logging.info("no functional impact")
        elif self.funcImpact == -2:
            logging.info("Documents found. Unable to get and analyse contents")
        else:
            logging.info("Content not found ")


class AnalyseGateway:
    def __init__(self, dciBook, data):
        self.data = data
        self.dciBook = dciBook

    def loadGWKeywords(self):  # tested ok
        f_keys = open('../user_input/GateWayKeywords.json', "r")
        keywords = json.load(f_keys)
        return keywords

    def myrange(self, start, stop, step):
        i = 0
        r = start
        while r <= stop:
            yield r
            i = i + 1
            r = round(start + i * step, 3)

    def findMiddle(self, input_list):
        middle = float(len(input_list)) / 2
        if middle % 2 != 0:
            return input_list[int(middle - .5)]
        else:
            return input_list[int(middle)]

    def getGWSignalValues(self, values):
        keywords = self.loadGWKeywords()
        # logging.info(keywords)
        invalidKeys = keywords["GatewayKeywords"]["InvalidValues"]

        # invalidKeys=[keywords[key] for key in keywords if key=="InvalidValues" ]
        # logging.info("invalid keys ",invalidKeys)
        validValues = {}
        invalidValues = {}
        data = values.copy()
        signalRange = ""
        logging.info("da", data)
        for d in values:
            for vals in invalidKeys:

                if re.findall(vals, d):
                    # logging.info("dt----",d,vals,re.findall(vals,d))
                    data.remove(d)
                    try:
                        invalidValues[d.split("=")[0].strip()] = d.split("=")[1]
                    except:
                        invalidValues[d.split(":")[0].strip()] = d.split(":")[1]
                    finally:
                        pass
                # else:
                # logging.info("df----",d,vals,re.findall(vals,d))
        logging.info("data ", data)
        if data:
            for i in data:
                if "[" in i and "]" in i:
                    logging.info("range signal ", i)
                    i = i.strip("[")
                    i = i.strip("]")
                    res = invalidValues['Resolution']
                    # logging.info("res ",res)
                    # try:
                    # logging.info("1",float(res))
                    # logging.info(i.split(";")[0])
                    initVal = 0
                    minVal = float(i.split(";")[0])
                    maxVal = float(i.split(";")[1])
                    logging.info("min ", minVal)
                    logging.info("max ", maxVal)
                    lstData = [i for i in self.myrange(minVal, maxVal, float(res))]
                    # logging.info("npDAta",lstData)
                    midV = lambda lstData: lstData[len(lstData) / 4:len(lstData) * 3 / 4]
                    midVal = self.findMiddle(lstData)
                    logging.info("mid Val", midVal)
                    validValues["init"] = initVal
                    validValues["min"] = minVal + float(res)
                    validValues["mid"] = midVal
                    validValues["max"] = maxVal

                    # except Exception as e:
                    #    logging.info(e)
                    # return[initVal,minVal,midVal,maxVal]

                else:
                    try:
                        validValues[i.split("=")[0].strip()] = i.split("=")[1]
                    except:
                        validValues[i.split(":")[0].strip()] = i.split(":")[1]
                    finally:
                        pass
                    pass
        else:

            logging.info("Invalid signal ")
            return -1

        logging.info("Valid Values ", validValues)
        logging.info("inValid Values ", invalidValues)
        return validValues

    def getFrameSignals(self, frame, pc):
        for sheet in self.dciBook.sheets:
            logging.info("1", sheet, sheet.name.strip())
            if (sheet.name.strip() == "MUX") or (sheet.name.strip() == "FILAIRE"):
                logging.info("2", sheet, sheet.name.strip())
                logging.info("Searching in Sheet...")
                maxrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                sheet_value = sheet.used_range.value
                # logging.info(maxrow)
                # logging.info(sheet)
                # searchResult = EI.searchDataInExcel(sheet, (maxrow, 9), frame)
                searchResult = EI.searchDataInExcelCache(sheet_value, (maxrow, 9), frame)
                logging.info("result--", searchResult)
                # logging.info("dci", signal)

                if searchResult["count"] > 0:
                    logging.info("!....Success....!")
                    result = {}
                    for cellPosition in searchResult["cellPositions"]:
                        p = 0
                        x, y = cellPosition
                        # logging.info("-------------", x, y, getDataFromCell(sheet, (x, y - 1)))
                        if (EI.getDataFromCell(sheet, (x, y - 7))) == "VSM" or (
                        EI.getDataFromCell(sheet, (x, y - 7))) == "BSI":
                            producedConsumed = (EI.getDataFromCell(sheet, (x, y + 2)))
                            if producedConsumed == pc.upper() or producedConsumed == pc.lower():
                                signal = (str(EI.getDataFromCell(sheet, (x, y - 6))))
                                Info = (str(EI.getDataFromCell(sheet, (x, y - 2))))
                                # logging.info("Info.....", Info)
                                # logging.info("signal...",signal)

                                # logging.info("Info.....", splitted)
                                '''for i in splitted:
                                    if value.lower() in i.lower():
                                        a = i.split("=")[1]
                                        logging.info(int(a, 2), producedConsumed)
                                        return (str(int(a, 2)))'''
                                if signal not in result:
                                    splitted = self.getGWSignalValues(Info.split("\n"))
                                    result[signal] = splitted

                            else:

                                p = 1
                    logging.info("result", result)
                    return result
                    if p == 1:
                        logging.info("Produced signal", frame)

                else:
                    logging.info("Signal not foundin Global Dci", frame)

    def createGWSteps(self):
        steps = []
        upFrameSignals = self.getFrameSignals(self.data["UpStreamFrame"], 'C')
        downFrameSignals = self.getFrameSignals(self.data["DownStreamFrame"], 'P')
        upNetwork = self.data["UpStreamNetwork"]
        downNetwork = self.data["DownStreamNetwork"]

        return steps
    # def getSignalValue(self):

#         #return insertrow,newStepData
#
# #
# # '''oldcontent="IF P_INFO_ACPK_MENU=MANEUVER AND P_INFO_ACPK_MANEUVER_STATE=IN_PROGRESS THEN P_SON_ACPK_ACTIVATION=DEMANDE ELSE P_SON_ACPK_ACTIVATION=PAS_DE_DEMANDE"
# # newcontent="IF (P_INFO_ACPK_MENU=MANEUVER AND P_INFO_ACPK_MANEUVER_STATE=IN_PROGRESS OR P_INFO_ACPK_MANEUVER_DIRECTION=FORWARD) AND P_INFO_ACPK_MANEUVER_TYPE=PARALLEL_ENTRY THEN P_SON_ACPK_ACTIVATION=DEMANDE AND P_SON_ACPK_MANEUVER_END=DEMANDE ELSE P_SON_ACPK_ACTIVATION=PAS_DE_DEMANDE AND P_SON_ACPK_MANEUVER_END=PAS_DE_DEMANDE"
# # testData="IF P_INFO_ACPK_MENU=10 AND P_INFO_ACPK_MANEUVER_STATE=1 THEN P_SON_ACPK_ACTIVATION=DEMANDE ELSE P_SON_ACPK_ACTIVATION=PAS_DE_DEMANDE"
# # '''
# EI.openExcel("C:/Users/10388/Downloads/BSI AUtomation/Input/[VSM]DCI_Global_21Q4_v1.xlsx")
# dciBook=xw.Book("C:/Users/10388/Downloads/BSI AUtomation/Input/[VSM]DCI_Global_21Q4_v1.xlsx")
# #alertDoc = EI.openAlertDoc()
# #dciBook = EI.openGlobalDCI()
# logging.info("dciBook Opened")
# #
# logging.info("alertDoc Opened")
# #ssFiches = EI.openSousFiches()
# logging.info("ssFiches Opened")
# values=['[-11;4]', 'Resolution : 0.1', 'InitValue=0x6E', 'Offset=-11.0', 'Est signé=false', 'Bit start=1', 'Taille=8.0', 'Valeur_Invalide_S=FF', 'Valeur_Interdite_S=0X97-0XFD', 'Unite_FR_U=m/s2']
# gateway={'UpStreamNetwork': 'FD3', 'UpStreamFrame': 'FD3_ASU_B_0A3', 'UpStreamSignal': '', 'DownStreamNetwork': 'FD8', 'DownStreamFrame': 'FD8_ASU_B_0A3', 'DownStreamSignal': '', 'TEMPO': '300ms', 'TEMPMIN': '10 ms', 'TEMPMAX': '5 ms'}
# dciBook=xw.Book(r"C:/Users/10388/Downloads/BSI AUtomation/Input/[VSM]DCI_Global_21Q4_v1.xlsx")
# gw=AnalyseGateway(dciBook, gateway)
# #gw.getGWSignalValues(values)
# a=gw.createSteps()
# '''
# tpBook=xw.Book(r"C:/Users/10388/Downloads/BSI AUtomation/Input/Tests_20_70_01272_19_01466_FSE_AEBS_V15_VSM.xlsm")
# alertDoc = xw.Book("C:/Users/10388/Downloads/BSI AUtomation/Input/SI_ALERT.xlsm")
# ssFiches=xw.Book(r"C:/Users/10388/Downloads/BSI AUtomation/Input/[VSM]Tests_01272_18_00824_ss_fiches.xlsm")
# # time.sleep(2)
# # #tpBook = EI.openTestPlan()
# # logging.info("tpBook Opened")
# EI.openExcel(ICF.getTestPlanMacro())
# # time.sleep(1)
# # # '''
# # prevDoc = input("Enter prev Doc path ")
# # logging.info("\n")
# # currDoc = input("Enter curr Doc path ")
# # logging.info("\n")
# # lst = []
# #
# # # number of elements as input
# # n = int(input("Enter number of testSheets : "))
# # logging.info("\n")
# #
# # # iterating till the range
# # for i in range(0, n):
# #     logging.info("\n")
# #     ele = str(input("Enter testsheet name :"))
# #     sheet = tpBook.sheets[ele]
# #     lst.append(sheet)  # adding the element
# # newReq=""
# # logging.info(lst)
# # reqName = input("Enter old requirement name to be checked in current doc ")
# # logging.info("\n")
# # reqVer = input("Enter old requirement version to be checked in current doc ")
# # logging.info("\n")
# # newReq=input("Enter new requirement version to be checked in current doc if any else press enter ")
# # testSheetName = input("Enter Test Sheet Name to be Modified")
# # logging.info("\n")
# # logging.info(testSheetName)
# # logging.info("\n")
# # newcontent = input("Enter new content ")
# # logging.info("\n")'''
# #
# prevDoc=r"C:/Users/10388/Downloads/BSI AUtomation/Input/[V7]SSD_HMIF_GROUND_LINK_HMI.docx"
# currDoc=r"C:/Users/10388/Downloads/BSI AUtomation/Input/[V8]SSD_HMIF_GROUND_LINK_HMI.docx"
# reqName="REQ-0632898"
# reqVer="B"
# newReq=""
# #
# lst=[tpBook.sheets['VSM20_GC_20_70_0002'],
#      tpBook.sheets['VSM20_GC_20_70_0002A'],
#      tpBook.sheets['VSM20_GC_20_70_0002B'],
#      tpBook.sheets['VSM20_GC_20_70_0002C'],
#      tpBook.sheets['VSM20_GC_20_70_0002D'],
#      tpBook.sheets['VSM20_GC_20_70_0002E'],
#      tpBook.sheets['VSM20_GC_20_70_0003'],
#      tpBook.sheets['VSM20_N1_20_70_0016'],
#      tpBook.sheets['VSM20_N1_20_70_0024'],
#      tpBook.sheets['VSM20_N1_20_70_0026'],
#      tpBook.sheets['VSM20_N1_20_70_0034'],
#      tpBook.sheets['VSM20_N1_20_70_0041'],
#      tpBook.sheets['VSM20_N1_20_70_0042'],
#      tpBook.sheets['VSM20_N1_20_70_0042A']]
# KMS.showWindow(tpBook.name.split('.')[0])
# time.sleep(1)
# TPM.selectArch()
# time.sleep(1)
# TPM.selectTpWritterProfile()
# time.sleep(1)
# TPM.selectToolbox()
# time.sleep(1)
# #testSheet=tpBook.sheets['VSM20_N1_20_64_0003']
# #
# #testSheet = tpBook.sheets[testSheetName]
# # #logging.info(testSheet)
# # #EI.activateSheet(tpBook, testSheet.name)
# # # if testSheet.range('C7').value == 'VALIDEE':
# # #     TPM.selectTestSheetModify()
# # #     logging.info("Testsheet Modified")
# logging.info(tpBook.name)
# logging.info(dciBook.name)
# logging.info(alertDoc.name)
# logging.info(ssFiches.name)
# newReq=""
# #
# # # rowData=[(21, [('P_INFO_ACPK_MANEUVER_TYPE', 'PARALLEL_ENTRY', 1)], [('P_SON_ACPK_MANEUVER_END', 'DEMANDE', 'C1', '')]), (25, [('P_INFO_ACPK_MANEUVER_TYPE', 'PARALLEL_ENTRY', 0)], [('P_SON_ACPK_MANEUVER_END', 'PAS_DE_DEMANDE', 'C1', '')])]
# # # stepData=[([('P_INFO_ACPK_MANEUVER_DIRECTION', 'FORWARD', 2), ('P_INFO_ACPK_MANEUVER_TYPE', 'PARALLEL_ENTRY', 1)], [('P_SON_ACPK_ACTIVATION', 'DEMANDE', 'C2', ''), ('P_SON_ACPK_MANEUVER_END', 'DEMANDE', 'C1', '')]), ([('P_INFO_ACPK_MANEUVER_DIRECTION', 'FORWARD', 0), ('P_INFO_ACPK_MANEUVER_TYPE', 'PARALLEL_ENTRY', 0)], [('P_SON_ACPK_ACTIVATION', 'PAS_DE_DEMANDE', 'C2', ''), ('P_SON_ACPK_MANEUVER_END', 'PAS_DE_DEMANDE', 'C1', '')])]
# analyseContents=AnalyseTestSheet(dciBook, tpBook, alertDoc, ssFiches, currDoc, prevDoc,lst,reqName, reqVer,newReq=newReq)
# content="""IF
# MODE_CONFIG_VHL = CLIENT OR APV

# THEN
# FA_DISPONIBLE_HAB = DISPONIBLE AND
# AFFM_FARC_MENU_FA = DEMANDE

# ELSE
# FA_DISPONIBLE_HAB = INDISPONIBLE AND
# AFFM_FARC_MENU_FA = PAS_DEMANDE

# newContentSteps=['ETAT_MT=MOTEUR_TOURNANT$FILTRE_DA_1|ETAT_PRINCIP_SEV=CONTACT$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4$PRM_TIME_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE', 'ETAT_MT=MOTEUR_TOURNANT$FILTRE_DA_1|ETAT_PRINCIP_SEV=CONTACT$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE', 'ETAT_MT=MOTEUR_TOURNANT$FILTRE_DA_1|ETAT_PRINCIP_SEV=DEM$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4$PRM_TIME_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE', 'ETAT_MT=MOTEUR_TOURNANT$FILTRE_DA_1|ETAT_PRINCIP_SEV=DEM$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE', 'ETAT_MT=ARRETE$FILTRE_DA_1|ETAT_PRINCIP_SEV=CONTACT$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4$PRM_TIME_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE', 'ETAT_MT=ARRETE$FILTRE_DA_1|ETAT_PRINCIP_SEV=CONTACT$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE', 'ETAT_MT=ARRETE$FILTRE_DA_1|ETAT_PRINCIP_SEV=DEM$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4$PRM_TIME_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE', 'ETAT_MT=ARRETE$FILTRE_DA_1|ETAT_PRINCIP_SEV=DEM$FILTRE_DA_2|DEFAUT_DIRECTION_ASSISTEE=DEMANDE_VOYANT_ROUGE|VITESSE_VEHICULE_ROUES>PRM_V_DA_FAULT_G4==_P_MSG_DA_DEFAUT_G4=DEMANDE|P_SON_DA_DEFAUT_G4=DEMANDE|P_MSG_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE|P_SON_DA_DEFAUT_G4_PARKING=PAS_DE_DEMANDE']
# analyseContents.TestAnalyse(content)
# # #data = AnalyseTestSheet(dciBook, tpBook, alertDoc, ssFiches, testSheet, newcontent)
# #stepData=[([('ETAT_MA', 'ENGAGE', 'ENGAGE', None, 'FONCTION')], [('P_SON_ACPK_ACTIVATION', 'DEMANDE', 'C2', None, None, ''), ('P_MSG_CP_DEFAUT', 'PAS_DE_DEMANDE', 136.0, None, None, '10 s'), ('P_INFO_ACPK_MANEUVER_DIRECTION', 'REVERSE_IN_PROGRESS', 1, None, None, None)]), ([('P_INFO_ACPK_MENU', 'MANEUVER', '0', None, None), ('P_INFO_ACPK_MANEUVER_STATE', 'IN_PROGRESS', '0', None, None), ('ETAT_MA', 'ENGAGE', '', None, 'FONCTION')], [('P_SON_ACPK_ACTIVATION', 'PAS_DE_DEMANDE', 'C2', None, None, ''), ('P_MSG_CP_DEFAUT', 'DEMANDE', 136.0, None, None, '10 s'), ('P_INFO_ACPK_MANEUVER_DIRECTION', 'REVERSE', 0, None, None, None)])]
# #rows=[28, 29, 30]
# #op=[('P_SON_ACPK_ACTIVATION', 'DEMANDE', 'C2', None, None, ''), ('P_MSG_CP_DEFAUT', 'PAS_DE_DEMANDE', 136.0, None, None, '10 s'), ('P_INFO_ACPK_MANEUVER_DIRECTION', 'REVERSE_IN_PROGRESS', 1, None, None, None)]
# #analyseContents.addOpData(rows, op)
# #analyseContents.addDataInSheet([],stepData)
# analyseContents.Analyse()
# # '''dciBook.close()
# # tpBook.close()
# # alertDoc.close()
# # ssFiches.close()
# '''try:
#
#      analyseContents.Analyse()
#      dciBook.close()
#      tpBook.close()
#      alertDoc.close()
#      ssFiches.close()
# except Exception as e:
#      logging.info(e)
#
#      dciBook.close()
#      tpBook.close()
#      alertDoc.close()
#      ssFiches.close()'''
