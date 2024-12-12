import json

# Open JSON file
import os
import logging
coordinateMap = {}
userInput = {}


def loadConfig():
    if os.path.isfile('../config/CoordinateMap.json'):
        # f_coordinateMap = open('../config/CoordinateMap.json', "r")
        with open('../config/CoordinateMap.json', "r") as f_coordinateMap:
            global coordinateMap
            coordinateMap = json.load(f_coordinateMap)
    else:
        logging.info("CoordinateMap.json File not found in /config folder")
        exit()

    if os.path.isfile('../user_input/UserInput.json'):
        with open('../user_input/UserInput.json', "r") as f_userInput:
            global userInput
            userInput = json.load(f_userInput)
    else:
        logging.info("UserInput.json File not found in /user_input folder")
        exit()


def getIEPath():
    return userInput["toolsPath"]["IE"]


def getDocInfoUrl():
    DocInfoPath = os.path.abspath(r'https://docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')
    return DocInfoPath
    # return userInput["toolsPath"]["DocInfoLink"]


def getExcelPath():
    ExcelPath1 = os.path.abspath(r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')
    ExcelPath2 = os.path.abspath(r'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE')
    ExcelPath3 = os.path.abspath(r'C:\Program Files (x86)\Microsoft Office 2016\Office16\EXCEL.EXE')
    if os.path.exists(ExcelPath1):
        logging.info("Excel Path is Present:", ExcelPath1)
        return ExcelPath1

    elif os.path.exists(ExcelPath2):
        logging.info("Excel Path is Present:", ExcelPath2)
        return ExcelPath2
    # return userInput["toolsPath"]["Excel"]

    elif os.path.exists(ExcelPath3):
        logging.info("Excel Path is Present:", ExcelPath3)
        # logging.info('************Excel Path is not Present*************')
        return ExcelPath3
    else:
        logging.info('************Excel Path is not Present*************')
        # TBD : Excel inot installed either in "Pogram Files" or in "Program Files (x86)" should we search the whole system ?


def getAnalyseDeEntrant():
    return userInput["toolsPath"]["AnalyseDeEntrant"]


def getTaskName():
    return userInput["toolsPath"]["AnalyseDeEntrant"]


def getTestPlan():
    return userInput["toolsPath"]["TestPlan"]


def getDCI():
    return userInput["toolsPath"]["InputDCI"]


def FetchTaskName():
    global inp_taskname
    for tname in userInput['taskDetails']:
        inp_taskname = tname['taskName']
        return inp_taskname


def gettrigram():
    for tname in userInput['taskDetails']:
        trigram = tname['trigram']
    return trigram

def getTestPlanMacro():
    # TestPlanMacroPath = os.path.abspath(r'..\Macros\MacrosValidationBSI.xlam')
    # return TestPlanMacroPath
    return userInput["toolsPath"]["TestPlanMacro"]


def getParamGlobal():
    return userInput["toolsPath"]["ParamGlobal"]


def getOutputFiles():
    return userInput["toolsPath"]["OutputFolder"]


def getQIAParam():
    return userInput["toolsPath"]["QIAParam"]


def getMessagerie():
    return userInput["toolsPath"]["BSI_Messagerie"]


def getUploadFiles():
    return userInput["toolsPath"]["UploadFiles"]


def getTaskDetails():
    return userInput["taskDetails"]


def getDocToDownload():
    dictToDownload = [
        {
            'Reference': userInput['Analyse_des_entrant']['Reference'],
            'Version': userInput['Analyse_des_entrant']['Version']
        },
        {
            'Reference': userInput['Qia_Param_Global']['Reference'],
            'Version': userInput['Qia_Param_Global']['Version']
        }
    ]
    return dictToDownload


def getSsdFolder():
    SsdFolder = os.path.abspath(r'..\Input_Files\SSD_folder')
    if not os.path.exists(SsdFolder):
        os.makedirs(SsdFolder)
    return SsdFolder


def getDicFolder():
    dciFolder = os.path.abspath(r'..\Input_Files\DCI_Functional_Files')
    if not os.path.exists(dciFolder):
        os.makedirs(dciFolder)
    return dciFolder


def getQIAParamCalibration():
    QIAParam = os.path.abspath(r'..\Input_Files\QIA_00949_11_06142_PARAM_GLOBAL.xlsm')
    if os.path.exists(QIAParam):
        logging.info("QIA param is Present:", QIAParam)
        return QIAParam
    else:
        return False


def getTPInitCoordinate():
    return coordinateMap["selectTPInit"]


def getArchCoordinate():
    return coordinateMap["selectArch"]


def getToolBoxCoordinate():
    return coordinateMap["selectTestSheetToolBox"]


def getTestSheetModifyCoordinate():
    return coordinateMap["selectTestSheetModify"]


def getTpWritterProfileCoordinate():
    return coordinateMap["selectTpWritterProfile"]


def getTPImpactCoordinate():
    return coordinateMap["selectTPImpact"]


def getAddThematiqueCoordinate():
    return coordinateMap["addThematique"]


def getSynthUpdateCoordinate():
    return coordinateMap["selectSynthUpdate"]


def getOption2Coordinate():
    return coordinateMap.coordinates.Option2


def getInputFolder():
    return userInput["toolsPath"]["InputFolder"]


def getDownloadFolder():
    return userInput["toolsPath"]["DownloadFolder"]


def getWebDriver():
    WebDriverPath = os.path.abspath(r'..\config\chromedriver.exe')
    return WebDriverPath
    # return userInput["toolsPath"]["WebDriver"]


def getBackLog():
    return userInput["getThematicLines"]

def getReqNameChange():
    return userInput["getReqNameChanges"]

def getGatewayReq():
    return userInput["getGatewayReqs"]

def getInterfaceReqNameChange():
    return userInput["getInterfaceReqNameChanges"]


def getCheckReportStatus():
    return userInput["check_report_for_pt"]


def getCredentials():
    return userInput["docInfo"]["username"], userInput["docInfo"]["password"]


def getAutoDownloadStatusAnalyzeDeEntrant():
    logging.info("Value in JSON autoDownloadStatus = ", userInput["getDownloadStatusAnalyzeDeEntrant"])
    return userInput["getDownloadStatusAnalyzeDeEntrant"]


def getAutoDownloadStatusInputDocument():
    logging.info("Value in JSON autoDownloadStatus = ", userInput["getDownloadStatusInputDocument"])
    return userInput["getDownloadStatusInputDocument"]


def getAutoDownloadStatusPreviousDocument():
    logging.info("Value in JSON autoDownloadStatus = ", userInput["getDownloadStatusPreviousDocument"])
    return userInput["getDownloadStatusPreviousDocument"]
