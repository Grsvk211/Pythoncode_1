import KeyboardMouseSimulator as KMS
import InputConfigParser as ICF
import ExcelInterface as EI
import time
import xlwings as xw
import threading
import InputConfigParser as ICF
import os
import pygetwindow as pgw
import pyautogui
import re
import logging
def addStepInINTIAL(sheet):
    sheet_value = sheet.used_range.value
    # CellValue = EI.searchDataInExcel(sheet, (26, 100), "CONDITIONS INITIALES")
    CellValue = EI.searchDataInExcelCache(sheet_value, (26, 100), "CONDITIONS INITIALES")
    x, y = CellValue["cellPositions"][0]
    sheet.range(x, y).select()
    KMS.rightClick()
    time.sleep(1)
    KMS.rightArrow()
    time.sleep(1)
    KMS.pressEnter()
    time.sleep(1)
    KMS.pressEnter()

def addStepInCORPS(sheet):
    sheet_value = sheet.used_range.value
    # CellValue = EI.searchDataInExcel(sheet, (26, 100), "CONDITIONS INITIALES")
    CellValue = EI.searchDataInExcelCache(sheet_value, (26, 100), "CONDITIONS INITIALES")
    x, y = CellValue["cellPositions"][0]
    logging.info("Adding step in corp at ",x,y)
    sheet.range(x, y).select()
    KMS.rightClick()
    time.sleep(1)
    KMS.rightArrow()
    time.sleep(1)
    KMS.downArrow()
    time.sleep(1)
    KMS.pressEnter()
    time.sleep(1)
    KMS.pressEnter()


def addStepInRETOUR(sheet):
    sheet_value = sheet.used_range.value
    # CellValue = EI.searchDataInExcel(sheet, (26, 100), "CONDITIONS INITIALES")
    CellValue = EI.searchDataInExcelCache(sheet_value, (26, 100), "CONDITIONS INITIALES")
    x, y = CellValue["cellPositions"][0]
    logging.info("Adding step in Retour at ", x, y)
    sheet.range(x, y).select()
    KMS.rightClick()
    time.sleep(1)
    KMS.rightArrow()
    time.sleep(1)
    KMS.downArrow()
    time.sleep(1)
    KMS.downArrow()
    time.sleep(1)
    KMS.pressEnter()
    time.sleep(1)
    KMS.pressEnter()


def addLineInStep(sheet, row, col):
    logging.info("Adding line in step ", row)
    sheet.range(row, col).select()
    KMS.rightClick()
    time.sleep(2)
    KMS.downArrow()
    time.sleep(2)
    KMS.rightArrow()
    time.sleep(2)
    KMS.pressEnter()

def selectToolbox(wb):
    # left, top, width, height = KMS.getCoordinatesByImage("../images/toolbox.png")
    # # KMS.moveMouse(left + 15, top + 15)
    # x_coordinates = ICF.getToolBoxCoordinate()['left']
    # y_coordinates = ICF.getToolBoxCoordinate()['top']
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)

    # ribbon = wb.macro("gestion_menu_fiche.CreationMenus")
    ribbon = wb.macro("TPMacros.Button_ToolBox")
    ribbon()

def selectArch(wb):
    # left, top, width, height = KMS.getCoordinatesByImage("../images/architecture.png")
    # # KMS.moveMouse(left + 15, top + 15)
    # x_coordinates = ICF.getArchCoordinate()['left']
    # y_coordinates = ICF.getArchCoordinate()['top']
    # # logging.info(left, top)
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)



    # Ruban as Main Module
    # ribbon = wb.macro('TestPlan_Macros.TriggerBtnArchi')
    # ribbon()
    UI_task_nameEI = ICF.FetchTaskName()
    logging.info("UI_task_nameEI", UI_task_nameEI)
    if UI_task_nameEI.split("_")[0] == 'F':
        # Ruban as Main Module
        ribbon = wb.macro('TPMacros.Button_Archi')
        [ribbon() for _ in range(2)]
    else:
        # Ruban as Main Module
        ribbon = wb.macro('TPMacros.Button_Archi')
        ribbon()


def selectTestSheetModify(wb):
    # left, top, width, height = KMS.getCoordinatesByImage("../images/test_sheet.png")
    # # KMS.moveMouse(left + 10, top + 10)
    # x_coordinates = ICF.getTestSheetModifyCoordinate()['left1']
    # y_coordinates = ICF.getTestSheetModifyCoordinate()['top1']
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # # KMS.moveMouse(left + 25, top + 90)
    # x1_coordinates = ICF.getTestSheetModifyCoordinate()['left2']
    # y1_coordinates = ICF.getTestSheetModifyCoordinate()['top2']
    # KMS.moveMouse(left + int(x1_coordinates), top + int(y1_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()

    # ribbon = wb.macro("Gestion_PT.modif_fiche")
    ribbon = wb.macro("TPMacros.Button_TSheetModify")
    ribbon()


def selectTestSheetAdd(wb):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")
    tpBook = EI.openTestPlan()
    KMS.showWindow(tpBook.name.split('.')[0])
    time.sleep(2)
    ribbon = wb.macro("TPMacros.TestSheetAdd")
    time.sleep(2)
    t1 = threading.Thread(target=doClick)
    t1.start()
    ribbon()
    # time.sleep(5)

def doClick():
    logging.info("Doclick----------")
    KMS.upArrow()
    time.sleep(2)
    logging.info("click up arrow")
    KMS.downArrow()
    time.sleep(2)
    logging.info("click down arrow")
    KMS.pressTab()
    time.sleep(2)
    KMS.pressEnter()
    time.sleep(2)
    KMS.rightArrow()
    time.sleep(2)
    KMS.pressEnter()
    time.sleep(2)
    # KMS.pressTab()
    # time.sleep(2)
    KMS.pressEnter()
    time.sleep(2)
    KMS.pressEnter()




# Insertion_etape_init
# Insertion_etape_corps_de_test
# Insertion_etape_retour_condition_init
def addInitialContionsStep(wb):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")
    ribbon = wb.macro("TPMacros.addInitialContionsStep")
    time.sleep(1)
    t1 = threading.Thread(target=doImpactPressEnter)
    t1.start()
    ribbon()


def addCorpDeTestStep(wb):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")
    ribbon = wb.macro("TPMacros.addCorpDeTest")
    time.sleep(1)
    t1 = threading.Thread(target=doImpactPressEnter)
    t1.start()
    ribbon()


def addRetourContionsStep(wb):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")
    ribbon = wb.macro("TPMacros.addRetourContionsStep")
    time.sleep(1)
    t1 = threading.Thread(target=doImpactPressEnter)
    t1.start()
    ribbon()


def selectTpWritterProfile(wb):
    # left, top, width, height = KMS.getCoordinatesByImage("../images/tp_writter_profile.png")
    # # KMS.moveMouse(left + 10, top + 10)
    # x_coordinates = ICF.getTpWritterProfileCoordinate()['left1']
    # y_coordinates = ICF.getTpWritterProfileCoordinate()['top1']
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # # KMS.moveMouse(left + 10, top + 200)
    # x1_coordinates = ICF.getTpWritterProfileCoordinate()['left2']
    # y1_coordinates = ICF.getTpWritterProfileCoordinate()['top2']
    # KMS.moveMouse(left + int(x1_coordinates), top + int(y1_coordinates))

    # ribbon = wb.macro('Ruban.RefreshRibbon')
    ribbon = wb.macro("TPMacros.Button_UserProfile")
    ribbon("Tab_VSIV_PT_En")

def selectTPInit(wb):

    # left, top, width, height = KMS.getCoordinatesByImage("../images/Tp_Init.png")
    # # KMS.moveMouse(left + 30, top + 30)
    # x_coordinates = ICF.getTPInitCoordinate()['left1']
    # y_coordinates = ICF.getTPInitCoordinate()['top1']
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # # KMS.moveMouse(left + 50, top + 120)
    # x1_coordinates = ICF.getTPInitCoordinate()['left2']
    # y1_coordinates = ICF.getTPInitCoordinate()['top2']
    # KMS.moveMouse(left + int(x1_coordinates), top + int(y1_coordinates))
    # time.sleep(1)

    # KMS.mouseClick()
    # KMS.showWindow("Tests")
    tpBook = EI.openTestPlan()
    KMS.showWindow(tpBook.name.split('.')[0])
    popupribbon = wb.macro("TPMacros.TriggerPopup")
    popupribbon("summary sheet")

    # Gestion_PT Main Module Fun Modif_PT
    # ribbon = wb.macro("TestPlan_Macros.TestPlanInitModif")
    ribbon = wb.macro("TPMacros.Button_TPModify")
    ribbon()


def selectTPImpact(wb):
    # left, top, width, height = KMS.getCoordinatesByImage("../images/Tp_Init.png")
    # # KMS.moveMouse(left + 30, top + 30)
    # x_coordinates = ICF.getTPImpactCoordinate()['left1']
    # y_coordinates = ICF.getTPImpactCoordinate()['top1']
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # KMS.moveMouse(left + 50, top + 175)
    # x1_coordinates = ICF.getTPImpactCoordinate()['left2']
    # y1_coordinates = ICF.getTPImpactCoordinate()['top2']
    # KMS.moveMouse(left + int(x1_coordinates), top + int(y1_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # KMS.pressEnter()
    # time.sleep(1)

    # Gestion_PT Main Module
    tpBook = EI.openTestPlan()
    KMS.showWindow(tpBook.name.split('.')[0])
    # ribbon = wb.macro("TestPlan_Macros.TestPlanInitImpact")
    ribbon = wb.macro("TPMacros.Button_TP_InitImpact")
    t1 = threading.Thread(target=doImpactPressEnter)
    t1.start()
    ribbon()


def doImpactPressEnter():
    time.sleep(3)
    KMS.pressEnter()

def addThematique(wb):
    # left, top, width, height = KMS.getCoordinatesByImage("../images/add_thematique.png")
    # # KMS.moveMouse(left + 10, top + 10)
    # x_coordinates = ICF.getAddThematiqueCoordinate()['left1']
    # y_coordinates = ICF.getAddThematiqueCoordinate()['top1']
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # # KMS.moveMouse(left + 25, top + 55)
    # x1_coordinates = ICF.getAddThematiqueCoordinate()['left2']
    # y1_coordinates = ICF.getAddThematiqueCoordinate()['top2']
    # KMS.moveMouse(left + int(x1_coordinates), top + int(y1_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # KMS.pressEnter()
    # time.sleep(1)

    # ribbon = wb.macro("TestPlan_Macros.TriggerThematicAdd")
    ribbon = wb.macro("TPMacros.Button_ThematicAdd")
    time.sleep(1)
    t1 = threading.Thread(target=doImpactPressEnter)
    t1.start()
    ribbon()

def selectSynthUpdate(wb):
    # left, top, width, height = KMS.getCoordinatesByImage("../images/tp_final.png")
    # # KMS.moveMouse(left + 10, top + 10)
    # x_coordinates = ICF.getSynthUpdateCoordinate()['left1']
    # y_coordinates = ICF.getSynthUpdateCoordinate()['top1']
    # KMS.moveMouse(left + int(x_coordinates), top + int(y_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()
    # time.sleep(1)
    # # KMS.moveMouse(left + 25, top + 90)
    # x1_coordinates = ICF.getSynthUpdateCoordinate()['left2']
    # y1_coordinates = ICF.getSynthUpdateCoordinate()['top2']
    # KMS.moveMouse(left + int(x1_coordinates), top + int(y1_coordinates))
    # time.sleep(1)
    # KMS.mouseClick()

    # KMS.showWindow("Excel")
    tpBook = EI.openTestPlan()
    KMS.showWindow(tpBook.name.split('.')[0])
    popupribbon = wb.macro("TPMacros.TriggerPopup")
    popupribbon("sysnthesis")
    # ribbon = wb.macro("TestPlan_Macros.TriggerBtnSynthesisUpdate")
    ribbon = wb.macro("TPMacros.Button_SynUpdate")
    ribbon()


def unProtectTestSheet(wb):
    ribbon = wb.macro("TPMacros.UnprotectWB")
    ribbon()


def TestSheetRemove(wb):
    ribbon = wb.macro("TPMacros.TestSheetRemove")
    t1 = threading.Thread(target=doImpactPressEnter)
    t1.start()
    ribbon()


def synchronizeSubSheet(wb):
    ribbon = wb.macro("TPMacros.syncSubSheets")
    t1 = threading.Thread(target=doImpactPressEnter)
    t1.start()
    ribbon()

def selectSynthUpdateFor_SF(wb):

    #sfBook = EI.openSousFiches()
   # KMS.showWindow(sfBook.name.split('.')[0])
    popupribbon = wb.macro("TPMacros.TriggerPopup")
    popupribbon("sysnthesis")
    # ribbon = wb.macro("TestPlan_Macros.TriggerBtnSynthesisUpdate")
    ribbon = wb.macro("TPMacros.Button_SynUpdate")
    ribbon()

# def addLineInStep(wb):
#     ribbon = wb.macro("TPMacros.addLineInStep")
#     ribbon()


def addVSMCheckReportTestStep(wb):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")
    ribbon = wb.macro("TPMacros.checkReport")
    time.sleep(3)
    dvcrts = threading.Thread(target=doVSMCheckReportTestStep)
    dvcrts.start()
    ribbon()

def doVSMCheckReportTestStep():
    # pop up - click on YES
    time.sleep(6)
    x = ICF.FetchTaskName().split('_')[1]
    flg = 0
    for j in pgw.getAllWindows():
        if 'ss_fiches' in j.title:
            flg = 1
    while True:
        active_window = pgw.getActiveWindow()
        if active_window is not None:
            if active_window.title == "Windows Security":
                time.sleep(6)
            else:
                logging.info("No Security Window")
                break
        else:
            logging.info("No Window to enter username and password")
    while True:
        active_window1 = pgw.getActiveWindow()
        if active_window1 is not None:
            if active_window1.title == "Check plan de test":
                logging.info("Active Window 1 Title:", active_window1.title)
                time.sleep(2)
                KMS.leftArrow()
                time.sleep(1)
                KMS.pressEnter()
                logging.info("clicked on Yes of excel popup")
                break
            else:
                logging.info("No Excel window")
                break
            # else:
            #     logging.info("No Excel Window")
            #     break
        else:
            logging.info("No Window to click on Ok")
    time.sleep(6)
    if (flg == 1):
        for j in range(0, 11):
            KMS.keyboard.press(KMS.Key.tab)
            time.sleep(1)
        time.sleep(4)
        pyautogui.typewrite(x)
        time.sleep(4)
        KMS.keyboard.press(KMS.Key.backspace)
        time.sleep(1)
        KMS.keyboard.press(KMS.Key.delete)
        time.sleep(1)
        pyautogui.typewrite(x[-1])
        time.sleep(4)
        KMS.keyboard.press(KMS.Key.tab)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'space')
        time.sleep(1)
        for j in range(0, 12):
            pyautogui.hotkey('shift', 'tab')
            time.sleep(1)

    # french window to click on first button top
    # select first button from the window for Checkreport

    # french window to click on first button top
    # select first button from the window for Checkreport
    time.sleep(6)
    KMS.pressEnter()
    # excel popup - click on Yes button - longue
    time.sleep(2)
    KMS.pressEnter()
    i = 0

    while True:
        active_window1 = pgw.getActiveWindow()
        if active_window1 is not None:
            if active_window1.title == "Microsoft Excel":
                logging.info("Active Window 1 Title:", active_window1.title)
                time.sleep(1)
                KMS.pressEnter()
                logging.info("clicked on OK of excel popup")
                i = i + 1
                if (flg == 0):
                    break
                if (i == 2):
                    break

        else:
            logging.info("No Window to click on Ok")

    while True:
        # here browser window to select the output folder
        active_window2 = pgw.getActiveWindow()
        if active_window2 is not None:
            if active_window2.title == "Browse For Folder":
                logging.info("Active Window 2 Title:", active_window2.title)
                KMS.keyboard.press(KMS.Key.tab)
                time.sleep(2)
                KMS.keyboard.press(KMS.Key.tab)
                time.sleep(0.5)
                KMS.keyboard.release(KMS.Key.tab)
                time.sleep(0.5)

                # towards _A folder & ts sub folder for output to save on Desktop
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                # sub folder - right arrow
                KMS.rightArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(1)
                KMS.downArrow()
                time.sleep(1)
                # press enter to select the checkreport op foder
                KMS.pressEnter()
                break
        # elif (active_window2.title == "Microsoft Excel"):
        #     logging.info("Active Window 2 Title:", active_window2.title)
        #     time.sleep(1)
        #     KMS.pressEnter()
        #     logging.info("clicked on OK of excel popup")
        else:
            logging.info("No Browser window")

    os.chdir("../Input_Files")
    go_to_input_dir1 = os.getcwd()
    input_dir_inside1 = str(go_to_input_dir1)
    logging.info(input_dir_inside1)

    # -------------------------------------
    # to capture exact vsm global param file
    list_of_input_files1 = os.listdir(input_dir_inside1)
    logging.info(list_of_input_files1)

    global vsm_global_param_file1
    for vsm_param1 in list_of_input_files1:
        match_vsm1 = re.match(r"PARAM_Global_VSM_(.*?)\.xlsm$", vsm_param1)

        if match_vsm1:
            vsm_global_param_file1 = str(vsm_param1)
            logging.info("Param global VSM :- ", vsm_param1)

        else:
            logging.info("Param Global file not found in input folder")

    # ----------------------------------
    # # here PARAM global file has to select by manual
    while True:
        # here browser window to select the output folder
        active_window3 = pgw.getActiveWindow()
        if active_window3 is not None:
            if active_window3.title == "Choix du fichier PARAM." or active_window3.title == "PARAM GLOBAL Selection":
                logging.info("Active Window 3 Title:", active_window3.title)
                time.sleep(1)
                # param file location path
                pyautogui.typewrite(input_dir_inside1)
                time.sleep(1)
                KMS.pressEnter()
                # -----------------------------
                # param_vsm_file = "PARAM_Global_VSM_01272_19_02001.xlsm"
                # param_vsm_global = vsm_global_param
                logging.info("VSMGlobal param file is:- ", vsm_global_param_file1)
                # param file name typing
                time.sleep(1)
                pyautogui.typewrite(vsm_global_param_file1)
                # -----------------------------
                time.sleep(1)
                KMS.pressEnter()
                time.sleep(1)
                logging.info("PARAM global file for VSM assigned by automation")
                break
        else:
            logging.info("No window")

    # C:/PSA popup window appears CLICK OK

    if (flg == 0):
        while True:
            # here browser window to select the output folder
            active_window4 = pgw.getActiveWindow()
            # logging.info("Active Window 4 Title:", active_window4.title)
            if active_window4 is not None:
                # logging.info("Active Window 4 Title:", active_window4.title)
                if active_window4.title == "Microsoft Excel":
                    logging.info("Active Window 4 Title:", active_window4.title)
                    time.sleep(2)
                    KMS.pressEnter()
                    time.sleep(1)

                    break
            else:
                logging.info("No window -4 of Microsoft Excel")
    time.sleep(2)
    if (flg == 0):
        clickVSMCheckRPTFinalindow(EI.getTestPlanAutomationMacro(), flg)
    # doTestStepToClickCheckFinalindow()


def clickVSMCheckRPTFinalindow(wb, flag, sfName='', ver=''):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")
    #ribbon = wb.macro("TPMacros.TriggerPopupToSubmitCheck")
    #time.sleep(3)
   # ribbon()
    # KMS.maximiseWindow()
    # doVSMclickCheckRPTFinalindow()
    # dtstccfw = threading.Thread(target=doTestStepToClickCheckFinalindow)
    # dtstccfw.start()

    active_window5 = pgw.getActiveWindow()
    time.sleep(5)

    if active_window5 is not None:
        # Get all windows with "Excel" in the title
        excel_windows = pyautogui.getWindowsWithTitle("Excel")
        logging.info("Present Active Window title is = ", active_window5.title)

        for each_excel_window in excel_windows:
            if each_excel_window.title.startswith("Tests_"):
                # Activate the desired window
                each_excel_window.activate()

            else:
                # Minimize the other windows
                each_excel_window.minimize()

        time.sleep(1)
        # Get the screen size
        width_of_screen, height_of_screen = pyautogui.size()
        origin_x = width_of_screen // 2
        origin_y = height_of_screen // 2
        time.sleep(1)
        pyautogui.click(origin_x, origin_y)
        time.sleep(1)
        pyautogui.click(origin_x, origin_y)
        time.sleep(1)
        KMS.mouseClick()
        time.sleep(0.5)
        KMS.mouseClick()
        logging.info("mouse cursor pointed at the center of screen and done click on final window")

        pyautogui.hotkey('alt', 'space')
        time.sleep(1)
        KMS.keyboard.press('c')
        logging.info("Final window closed")

        time.sleep(2)

        # activate other excels
        for other_excel in excel_windows:
            if other_excel.title.startswith("Tests_"):
                other_excel.minimize()
                time.sleep(2)
            else:
                other_excel.activate()
                other_excel.maximize()
                time.sleep(2)
                pyautogui.hotkey('alt', 'space')
                time.sleep(1)
                KMS.keyboard.press('c')
                logging.info("One bilian excel closed")
                time.sleep(6)
                KMS.pressEnter()
                time.sleep(2)

        # pyautogui.hotkey('alt', 'space')
        # time.sleep(1)
        # KMS.keyboard.press('c')
        # logging.info("One final excel (bilian etc but not Tests file) closed")
        # time.sleep(6)
        # KMS.pressEnter()
        # time.sleep(0.5)
        # pyautogui.hotkey('alt', 'space')
        # time.sleep(1)
        # KMS.keyboard.press('c')
        # logging.info("Another final excel (bilian etc but not Tests file) closed")
        # time.sleep(6)
        # KMS.pressEnter()
        if (flag == 1):
            for other_excel in excel_windows:
                if other_excel.title.startswith("Tests_"):
                    # if(other_excel.title.endswith('ss_fiches')!=True) :
                    other_excel.activate()
                    other_excel.maximize()
                    time.sleep(2)
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(1)
                    KMS.keyboard.press('c')
                    logging.info("ss fiche closed")
                    time.sleep(6)
                    KMS.pressEnter()
    else:
        logging.info("No Active Final Window (5)")
    time.sleep(6)
    if (flag == 1):
        SelectWriteTask = ICF.FetchTaskName()
        # macro = EI.getTestPlanAutomationMacro()
        # ICF.FetchTaskName()
        # macro = EI.getTestPlanAutomationMacro()

        # ICF.FetchTaskName()

        taskArch = SelectWriteTask.split("_")[0]

        if taskArch == "F":

            Arch = "BSI"

        else:

            Arch = "VSM"
        home_dir = os.path.expanduser("~")

        p1 = os.path.join(home_dir, "Desktop")
        logging.info("home dir", p1)
        # os.chdir(p1)
        bsi_checkreport_folder1 = "_Output_CHECK_SF\\BSI"
        # if not os.path.exists(bsi_checkreport_folder1):
        #    os.makedirs(bsi_checkreport_folder1)
        vsm_checkreport_folder3 = "_Output_CHECK_SF\\VSM"
        # if not os.path.exists(vsm_checkreport_folder3):
        # os.makedirs(vsm_checkreport_folder3)
        func_ref = ICF.getTaskDetails()[0]['referentiel']
        folder = bsi_checkreport_folder1 if Arch == 'BSI' else vsm_checkreport_folder3
        tpbook = EI.findInputFiles()[1]

        os.chdir(p1)
        logging.info(f"folder+++ {folder}\\Bilan total {sfName.split('.')[0]}.xlsx")
        if (os.path.exists(folder +  r'\Bilan total ' + sfName.split('.')[0] + '.xlsx')):
        # os.rename(folder +  '\\Bilan total ' + sfName.split('.')[0] + '.xlsx',
        #          folder + '\\CHECK_SS_FICHE_Tests' +
        #               tpbook.split('.')[0][:-6] +ver+ '_VSM.xlsx')
            os.rename(f"{folder}\\Bilan total {sfName.split('.')[0]}.xlsx",
                  f"{folder}\\CHECK_SS_FICHE_{tpbook.split('.')[0][:-6]}{int(ver)}_VSM.xlsx")

        logging.info("end$$$$$$$$$$$$$$$$$$$$$$")
        # sfbook.save()
        # sfbook.close()
        # tpbook.close()


def addBSICheckReportTestStep(wb):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")
    ribbon = wb.macro("TPMacros.checkReport")
    time.sleep(3)
    dbcrts = threading.Thread(target=doBSICheckReportTestStep)
    dbcrts.start()
    ribbon()


def doBSICheckReportTestStep():
    # pop up - click on YES
    time.sleep(6)
    x = ICF.FetchTaskName().split('_')[1]
    flg = 0
    for j in pgw.getAllWindows():
        if 'ss_fiches' in j.title:
            flg = 1
    while True:
        active_window = pgw.getActiveWindow()
        if active_window is not None:
            if active_window.title == "Windows Security":
                time.sleep(6)
            else:
                logging.info("No Security Window")
                break
        else:
            logging.info("No Window to enter username and password")
    while True:
        active_window1 = pgw.getActiveWindow()
        if active_window1 is not None:
            if active_window1.title == "Check plan de test":
                logging.info("Active Window 1 Title:", active_window1.title)
                time.sleep(2)
                KMS.leftArrow()
                time.sleep(1)
                KMS.pressEnter()
                logging.info("clicked on Yes of excel popup")
                break
            else:
                logging.info("No Excel Window")
                break
            # else:
            #     logging.info("No Excel Window")
            #     break
        else:
            logging.info("No Window to click on Ok")
    time.sleep(6)
    if (flg == 1):
        for j in range(0, 11):
            KMS.keyboard.press(KMS.Key.tab)
            time.sleep(1)
        time.sleep(4)
        pyautogui.typewrite(x)
        time.sleep(4)
        KMS.keyboard.press(KMS.Key.backspace)
        time.sleep(1)
        KMS.keyboard.press(KMS.Key.delete)
        time.sleep(1)
        pyautogui.typewrite(x[-1])
        time.sleep(4)
        KMS.keyboard.press(KMS.Key.tab)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'space')
        time.sleep(1)
        for j in range(0, 12):
            pyautogui.hotkey('shift', 'tab')
            time.sleep(1)

    # french window to click on first button top
    # select first button from the window for Checkreport
    time.sleep(2)
    KMS.pressEnter()
    # excel popup - click on Yes button - longue
    time.sleep(6)
    KMS.pressEnter()
    i = 0

    while True:
        bsirpt_window1 = pgw.getActiveWindow()
        bsirpt_window1_title = pgw.getActiveWindowTitle()
        if bsirpt_window1 is not None:
            if bsirpt_window1_title == "Microsoft Excel":
                logging.info("BSI RPT Active Window 1 Title:", bsirpt_window1_title)
                time.sleep(1)
                KMS.pressEnter()
                time.sleep(6)
                logging.info("clicked on OK of excel popup")
                i = i + 1
                if (flg == 0):
                    break
                if (i == 2):
                    break
        else:
            logging.info("No Window to click on Ok")

    while True:
        # here browser window to select the output folder
        bsirpt_window2 = pgw.getActiveWindow()
        bsirpt_window2_title = pgw.getActiveWindowTitle()
        if bsirpt_window2 is not None:
            if bsirpt_window2_title == "Browse For Folder":
                logging.info("BSI RPT Active Window 2 Title:", bsirpt_window2_title)
                KMS.keyboard.press(KMS.Key.tab)
                time.sleep(2)
                KMS.keyboard.press(KMS.Key.tab)
                time.sleep(0.5)
                KMS.keyboard.release(KMS.Key.tab)
                time.sleep(0.5)

                # towards _A folder & ts sub folder for output to save on Desktop
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                KMS.downArrow()
                time.sleep(0.5)
                # sub folder - right arrow
                KMS.rightArrow()
                time.sleep(0.5)
                KMS.downArrow()
               # if (flg == 1):
                #    time.sleep(1)
                 #   KMS.downArrow()
                time.sleep(0.5)
                # press enter to select the checkreport op foder
                KMS.pressEnter()
                break
            # elif (bsirpt_window2_title == "Microsoft Excel"):
            #     logging.info("Active Window 2 Title:", bsirpt_window2_title)
            #     time.sleep(1)
            #     KMS.pressEnter()
            #     logging.info("clicked on OK of excel popup")
        else:
            logging.info("No Browser window-2")

    # -------------------------------------
    os.chdir("../Input_Files")
    bsi_input_dir1 = os.getcwd()
    bsi_input_dir_path1 = str(bsi_input_dir1)
    logging.info("List of files in input folder (BSI)", bsi_input_dir_path1)

    # to capture exact vsm global param file
    list_of_bsi_input_files1 = os.listdir(bsi_input_dir_path1)
    logging.info("list of input files :- ", list_of_bsi_input_files1)

    global bsi_global_param_file1
    for bsi_param1 in list_of_bsi_input_files1:
        match_bsi1 = re.match(r"PARAM_Global_BSI_(.*?)\.xlsm$", bsi_param1)

        if match_bsi1:
            bsi_global_param_file1 = str(bsi_param1)
            logging.info("Param global VSM :- ", bsi_param1)

        else:
            logging.info("Param Global file not found in input folder")

    # ----------------------------------

    # # here PARAM global file has to select by manual
    while True:
        # here browser window to select the output folder
        bsirpt_window3 = pgw.getActiveWindow()
        bsirpt_window3_title = pgw.getActiveWindowTitle()
        if bsirpt_window3 is not None:
            if bsirpt_window3_title == "Choix du fichier PARAM." or bsirpt_window3_title == "PARAM GLOBAL Selection":
                logging.info("Active Window 3 Title:", bsirpt_window3_title)
                time.sleep(1)
                # param file location path
                pyautogui.typewrite(bsi_input_dir_path1)
                time.sleep(1)
                KMS.pressEnter()
                # param_vsm_file = "PARAM_Global_BSI_00949_11_00178.xlsm"
                time.sleep(1)
                logging.info("BSI_Global param file is:- ", bsi_global_param_file1)
                # param file name typing
                pyautogui.typewrite(bsi_global_param_file1)
                time.sleep(1)
                KMS.pressEnter()
                time.sleep(1)
                logging.info("PARAM global file for VSM assigned by automation")
                break
        else:
            logging.info("No window-3")

    # C:/PSA popup window appears CLICK OK
    # if(flg==0):
    #     while True:
    #         time.sleep(1)
    #         # here browser window to select the output folder
    #         pgw.getActiveWindow()
    #         bsirpt_window4_title = pgw.getActiveWindowTitle()
    #
    #         # highlight the Test file to process
    #         all_excel_windows = pyautogui.getWindowsWithTitle("Excel")
    #         for test_excel_window in all_excel_windows:
    #             if test_excel_window.title.startswith("Tests_"):
    #                 # Activate the desired window
    #                 test_excel_window.activate()
    #                 test_excel_window.maximize()
    #                 logging.info("BSI RPT Window 4 Title:", bsirpt_window4_title)
    #                 time.sleep(2)
    #                 KMS.pressEnter()
    #                 time.sleep(1)
    #                 break
    #         break
    if (flg == 0):
        while True:
            # here browser window to select the output folder
            active_window4 = pgw.getActiveWindow()
            # logging.info("Active Window 4 Title:", active_window4.title)
            if active_window4 is not None:
                # logging.info("Active Window 4 Title:", active_window4.title)
                if active_window4.title == "Microsoft Excel":
                    logging.info("Active Window 4 Title:", active_window4.title)
                    time.sleep(2)
                    KMS.pressEnter()
                    time.sleep(1)

                    break
            else:
                logging.info("No window -4 of Microsoft Excel")
    # while True:
    #     bsirpt_window5_title = pgw.getActiveWindowTitle()
    #     if(bsirpt_window5_title=='Check plan de test'):
    #         KMS.leftArrow()
    #         time.sleep(1)
    #         KMS.pressEnter()
    #         break
    #  test_excel_window.activate()
    # break
    time.sleep(2)
    if (flg == 0):
        clickBSICheckRPTFinalindow(EI.getTestPlanAutomationMacro(), flg)


def clickBSICheckRPTFinalindow(wb, flag, sfName='', ver=''):
    # ribbon = wb.macro("Gestion_PT.nvx_nom_fiche")
    # ribbon = wb.macro("Gestion_PT.ajout_fiche")

    logging.info("testing last window popup")
    #ribbon = wb.macro("TPMacros.TriggerPopupToSubmitCheck")
    #time.sleep(3)
    #ribbon()
    #time.sleep(3)

    # KMS.maximiseWindow()
    # doClickBSICheckRPTFinalindow()
    # dtstccfw = threading.Thread(target=doTestStepToClickCheckFinalindow)
    # dtstccfw.start()

    time.sleep(2)
    bsirpt_window5 = pgw.getActiveWindow()
    bsirpt_window5_title = pgw.getActiveWindowTitle()

    if bsirpt_window5 is not None:
        # Get all windows with "Excel" in the title
        excel_windows = pyautogui.getWindowsWithTitle("Excel")
        logging.info("Present Active Window title is = ", bsirpt_window5_title)

        for each_excel_window in excel_windows:
            if each_excel_window.title.startswith("Tests_"):
                # Activate the desired window
                each_excel_window.activate()

            else:
                # Minimize the other windows
                each_excel_window.minimize()

        time.sleep(1)
        # Get the screen size
        width_of_screen, height_of_screen = pyautogui.size()
        origin_x = width_of_screen // 2
        origin_y = height_of_screen // 2
        time.sleep(1)
        pyautogui.click(origin_x, origin_y)
        time.sleep(1)
        pyautogui.click(origin_x, origin_y)
        time.sleep(1)
        KMS.mouseClick()
        KMS.mouseClick()
        logging.info("mouse cursor pointed at the center of screen and done click on final window")

        pyautogui.hotkey('alt', 'space')
        time.sleep(1)
        KMS.keyboard.press('c')
        logging.info("Final window closed")

        time.sleep(2)
        # flag=0
        # activate other excels
        for other_excel in excel_windows:
            if other_excel.title.startswith("Tests_"):
                # if(other_excel.title.endswith('ss_fiches')!=True) :
                other_excel.minimize()
                time.sleep(2)
                # else:
                #     flag=1
                #     other_excel.activate()
                #     other_excel.maximize()
                #     time.sleep(2)

            else:
                other_excel.activate()
                other_excel.maximize()
                time.sleep(2)
                pyautogui.hotkey('alt', 'space')
                time.sleep(1)
                KMS.keyboard.press('c')
                logging.info("One bilian excel closed")
                time.sleep(6)
                KMS.pressEnter()
                time.sleep(2)

        # pyautogui.hotkey('alt', 'space')
        # time.sleep(1)
        # KMS.keyboard.press('c')
        # logging.info("One bilian excel closed")
        # time.sleep(6)
        # KMS.pressEnter()
        # time.sleep(1)
        # pyautogui.hotkey('alt', 'space')
        # time.sleep(1)
        # KMS.keyboard.press('c')
        # logging.info("One bilian excel closed")
        # time.sleep(6)
        # KMS.pressEnter()

        if (flag == 1):
            for other_excel in excel_windows:
                if other_excel.title.startswith("Tests_"):
                    # if(other_excel.title.endswith('ss_fiches')!=True) :
                    other_excel.activate()
                    other_excel.maximize()
                    time.sleep(2)
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(1)
                    KMS.keyboard.press('c')
                    logging.info("ss fiche closed")
                    time.sleep(6)
                    KMS.pressEnter()

    else:
        logging.info("No Active Final Window (5)")
    time.sleep(6)
    if (flag == 1):
        SelectWriteTask = ICF.FetchTaskName()
        # macro = EI.getTestPlanAutomationMacro()
        # ICF.FetchTaskName()
        # macro = EI.getTestPlanAutomationMacro()

        # ICF.FetchTaskName()

        taskArch = SelectWriteTask.split("_")[0]

        if taskArch == "F":

            Arch = "BSI"

        else:

            Arch = "VSM"
        home_dir = os.path.expanduser("~")
        p1 = os.path.join(home_dir, "Desktop")
        # os.chdir(p1)
        bsi_checkreport_folder1 = "_Output_CHECK_SF\\BSI"
        # if not os.path.exists(bsi_checkreport_folder1):
        #    os.makedirs(bsi_checkreport_folder1)
        vsm_checkreport_folder3 = "_Output_CHECK_SF\\VSM"
        # if not os.path.exists(vsm_checkreport_folder3):
        # os.makedirs(vsm_checkreport_folder3)
        func_ref = ICF.getTaskDetails()[0]['referentiel']
        folder = bsi_checkreport_folder1 if Arch == 'BSI' else vsm_checkreport_folder3
        tpbook = EI.findInputFiles()[1]
        sfbook = EI.findInputFiles()[6]
        os.chdir(p1)
        if(os.path.exists(folder+"\\"+'Bilan total '+ sfName.split('.')[0]+'.xlsx')):
          os.rename(f"{folder}\\Bilan total {sfName.split('.')[0]}.xlsx",
                  f"{folder}\\CHECK_SS_FICHE_{tpbook.split('.')[0][:-6]}{int(ver)}_BSI.xlsx")        # sfbook.save()
        # sfbook.close()
        # tpbook.close()


def addNewLine(sheet, row, col):
    logging.info("Adding New Row ==>", row)
    coord = row, col
    sheet.range(coord).select()
    KMS.addNewRow()
    time.sleep(1)


# def selectTPInitModify():
#     left, top, width, height = KMS.getCoordinatesByImage("../images/Tp_Init_modify.png")
#     KMS.moveMouse(left + 20, top + 10)
#     time.sleep(1)
#     KMS.mouseClick()


# def selectTPFinal():
#     left, top, width, height = KMS.getCoordinatesByImage("../images/tp_final.png")
#     KMS.moveMouse(left + 10, top + 10)
#     time.sleep(1)
#     KMS.mouseClick()

# def selectTestSheet():
#     left, top, width, height = KMS.getCoordinatesByImage("../images/test_sheet.png")
#     KMS.moveMouse(left + 10, top + 10)
#     time.sleep(1)
#     KMS.mouseClick()
