import ExcelInterface as EI
import pyautogui

import time
import xlwings as xw
import threading
import InputConfigParser as ICF
import os
import pygetwindow as pgw
import pyautogui
import re
import json
import time
from datetime import date
import KeyboardMouseSimulator as KMS
import TestPlanMacros as TPM
import InputConfigParser as ICF
import threading
import pygetwindow as pgw
import InputConfigParser as ICF
import logging

# ICF.loadConfig()


def check_sf_report():
    ICF.loadConfig()
    home_dir = os.path.expanduser("~")
    p1 = os.path.join(home_dir, "Desktop")

    if not os.path.exists(p1):
        os.makedirs(p1)
        logging.info("Desktop folder created!")
    else:
        logging.info("Desktop folder already exists!")

    os.chdir(p1)
    bsi_checkreport_folder1 = "_Output_CHECK_SF\\BSI"
    if not os.path.exists(bsi_checkreport_folder1):
        os.makedirs(bsi_checkreport_folder1)
    vsm_checkreport_folder3 = "_Output_CHECK_SF\\VSM"
    if not os.path.exists(vsm_checkreport_folder3):
        os.makedirs(vsm_checkreport_folder3)


    output_dir_path = os.path.join(ICF.getOutputFiles())
    os.chdir(output_dir_path)


#try:
    tpBook = EI.openTestPlan()

    sheet = tpBook.sheets['Sommaire']
    version = EI.getDataFromCell(sheet, 'A6')
    tes = "(?:BSI|VSM)[0-9]{2}_SF_[0-9]{2}_[0-9]{2}_[0-9]{4}[A-Z]?"
    flag = 0
    for i in tpBook.sheet_names:
       # sf = re.findall(tes, i)
        if i.find('_SF_') != -1:
            flag = 1
            break
        else:
            pass
    if (flag == 0):
        return 0

    # tpBook.close()
    tpBook.save()
    tpBook.close()
    sfBook = EI.openSousFiches()
    sf_name = sfBook.name
    # sfBook.activate()
    KMS.showWindow(sfBook.name.split('.')[0])
    SelectWriteTask = ICF.FetchTaskName()
    macro = EI.getTestPlanAutomationMacro()

    taskArch = SelectWriteTask.split("_")[0]

    if taskArch == "F":
        Arch = "BSI"
    else:
        Arch = "VSM"

    time.sleep(2)
    TPM.selectArch(macro)

    # call check function
    if Arch == 'BSI':
        TPM.addBSICheckReportTestStep(macro)
        TPM.clickBSICheckRPTFinalindow(macro, 1, sf_name, version)
    else:
        TPM.addVSMCheckReportTestStep(macro)
        TPM.clickVSMCheckRPTFinalindow(macro, 1, sf_name, version)

    logging.info("opening test plan")
    logging.info(f"ICF.getInputFolder() {ICF.getInputFolder()}")
    input_path = os.path.join(ICF.getInputFolder())
    logging.info(f"input_path {input_path}")
    os.chdir(input_path)
    time.sleep(3)
    tpBook = EI.openTestPlan()
    tpBook.activate()
    for sheet in tpBook.sheets:
        if ('SF' in sheet.name):
            value = EI.getDataFromCell(sheet, 'C7')
            if (value == 'EN COURS'):
                 return 0

    macro = EI.getTestPlanAutomationMacro()
    t1 = threading.Thread(target=synch_aut)
    t1.start()
    TPM.synchronizeSubSheet(macro)
    #except Exception as e:
    #  exc_type, exc_obj, exc_tb = sys.exc_info()
    #  exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
     # logging.info(
      #    f"\nSomething went wrong in processing check report: {e} line no. {exc_tb.tb_lineno} file name: {exp_fname}********************")


def synch_aut():
    flg=0
    time.sleep(6)
    KMS.pressEnter()
    time.sleep(4)
    # KMS.pressEnter()
    output_dir_path = os.path.join(ICF.getOutputFiles())
    os.chdir(output_dir_path)
    os.chdir("..\Input_Files")
    #bsi_input_dir1 = os.getcwd()
    bsi_input_dir1=ICF.getInputFolder()
    bsi_input_dir_path1 = str(bsi_input_dir1)
    logging.info("List of files in input folder (BSI)", bsi_input_dir_path1)
    # os.chdir("..")
    while True:
        # here browser window to select the output folder
        bsirpt_window3 = pgw.getActiveWindow()
        bsirpt_window3_title = pgw.getActiveWindowTitle()
        logging.info(bsirpt_window3_title)
        if bsirpt_window3 is not None:
            if (bsirpt_window3_title == "Check Test book"):
              time.sleep(2)
              KMS.pressEnter()
              break
            else:
              break

    while True:
        # here browser window to select the output folder
        bsirpt_window2 = pgw.getActiveWindow()
        bsirpt_window2_title = pgw.getActiveWindowTitle()
        logging.info(bsirpt_window2_title)
        if bsirpt_window2 is not None:
            if bsirpt_window2_title == "Choose the Sub-sheets file !" or bsirpt_window2_title=="SubSheets TestBook Selection" :
                logging.info("Active Window 5 Title:", bsirpt_window2_title)
                time.sleep(6)
                # param file location path
                pyautogui.typewrite(bsi_input_dir_path1)
                time.sleep(4)
                KMS.pressEnter()
                # param_vsm_file = "PARAM_Global_BSI_00949_11_00178.xlsm"
                time.sleep(1)
                # logging.info("BSI_Global param file is:- ", )
                # param file name typing
                pyautogui.typewrite(EI.findInputFiles()[6])
                time.sleep(1)
                KMS.pressEnter()
                time.sleep(1)
                logging.info("Sous file assigned by automation")
                break
            elif (bsirpt_window2_title=="Microsoft Excel"):
                flg=1
                time.sleep(1)
                KMS.rightArrow()
                time.sleep(1)
                KMS.pressEnter()
                time.sleep(2)
                break

        else:
            logging.info("No window-5")
    while True:
        # here browser window to select the output folder
        bsirpt_window4 = pgw.getActiveWindow()
        bsirpt_window4_title = pgw.getActiveWindowTitle()
        if bsirpt_window4 is not None:
            if bsirpt_window4_title == "Microsoft Excel":
                time.sleep(1)
                KMS.rightArrow()
                time.sleep(1)
                KMS.pressEnter()
                time.sleep(2)
                break
        else:
            logging.info("No window-6")

