import datetime
import ExcelInterface as EI
import re
import logging
date_time = datetime.datetime.now()
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil
import os
import InputConfigParser as ICF
#from docx import Document
import docx
import sys
import tkinter as tk
from tkinter import scrolledtext
#from docx.enum.text import WD_BREAK


def configChromeVersion():
    service = Service(executable_path=ICF.getWebDriver())
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    try:
        chrome_driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"There is problem in connecting with chrome please check the chrome version and update the chrome driver with latest version: {type(e).__name__}")
        return -1


def docInfo(refnum, reqVer="", newSSdFolder=False,docType=('doc',)):
    # service = Service(executable_path=ICF.getWebDriver())
    # driver = webdriver.Chrome(service=service)
    username, password = ICF.getCredentials()
    driver.get(
        'https://' + username + ':' + password + '@docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')
    # page = driver.get('https://SC38437:asd321AK@docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')

    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: driver.execute_script('return document.readyState')=='complete')

    expand_menu = driver.find_element(By.ID, "ext-gen24")  # sidebar
    time.sleep(1)
    expand_menu.click()

    time.sleep(1)
    ref_search = driver.find_element(By.ID, "txtRef")  # text bar
    ref_search.click()
    time.sleep(1)
    ref_search.send_keys(refnum)
    time.sleep(1)
    ref_search.send_keys(Keys.RETURN)

    wait.until(lambda driver: driver.execute_script('return document.readyState')=='complete')
    # clickVerTab()
    content = driver.find_elements(By.CLASS_NAME, "x-tab-strip-text")  # versions switch
    for i in content:
        if i.text=="Versions":
          #  print(i.text)
            i.click()
    # waitUntilElementFinishedLoading()
    time.sleep(5)
    try:
        if reqVer!="":
            docdownload(refnum, reqVer, newSSdFolder,docType)
        else:
            lst_ver = []
            version_element = driver.find_elements(By.XPATH, (
                "//div[@class='x-grid3-cell-inner x-grid3-col-docVersion']"))  # version table
            for x in version_element:
                actualVer = (float(x.get_attribute('innerHTML')))
                lst_ver.append(actualVer)
           # print(lst_ver)
            ver_num = []
            for num in lst_ver:
                val_ver = str(num).split('.')
                if int(val_ver[1])==0:
                    ver_num.append(float(val_ver[0]))
            docdownload(refnum, max(ver_num), newSSdFolder,docType)
           # print(max(ver_num))
    except:
        pass


def docdownload(refnum, reqVer, newSSdFolder=False,docType=('doc',)):
    table_body = driver.find_elements(By.CLASS_NAME, "x-grid3-scroller")  # slide bar
    wait = WebDriverWait(driver, 15)
   # print("table_body------>", table_body)
    for tbody in table_body:
        # print(tbody)
        # time.sleep(2)
        if not tbody.find_elements(By.CLASS_NAME,
                                   "x-grid3-col x-grid3-cell x-grid3-td-docViewIcon x-grid3-cell-first "):
            tables = tbody.find_elements(By.TAG_NAME, "table")
            for each_table in tables:
                try:
                    t_body = each_table.find_element(By.TAG_NAME, "tbody")
                    each_row = t_body.find_element(By.TAG_NAME, "tr")
                    all_data = each_row.find_elements(By.TAG_NAME, "td")
                    lst_elements = []
                    txt_elements = []
                    for each_data in all_data:
                        # time.sleep(2)
                        ver_txt = each_data
                        lst_elements.append(ver_txt)
                        txt_elements.append(ver_txt.text)
                  #  print(txt_elements)
                    if len(lst_elements) >= 3:
                        if str(float(reqVer))==lst_elements[0].text:
                          #  print("Version matched")
                            url = lst_elements[2].find_element(By.TAG_NAME, "a")
                            a = EC.element_to_be_clickable(url)
                            wait.until(a).click()
                            count = 0
                            # after the click
                            attachment = wait.until(EC.element_to_be_clickable((By.ID, "attachementPanel")))
                            attached_elements = attachment.find_elements(By.TAG_NAME, ('a'))
                            toggle = attachment.find_elements(By.XPATH, "//img[@title='Click to collapse.']")

                            # words = ["xls", "doc", "rtf"]
                            words = docType

                            for attached in attached_elements[0:]:
                                if attached not in toggle:
                                 #   print('file_name', attached.text)
                                    for w in words:
                                        if w in attached.text:
                                            count = count + 1
                                            attached.click()
                                            # download_wait(ICF.getDownloadFolder())

                            # time.sleep(30)
                            sortFiles(count, reqVer, refnum, newSSdFolder)
                            driver.close()
                            break
                        #print("Version not matched")
                except exceptions.StaleElementReferenceException as e:
                    print(e)


def moveFile(x, newSSdFolder=False):
    if newSSdFolder is True:
        dst = ICF.getDicFolder()
    else:
        dst = ICF.getInputFolder()
    src = ICF.getDownloadFolder()
    sorc = (os.path.join(src, x))
    dset = (os.path.join(dst, x))
   # print('src====>', sorc)
    #print('dst====>', dset)
    #print("Before moving file:")
   # print(os.listdir(dst))
    if os.path.isfile(dset):
        os.remove(dset)
       # print(x, 'deleted in', dst)
    dest = shutil.move(sorc, dset)
    #print('moved file:', dest)
    #print('file Moved successfully')


def sortFiles(count, reqVer, refnum, newSSdFolder=False):
   # print("In sortFiles function - ")
    while any(filename.endswith('.crdownload') or filename.endswith('.tmp') for filename in
              os.listdir(ICF.getDownloadFolder())):
        time.sleep(1)
    a = -1
    path = ICF.getDownloadFolder()
    dst = ICF.getInputFolder()
    name_list = os.listdir(path)
    full_list = [os.path.join(path, i) for i in name_list]
    time_sorted_list = sorted(full_list, key=os.path.getmtime)
    sorted_filename_list = [os.path.basename(i) for i in time_sorted_list]
    sorted_filename_list.reverse()

    for x in (sorted_filename_list[0:count]):

        if (os.path.splitext(x)[1]==".docx") or (os.path.splitext(x)[1]==".doc") or (
                os.path.splitext(x)[1]==".docm") or (os.path.splitext(x)[1]==".rtf"):
            if refnum not in x:
               # print(refnum)
                newFile = '[' + 'V' + str(float(reqVer)) + ']' + '[' + str(refnum) + ']' + x
            else:
                newFile = '[' + 'V' + str(float(reqVer)) + ']' + x
          #  print("- ", newFile, path + '\\' + x, path + '\\' + newFile)
            try:
                os.rename(path + '\\' + x, path + '\\' + newFile)
            except:
                os.remove(path + '\\' + newFile)
                time.sleep(3)
                os.rename(path + '\\' + x, path + '\\' + newFile)
           # print("f", newFile)
            moveFile(newFile, newSSdFolder)
            a = dst + '\\' + newFile
        else:
           # print("x", x)
            moveFile(x, newSSdFolder)
            # a = path + '\\' + x
    return a


# docInfo("01272_22_00311")
# these function is used for to remove the .crdownload, .crdownloads , .temp and .tmp files in the download folder
def delete_files_with_extensions():
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    extensions = [".crdownload", ".crdownloads", ".temp", ".tmp"]
    # Traverse through the files in the Downloads folder and its sub-folders
    for root, _, files in os.walk(downloads_path):
        # For each file, check if its extension is in the list of extensions to delete
        for file in files:
            _, file_extension = os.path.splitext(file)
            if file_extension.lower() in extensions:
                # If the file's extension matches, delete the file
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                   # print(f"Deleted: {file_path}")
                except Exception as e:
                    print(f"Error deleting {file_path}: {e}")


def startDocumentDownload(fepsRefVers, allowDownload=True, newSSdFolder=False,docType=('doc',)):
   # print("startDocumentDownload +", fepsRefVers)
    fepsRefVers = [item for item in fepsRefVers if item != (None, '')]
    fepsRefVers = [[item[0].strip(), item[1]] for item in fepsRefVers]
   # print("after remove none elements--->", fepsRefVers)
    if allowDownload:
       # print("Allowed", allowDownload)
        for fepdDoc in fepsRefVers:
           # print("Startdownload = ", fepdDoc)
            ref, ver = fepdDoc
            if ref!="":
                delete_files_with_extensions()
                service = Service(executable_path=ICF.getWebDriver())
                options = webdriver.ChromeOptions()
                options.add_experimental_option('excludeSwitches', ['enable-logging'])
                global driver
                driver = webdriver.Chrome(service=service, options=options)
               # print("--", ref, ver)
                docInfo(ref, ver.replace("V", ""), newSSdFolder,docType)
                time.sleep(3)
            else:
                print("reference number is empty")


def getCellAbsVal(sheet, row, col):
    for i in range(row, 0, -1):
        cellVal = EI.getDataFromCell(sheet, f"{col}{i}")
        if cellVal is not None:
            return str(cellVal)
    return None


def downloadSSD(tpBook):
    global summarySheet
    try:
        summarySheet = tpBook.sheets["Sommaire"]
    except Exception as e:
        print(f"TestPlan or Sommaire sheet not found! in Input Path.")
        exit(1)
    referenceList = []
    # Define the document types you want to check for
    document_types = ["ssd", "eead", "nt", "tfd"]
    # get all reference numbers for ssd files
    nrows = summarySheet.used_range.last_cell.row
    for i in range(6, nrows):
        typeVal = EI.getDataFromCell(summarySheet, f"E{i}")
        for a in document_types:
           # print("a-------->",a)
            # if typeVal is not None and re.match("ssd", typeVal.lower()):
            if typeVal is not None and re.match(a, typeVal.lower()):
               # print("typeVal---->",typeVal)
                referenceNumber = getCellAbsVal(summarySheet, i, "F")
                refver = getCellAbsVal(summarySheet, i, "G")
               # print('refver-->', refver)
                referenceList.append(referenceNumber)
   # print("referenceList --> ", referenceList)

    # remove duplicate elements from list
    referenceList = [*(set(referenceList))]

   # print("After Duplicate removal referenceList --> ", referenceList)
    # iterate over all reference numbers and download respective documents
    for referenceNum in referenceList:
        startDocumentDownload([[referenceNum, ""]], True, False)


# def getInternalflow():
def download_documents(tpBook, docType=('doc',), document=('dci',), newssdFolder=False):
    global summarySheet
    try:
        summarySheet = tpBook.sheets["Sommaire"]
    except Exception as e:
        print(f"TestPlan or Sommaire sheet not found! in Input Path.{type(e).__name__}")
        exit(1)
    referenceList = []
    # Define the document types you want to check for
    # document = ["dci"]
    # get all reference numbers for ssd files
    nrows = summarySheet.used_range.last_cell.row
    for i in range(6, nrows):
        typeVal = EI.getDataFromCell(summarySheet, f"E{i}")
        for a in document:

            # if typeVal is not None and re.match("ssd", typeVal.lower()):
            if typeVal is not None and re.match(a, typeVal.lower()):
                referenceNumber = getCellAbsVal(summarySheet, i, "F")
                refver = getCellAbsVal(summarySheet, i, "G")

                referenceList.append((referenceNumber, refver))

    referenceList = [*(set(referenceList))]
    print("referenceList--------->",referenceList)

    # iterate over all reference numbers and download respective documents
    for referenceNum, refver in referenceList:
        startDocumentDownload([[referenceNum, refver]], True, newssdFolder, docType)


if __name__=="__main__":
    ICF.loadConfig()
    #ain()

# C:\Users\vgajula\Documents\17-11-2023\Input_Files_Supporting_Req\Keyword_file_folder
# BRAKE_REQUEST

# if __name__ == '__main__':
#     ICF.loadConfig()
#     PT = EI.findInputFiles()[1]
#     print("PT---->",PT)
#     tpBook = EI.openExcel(ICF.getInputFolder() + "\\" + PT)
#     DCIWB.download_documents(tpBook, ('xls',), ('dci'), True)