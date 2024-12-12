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
import logging
import ExcelInterface as EI

def configChromeVersion():
    service = Service(executable_path=ICF.getWebDriver())
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    try:
        chrome_driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        logging.info(f"There is problem in connecting with chrome please check the chrome version and update the chrome driver with latest version: {type(e).__name__}")
        return -1


def docInfo(refnum, reqVer="", newSSdFolder=False):
    # service = Service(executable_path=ICF.getWebDriver())
    # driver = webdriver.Chrome(service=service)
    username, password = ICF.getCredentials()
    driver.get('https://' + username + ':' + password + '@docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')
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
            logging.info(i.text)
            i.click()
    # waitUntilElementFinishedLoading()
    time.sleep(5)
    if reqVer!="":
        docdownload(refnum, reqVer, newSSdFolder)
    else:
        lst_ver = []
        version_element = driver.find_elements(By.XPATH, (
            "//div[@class='x-grid3-cell-inner x-grid3-col-docVersion']"))  # version table
        for x in version_element:
            actualVer = (float(x.get_attribute('innerHTML')))
            lst_ver.append(actualVer)
        logging.info(lst_ver)
        ver_num = []
        for num in lst_ver:
            val_ver = str(num).split('.')
            if int(val_ver[1])==0:
                ver_num.append(float(val_ver[0]))
        docdownload(refnum, max(ver_num), newSSdFolder)
        logging.info(max(ver_num))

def docInfoDossier(refnum, reqVer=""):
    # service = Service(executable_path=ICF.getWebDriver())
    # driver = webdriver.Chrome(service=service)
    isdossier = 0
    username, password = ICF.getCredentials()
    driver.get(
        'https://' + username + ':' + password + '@docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')
    # page = driver.get('https://SC38437:asd321AK@docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')

    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')

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

    wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
    # clickVerTab()
    content = driver.find_elements(By.CLASS_NAME, "x-tab-strip-text")  # versions switch
    for i in content:
        if i.text == "Dossier":
            isdossier = 1
            logging.info(i.text)
            i.click()
    # waitUntilElementFinishedLoading()
    time.sleep(5)
    refNumer = ""
    logging.info("isdossier >> ", isdossier)
    if isdossier != 1:
        logging.info("^^^^^^^^^^^^^^")
        return -1, refNumer

    downloadedFile, refNumer = QIAdocdownload(refnum, reqVer)
    logging.info("+++++++++++++++++++++ ", downloadedFile)
    return downloadedFile, refNumer

def waitUntilElementFinishedLoading():
    # table = driver.find_attribute("ext-gen401")
    table = driver.find_element(By.ID, value='ext-gen401')

    num_children = len(table.find_elements_by_xpath("./*"))
    time.sleep(1)
    while True:
        if num_children!=len(table.find_elements_by_xpath("./*")):
            time.sleep(1)
        else:
            break


def docdownload(refnum, reqVer, newSSdFolder=False):
    table_body = driver.find_elements(By.CLASS_NAME, "x-grid3-scroller")  # slide bar
    wait = WebDriverWait(driver, 15)
    for tbody in table_body:
        # logging.info(tbody)
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
                    logging.info(txt_elements)
                    if len(lst_elements) >= 3:
                        if str(float(reqVer))==lst_elements[0].text:
                            logging.info("Version matched")
                            url = lst_elements[2].find_element(By.TAG_NAME, "a")
                            a = EC.element_to_be_clickable(url)
                            # a = driver.get('http://docinfogroupe.inetpsa.com/ead/doc/ref.'+refnum+'/v.vc/fiche')
                            wait.until(a).click()
                            count = 0
                            # after the click
                            attachment = wait.until(EC.element_to_be_clickable((By.ID, "attachementPanel")))
                            attached_elements = attachment.find_elements(By.TAG_NAME, ('a'))
                            toggle = attachment.find_elements(By.XPATH, "//img[@title='Click to collapse.']")

                            # words = ["xls", "doc", "rtf"]
                            words = ["doc", "rtf"]

                            for attached in attached_elements[0:]:
                                if attached not in toggle:
                                    logging.info('file_name', attached.text)
                                    for w in words:
                                        if w in attached.text:
                                            count = count + 1
                                            attached.click()
                                            # download_wait(ICF.getDownloadFolder())

                            # time.sleep(30)
                            sortFiles(count, reqVer, refnum, newSSdFolder)
                            driver.close()
                            break
                        logging.info("Version not matched")
                except exceptions.StaleElementReferenceException as e:
                    logging.info(e)


# def moveFile(x):
#     # time.sleep(15)
#     # src = r"C:\Users\pragupathy\Downloads"
#     # dst = r"C:\Users\pragupathy\Downloads\input"
#     src = ICF.getDownloadFolder()
#     dst = ICF.getInputFolder()
#     sorc = src + '\\' + x
#     dest = dst + '\\' + x
#     shutil.move(sorc, dest)


# def register(flag):
#     return flag

def QIAdocdownload(refnum, reqVer):
    table_body = driver.find_elements(By.CLASS_NAME, "x-grid3-scroller") #slide bar
    wait = WebDriverWait(driver, 15)
    downloadedFileName = ""
    refNumer = ""
    for tbody in table_body:
        # logging.info(tbody)
        # time.sleep(2)
        if not tbody.find_elements(By.CLASS_NAME,
                                   "x-grid3-col x-grid3-cell x-grid3-td-docTitle  "):
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
                    logging.info(txt_elements, "txt_elements")
                    if len(txt_elements) != '':
                        if txt_elements[4].find("QIA") != -1 or txt_elements[4].find("IQ") != -1:
                            logging.info("txt_elements[4] >> ", txt_elements[4])
                            downloadedFileName = txt_elements[4]
                            refNumer = txt_elements[1]
                            logging.info("txt_elements[1] ", txt_elements[1])
                            url = lst_elements[4].find_element(By.TAG_NAME, "a")
                            a = EC.element_to_be_clickable(url)
                            wait.until(a).click()
                            count = 0
                            # after the click
                            attachment = wait.until(EC.element_to_be_clickable((By.ID, "attachementPanel")))
                            attached_elements = attachment.find_elements(By.TAG_NAME, ('a'))
                            toggle = attachment.find_elements(By.XPATH, "//img[@title='Click to collapse.']")

                            words = ["xls"]

                            for attached in attached_elements[0:]:
                                if attached not in toggle:
                                    logging.info(attached.text)
                                    for w in words:
                                        if w in attached.text:
                                            count = count + 1
                                            attached.click()
                                            # download_wait(ICF.getDownloadFolder())

                            # time.sleep(30)
                            time.sleep(5)
                            downloadedFileList = sortFilesQIA(count, reqVer, refnum)
                            logging.info("downloadedFileList >> ", downloadedFileList)
                            driver.close()
                            break
                except exceptions.StaleElementReferenceException as e:
                    logging.info(e)

    return downloadedFileList, refNumer


def moveFile(x, newSSdFolder=False):
    if newSSdFolder is True:
        dst = ICF.getSsdFolder()
    else:
        dst = ICF.getInputFolder()
    src = ICF.getDownloadFolder()
    sorc = (os.path.join(src, x))
    dset = (os.path.join(dst, x))
    logging.info('src====>', sorc)
    logging.info('dst====>', dset)
    logging.info("Before moving file:")
    logging.info(os.listdir(dst))
    if os.path.isfile(dset):
        os.remove(dset)
        logging.info(x, 'deleted in', dst)
    dest = shutil.move(sorc, dset)
    logging.info('moved file:', dest)
    logging.info('file Moved successfully')


def renameFiles(reqVer):
    # dst = r"C:\Users\pragupathy\Downloads\input"
    # or (os.path.splitext(f)[1] == ".docx")
    dst = ICF.getInputFolder()
    words = ["xls", "doc", "rtf"]
    for f in os.listdir(dst):
        try:
            if str('[') not in f:
                for w in words:
                    if w in os.path.splitext(f)[1]:
                        newFile = '[' + 'V' + str(int(reqVer)) + ']' + f
                        os.rename(dst + '\\' + f, dst + '\\' + newFile)
                        logging.info("f", newFile)
            else:
                logging.info("Version no. already added")
        except WindowsError:
            logging.info("File not Found")
    return dst + "/" + newFile


def download_wait(path_to_downloads):
    # this function is not needed now
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(30)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds


def sortFiles(count, reqVer, refnum, newSSdFolder=False):
    logging.info("In sortFiles function - ")
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
                logging.info(refnum)
                newFile = '[' + 'V' + str(float(reqVer)) + ']' + '[' + str(refnum) + ']' + x
            else:
                newFile = '[' + 'V' + str(float(reqVer)) + ']' + x
            logging.info("- ", newFile, path + '\\' + x, path + '\\' + newFile)
            try:
                os.rename(path + '\\' + x, path + '\\' + newFile)
            except:
                os.remove(path + '\\' + newFile)
                time.sleep(3)
                os.rename(path + '\\' + x, path + '\\' + newFile)
            logging.info("f", newFile)
            moveFile(newFile, newSSdFolder)
            a = dst + '\\' + newFile
        else:
            logging.info("x", x)
            moveFile(x, newSSdFolder)
            # a = path + '\\' + x
    return a

def sortFilesQIA(count, reqVer, refnum):
    logging.info("In sortFiles function QIA - ")
    while any(filename.endswith('.crdownload') for filename in os.listdir(ICF.getDownloadFolder())):
        time.sleep(1)
    a = -1
    movedFile = []
    path = ICF.getDownloadFolder()
    dst = ICF.getInputFolder()
    name_list = os.listdir(path)
    full_list = [os.path.join(path, i) for i in name_list]
    time_sorted_list = sorted(full_list, key=os.path.getmtime)
    sorted_filename_list = [os.path.basename(i) for i in time_sorted_list]
    sorted_filename_list.reverse()

    for x in (sorted_filename_list[0:count]):
        if(os.path.splitext(x)[1] != ".docx") or (os.path.splitext(x)[1] != ".doc") or (
                os.path.splitext(x)[1] != ".docm") or (os.path.splitext(x)[1] != ".rtf"):
            logging.info("x -- ", x)
            movedFile.append(x)
            moveFile(x)
    return movedFile

def getReference():
    reference = []
    for ref in ICF.getDocToDownload():
        reference.append(ref["Reference"])
        logging.info("Reference = " + str(reference))
    return reference


def getVersion():
    version = []
    for ref in ICF.getDocToDownload():
        version.append(ref["Version"])
        logging.info("Version = " + str(version))
    return version



def getCellAbsVal(sheet, row, col):
    for i in range(row, 0, -1):
        cellVal = EI.getDataFromCell(sheet, f"{col}{i}")
        if cellVal is not None:
            return str(cellVal)
    return None


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
                    logging.info(f"Deleted: {file_path}")
                except Exception as e:
                    logging.info(f"Error deleting {file_path}: {e}")


def startDocumentDownload(fepsRefVers, allowDownload=True, newSSdFolder=False):
    print("startDocumentDownload +", fepsRefVers)
    fepsRefVers = [item for item in fepsRefVers if item != (None, '')]
    fepsRefVers = [[item[0].strip(), item[1]] for item in fepsRefVers]
    logging.info("after remove none elements--->", fepsRefVers)
    if allowDownload:
        logging.info("Allowed", allowDownload)
        for fepdDoc in fepsRefVers:
            logging.info("Startdownload = ", fepdDoc)
            ref, ver = fepdDoc
            if ref!="":
                delete_files_with_extensions()
                service = Service(executable_path=ICF.getWebDriver())
                options = webdriver.ChromeOptions()
                options.add_experimental_option('excludeSwitches', ['enable-logging'])
                global driver
                driver = webdriver.Chrome(service=service, options=options)
                logging.info("--", ref, ver)
                docInfo(ref, ver.replace("V", ""), newSSdFolder)
                time.sleep(3)
            else:
                logging.info("reference number is empty")


def startDocumentDownloadFilesFromDossier(fepsRefVers, allowDownload=True):
    logging.info("startDocumentDownload +", fepsRefVers)
    dwnldFile = ""
    refNumer = ""
    if allowDownload:
        logging.info("Allowed", allowDownload)
        for fepdDoc in fepsRefVers:
            logging.info("Startdownload = ", fepdDoc)
            ref, ver = fepdDoc
            if ref != "":
                service = Service(executable_path=ICF.getWebDriver())
                options = webdriver.ChromeOptions()
                options.add_experimental_option('excludeSwitches', ['enable-logging'])
                global driver
                driver = webdriver.Chrome(service=service, options=options)
                logging.info("--", ref, ver)
                dwnldFile, refNumer = docInfoDossier(ref, ver.replace("V", ""))
                logging.info("dwnldFiledwnldFiledwnldFile +++++++++ ", dwnldFile)
                return dwnldFile, refNumer
            else:
                logging.info("reference number is empty")
                return -1, refNumer


if __name__=="__main__":
    # ICF.loadConfig()
    startDocumentDownload([("01272_18_01377", "3.0")])

    # [("21074_23_00186", "1.0")]


