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
import re
import InputConfigParser as ICF
import logging
# import json


ICF.loadConfig()
# credentials = ["SC58337","Zoro1599"]
# downloadFolder = r"C:\Users\yjagtap\Downloads"
# destinationFolder = r"C:\Users\yjagtap\Desktop\QIA\docs"
# docToDownload = []
# webDriverAddr = r"C:\Users\yjagtap\Desktop\QIA\chromedriver_win32\chromedriver.exe"

# credentials = [input("Enter username of DocInfogroup: "),
#                input("Enter password of DocInfogroup: ")]
# # The above code is commented out, so it is not doing anything. It appears to be setting the value
# of the `downloadFolder` variable to the result of a function call to `ICF.getDownloadFolder()`.
# The code also includes commented out lines that suggest there may be additional functionality
# related to a `destinationFolder`, a list of `docToDownload`, and a `webDriverAddr`.
downloadFolder = ICF.getDownloadFolder()
destinationFolder = ICF.getInputFolder()
outputFolder = ICF.getOutputFiles()
docToDownload = []
webDriverAddr = ICF.getWebDriver()


def docInfo(refnum, reqVer="", words=["xls"]):
    # service = Service(executable_path=ICF.getWebDriver())
    # driver = webdriver.Chrome(service=service)
    username, password = ICF.getCredentials()
    # logging.info(username, password)
    driver.get(
        'https://' + username + ':' + password + '@docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')
    # page = driver.get('https://SC38437:asd321AK@docinfogroupe.psa-peugeot-citroen.com/ead/accueil_init.action')

    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: driver.execute_script(
        'return document.readyState') == 'complete')

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

    wait.until(lambda driver: driver.execute_script(
        'return document.readyState') == 'complete')
    # clickVerTab()
    content = driver.find_elements(
        By.CLASS_NAME, "x-tab-strip-text")  # versions switch
    for i in content:
        if i.text == "Versions":
            logging.info(i.text)
            i.click()
    waitUntilElementFinishedLoading()
    if reqVer != "":
        if reqVer == "latest":
            docdownloadLatest(refnum, reqVer, words)
        else:
            docdownload(refnum, reqVer, words)
    else:
        docdownloadAll(refnum, "", words)


def waitUntilElementFinishedLoading():
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'x-grid3-row')))
    time.sleep(1)


def docdownload(refnum, reqVer, words=["xls"]):
    table_body = driver.find_elements(
        By.CLASS_NAME, "x-grid3-scroller")  # slide bar
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
                        if str(float(reqVer)) == lst_elements[0].text:
                            logging.info("Version matched")
                            url = lst_elements[2].find_element(
                                By.TAG_NAME, "a")
                            a = EC.element_to_be_clickable(url)
                            wait.until(a).click()
                            count = 0
                            # after the click
                            attachment = wait.until(
                                EC.element_to_be_clickable((By.ID, "attachementPanel")))
                            attached_elements = attachment.find_elements(
                                By.TAG_NAME, ('a'))
                            toggle = attachment.find_elements(
                                By.XPATH, "//img[@title='Click to collapse.']")

                            for attached in attached_elements[0:]:
                                if attached not in toggle:
                                    logging.info(attached.text)
                                    for w in words:
                                        if w in attached.text:
                                            count = count + 1
                                            attached.click()
                                            # download_wait(downloadFolder)

                            # time.sleep(30)
                            sortFiles(count, reqVer, refnum)
                            driver.close()
                            break
                        logging.info("Version not matched")
                except exceptions.StaleElementReferenceException as e:
                    logging.info(e)


def docdownloadLatest(refnum, reqVer, words=["xls"]):

    table_body = driver.find_elements(
        By.CLASS_NAME, "x-grid3-scroller")  # slide bar
    wait = WebDriverWait(driver, 15)
    linkList = []
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
                    # logging.info(txt_elements)
                    if len(lst_elements) >= 3:
                        url = lst_elements[2].find_element(
                            By.TAG_NAME, "a").get_attribute('href')
                        linkList.append(url)
                except exceptions.StaleElementReferenceException as e:
                    logging.info(e)

    logging.info("total links:", len(linkList))

    link = linkList[0]
    logging.info(f"downloading latest version of {refnum}, {link}, {linkList}")
    driver.get(link)
    count = 0
    # after the click
    wait.until(lambda driver: driver.execute_script(
        'return document.readyState') == 'complete')
    attachment = wait.until(
        EC.element_to_be_clickable((By.ID, "attachementPanel")))
    attached_elements = attachment.find_elements(By.TAG_NAME, ('a'))
    toggle = attachment.find_elements(
        By.XPATH, "//img[@title='Click to collapse.']")

    version = re.findall(r'\d+\.\d+', link)
    if reqVer == 'latest':
        version = "0.0"
    elif len(version) > 0:
        version = version[0]
    else:
        version = "0.0"

    for attached in attached_elements[0:]:
        if attached not in toggle:
            logging.info(attached.text)
            for w in words:
                if w in attached.text:
                    count = count + 1
                    attached.click()
                    # download_wait(downloadFolder)

    # time.sleep(30)
    sortFiles(count, version, refnum)


def docdownloadAll(refnum, reqVer, words=["xls"]):
    table_body = driver.find_elements(
        By.CLASS_NAME, "x-grid3-scroller")  # slide bar
    wait = WebDriverWait(driver, 15)
    linkList = []
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
                    # logging.info(txt_elements)
                    if len(lst_elements) >= 3:
                        url = lst_elements[2].find_element(
                            By.TAG_NAME, "a").get_attribute('href')
                        linkList.append(url)
                except exceptions.StaleElementReferenceException as e:
                    logging.info(e)

    logging.info("total links:", len(linkList))
    for link in linkList:
        driver.get(link)
        count = 0
        # after the click
        wait.until(lambda driver: driver.execute_script(
            'return document.readyState') == 'complete')
        attachment = wait.until(
            EC.element_to_be_clickable((By.ID, "attachementPanel")))
        attached_elements = attachment.find_elements(By.TAG_NAME, ('a'))
        toggle = attachment.find_elements(
            By.XPATH, "//img[@title='Click to collapse.']")

        version = re.findall(r'\d+\.\d+', link)

        if len(version) > 0:
            version = version[0]
        else:
            version = "0.0"

        for attached in attached_elements[0:]:
            if attached not in toggle:
                logging.info(attached.text)
                for w in words:
                    if w in attached.text:
                        count = count + 1
                        attached.click()
                        # download_wait(downloadFolder)

        # time.sleep(30)
        sortFiles(count, version, refnum)


def moveFile(x):
    # time.sleep(15)
    # src = r"C:\Users\pragupathy\Downloads"
    # dst = r"C:\Users\pragupathy\Downloads\input"
    src = downloadFolder
    dst = destinationFolder
    sorc = src + '\\' + x
    dest = dst + '\\' + x
    shutil.move(sorc, dest)


def renameFiles(reqVer, words=["xls", "doc", "rtf"]):
    # dst = r"C:\Users\pragupathy\Downloads\input"
    # or (os.path.splitext(f)[1] == ".docx")
    dst = destinationFolder
    pattern = re.compile(r'(?:VSM|BSI)_ANALYSE_DES_ENTRANT', re.IGNORECASE)
    for f in os.listdir(dst):
        try:
            if re.match(pattern, f):
                continue
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


def sortFiles(count, reqVer, refnum):
    logging.info("In sortFiles function - ")
    while any(filename.endswith('.crdownload') or filename.endswith('.tmp') for filename in os.listdir(downloadFolder)):
        time.sleep(1)
    a = -1
    path = downloadFolder
    dst = destinationFolder
    name_list = os.listdir(path)
    full_list = [os.path.join(path, i) for i in name_list]
    time_sorted_list = sorted(full_list, key=os.path.getmtime)
    sorted_filename_list = [os.path.basename(i) for i in time_sorted_list]
    sorted_filename_list.reverse()

    for x in (sorted_filename_list[0:count]):
        if (os.path.splitext(x)[1] == ".docx") or (os.path.splitext(x)[1] == ".doc") or (
                os.path.splitext(x)[1] == ".docm") or (".xls" in os.path.splitext(x)[1]):
            if refnum not in x:
                logging.info(refnum)
                newFile = '[' + 'V' + str(float(reqVer)) + ']' + \
                    '[' + str(refnum) + ']' + x
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
            moveFile(newFile)
            a = dst + '\\' + newFile
        else:
            # newFile = f'[{refnum}][{str(float(reqVer))}]{x}'
            # try:
            #     os.rename(path + '\\' + x, path + '\\' + newFile)
            # except:
            #     os.remove(path + '\\' + newFile)
            #     time.sleep(3)
            #     os.rename(path + '\\' + x, path + '\\' + newFile)
            logging.info("x", x)
            moveFile(x)
            # a = path + '\\' + x
    return a


def getReference():
    reference = []
    for ref in docToDownload:
        reference.append(ref["Reference"])
        logging.info("Reference = " + str(reference))
    return reference


def getVersion():
    version = []
    for ref in docToDownload:
        version.append(ref["Version"])
        logging.info("Version = " + str(version))
    return version


# docInfo("01272_22_00311")


def startDocumentDownload(fepsRefVers, allowDownload=True, typesOffiles=["xls"]):
    logging.info("startDocumentDownload +", fepsRefVers)
    if allowDownload:
        logging.info("Allowed", allowDownload)
        for fepdDoc in fepsRefVers:
            try:

                ref, ver = fepdDoc
                logging.info(f"Startdownload = , {ref} {ver}")
                if ref != "":
                    service = Service(executable_path=webDriverAddr)
                    options = webdriver.ChromeOptions()
                    options.add_experimental_option(
                        'excludeSwitches', ['enable-logging'])
                    global driver
                    driver = webdriver.Chrome(service=service, options=options)
                    logging.info("--", ref, ver)
                    docInfo(ref, ver.replace("V", ""), words=typesOffiles)
                    time.sleep(3)
                else:
                    logging.info("reference number is empty")
            except Exception as e:
                logging.info(e)
                continue


userInput = {}
if __name__ == "__main__":
    startDocumentDownload([["00949_17_00725", ""]])
# C: \Users\skokane\VSM_BSI_QIA_PT\user_input\UserInput.json



