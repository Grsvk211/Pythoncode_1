from docx import Document
import shutil
import datetime
import shutil
import docx
import logging
import sys
import ExcelInterface as EI
import os
import re
import InputConfigParser as ICF
import WordDocInterface as WDI
import web_interface as WI
import NewRequirementHandler as NRH
import threading
import pygetwindow as pgw
import time
import KeyboardMouseSimulator as KMS
import logging
from itertools import zip_longest
from DCI_download_webinterface import startDocumentDownload
from itertools import chain
from docx import Document
from docx.enum.text import WD_BREAK
date_time = datetime.datetime.now()


def extract_text_from_corrupted_docx(file_path):
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    except Exception as e:
        print("Error opening corrupted Word document:", e)
        return None

# def save_copy_of_corrupted_docx(file_path, copied_file_path):
#     try:
#         shutil.copyfile(file_path, copied_file_path)
#         print("Copied corrupted Word document to:", copied_file_path)
#     except Exception as e:
#         print("Error copying corrupted Word document:", e)
#
# corrupted_text = extract_text_from_corrupted_docx(corrupted_docx_path)
# if corrupted_text:
#     print("Text extracted from corrupted Word document:")
#     print(corrupted_text)
# else:
#     print("Failed to extract text from corrupted Word document.")
#
# # Save a copy of the corrupted Word document with the same name but in another file
# copied_docx_path = "path/to/your/corrupted_copy.docx"
# save_copy_of_corrupted_docx(corrupted_docx_path, copied_docx_path)


def getReqVer(req):
    if req.find('(') != -1:
        new_reqName = req.split("(")[0].split()[0] if len(req.split("(")) > 0 else ""
        new_reqVer = req.split("(")[1].split(")")[0] if len(req.split("(")) > 1 else ""
    else:
        new_reqName = req.split()[0] if len(req.split()) > 0 else ""
        new_reqVer = req.split()[1] if len(req.split()) > 1 else ""
    return new_reqName.strip(), new_reqVer.strip()


def main():
    table_lists = []
    Doc_lists = []
    # Doc= r"C:\Users\vgajula\Documents\4-3-2024\ISA\Input_Files\[V14.0][02017_19_02186]SSD_HMIF_LEGACY_DRIVING_ASSISTANCE_HMI.docx"
    Doc = r"C:\Users\vgajula\Downloads\table.docx"
    req = 'REQ-0580241 E'
    ReqName, ReqVer = getReqVer(req)
    print("ReqName, ReqVer------------->", ReqName, ReqVer)
    try:
        RqTable, Content, Doc = WDI.getContent(Doc, ReqName, ReqVer)
        print("RqTable, Content, Doc-------------->", RqTable, Content, Doc)
        # table_lists.append(RqTable)
        # Doc_lists.append(Doc)
        # # if Content != None and Content != -1 and Content != '':
        # #     break
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(
            f"\nSomething went wrong getting table {ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")


if __name__ == "__main__":
    main()
    