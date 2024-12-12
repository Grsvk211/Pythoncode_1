import ctypes
import ExcelInterface as EI
import KeyboardMouseSimulator as KMS
import time
import InputConfigParser as ICF
from tkinter import *
import webbrowser
import re

import logging

def callback(url):
    webbrowser.open_new(url)


def getDocRefandVersion(tpBook, vPT):
    references = []
    versions = []
    ref = 6
    ver = 7
    maxrow = tpBook.sheets['Sommaire'].range('A' + str(tpBook.sheets['Sommaire'].cells.last_cell.row)).end('up').row
    CellValue = EI.searchDataInSheet(tpBook.sheets['Sommaire'], (1, maxrow), vPT)
    for i in CellValue["cellPositions"]:
        logging.info("i = ", vPT, i, str(int(tpBook.sheets['Sommaire'].range(i).value)))
        if str(int(tpBook.sheets['Sommaire'].range(i).value)) == vPT:
            x, y = i
            logging.info("x, y =", x, y)
    if x is not None:
        if tpBook.sheets['Sommaire'].range(x, 1).merge_cells is True:
            cellRange = tpBook.sheets['Sommaire'].range(x, 1).merge_area
            rlo = cellRange.row
            rhi = cellRange.last_cell.row
            logging.info("Merged Area =", rlo, rhi)

            links = []
            for i in range(rlo, rhi + 1):
                reference = tpBook.sheets['Sommaire'].range(i, ref).value
                version = tpBook.sheets['Sommaire'].range(i, ver).value
                # if reference not in references:
                if reference:
                    references.append(reference)
                    versions.append(version)

        else:
            # reference = tpBook.sheets['Sommaire'].range(i, ref).value
            # version = tpBook.sheets['Sommaire'].range(i, ver).value
            reference = tpBook.sheets['Sommaire'].range(x, ref).value
            version = tpBook.sheets['Sommaire'].range(x, ver).value
            if reference:
            # if reference not in references:
                references.append(reference)
                versions.append(version)

        # used to remove the spaces in reference(ex:-12345_12_12345\n --> we get o/p as 12345_12_12345)
        references = [s.rstrip() for s in references if s is not None]
        logging.info("reference--->", references)

        # used to add the V and float if it is not having means in summary (ex:- 23 o/p as V23.0)
        logging.info("versions--->", versions)
        versions = [s for s in versions if s is not None]
        versions = ['V' + str(s) if isinstance(s, float) else ('V' + str(s) if 'V' not in s else str(s)) for s in
                    versions]
        logging.info("versions--->", versions)

        # Create a set to track the unique combinations
        # from line 66 - 78 we are going to remove the duplicates of the version and references same in the summary tab
        unique_combinations = set()

        TRef_unique = []
        VRef_unique = []

        # Iterate over the T-Ref and V-Ref lists simultaneously
        for tref, vref in zip(references, versions):
            combination = (tref, vref)
            if combination not in unique_combinations:
                # Add the combination to the set and the result lists
                unique_combinations.add(combination)
                TRef_unique.append(tref)
                VRef_unique.append(vref)

        logging.info("T-Ref =", TRef_unique)
        references = TRef_unique
        logging.info("V-Ref =", VRef_unique)
        versions = VRef_unique

        # to remove the none in the version and reference previous tuple
        t_reff = []
        t_verr = []
        v = None
        for i in references:
            if i != v:
                t_reff.append(i)

        for j in versions:
            if j != v:
                t_verr.append(j)

        t_ref = tuple(t_reff)
        t_ver = tuple(t_verr)

        logging.info("T_REF---->", t_ref)
        logging.info("T_VER---->", t_ver)

        logging.info("Previous Version Documents to be downloaded = ", list(zip(t_ref, t_ver)))

    return list(zip(t_ref, t_ver))



def getDocLinks(tpBook, vPT):
    maxrow = tpBook.sheets['Sommaire'].range('A' + str(tpBook.sheets['Sommaire'].cells.last_cell.row)).end('up').row
    # CellValue = EI.searchDataInExcel(tpBook.sheets['Sommaire'], (1, maxrow), vPT)

    sheet_value = tpBook.sheets['Sommaire'].used_range.value
    CellValue = EI.searchDataInExcelCache(sheet_value,  (1, maxrow), vPT)

    for i in CellValue["cellPositions"]:
        logging.info("i = ", vPT, i, str(int(tpBook.sheets['Sommaire'].range(i).value)))
        if str(int(tpBook.sheets['Sommaire'].range(i).value)) == vPT:
            x, y = i
            logging.info("x, y =", x, y)
    if x is not None:
        if tpBook.sheets['Sommaire'].range(x, 1).merge_cells is True:
            cellRange = tpBook.sheets['Sommaire'].range(x, 1).merge_area
            rlo = cellRange.row
            rhi = cellRange.last_cell.row
            logging.info("Merged Area =", rlo, rhi)
            col = 7
            links = []
            for i in range(rlo, rhi + 1):
                link = tpBook.sheets['Sommaire'].range(i, col).hyperlink
                if link not in links:
                    links.append(link)
        else:
            col = 7
            links = []
            try:
                link = tpBook.sheets['Sommaire'].range(x, col).hyperlink
            except:
                link = "-"
            if link not in links:
                links.append(link)
        logging.info(links)
        return links


def showLinkPopUp(docLinks, vPT):
    root = Tk()
    frame = Frame(root)
    frame.pack()
    root.title("Links of documents used in testplan version " + vPT)
    link = []
    for docLink in docLinks:
        link.append(Label(root, text=docLink + " \n", fg="blue", cursor="hand2"))
    for i in range(len(link)):
        dclink = docLinks[i]

        def temp(event, link=dclink):
            event.widget.focus_set()  # give keyboard focus to the label
            event.widget.bind('<Key>', callback(link))

        link[i].pack()
        link[i].bind("<Button-1>", temp)
    root.after(5000, lambda: root.destroy())
    root.mainloop()


# ctypes.windll.user32.MessageBoxW(0, "All documents downloaded", "Aptest", 1)
