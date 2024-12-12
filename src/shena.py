import sys

import ExcelInterface as EI
import os
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import time
from lexer import Lexer
from thmParser import Parser
import re
import collections
# from openpyxl import load_workbook
import json
import xlwings as xw
import InputConfigParser as ICF
import logging
def grepThematicsCode(rawThematics):
    start_time = time.time()
    logging.info(f'grepThematicsCode start time: {start_time}')

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
        # logging.info("a - ", a)
        if re.search("[a-zA-Z0-9]{3}[(][0-9]{2}[)]", a) is not None:
            a = a.replace("(", "_")
            a = a.replace(")", "")
            a = a.strip()
            thematics_code.append(a)
            logging.info("thematics_code = ", thematics_code)
        else:
            if a.find("AND") == 0:
                if (thematics_code[-1] != "AND"):
                    thematics_code.append(a)
            if a.find("OR") == 0:
                if (thematics_code[-1] != "OR"):
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
    if len(re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", thematics_code[0])) == 0:
        if thematics_code[0] != "(":
            # logging.info("removing first element", thematics_code[0], re.findall("[a-zA-Z0-9]{3}_[0-9]{2}", thematics_code[0]))
            thematics_code.remove(thematics_code[0])
    # logging.info("Thematic = ", thematics_code)
    # logging.info("Thematic code final(1) = ", ''.join(thematics_code))
    openBracket = []
    closeBracket = []
    for i in range(len(thematics_code)):
        if thematics_code[i] == "(":
            openBracket.append(i)
        if thematics_code[i] == ")":
            closeBracket.append(i)
    # logging.info("Indices = ", openBracket, closeBracket)
    # logging.info("Thematic code = ", thematics_code)
    for n, i in enumerate(thematics_code):
        # logging.info("N & i", n, i)
        if i.find('_') != -1:
            # logging.info("thm code",thematics_code[n+1])
            if n < (len(thematics_code) - 1):
                if thematics_code[n + 1].find('_') != -1:
                    thematics_code[n + 1] = ',' + thematics_code[n + 1]
    reducedThm = ' '.join(thematics_code)
    logging.info("Thematic code final(2)1 = ", reducedThm)
    reducedThm = reducedThm.replace("( )", "")
    logging.info("Thematic code final(2)2 = ", reducedThm)
    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'grepThematicsCode end execution time: {execution_time}')

    return reducedThm


def remove_trailing_and_or(input_string):
    if input_string.strip().endswith("'"):
        input_string=input_string.strip()[:-1]
    while input_string.strip().endswith('AND') or input_string.strip().endswith('OR'):
        if input_string.strip().endswith('AND'):
            input_string = input_string.strip()[:-3]
        elif input_string.strip().endswith('OR'):
            input_string = input_string.strip()[:-2]
    return input_string


def createCombination(data):
    start_time = time.time()
    logging.info(f'createCombination start time: {start_time}')

    lexer = Lexer().get_lexer()
    tokens = lexer.lex(data)
    '''
    for token in tokens:
        logging.info(token)'''

    pg = Parser()
    pg.parse()
    parser = pg.get_parser()
    combinations = parser.parse(tokens).eval()

    end_time = time.time()
    execution_time = end_time - start_time
    logging.info(f'createCombination end execution time: {execution_time}')

    return combinations


if __name__ == "__main__":
    rawThematics = input("Enter the Raw thematics to get the combination lines: ")
    # rawThematics = "AND ( CLI TYPE_AEE_LEVE_VITRES (CLI_02 MUX)  AND IWV OPTION_LV_AR_ELEC (IWV_00 WITHOUT)  AND IWY OPTION_LV_AP (IWY_01 AVEC_AP)  AND LNG TYPE_SIDE_DOORS_ARCHI (LNG_02 2_DCU_AV , LNG_03 4_DCU_AV_AR)   AND LYQ TYPE_DIVERSITY (LYQ_01 BEFORE_FUNCT_CODIF)  )  OR ( CLI TYPE_AEE_LEVE_VITRES (CLI_02 MUX)   AND DLE REAR WINDOWS LIFTER (DLE_00 WITHOUT , DLE_10 MANUAL)  AND IWY OPTION_LV_AP (IWY_01 AVEC_AP)  AND LNG TYPE_SIDE_DOORS_ARCHI (LNG_02 2_DCU_AV , LNG_03 4_DCU_AV_AR)  AND LYQ TYPE_DIVERSITY (LYQ_02 FUNCT_CODIF)  ) "
    data = grepThematicsCode(rawThematics)
    print("reducedThm--------------->", data)
    if data.strip().endswith('OR') or data.strip().endswith('AND'):
        data = remove_trailing_and_or(data)
    effectiveExpression = ''
    if len(data.strip()) != 0:
        effectiveExpression = createCombination(data)
    print("effectiveExpression------------->", effectiveExpression)
