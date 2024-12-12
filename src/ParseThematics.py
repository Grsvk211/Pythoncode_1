import logging
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 22:55:56 2021

@author: 10388
"""

from lexer import Lexer
from thmParser import Parser


# import IM018 as im


def createCombination(data):
    lexer = Lexer().get_lexer()
    tokens = lexer.lex(data)
    logging.info(f"tokens {tokens}")
    '''
    for token in tokens:
        logging.info(token)'''

    pg = Parser()
    logging.info("parser obj created")
    pg.parse()
    logging.info("Calling parse method")
    parser = pg.get_parser()
    logging.info("parser built")
    combinations = parser.parse(tokens).eval()
    # logging.info("token evaluated",parser.parse(tokens).eval())
    return combinations

# data = "( AZC_00 ) AND ( ( AJD_01,AJD_02) AND ( ( AJO_01 ) AND ( LYQ_01 ) OR ( ALN_00 ) AND ( LYQ_02 ) AND ( DUB_21 ,DUB_23 ,DUB_24 ) ) )"
# # # data = " ( AZC_00 ) AND ( ( AJA_01 ) AND ( AJO_01 ) AND ( LYQ_01 ) OR ( DUB_21 ) AND ( LYQ_02 ) OR( DUB_24 ) AND ( LYQ_03 ) )"
# logging.info(createCombination(data))
# for comb in combList
