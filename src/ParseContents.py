import logging
# -*- coding: utf-8 -*-
"""
Created on Mon Dec  6 16:09:47 2021

@author: 10388
"""

from ContentLexer import Lexer
from ContentParser import Parser


def createSteps(data):
    lexer=Lexer().get_lexer()
    tokens=lexer.lex(data)
    '''   
    for token in tokens:
        logging.info(token)
    
    '''
    pg = Parser()
    logging.info("parser obj created")
    pg.parse()
    logging.info("Calling parse method")
    parser = pg.get_parser()
    logging.info("parser built")
    result=parser.parse(tokens).eval()
    #logging.info(result.split('\n'))
    return result