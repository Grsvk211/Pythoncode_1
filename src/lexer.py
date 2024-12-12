# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 22:58:31 2021

@author: 10388
"""

from rply import LexerGenerator

'''
class Lexer():
    def __init__(self):
        self.lexer=LexerGenerator()
    def _add_tokens(self):
        #Parenthesis
        self.lexer.add('OPEN_PAREN', r'\(')
        self.lexer.add('CLOSE_PAREN',r'\)')
        #Operators
        self.lexer.add('AND',r'AND')
        self.lexer.add('OR',r'OR')
        #Conditions
        self.lexer.add('IF',r'IF')
        self.lexer.add('THEN',r'THEN')
        self.lexer.add('ELSE',r'ELSE')
        self.lexer.add('EQUALTO',r'=')
        self.lexer.add('SIGNAL',r'\D+')
        self.lexer.add('VALUE',r'\d+')
        #ignore spaces
        self.lexer.ignore('\s+')
    
    def get_lexer(self):
        self._add_tokens()
        return self.lexer.build()
'''


class Lexer():
    def __init__(self):
        self.lexer = LexerGenerator()

    def _add_tokens(self):
        # self.lexer.add('CREATE',r'createCombination')
        # Parenthesis
        self.lexer.add('OPEN_PAREN', r'\(')
        self.lexer.add('CLOSE_PAREN', r'\)')
        # Operators
        self.lexer.add('AND', r'AND')
        self.lexer.add('OR', r'OR')
        # Conditions
        # self.lexer.add('IF',r'IF')
        # self.lexer.add('THEN',r'THEN')
        # self.lexer.add('ELSE',r'ELSE')
        # self.lexer.add('EQUALTO',r'=')
        self.lexer.add('SIGNAL', r'\w+')
        self.lexer.add('COMA', r'\,')
        # self.lexer.add('VALUE',r'\d+')
        # ignore spaces
        self.lexer.ignore(r'\s+')

    def get_lexer(self):
        self._add_tokens()
        return self.lexer.build()
