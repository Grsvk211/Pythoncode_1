# -*- coding: utf-8 -*-
"""
Created on Mon Dec  6 14:41:38 2021

@author: 10388
"""

from rply import LexerGenerator


class Lexer():
    def __init__(self):
        self.lexer = LexerGenerator()

    def _add_tokens(self):
        # Parenthesis
        self.lexer.add('OPEN_PAREN', r'\(')
        self.lexer.add('CLOSE_PAREN', r'\)')
        self.lexer.add('OPEN_SQPAREN', r'\[')
        self.lexer.add('CLOSE_SQPAREN', r'\]')
        self.lexer.add('OPEN_CURLY', r'\{')
        self.lexer.add('CLOSE_CURLY', r'\}')
        # Operators
        self.lexer.add('AND', r'AND')
        self.lexer.add('OR', r'OR')
        # Conditions
        self.lexer.add('IF', r'IF')
        self.lexer.add('THEN', r'THEN')
        self.lexer.add('ELSE', r'ELSE')
        self.lexer.add('FOR', r'FOR')
        self.lexer.add('EQUALTO', r'=')
        self.lexer.add('NOT', r'≠')
        self.lexer.add('NOT', r'NOT')
        self.lexer.add('GREAT', r'>')
        self.lexer.add('LESS', r'<')
        self.lexer.add('SIGNAL', r"\w+'*\w+")
        # self.lexer.add('COMA',r'\,')
        # self.lexer.add('VALUE',r'\d+')

        # ignore spaces
        self.lexer.ignore('\s+')

    def get_lexer(self):
        self._add_tokens()
        return self.lexer.build()
