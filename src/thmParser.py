# -*- coding: utf-8 -*-
"""
Created on Wed Dec  1 10:09:48 2021

@author: 10388
"""

from rply import ParserGenerator
from thmAST import Signal, And, Or, Coma, And_Paren


class Parser():

    def __init__(self):
        self.pg = ParserGenerator(
            # A list of all token names, accepted by the parser.
            ['SIGNAL', 'OPEN_PAREN', 'CLOSE_PAREN',
             'AND', 'OR', 'COMA'],
            # A list of precedence rules with ascending precedence, to
            # disambiguate ambiguous production rules.
            precedence=[('left', ['OPEN_PAREN', 'CLOSE_PAREN']), ('left', ['COMA', 'AND', 'OR'])]
            # precedence=[('left', ['COMA','AND','OR'])]
        )

    def parse(self):

        '''@self.pg.production('program : CREATE OPEN_PAREN expression CLOSE_PAREN')
        def program(p):
            #logging.info("here")
            return createLines(p[2])'''

        @self.pg.production('expression : OPEN_PAREN expression CLOSE_PAREN')
        def expression_paren(p):
            # logging.info("paren ",p[1])
            return p[1]

        @self.pg.production('expression : expression AND OPEN_PAREN expression CLOSE_PAREN')
        def expression_and_paren(p):
            return And_Paren(p[0], p[3])

        # @self.pg.production('expression : expression AND OPEN_PAREN OPEN_PAREN expression CLOSE_PAREN')
        # def expression_and_paren(p):
        #     return And_Paren(p[0], p[4])

        @self.pg.production('expression : OPEN_PAREN expression CLOSE_PAREN AND expression')
        def expression_paren_and(p):
            return And_Paren(p[1], p[4])

        # @self.pg.production('expression : OPEN_PAREN expression CLOSE_PAREN CLOSE_PAREN AND expression')
        # def expression_paren_and(p):
        #     return And_Paren(p[1], p[5])

        @self.pg.production('expression : expression AND expression')
        @self.pg.production('expression : expression OR expression')
        @self.pg.production('expression : expression COMA expression')
        # @pg.production('expression : expression DIV expression')
        def expression_op(p):
            # logging.info(p[0],p[2])
            left = p[0]
            right = p[2]
            if p[1].gettokentype() == 'AND':
                return And(left, right)
            elif p[1].gettokentype() == 'OR':
                return Or(left, right)
            elif p[1].gettokentype() == 'COMA':
                return Coma(left, right)
            # elif p[1].gettokentype() == 'DIV':
            #   return Div(left, right)
            else:
                raise AssertionError('Oops, this should not be possible!')

        @self.pg.production('expression : SIGNAL')
        def expression_signal(p):
            # p is a list of the pieces matched by the right hand side of the
            # rule
            # logging.info("signal ",p[0])
            return Signal((p[0].getstr()))

        @self.pg.error
        def error_handle(token):
            raise ValueError(token)

    def get_parser(self):
        return self.pg.build()
