# -*- coding: utf-8 -*-
"""
Created on Mon Dec  6 16:11:25 2021

@author: Sam Joel
"""

from rply import ParserGenerator
from contentAST import Signal, And, Or, Equalto, If, elseIf, And_Paren, Not, Time_Max, Time_Min, Great, Less, Add, Sub


class Parser():

    def __init__(self):
        self.pg = ParserGenerator(
            # A list of all token names, accepted by the parser.
            ['SIGNAL', 'OPEN_PAREN', 'CLOSE_PAREN', 'OPEN_SQPAREN', 'CLOSE_SQPAREN', 'OPEN_CURLY', 'CLOSE_CURLY',
             'AND', 'OR', 'IF', 'THEN', 'ELSE', 'CHANGE', 'EQUALTO', 'NOT', 'GREAT', 'LESS', 'FOR', 'PLUS', 'MINUS'
             ],
            # A list of precedence rules with ascending precedence, to
            # disambiguate ambiguous production rules.
            precedence=[
                ('left', ['AND', 'OR']),
                ('left', ['EQUALTO', 'NOT', 'GREAT', 'LESS', 'PLUS', 'MINUS']),
                ('left', 'FOR'),
                ('left', ['OPEN_CURLY', 'CLOSE_CURLY']),
                ('left', ['OPEN_SQPAREN', 'CLOSE_SQPAREN']),
                ('left', ['OPEN_PAREN', 'CLOSE_PAREN'])
            ]
            # precedence=[('left', ['COMA','AND','OR'])]
        )

    def parse(self):

        '''
        @self.pg.production('program : CREATE OPEN_PAREN expression CLOSE_PAREN')
        def program(p):
            #logging.info("here")
            return createLines(p[2])'''

        @self.pg.production('expression : OPEN_PAREN expression CLOSE_PAREN')
        def expression_paren(p):
            # logging.info("paren ",p[1])
            return p[1]

        @self.pg.production('expression : OPEN_SQPAREN expression CLOSE_SQPAREN')
        def expression_sqparen(p):
            # logging.info("paren ",p[1])
            return p[1]

        @self.pg.production('expression : OPEN_CURLY expression CLOSE_CURLY')
        def expression_curly(p):
            # logging.info("paren ",p[1])
            return p[1]

        @self.pg.production('expression : expression AND OPEN_PAREN expression CLOSE_PAREN')
        def expression_and_paren(p):
            return And_Paren(p[0], p[3])

        @self.pg.production('expression : OPEN_PAREN expression CLOSE_PAREN AND expression')
        def expression_paren_and(p):
            return And_Paren(p[1], p[4])

        @self.pg.production('expression : IF expression THEN expression')
        def expression_if(p):
            # logging.info("paren ",p[1])
            return If(p[1], p[3])

        @self.pg.production('expression : IF expression THEN expression ELSE expression')
        def expression_elseif(p):
            # logging.info("paren ",p[1])
            return elseIf(p[1], p[3], p[5])

        @self.pg.production('expression : expression NOT EQUALTO expression')
        def expression_not(p):
            left = p[0]
            right = p[3]
            if p[2].gettokentype() == 'EQUALTO':
                return Not(left, right)

        @self.pg.production('expression : expression FOR GREAT expression')
        @self.pg.production('expression : expression FOR LESS expression')
        def expression_for_time(p):
            left = p[0]
            right = p[3]
            if p[2].gettokentype() == 'GREAT':
                return Time_Max(left, right)
            if p[2].gettokentype() == 'LESS':
                return Time_Min(left, right)

        @self.pg.production('expression : expression AND expression')
        @self.pg.production('expression : expression OR expression')
        @self.pg.production('expression : expression EQUALTO expression')
        @self.pg.production('expression : expression NOT expression')
        @self.pg.production('expression : expression GREAT expression')
        @self.pg.production('expression : expression LESS expression')
        @self.pg.production('expression : expression PLUS expression')
        @self.pg.production('expression : expression MINUS expression')
        # @self.pg.production('expression : expression GEQUAL expression')
        # @self.pg.production('expression : expression LEQUAL expression')
        # @pg.production('expression : expression DIV expression')
        def expression_op(p):
            # logging.info(p[0],p[2])
            left = p[0]
            right = p[2]
            if p[1].gettokentype() == 'AND':
                return And(left, right)
            elif p[1].gettokentype() == 'OR':
                return Or(left, right)
            elif p[1].gettokentype() == 'EQUALTO':
                return Equalto(left, right)
            elif p[1].gettokentype() == 'NOT':
                return Not(left, right)
            elif p[1].gettokentype() == 'GREAT':
                return Great(left, right)
            elif p[1].gettokentype() == 'LESS':
                return Less(left, right)
            elif p[1].gettokentype() == 'PLUS':
                return Add(left, right)
            elif p[1].gettokentype() == 'MINUS':
                return Sub(left, right)
            # elif p[1].gettokentype() == 'DIV':
            #   return Div(left, right)
            else:
                raise AssertionError('Oops, this should not be possible!')

        @self.pg.production('expression : SIGNAL')
        def expression_signal(p):
            # p is a list of the pieces matched by the right hand side of the
            # rule
            return Signal((p[0].getstr()))

        @self.pg.error
        def error_handle(token):
            raise ValueError(token)

    def get_parser(self):
        return self.pg.build()
