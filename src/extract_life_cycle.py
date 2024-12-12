import ply.lex as lex
import ply.yacc as yacc
import logging

showTokens = False
compileStatus = True
inputCondition = []
outputCondition = []

# Tokenize the string
tokens = [
    'EXPRESSION',
    'AND',
    'OR',
    'VALUE'
]


def t_EXPRESSION(t):
    r'[a-zA-Z0-9_]+\s?=\s?[a-zA-Z0-9_]+'
    t.type = 'EXPRESSION'
    logging.info(t) if showTokens else ()
    return t


def t_AND(t):
    r'AND'
    t.type = 'AND'
    logging.info(t) if showTokens else ()
    return t


def t_OR(t):
    r'OR'
    t.type = 'OR'
    logging.info(t) if showTokens else ()
    return t


def t_value(t):
    r'[a-zA-Z0-9_]+'
    t.type = 'VALUE'
    logging.info(t) if showTokens else ()
    return t


t_ignore = ' \t\n+'


def t_error(t):
    global compileStatus
    compileStatus = False
    logging.info("Illegal character '%s'" % t.value[0])
    t.lexer.skip(1)


def reinit():
    global inputCondition, outputCondition, compileStatus
    inputCondition = []
    outputCondition = []
    compileStatus = True


lexer = lex.lex()


def extract_condn(s):
    """
    The function extracts conditions from a string and returns them in a nested list format.
    
    :param s: The input string containing the condition to be extracted and parsed
    :return: The function `extract_condn` returns a list of lists, where each inner list contains pairs
    of conditions in the form of [lhs, rhs]. The outer list contains groups of conditions that are
    connected by either 'AND' or 'OR' operators, with conditions separated by AND in same group.
    """
    s = s.replace("==", "=")
    
    logging.info("==========LEXER STARTED===========")
    lexer.input(s)
    expr_list = []
    lastConj= None
    lastToken = None

    for token in lexer:
        if token.type == 'VALUE' and lastToken == None:
            continue
        logging.info(token.type, token.value)
        if token.type == 'AND':
            lastConj = 'AND'
        elif token.type == 'OR':
            lastConj = 'OR'

        if token.type == 'EXPRESSION':
            if lastConj != 'AND':
                expr_list.append([])
            expr_list[-1].append(token.value)
        
        if token.type == 'VALUE':
            if lastToken != 'OR':
                expr_list[-1][-1] += " " + token.value
            else:
                expr_list[-1].append(token.value)
        
        lastToken = token.type
    
    logging.info("==========LEXER COMPLETED===========")

    output_list = []
    lastCondn = ""

    for tup in expr_list:
        list_pair = []
        buffer_list = []
        for condn in tup:
            if "=" in condn:
                lhs, rhs = condn.split("=")
                lastCondn = (lhs.strip(), rhs.strip())
                list_pair.append(lastCondn)
            else:
                lastCondn = (lastCondn[0], condn.strip())
                buffer_list.append(lastCondn)
        output_list.append(list_pair)
        if len(buffer_list)>0:
            output_list.append(buffer_list)
    
    return output_list


if __name__=="__main__":

#     s = '''
# ETAT_PRINCIP_SEV == Arret
# AND
# CDE_PDV_CAN_HS4 == Wakeup
# OR
# ETAT_PRINCIP_SEV == Contact OR DEM
# AND
# ETAT_GMP == MOTEUR_TOURNANT
# OR
# ETAT_PRINCIP_SEV == Contact OR DEM
# AND
# ETAT_GMP == MOTEUR_NON_TOURNANT
# '''

    s = '\nETAT_PRINCIP_SEV == Arret\nAND\nCDE_PDV_CAN_HS4 == Wakeup\nOR\nETAT_PRINCIP_SEV == Contact OR DEM\nAND\nETAT_GMP == MOTEUR_TOURNANT\nOR\nETAT_PRINCIP_SEV == Contact OR DEM\nAND\nETAT_GMP == MOTEUR_NON_TOURNANT\n\n'

    output_list = extract_condn(s)
    logging.info('output_list----->', output_list)

    # for lst in output_list:
    #     logging.info(lst)


        


