import logging
# -*- coding: utf-8 -*-
"""
Created on Tue Dec  7 23:38:25 2021

@author: 10388
"""


class Signal():
    def __init__(self, value):
        self.value = value

    def eval(self):
        # logging.info("sig",self.value)
        return str(self.value)


class BinaryOp():
    def __init__(self, left, right):
        self.left = left
        self.right = right


class And_Paren(BinaryOp):
    def eval(self):
        if self.right.eval().find('\n') != -1:
            # logging.info("Right")
            right = self.right.eval().split('\n')
            for i in range(0, len(right)):
                right[i] = right[i] + '|' + str(self.left.eval())
            return '\n'.join(right)
        if self.left.eval().find('\n') != -1:
            # logging.info("left")
            left = self.left.eval().split('\n')
            for i in range(0, len(left)):
                left[i] = left[i] + '|' + str(self.right.eval())
            return '\n'.join(left)

        return self.left.eval() + '|' + self.right.eval()


class And(BinaryOp):
    def eval(self):

        if (self.right.eval().find('\n') != -1):
            # logging.info("Right")
            left = self.left.eval().split('\n')
            right = self.right.eval().split('\n')
            result = []
            for i in range(0, len(left)):
                for j in range(0, len(right)):
                    result.append(left[i] + '|' + right[j])
            return '\n'.join(result)
        if (self.left.eval().find('\n') != -1):
            # logging.info("AND_LEFT")
            left = self.left.eval().split('\n')
            right = self.right.eval().split('\n')
            result = []
            for i in range(0, len(left)):
                for j in range(0, len(right)):
                    result.append(left[i] + '|' + right[j])
            # logging.info("AND_LEFT",left)
            return '\n'.join(result)
        if (self.left.eval().find('=') != -1) and (
                (self.right.eval().find('=') == -1) and (self.right.eval().find('<') == -1) and (
                self.right.eval().find('>') == -1) and (self.right.eval().find('!') == -1)):
            left = self.left.eval().split('=')
            # logging.info("left",left)
            left.remove(left[-1])
            left = '='.join(left)
            # logging.info("left2",left)
            return self.left.eval() + '|' + left + '=' + self.right.eval()
        if (self.left.eval().find('!') != -1) and (
                (self.right.eval().find('=') == -1) and (self.right.eval().find('<') == -1) and (
                self.right.eval().find('>') == -1) and (self.right.eval().find('!') == -1)):
            left = self.left.eval().split('!')
            # logging.info("left",left)
            left.remove(left[-1])
            left = '!'.join(left)
            # logging.info("left2",left)
            return self.left.eval() + '|' + left + '!' + self.right.eval()
        if (self.left.eval().find('<') != -1) and (
                (self.right.eval().find('=') == -1) and (self.right.eval().find('<') == -1) and (
                self.right.eval().find('>') == -1) and (self.right.eval().find('!') == -1)):
            left = self.left.eval().split('<')
            # logging.info("left",left)
            left.remove(left[-1])
            left = '<'.join(left)
            # logging.info("left2",left)
            return self.left.eval() + '|' + left + '<' + self.right.eval()
        if (self.left.eval().find('>') != -1) and (
                (self.right.eval().find('=') == -1) and (self.right.eval().find('<') == -1) and (
                self.right.eval().find('>') == -1) and (self.right.eval().find('!') == -1)):
            left = self.left.eval().split('>')
            # logging.info("left",left)
            left.remove(left[-1])
            left = '>'.join(left)
            logging.info("left2", left)
            return self.left.eval() + '|' + left + '>' + self.right.eval()
            '''for i in range(0,len(right)):
                right[i]=right[i]+'|'+str(self.left.eval())
            return ( '\n'.join(right))
        if self.left.eval().find('\n')!=-1:
            logging.info("left")
            left=self.left.eval().split('\n')
            for i in range(0,len(left)):str(self.right.eval()
                left[i]=left[i]+'|'+str(self.right.eval())
            return '\n'.join(left)'''

        return (str(self.left.eval()) + '|' + str(self.right.eval()))


class Or(BinaryOp):
    def eval(self):
        # if (self.left.eval() is not None and self.right.eval() is not None):
        if (self.left.eval().find('=') != -1) and (self.right.eval().find('=') == -1):
            left = self.left.eval().split('=')
            left.remove(left[-1])
            left = '='.join(left)
            return self.left.eval() + '\n' + left + '=' + self.right.eval()
        if (self.left.eval().find('!') != -1) and (self.right.eval().find('!') == -1):
            left = self.left.eval().split('!')
            left.remove(left[-1])
            left = '!'.join(left)
            return self.left.eval() + '\n' + left + '!' + self.right.eval()
        if (self.left.eval().find('<') != -1) and (self.right.eval().find('<') == -1):
            left = self.left.eval().split('<')
            # rint("left",left)
            left.remove(left[-1])
            left = '<'.join(left)
            # logging.info("left2",left)
            return self.left.eval() + '\n' + left + '<' + self.right.eval()
        if (self.left.eval().find('>') != -1) and (self.right.eval().find('>') == -1):
            left = self.left.eval().split('>')
            # logging.info("left",left)
            left.remove(left[-1])
            left = '>'.join(left)
            # logging.info("left2",left)
            return self.left.eval() + '\n' + left + '>' + self.right.eval()

        return (str(self.left.eval()) + '\n' + str(self.right.eval()))


class Equalto(BinaryOp):
    def eval(self):
        '''if ((self.left.eval().find('\n')!=-1)or(self.left.eval().find('|')!=-1)):
            left=self.left.eval()
            for l in left:
                if l !='|' and l != '\n':
                    left=left.replace(l,l+'='+self.right.eval())
            return left
        if ((self.right.eval().find('\n')!=-1)or(self.right.eval().find('|')!=-1)):
            logging.info("here",self.right.eval())
            right=self.right.eval()
            for l in right:
                if l !='|' and l != '\n':
                    right=right.replace(l,l+'='+self.left.eval())
            logging.info("complex equal")
            return right'''

        # logging.info("simple equal")
        return (str(self.left.eval()) + '=' + str(self.right.eval()))


class Not(BinaryOp):
    def eval(self):
        '''if ((self.left.eval().find('\n')!=-1)or(self.left.eval().find('|')!=-1)):
            left=self.left.eval()
            for l in left:
                if l !='|' and l != '\n':
                    left=left.replace(l,l+'!'+self.right.eval())
            return left
        if ((self.right.eval().find('\n')!=-1)or(self.right.eval().find('|')!=-1)):
            right=self.right.eval()
            for l in right:
                if l !='|' and l != '\n':
                    right=right.replace(l,l+'!'+self.left.eval())
            return right'''
        return (str(self.left.eval()) + '!' + str(self.right.eval()))


class Great(BinaryOp):
    def eval(self):
        return self.left.eval() + '>' + self.right.eval()


class Less(BinaryOp):
    def eval(self):
        return self.left.eval() + '<' + self.right.eval()


class Time_Min(BinaryOp):
    def eval(self):
        # logging.info("time added")
        if self.left.eval().find('\n'):
            left = self.left.eval().split('\n')
            result = []
            for l in left:
                result.append(l + '#' + self.right.eval())
            return '\n'.join(result)
        if self.left.eval().find('|'):
            left = self.left.eval().split('|')
            result = []
            for l in left:
                result.append(l + '#' + self.right.eval())
            return '|'.join(result)
        return self.left.eval() + "#" + self.right.eval()


class Add(BinaryOp):
    def eval(self):
        if self.left.eval().find('\n'):
            left = self.left.eval().split('\n')
            result = []
            for l in left:
                result.append(l + '+' + self.right.eval())
            return '\n'.join(result)
        if self.left.eval().find('|'):
            left = self.left.eval().split('|')
            result = []
            for l in left:
                result.append(l + '+' + self.right.eval())
            return '|'.join(result)
        return self.left.eval() + "+" + self.right.eval()


class Sub(BinaryOp):
    def eval(self):
        if self.left.eval().find('\n'):
            left = self.left.eval().split('\n')
            result = []
            for l in left:
                result.append(l + '-' + self.right.eval())
            return '\n'.join(result)
        if self.left.eval().find('|'):
            left = self.left.eval().split('|')
            result = []
            for l in left:
                result.append(l + '-' + self.right.eval())
            return '|'.join(result)
        return self.left.eval() + "-" + self.right.eval()


class Time_Max(BinaryOp):
    def eval(self):
        # logging.info("time added")
        # logging.info("right ",self.right.eval())
        if self.left.eval().find('\n'):
            left = self.left.eval().split('\n')
            result = []
            for l in left:
                result.append(l + '$' + self.right.eval())
            return '\n'.join(result)
        if self.left.eval().find('|'):
            left = self.left.eval().split('|')
            result = []
            for l in left:
                result.append(l + '$' + self.right.eval())
            return '|'.join(result)
        return self.left.eval() + "$" + self.right.eval()


class If():
    def __init__(self, ip, op):
        self.ip = ip
        self.op = op

    def eval(self):

        if self.ip.eval().find('\n'):
            ipList = self.ip.eval().split('\n')
            opList = self.op.eval().split('\n')
            result = []
            for o in opList:
                for i in ipList:
                    result.append(i + '==' + o)
            return '\n'.join(result)

        return '\n' + self.ip.eval() + '==' + self.op.eval()


class elseIf():
    def __init__(self, ip, op, falseOp):
        self.ip = ip
        self.op = op
        self.falseOp = falseOp

    def eval(self):
        # split output with \n  remove duplicate elements
        '''with open("C:/Users/10388/CompileRequirement/result.txt", "w") as file:
            # write to file
            file.writelines(self.value.eval())'''

        if self.ip.eval().find('\n') != -1:
            # logging.info("Right")
            right = self.ip.eval().split('\n')
            fright = right.copy()
            iright = right.copy()
            for n, i in enumerate(right):
                for j in i:
                    if j == '=':
                        i = i.replace(j, '!')
                    if j == '!':
                        i = i.replace(j, '=')
                fright[n] = i
            for i in range(0, len(right)):
                iright[i] = right[i] + '==' + str(self.op.eval()) + '\n' + fright[i] + '==' + self.falseOp.eval()
            result = '\n'.join(iright)
            '''for i in result:
                logging.info("i",i)
                for j in i:
                    if '=' in  j:
                        j =j.replace(j,'!')
                result=result.replace(i,j)'''

            return result
        '''
        if self.left.eval().find('\n')!=-1:
            logging.info("left")
            left=self.left.eval().split('\n')
            for i in range(0,len(left)):
                left[i]=left[i]+str(self.right.eval())
            return '\n'.join(left)'''

        # logging.info(self.ip.eval()+'=='+self.op.eval())
        falseip = self.ip.eval()
        for fip in falseip:
            if fip == '=':
                falseip = falseip.replace(fip, "!")
            if fip == '!':
                falseip = falseip.replace(fip, "=")
        tIp = self.ip.eval() + '==' + self.op.eval()
        falip = falseip + '==' + self.falseOp.eval()
        result = tIp + '\n' + falip
        # return self.ip.eval()+'=='+self.op.eval()+'\n'+self.ip.eval()+'!='+self.falseOp.eval()
        return result
