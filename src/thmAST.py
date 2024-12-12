import logging
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  1 09:49:40 2021

@author: 10388
"""
c = 0


class Signal():
    def __init__(self, value):
        self.value = value

    def eval(self):
        return str(self.value)


class BinaryOp():
    def __init__(self, left, right):
        self.left = left
        self.right = right
        # logging.info("self.left = ", self.left, self.right)


# class And_Paren(BinaryOp):
#     def eval(self):
#         if self.left.eval().find('\n')!=-1:
#             #logging.info("left")
#             left=self.left.eval().split('\n')
#             # logging.info("left", left)
#             for i in range(0,len(left)):
#                 left[i]=left[i]+'|'+str(self.right.eval())
#             return '\n'.join(left)
#         if self.right.eval().find('\n')!=-1:
#             #logging.info("Right")
#             right=self.right.eval().split('\n')
#             # logging.info("Right", right)
#             for i in range(0,len(right)):
#                 right[i]=right[i]+'|'+str(self.left.eval())
#             return ( '\n'.join(right))
#
#
#         return self.left.eval()+'|'+self.right.eval()
class And_Paren(BinaryOp):
    def eval(self):

        if self.right.eval().find('\n') != -1:
            # logging.info("Right")
            right = self.right.eval().split('\n')
            res = []
            for i in range(0, len(right)):
                if self.left.eval().find('\n') != -1:
                    left = self.left.eval().split('\n')
                    for j in range(0, len(left)):
                        res.append(right[i] + '|' + str(left[j]))
                else:
                    res.append(right[i] + "|" + str(self.left.eval()))
            return '\n'.join(res)
        if self.left.eval().find('\n') != -1:
            # logging.info("left")
            left = self.left.eval().split('\n')
            res = []
            for i in range(0, len(left)):
                if self.right.eval().find('\n') != -1:
                    right = self.right.eval().split('\n')
                    for j in range(0, len(right)):
                        res.append(left[i] + '|' + str(right[j]))
                else:
                    res.append(left[i] + "|" + str(self.right.eval()))
            return '\n'.join(res)

        return self.left.eval() + '|' + self.right.eval()


class And(BinaryOp):
    def eval(self):
        # if self.right.eval().find('\n')!=-1:
        #     #logging.info("Right")
        #     right=self.right.eval().split('\n')
        #     for i in range(0,len(right)):
        #         right[i]=right[i]+"|"+str(self.left.eval())
        #     return ( '\n'.join(right))
        # if self.left.eval().find('\n')!=-1:
        #     #logging.info("left")
        #     left=self.left.eval().split('\n')
        #     for i in range(0,len(left)):
        #         left[i]=left[i]+"|"+str(self.right.eval())
        #     return '\n'.join(left)
        return (str(self.left.eval()) + "|" + str(self.right.eval()))


class Or(BinaryOp):
    def eval(self):
        return (str(self.right.eval()) + '\n' + str(self.left.eval()))


class Coma(BinaryOp):
    def eval(self):
        return (str(self.right.eval()) + '\n' + str(self.left.eval()))


class createLines():
    def __init__(self, value):
        self.value = value

    def eval(self):
        global c
        c = c + 1
        # split output with \n  remove duplicate elements
        with open("C:/Users/10388/CompileRequirement/result.txt", "w") as file:
            # write to file
            file.writelines(self.value.eval())
        logging.info(c)
        logging.info(self.value.eval())
