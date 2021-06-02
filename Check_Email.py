#!/usr/bin/python
# -*- coding: utf-8 -*-

class Mail():

    def __init__(self):
        pass
    def CheckMail(self, a):

        val = False
        list1 = ['gmail.com', 'yahoo.com', 'msn.com']
        a = a.split("@")
        b = a[1]
        for i in range(len(list1)):
            if (list1[i] in b):
                val = True
                break
            else:
                val = False
                break
        return val
