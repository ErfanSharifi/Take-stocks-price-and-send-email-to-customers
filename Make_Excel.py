#!/usr/bin/python
# -*- coding: utf-8 -*-


import xlsxwriter

class Excel():

    def __init__(self):
        pass

    def Make(self,g, lis):
        
        name = str(g)
        lis = lis

        row = 1
        column = 0

        workbook = xlsxwriter.Workbook('C:/Users/Erfan/OneDrive/Projects/SSM_Test/Outputs/%s.xlsx'%name)
        worksheet = workbook.add_worksheet()
        worksheet.write_string(0, 0 , "نام سهم")
        worksheet.write_string(0, 1 , "تعداد سهم")
        worksheet.write_string(0, 2 , "قیمت روز")
        worksheet.write_string(0, 3 , "ارزش سهم")
        worksheet.write_string(0, 4, " کل دارایی")
        for name, num, price, val in (lis):

            worksheet.write(row, column, name)
            worksheet.write(row, column + 1, num)
            worksheet.write(row, column + 2, price)
            worksheet.write(row, column + 3, int(val))
            row +=1
        # Write a total using a formula.
        
        worksheet.write_formula(row, 4, '=SUM(D2:D100)')
        workbook.close()




