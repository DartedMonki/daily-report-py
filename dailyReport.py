#!/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "Your Name"
__version__ = "0.1.0"
__license__ = "MIT"

import xlsxwriter
import openpyxl
def main():
 
    print('Input:')
    inputStock = int(input())
    print('Output:')
    outputStock = int(input())
    print('Tanggal:')
    date = input()
    dateInt = int(date)


    yesterdayWorkbook = openpyxl.load_workbook('report-daily-tanggal-{}.xlsx'.format(dateInt-1))
    yesterdayWorksheet = yesterdayWorkbook.active
    stockAwal = yesterdayWorksheet.cell(row = dateInt+3, column = 3)
    stockAwalInt = int(stockAwal.value)

    workbook = xlsxwriter.Workbook('report-daily-tanggal-{}.xlsx'.format(date))

    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Tanggal')
    worksheet.write('A2', date)
    worksheet.write('B1', 'Input')
    worksheet.write('B2', inputStock)
    worksheet.write('C1', 'Output')
    worksheet.write('C2', outputStock)

    # header
    worksheet.write('A4', 'Tanggal')
    worksheet.write('B4', 'Stock Awal')
    worksheet.write('C4', 'Stock Akhir')

    i = 5
    # iterate last xlsx to populate yesterday datas
    while yesterdayWorksheet.cell(row = i, column = 1).value is not None:
        worksheet.write('A{}'.format(i), yesterdayWorksheet.cell(row = i, column = 1).value)
        worksheet.write('B{}'.format(i), yesterdayWorksheet.cell(row = i, column = 2).value)
        worksheet.write('C{}'.format(i), yesterdayWorksheet.cell(row = i, column = 3).value)
        i += 1
    
    todayRow = i
    worksheet.write('A{}'.format(todayRow), date)
    worksheet.write('B{}'.format(todayRow), stockAwalInt)
    worksheet.write('C{}'.format(todayRow), stockAwalInt + inputStock - outputStock)
    
    workbook.close()


if __name__ == "__main__":
    """ This is executed when run from the command line """
    main()