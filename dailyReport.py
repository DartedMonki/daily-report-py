#!/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "Your Name"
__version__ = "0.1.0"
__license__ = "MIT"

import xlsxwriter
import openpyxl
import os
def main():
    def write_sheet(sheetname, date, inputStock, outputStock, yesterdayWorksheet, isWithToday):
        worksheet = workbook.add_worksheet(sheetname)
        # top header
        worksheet.write('A1', 'Tanggal')
        worksheet.write('B1', 'Input')
        worksheet.write('C1', 'Output')

        worksheet.write('A2', date)
        worksheet.write('B2', inputStock)
        worksheet.write('C2', outputStock)

        # bottom header
        worksheet.write('A4', 'Tanggal')
        worksheet.write('B4', 'Stock Awal')
        worksheet.write('C4', 'Stock Akhir')

        i = 5
        # iterate previous xlsx to populate yesterday datas
        while yesterdayWorksheet.cell(row = i, column = 1).value is not None:
            worksheet.write('A{}'.format(i), yesterdayWorksheet.cell(row = i, column = 1).value)
            worksheet.write('B{}'.format(i), yesterdayWorksheet.cell(row = i, column = 2).value)
            worksheet.write('C{}'.format(i), yesterdayWorksheet.cell(row = i, column = 3).value)
            i += 1
        
        if isWithToday:
            todayRow = i
            worksheet.write('A{}'.format(todayRow), date)
            worksheet.write('B{}'.format(todayRow), stockAwalInt)
            worksheet.write('C{}'.format(todayRow), stockAwalInt + inputStock - outputStock)
 
    print('Input:')
    inputStock = int(input())
    print('Output:')
    outputStock = int(input())
    print('Tanggal:')
    date = input()
    dateInt = int(date)
    print('Nama Sheet:')
    sheetNameInput = input()

    newFileName = 'report-daily-tanggal-{}.xlsx'.format(date)
    yesterdayWorkbook = openpyxl.load_workbook('report-daily-tanggal-{}.xlsx'.format(dateInt-1))

    # check if data already exists. use exixting data
    files = [f for f in os.listdir('.') if os.path.isfile(f)]
    if newFileName in files:
        yesterdayWorkbook = openpyxl.load_workbook(newFileName)

    yesterdayWorksheet = yesterdayWorkbook[sheetNameInput]
    stockAwal = yesterdayWorksheet.cell(row = dateInt+3, column = 3)
    stockAwalInt = int(stockAwal.value)

    workbook = xlsxwriter.Workbook(newFileName)
    write_sheet(yesterdayWorksheet.title, date, inputStock, outputStock, yesterdayWorksheet, True)

    # construct unedited yesterday sheets
    yesterdaySheetnames = yesterdayWorkbook.sheetnames
    filteredYesterdaySheetnames = filter(lambda x: x != sheetNameInput, yesterdaySheetnames)
    for yesterdaySheetname in filteredYesterdaySheetnames:
        currentWorksheet = yesterdayWorkbook[yesterdaySheetname]
        write_sheet(yesterdaySheetname, currentWorksheet.cell(row = 2, column = 1).value, currentWorksheet.cell(row = 2, column = 2).value, currentWorksheet.cell(row = 2, column = 3).value, currentWorksheet, False)
    
    workbook.close()


if __name__ == "__main__":
    """ This is executed when run from the command line """
    main()