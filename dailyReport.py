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
    def write_sheet(sheetname, date, inputStock, outputStock, yesterdayWorksheet, todayNotes, isWithToday):
        worksheet = workbook.add_worksheet(sheetname)
        # top header
        worksheet.write('A1', 'Tanggal')
        worksheet.write('B1', 'Input')
        worksheet.write('C1', 'Output')

        worksheet.write('A2', date)
        worksheet.write('B2', inputStock)
        worksheet.write('C2', outputStock)

        # bottom header
        worksheet.write('A4', 'Nomor')
        worksheet.write('B4', 'Stock Awal')
        worksheet.write('C4', 'Stock Akhir')
        worksheet.write('D4', 'Tanggal')
        worksheet.write('E4', 'Notes')

        i = 5
        # iterate previous xlsx to populate yesterday datas
        while yesterdayWorksheet.cell(row = i, column = 1).value is not None:
            worksheet.write('A{}'.format(i), yesterdayWorksheet.cell(row = i, column = 1).value)
            worksheet.write('B{}'.format(i), yesterdayWorksheet.cell(row = i, column = 2).value)
            worksheet.write('C{}'.format(i), yesterdayWorksheet.cell(row = i, column = 3).value)
            worksheet.write('D{}'.format(i), yesterdayWorksheet.cell(row = i, column = 4).value)
            worksheet.write('E{}'.format(i), yesterdayWorksheet.cell(row = i, column = 5).value)
            i += 1
        
        if isWithToday:
            todayRow = i
            worksheet.write('A{}'.format(todayRow), yesterdayWorksheet.cell(row = i-1 , column = 1).value + 1)
            worksheet.write('B{}'.format(todayRow), stockAwalInt)
            worksheet.write('C{}'.format(todayRow), stockAwalInt + inputStock - outputStock)
            worksheet.write('D{}'.format(todayRow), date)
            worksheet.write('E{}'.format(todayRow), todayNotes)
 
    print('Input:')
    inputStock = int(input())
    print('Output:')
    outputStock = int(input())
    print('Tanggal:')
    date = input()
    dateInt = int(date)
    print('Nama Sheet:')
    sheetNameInput = input()
    print('Notes:')
    todayNotes = input()

    newFileName = 'report-daily-tanggal-{}.xlsx'.format(date)
    yesterdayWorkbook = openpyxl.load_workbook('report-daily-tanggal-{}.xlsx'.format(dateInt-1))

    # check if data already exists. use exixting data
    files = [f for f in os.listdir('.') if os.path.isfile(f)]
    if newFileName in files:
        yesterdayWorkbook = openpyxl.load_workbook(newFileName)

    yesterdayWorksheet = yesterdayWorkbook[sheetNameInput]
    latestRow = 5
    while yesterdayWorksheet.cell(row = latestRow, column = 1).value is not None:
        latestRow += 1
    stockAwal = yesterdayWorksheet.cell(row = latestRow - 1, column = 3)
    stockAwalInt = int(stockAwal.value)

    workbook = xlsxwriter.Workbook(newFileName)
    write_sheet(yesterdayWorksheet.title, date, inputStock, outputStock, yesterdayWorksheet, todayNotes, True)

    # construct unedited yesterday sheets
    yesterdaySheetnames = yesterdayWorkbook.sheetnames
    filteredYesterdaySheetnames = filter(lambda x: x != sheetNameInput, yesterdaySheetnames)
    for yesterdaySheetname in filteredYesterdaySheetnames:
        currentWorksheet = yesterdayWorkbook[yesterdaySheetname]
        write_sheet(yesterdaySheetname, currentWorksheet.cell(row = 2, column = 1).value, currentWorksheet.cell(row = 2, column = 2).value, currentWorksheet.cell(row = 2, column = 3).value, currentWorksheet, "", False)
    
    workbook.close()


if __name__ == "__main__":
    """ This is executed when run from the command line """
    main()