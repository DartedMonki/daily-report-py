#!/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "Your Name"
__version__ = "0.1.0"
__license__ = "MIT"

import openpyxl
import os
def main():
    def write_sheet(sheetname, date, inputStock, outputStock, yesterdayWorksheet, todayNotes):
        worksheet = workbook.create_sheet(sheetname)
        # top header
        worksheet['A1'] = 'Tanggal'
        worksheet['B1'] = 'Input'
        worksheet['C1'] = 'Output'

        worksheet['A2'] = date
        worksheet['B2'] = inputStock
        worksheet['C2'] = outputStock

        # bottom header
        worksheet['A4'] = 'Nomor'
        worksheet['B4'] = 'Stock Awal'
        worksheet['C4'] = 'Stock Akhir'
        worksheet['D4'] = 'Tanggal'
        worksheet['E4'] = 'Notes'

        i = 5
        # iterate previous xlsx to populate yesterday datas
        while yesterdayWorksheet.cell(row = i, column = 1).value is not None:
            worksheet['A{}'.format(i)] = yesterdayWorksheet.cell(row = i, column = 1).value
            worksheet['B{}'.format(i)] = yesterdayWorksheet.cell(row = i, column = 2).value
            worksheet['C{}'.format(i)] = yesterdayWorksheet.cell(row = i, column = 3).value
            worksheet['D{}'.format(i)] = yesterdayWorksheet.cell(row = i, column = 4).value
            worksheet['E{}'.format(i)] = yesterdayWorksheet.cell(row = i, column = 5).value
            i += 1
        
        todayRow = i
        worksheet['A{}'.format(todayRow)] = yesterdayWorksheet.cell(row = i-1 , column = 1).value + 1
        worksheet['B{}'.format(todayRow)] = stockAwalInt
        worksheet['C{}'.format(todayRow)] = stockAwalInt + inputStock - outputStock
        worksheet['D{}'.format(todayRow)] = date
        worksheet['E{}'.format(todayRow)] = todayNotes
 
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

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    write_sheet(yesterdayWorksheet.title, date, inputStock, outputStock, yesterdayWorksheet, todayNotes)

    # construct unedited yesterday sheets
    yesterdaySheetnames = yesterdayWorkbook.sheetnames
    filteredYesterdaySheetnames = filter(lambda x: x != sheetNameInput, yesterdaySheetnames)
    for yesterdaySheetname in filteredYesterdaySheetnames:
        currentWorksheet = yesterdayWorkbook[yesterdaySheetname]
        currentWorksheet._parent = workbook
        workbook._add_sheet(currentWorksheet)
    
    workbook.save(newFileName)

if __name__ == "__main__":
    """ This is executed when run from the command line """
    main()