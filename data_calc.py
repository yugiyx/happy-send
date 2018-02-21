import glob
import xlrd
import xlwt
from xlutils.copy import copy

# Add the formula to the single raw data and analysize it
file_list = glob.glob('*pin*.xls')
i = 1
for name in file_list:
    print('Process', i, name)
    book = xlrd.open_workbook(name)
    wbook = copy(book)
    wsheet = wbook.get_sheet(0)

    wsheet.write(2, 11, 'IN')
    wsheet.write(2, 12, 'OUT')
    wsheet.write(2, 13, 'AVERAGE')
    wsheet.write(2, 14, 'GMAX')
    wsheet.write(2, 15, 'GMIN')
    wsheet.write(2, 16, 'GF')
    wsheet.write(2, 17, 'GF+')
    wsheet.write(2, 18, 'GF-')
    wsheet.write(2, 19, 'MAX NF')
    wsheet.write(2, 20, 'TILT')
    wsheet.write(1, 11, 'SLOPE')
    wsheet.write(3, 11, xlwt.Formula('B49'))
    wsheet.write(3, 12, xlwt.Formula('B50'))
    wsheet.write(3, 13, xlwt.Formula('AVERAGE(G2:G48)'))
    wsheet.write(3, 14, xlwt.Formula('MAX(G2:G48)'))
    wsheet.write(3, 15, xlwt.Formula('MIN(G2:G48)'))
    wsheet.write(3, 16, xlwt.Formula('O4-P4'))
    wsheet.write(3, 17, xlwt.Formula('O4-N4'))
    wsheet.write(3, 18, xlwt.Formula('P4-N4'))
    wsheet.write(3, 19, xlwt.Formula('MAX(H2:H48)'))
    wsheet.write(3, 20, xlwt.Formula('M2*(B48-B2)'))
    wsheet.write(1, 12, xlwt.Formula('SLOPE(G2:G49,B2:B49)'))
    wbook.save(name)
    i += 1

# Collect all analysis data to a sheet
file_list = glob.glob('*pin*.xls')
i = 1
for name in file_list:
    print('Process', i, name)
    xlsfile = name
    book = xlrd.open_workbook(xlsfile, formatting_info=True)
    sheet1 = book.sheet_by_index(0)
    pin = sheet1.cell(3, 11).value
    pout = sheet1.cell(3, 12).value
    gain_average = sheet1.cell(3, 13).value
    gain_max = sheet1.cell(3, 14).value
    gain_min = sheet1.cell(3, 15).value
    gf = sheet1.cell(3, 16).value
    gf_p = sheet1.cell(3, 17).value
    gf_m = sheet1.cell(3, 18).value
    nf = sheet1.cell(3, 19).value
    tilt = sheet1.cell(3, 20).value
    print(pin)
    book = xlrd.open_workbook('test_result_all.xls')
    wbook = copy(book)
    wsheet = wbook.get_sheet(0)
    wsheet.write(i, 0, name)
    wsheet.write(i, 1, pin)
    wsheet.write(i, 2, pout)
    wsheet.write(i, 3, gain_average)
    wsheet.write(i, 4, gain_max)
    wsheet.write(i, 5, gain_min)
    wsheet.write(i, 6, gf)
    wsheet.write(i, 7, gf_p)
    wsheet.write(i, 8, gf_m)
    wsheet.write(i, 9, nf)
    wsheet.write(i, 10, tilt)
    wbook.save('test_result_all.xls')
    i += 1
