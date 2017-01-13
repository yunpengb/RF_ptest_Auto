import xlwt
from datetime import datetime
from time import sleep
import xlrd
from xlutils.copy import copy

style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY,HH:MM:SS')

sheetName = 'Test Results'
xls_now = 't4.xls'

#=== add sheet
wb = xlwt.Workbook()
ws = wb.add_sheet('Test Results')
wb.save(xls_now)
sleep(5)

def add_empty_xls(sheet_name,xls_name):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)
    wb.save(xls_name)
    sleep(5)

def add_head(xls_in,x):
    # open an exist xls
    oldwb = xlrd.open_workbook(xls_in,formatting_info=True)
    #print oldwb
    newWb = copy(oldwb)
    #print newWb
    newWs = newWb.get_sheet(0)
    
    newWs.write(x,0, datetime.now(), style1)
    newWs.write(x,1, 'Bandwidth', style1)
    newWs.write(x,2, 'Freq', style1)
    newWs.write(x,3, 'Tx_Power', style1)
    newWs.write(x,4, 'Aclr_Low', style1)
    newWs.write(x,5, 'Aclr_High', style1)
    newWs.write(x,6, 'Tx_EVM(64QAM).', style1)
    newWs.write(x,7, 'Rx_EVM.', style1)
    newWb.save(xls_in)
    print "save xls OK"

add_empty_xls(sheetName,xls_now)
add_head(xls_now,0)
add_head(xls_now,5)

