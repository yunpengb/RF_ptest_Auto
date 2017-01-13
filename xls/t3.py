import xlwt
from datetime import datetime
from time import sleep
import xlrd
from xlutils.copy import copy

style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY,HH:MM:SS')

xls_now = 't3.xls'

#=== add sheet
wb = xlwt.Workbook()
ws = wb.add_sheet('Test Results')
wb.save(xls_now)
sleep(5)

def add_head(xls_in,x,first):
    # open an exist xls
    oldwb = xlrd.open_workbook(xls_in,formatting_info=True)
    #print oldwb
    newWb = copy(oldwb)
    #print newWb
    newWs = newWb.get_sheet(0)
    if first == True:
        i = 0
    else:
        i = 2
    
    newWs.write(x+i, 0, datetime.now(), style1)
    newWs.write(x+i, 1, 'Bandwidth', style1)
    newWs.write(x+i, 2, 'Freq', style1)
    newWs.write(x+i, 3, 'Tx_Power', style1)
    newWs.write(x+i, 4, 'Aclr_Low', style1)
    newWs.write(x+i, 5, 'Aclr_High', style1)
    newWs.write(x+i, 6, 'Tx_EVM(64QAM).', style1)
    newWs.write(x+i, 7, 'Rx_EVM.', style1)
    newWb.save(xls_in)
    print "save xls OK"

add_head(xls_now,0,1)
add_head(xls_now,5,0)

