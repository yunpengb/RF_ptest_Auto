import xlwt
from datetime import datetime
from time import sleep
import xlrd
from xlutils.copy import copy

style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY,HH:MM:SS')

xls_now = 't2.xls'

#=== add sheet
wb = xlwt.Workbook()
ws = wb.add_sheet('Test Results')

#=== add title
ws.write(0, 0, datetime.now(), style1)
ws.write(0, 1, 'Bandwidth', style1)
ws.write(0, 2, 'Freq', style1)
ws.write(0, 3, 'Tx_Power', style1)
ws.write(0, 4, 'Aclr_Low', style1)
ws.write(0, 5, 'Aclr_High', style1)
ws.write(0, 6, 'Tx_EVM(64QAM).', style1)
ws.write(0, 7, 'Rx_EVM.', style1)


wb.save(xls_now)
sleep(5)

#=== open old xls
oldwb = xlrd.open_workbook(xls_now,formatting_info=True)
print oldwb

newWb = copy(oldwb)
print newWb
newWs = newWb.get_sheet(0)

#=== add new info to new xls
newWs.write(0+5, 0, datetime.now(), style1)
newWs.write(0+5, 1, 'Bandwidth', style1)
newWs.write(0+5, 2, 'Freq', style1)
newWs.write(0+5, 3, 'Tx_Power', style1)
newWs.write(0+5, 4, 'Aclr_Low', style1)
newWs.write(0+5, 5, 'Aclr_High', style1)
newWs.write(0+5, 6, 'Tx_EVM(64QAM).', style1)
newWs.write(0+5, 7, 'Rx_EVM.', style1)

newWb.save(xls_now)
print "save with same name OK"