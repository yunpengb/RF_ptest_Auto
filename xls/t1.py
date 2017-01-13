import xlwt
from datetime import datetime

style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY,HH:MM:SS')

#=== add sheet
wb = xlwt.Workbook()
ws = wb.add_sheet('Test Results')

#=== add title
ws.write(0, 0, datetime.now(), style1)
ws.write(1, 1, 'No.', style1)
ws.write(1, 2, 'Tag', style1)
ws.write(1, 3, 'Tx_Power', style1)
ws.write(1, 4, 'Aclr_Low', style1)
ws.write(1, 5, 'Aclr_High', style1)
ws.write(1, 6, 'Tx_EVM(64QAM).', style1)
ws.write(1, 7, 'Rx_EVM.', style1)


wb.save('t1.xls')