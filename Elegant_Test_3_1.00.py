import xlrd                   
import xlwt
import visa
import time
import serial
import string
import binascii

pin_offset = 0.4 
pout_offset = 1.1 
pin_set_values = []
test_values_buffer = []
pout_test_values1 = []
pout_test_values2 = []
pout_test_values3 = []
pm_report_values = []

i = 0
delay = 0.3
rm = visa.ResourceManager()
pm_gpib = rm.open_resource("GPIB0::21::INSTR")
voa_gpib = rm.open_resource("GPIB0::28::INSTR")
sw = rm.open_resource("ASRL7::INSTR")
sw.baud_rate = 115200
com1 = serial.Serial('COM1',115200)

xlsfile = 'test case.xls'   
book = xlrd.open_workbook(xlsfile)     
sheet_name=book.sheet_names() [0]       
sheet1=book.sheet_by_name(sheet_name) 
pin_set_values = sheet1.col_values(0)   



voa_gpib.write(':INP:ATT 24.76')   #设置入光为-10
time.sleep(delay)
sw.write("setswch 2 1")
time.sleep(delay)
sw.write("setswch 3 1")
time.sleep(delay)

for x in pin_set_values:
	if x >= 4096:
		tx_data ='44 A0 80 AA 00 20 40 00 00 00 ' + hex(int(x))[2:4] + ' ' + hex(int(x))[4:] 
	elif x< 4096 and x>= 256:
		tx_data ='44 A0 80 AA 00 20 40 00 00 00 ' + '0'+hex(int(x))[2:3] + ' ' + hex(int(x))[3:] 
	elif x < 256 and x >= 16:
		tx_data ='44 A0 44 ' + '00 ' + hex(int(x))[2:4]
	elif x < 16 and x >= 0:
		tx_data ='44 A0 44 ' + '00 0' + hex(int(x))[2:4]
	else :
		tx_data ='44 A0 44 ' + hex(65536 + int(x))[2:4] + ' ' + hex(65536 + int(x))[4:]  
	print('Set',tx_data)	
	tx_data = bytes.fromhex(tx_data) 
	com1.write(tx_data)

	time.sleep(delay)
	com1.reset_input_buffer()

	tx_data = bytes.fromhex('44 A1 0C 01')  ##pdtotal 0C or msout 20
	com1.write(tx_data)	
	time.sleep(delay)	

	dut_rd_data = str(binascii.b2a_hex(com1.read(2)))[2:-1]
	if dut_rd_data[0] == '0':
		dut_rd_data = "%.2f" %(0.1 * int(dut_rd_data[:],16)) 
	else:
		dut_rd_data = "%.2f" %(-0.1 * int(hex(65536-int(dut_rd_data[:],16)),16)) 
	pout_test_values1.append(float(dut_rd_data[:]))

	pwr_pm_rprt = float("%.2f" %float(pm_gpib.query('read1:pow?'))) 
	pm_report_values.append(pwr_pm_rprt + pout_offset)

print (pin_set_values)
print (pout_test_values1)
print (pm_report_values)

wbook = xlwt.Workbook()
sht = wbook.add_sheet('test_result',cell_overwrite_ok=True)
nwrows = 0
nwcols = 0
for x in pin_set_values:
	sht.write(nwrows,nwcols,x)
	sht.write(nwrows,nwcols+1,pout_test_values1[i])
	sht.write(nwrows,nwcols+2,pm_report_values[i])
	nwrows += 1
	i += 1
	wbook.save('test result3.xls')  
print('Save data.Test finish.')