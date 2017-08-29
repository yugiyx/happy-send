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
pin_test_values = []
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


# tx_data = bytes.fromhex('44 A0 40 00')   #设置模式00
# com1.write(tx_data)
# time.sleep(delay)
# tx_data = bytes.fromhex('44 A1 41 11')   #设置增益control set 17
# com1.write(tx_data)
# time.sleep(delay)
# tx_data = bytes.fromhex('44 A0 46 00 64')   #设置增益为10
# com1.write(tx_data)
# time.sleep(delay)
# com1.reset_input_buffer()

for x in pin_set_values:
	print('Set power to',x)
	sw.write("setswch 3 0")
	time.sleep(delay)
	sw.write("setswch 2 0")
	time.sleep(delay)
	#
	pwr_pm_rprt = float("%.2f" %float(pm_gpib.query('read1:pow?'))) 
	time.sleep(delay)
	att_voa_set = float("%.2f" %float(voa_gpib.query(':INP:ATT?')))
	time.sleep(delay)
	#
	att_voa_set = att_voa_set-(x-(pwr_pm_rprt + pin_offset))
	att_voa_set = float("%.2f" %att_voa_set) 
	if (att_voa_set >= 0 and att_voa_set <= 60):
		att_voa_set = str(att_voa_set)
		data = ':INP:ATT ' + att_voa_set
		voa_gpib.write(data)
		time.sleep(delay)
		print('Set power value OK.')
	else:
		print('Set power value failed.')	
	
	sw.write("setswch 2 1")
	time.sleep(delay)
	sw.write("setswch 3 1")
	time.sleep(delay)
	#pd out= 0C 
	tx_data = bytes.fromhex('44 A1 0C 01')
	com1.write(tx_data)
	time.sleep(delay)
	dut_rd_data = str(binascii.b2a_hex(com1.read(2)))[2:-1]
	print(dut_rd_data)
	if x >= -10:
		dut_rd_data = "%.2f" %(0.1 * int(dut_rd_data[:],16)) #正数十六进制转十进制结果
	else:
		dut_rd_data = "%.2f" %(-0.1 * int(hex(65536-int(dut_rd_data[:],16)),16)) #负数十六进制转十进制结果
	pin_test_values.append(float(dut_rd_data[:]))
	
	pwr_pm_rprt = float("%.2f" %float(pm_gpib.query('read1:pow?'))) 
	pm_report_values.append(pwr_pm_rprt + pout_offset)

print (pin_set_values)
print (pin_test_values)
print (pm_report_values)

wbook = xlwt.Workbook()
sht = wbook.add_sheet('test_result',cell_overwrite_ok=True)
nwrows = 0
nwcols = 0
for x in pin_set_values:
	sht.write(nwrows,nwcols,x)
	sht.write(nwrows,nwcols+1,pin_test_values[i])
	sht.write(nwrows,nwcols+2,pm_report_values[i])
	nwrows += 1
	i += 1
	wbook.save('test result2.xls')  
print('Save data.Test finish.')