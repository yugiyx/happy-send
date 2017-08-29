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
pin_test_values1 = []
pin_test_values2 = []
pin_test_values3 = []
i = 0
delay = 0.2
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

for x in pin_set_values:
	print('Set power to',x)
	sw.write("setswch 3 0")
	time.sleep(delay)
	sw.write("setswch 2 0")
	time.sleep(delay)
	#
	pwr_pm_rprt = float("%.3f" %float(pm_gpib.query('read1:pow?'))) 
	time.sleep(delay)
	att_voa_set = float("%.3f" %float(voa_gpib.query(':INP:ATT?')))
	time.sleep(delay)
	#
	att_voa_set = att_voa_set-(x-(pwr_pm_rprt + pin_offset))
	att_voa_set = float("%.3f" %att_voa_set) 
	if (att_voa_set >= 0 and att_voa_set <= 60):
		att_voa_set = str(att_voa_set)
		data = ':INP:ATT ' + att_voa_set
		voa_gpib.write(data)
		time.sleep(delay)
		print('Set power value OK.')
	else:
		print('Set power value failed.')	
	#
	sw.write("setswch 2 1")
	time.sleep(delay)
	#pd3 14 pd6 1A pd1 10 pd2 12
	t = 1
	test_values_buffer = []
	while t <= 3:
		tx_data = bytes.fromhex('44 A1 12 01')
		com1.write(tx_data)
		time.sleep(delay)
		dut_rd_data = str(binascii.b2a_hex(com1.read(2)))[2:-1]
		if x > 0:
			dut_rd_data = "%.2f" %(0.1 * int(dut_rd_data[:],16)) #正数十六进制转十进制结果
		else:
			dut_rd_data = "%.2f" %(-0.1 * int(hex(65536-int(dut_rd_data[:],16)),16)) #负数十六进制转十进制结果
		print(float(dut_rd_data[:]))		
		test_values_buffer.append(float(dut_rd_data[:]))
		t += 1
		time.sleep(0.5)		

	pin_test_values1.append(test_values_buffer[0])
	pin_test_values2.append(test_values_buffer[1])
	pin_test_values3.append(test_values_buffer[2])


	print (pin_test_values1)
	print (pin_test_values2)
	print (pin_test_values3)

# print (pin_set_values)
# print (pin_test_values1)
# print (pin_test_values2)
# print (pin_test_values3)

wbook = xlwt.Workbook()
sht = wbook.add_sheet('test_result',cell_overwrite_ok=True)
nwrows = 0
nwcols = 0
for x in pin_set_values:
	sht.write(nwrows,nwcols,x)
	sht.write(nwrows,nwcols+1,pin_test_values1[i])
	sht.write(nwrows,nwcols+2,pin_test_values2[i])
	sht.write(nwrows,nwcols+3,pin_test_values3[i])		
	nwrows += 1
	i += 1
	wbook.save('test result1.xls')  
print('Save data.Test finish.')