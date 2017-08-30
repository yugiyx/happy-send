import xlrd                     
import xlwt
import visa
import time
import serial
import string
import binascii

pin_offset = 0.4 
pout_offset = 1.1 
i = 0
delay = 0.2
rm = visa.ResourceManager()
osa_gpib = rm.open_resource("GPIB0::20::INSTR")
pm_gpib = rm.open_resource("GPIB0::21::INSTR")
voa_gpib = rm.open_resource("GPIB0::28::INSTR")
sw = rm.open_resource("ASRL7::INSTR")
sw.baud_rate = 115200
com1 = serial.Serial('COM1',115200)

#打开表格获取参数  
xlsfile = 'test case.xls'   
book = xlrd.open_workbook(xlsfile)     
sheet_name=book.sheet_names()[0]      
sheet1=book.sheet_by_name(sheet_name) 
pin_set_values = sheet1.col_values(0)   
gain_set_values = sheet1.col_values(1)
print(gain_set_values)
case_num = list(range(len(pin_set_values)))
print(case_num)

tx_data = bytes.fromhex('44 A0 40 00')   #设置模式00 01 02
com1.write(tx_data)
time.sleep(delay)
tx_data = bytes.fromhex('44 A0 41 11')   #设置增益control set 17
com1.write(tx_data)
time.sleep(delay)

for x in case_num:
	#调试用
	# f = open('osa_data.txt', 'r')
	# osa_raw_data = f.read()
	# f.close()
	#以下为获得原始数据osa_raw_data所需要的测试步骤程序
	#设置增益
	set_gain = int(gain_set_values[x])
	print('case',x+1,'start')
	print('Set gain to',0.1*set_gain)	
	# print(set_gain)
	if set_gain < 256 and set_gain >= 16:
		tx_data ='44 A0 46 ' + '00 ' + hex(set_gain)[2:4]
	elif set_gain < 16 and set_gain >= 0:
		tx_data ='44 A0 46 ' + '00 0' + hex(set_gain)[2:4]
	else :
		tx_data ='44 A0 46 ' + hex(65536 + set_gain)[2:4] + ' ' + hex(65536 + set_gain)[4:]  
	# print(tx_data)	
	tx_data = bytes.fromhex(tx_data)   #设置增益为10
	com1.write(tx_data)
	time.sleep(delay)
	#设置tilt
	#
	#设置中间级插损	
	#
	#设置入光功率
	print('Set power to',pin_set_values[x])
	sw.write("setswch 3 0")
	time.sleep(delay)
	sw.write("setswch 2 0")
	time.sleep(delay)
	
	pwr_pm_rprt = float("%.3f" %float(pm_gpib.query('read1:pow?'))) 
	time.sleep(delay)
	att_voa_set = float("%.3f" %float(voa_gpib.query(':INP:ATT?')))
	time.sleep(delay)
	#
	att_voa_set = att_voa_set-(pin_set_values[x]-(pwr_pm_rprt + pin_offset))
	att_voa_set = float("%.3f" %att_voa_set) 
	if (att_voa_set >= 0 and att_voa_set <= 60):
		att_voa_set = str(att_voa_set)
		data = ':INP:ATT ' + att_voa_set
		voa_gpib.write(data)
		time.sleep(delay)
		print('Set power value OK.')
	else:
		print('Set power value failed.')	
	#扫描TraceA
	sw.write("setswch 4 1")
	time.sleep(delay)		
	voa_gpib.write()#激活A
	voa_gpib.write()#设置A Write
	voa_gpib.write()#扫描A
	voa_gpib.write()#设置A Fix	
	#扫描TraceB
	sw.write("setswch 2 1")
	time.sleep(delay)
	sw.write("setswch 3 1")
	time.sleep(delay)

	voa_gpib.write()#激活B
	voa_gpib.write()#设置B Write
	voa_gpib.write()#扫描B
	voa_gpib.write()#设置B Fix
	sw.write("setswch 3 0")
	time.sleep(delay)	
	sw.write("setswch 4 0")
	time.sleep(delay)	
	#获取数据
	voa_gpib.write()#执行EDFA-NF程序
	osa_raw_data = voa_gpib.query()#获取osa_raw_data

	#以上为获得原始数据osa_raw_data所需要的测试步骤程序
	写数据到表格
	wbook = xlwt.Workbook()
	sht = wbook.add_sheet('test_result',cell_overwrite_ok=True)
	print('Channel count',osa_raw_data[5:7],len(osa_raw_data))
	i = 0
	m = 8
	n = 24
	nwrows = 1
	nwcols = 1
	while i < 48:
		j = 0
		osa_data_buffer = []	
		while j < 7:
			if j == 0:
				osa_data_buffer.append('%.2f' %((float(osa_raw_data[m:n]))*10**9))
			else:
				osa_data_buffer.append('%.3f' %(float(osa_raw_data[m:n])))
			m += 17
			n += 17
			j += 1
		i += 1
		# print('Write',i,osa_data_buffer)
	#写入excel
		sht.write(nwrows,nwcols-1,i)
		sht.write(nwrows,nwcols+0,float(osa_data_buffer[0]))
		sht.write(nwrows,nwcols+1,float(osa_data_buffer[1]))
		sht.write(nwrows,nwcols+2,float(osa_data_buffer[2]))	
		sht.write(nwrows,nwcols+3,float(osa_data_buffer[3]))
		sht.write(nwrows,nwcols+4,float(osa_data_buffer[4]))
		sht.write(nwrows,nwcols+5,float(osa_data_buffer[5]))
		sht.write(nwrows,nwcols+6,float(osa_data_buffer[6]))
		nwrows += 1
	# 写入表头
		sht.write(0,0,'ch num')
		sht.write(0,1,'center wl')
		sht.write(0,2,'input_power')	
		sht.write(0,3,'output_power')
		sht.write(0,4,'ase')		
		sht.write(0,5,'resoln')
		sht.write(0,6,'gain')	
		sht.write(0,7,'nf')

	save_name = 'test_result_' + str(x+1) +'_input_' +  str(pin_set_values[x]) + '_gain' +  str(0.1*set_gain) + '.xls'
	print(save_name)
	wbook.save(save_name)  