import xlrd                     
import xlwt

#打开表格获取参数  
xlsfile = 'test case.xls'   
book = xlrd.open_workbook(xlsfile)     
sheet_name=book.sheet_names() [0]      
sheet1=book.sheet_by_name(sheet_name) 
input_power = sheet1.col_values(0)   
gain = sheet1.col_values(1)

case_num = list(range(len(input_power)))

for x in case_num:
	#以下为获得原始数据osa_raw_data所需要的测试步骤程序
	f = open('osa_data.txt', 'r')
	osa_raw_data = f.read()
	f.close()
	#以上为获得原始数据osa_raw_data所需要的测试步骤程序
	# 写数据到表格
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

	save_name = 'test_result_' + str(x+1) +'_input_' +  str(input_power[x]) + '_gain' +  str(gain[x]) + '.xls'
	print(save_name)
	wbook.save(save_name)  


