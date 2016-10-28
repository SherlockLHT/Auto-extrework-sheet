#coding:utf-8
"""Auto make extra work sheet"""

__author__ = 'Sherlock_Lin'

import pycurl
import StringIO
import re
import os
import xlwt
import xlrd
import time
import random
import sys
import itertools

reason = [u'支援工廠', u'CODE編寫',u'程式驗證']
row_one_value = [u'工號', u'加班日期', u'加班類型', u'開始日期', u'結束日期', u'用餐否', u'跨天否', u'加班時數', u'加班原因']
def help():
	print '-help	help'
	print '-o:make personal extra work sheet'
	print '		e.g. python mypycurl.py -o user password'
	print "-s:statistic everyone's extra work sheet"
	print '		e.g. python mypycurl.py -s '
def myPycurl():
	try:
		bStatus = True
		curl = pycurl.Curl()
		str = StringIO.StringIO()
		user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'
		curl.setopt(pycurl.URL, url)
		curl.setopt(pycurl.WRITEFUNCTION, str.write)
		curl.setopt(pycurl.FOLLOWLOCATION, 1)
		curl.setopt(pycurl.MAXREDIRS, 2)
		curl.setopt(pycurl.TIMEOUT, 5)
		curl.setopt(pycurl.CONNECTTIMEOUT, 60)
		curl.setopt(pycurl.USERAGENT, user_agent)
		curl.setopt(pycurl.PROXY, 'http://proxy1.local:80')
		curl.setopt(pycurl.PROXYUSERPWD, userpassword)
		curl.setopt(pycurl.USERPWD, userpassword)
		curl.setopt(pycurl.PROXYAUTH, pycurl.HTTPAUTH_NTLM)  # 代理服務器驗證
		curl.setopt(pycurl.HTTPAUTH, pycurl.HTTPAUTH_NTLM)  # http驗證
		# curl.setopt(pycurl.SSL_VERIFYPEER,False)	#有些網站需要SSL驗證
		curl.perform()
	except Exception as e:
		print e
		bStatus = False
		raise e
	finally:
		if bStatus:
			timeProcess(str.getvalue().decode('utf-8'))
		str.close()
		curl.close()
def timeProcess(data = u''):
	try:
		if u'' == data:
			return False
		pattern = re.compile(ur'<td>S\d{8}</td>'
							 ur'<td>[\u4e00-\u9fa5]{2,3}</td>'
							 ur'<td>\d{2}\/\d{2}\([\u4e00-\u9fa5]{2}\)</td>'
							 ur'<td>\d{2}\/\d{2} \d{2}:\d{2}:\d{2}  </td>'
							 ur'<td>\d{2}\/\d{2} \d{2}:\d{2}:\d{2}  </td>'
							 ur'<td>\d{2}\/\d{2} \d{2}:\d{2}:\d{2}  </td>'
							 ur'<td>\d{2}\/\d{2} \d{2}:\d{2}:\d{2}  </td>')
		items = re.findall(pattern, data)
		all_result = []
		for item in items:
			row_item = item.split('</td><td>')
			row_result = []
			row_result.append(row_item[0].replace('<td>',''))
			row_result.append(time.strftime('%Y') + row_item[2].replace('/', '').split('(')[0])
			if row_item[2].split('(')[1].split(')')[0] in u'週一週二週三週四週五周一周二周三周四周五':
				type = 1
			else:
				type = 2
			row_result.append(str(type))
			if 1 == type:
				row_result.append('1800')
				start_hour = 18
				start_min = 0
			else:
				temp = row_item[5].split(' ')[1].split(':')
				if int(temp[1]) < 30:
					start_min = 30
					row_result.append(temp[0] + '30')
					start_hour = int(temp[0])
				else:
					start_min = 0
					row_result.append(str(int(temp[0])+1).zfill(2) + '00')
					start_hour = int(str(int(temp[0])+1).zfill(2))

			temp = row_item[6].split(' ')[1].split(':')
			end_hour = int(temp[0])
			if int(temp[1]) < 30:
				end_min = 0
				row_result.append(temp[0] + '00')
			else:
				end_min = 30
				row_result.append(temp[0] + '30')
			row_result.append('N')
			row_result.append('N')
			row_result.append(str(end_hour - start_hour + 0.5 * (end_min - start_min) / 30))
			row_result.append(random.sample(reason,1))
			if 2300 < int(row_result[4]):
				row_result[4] = u'2300'
			if '0.0' == row_result[7]:
				continue
			print row_result
			all_result.append(row_result)
		writeFile(data = all_result)
	except Exception as e:
		print e
		raise
	finally:
		pass
def writeFile(data = [],file = 'file.xls'):
	try:
		if [] == data:
			return False
		if os.path.exists(file):
			os.remove(file)
		work = xlwt.Workbook('utf-8')
		sheet = work.add_sheet('sheet1', cell_overwrite_ok=True)
		style = xlwt.XFStyle()
		font = xlwt.Font()
		font.name = u'新細明體'
		font.height = 0x00F0  # 字號*20
		style.font = font
		style.num_format_str = '@'  # 數字格式為文本
		index = 0
		for value in row_one_value:
			sheet.write(0, index, value, style)
			index = index + 1
		row_index = 1
		colum_index = 0
		for row_value in data:
			for value in row_value:
				sheet.write(row_index,colum_index,value,style)
				colum_index = colum_index + 1
			row_index = row_index + 1
			colum_index = 0
		work.save(file)
	except Exception as e:
		print e
		raise e
	finally:
		pass
def arrange():
	try:
		files = os.listdir(os.getcwd())
		xls_file = []
		for file in files:
			if os.path.isfile(file) and '.xls' == os.path.splitext(file)[1]:
				xls_file.append(file)
		if not xls_file:
			return False
		all_data = []
		for file in xls_file:
			file_data = xlrd.open_workbook(file)
			table = file_data.sheet_by_index(0)
			file_row_data = table.row_values(0)
			if not file_row_data == row_one_value:
				return False
			for i in range(table.nrows):
				value = table.row_values(i)
				if not row_one_value == value:
					all_data.append(value)
		per_month_data = []
		curr_month_data = []
		month_data = []
		for i in all_data:
			month_data.append(i[1][4:6])
		month_data = list(set(month_data))
		for i in all_data:
			if month_data[0] == i[1][4:6]:
				per_month_data.append(i)
			else:
				curr_month_data.append(i)
		if per_month_data:
			writeFile(data=per_month_data,file=month_data[0]+'.xls')
		if curr_month_data:
			writeFile(data=curr_month_data, file=month_data[1] + '.xls')
		for i in curr_month_data:
			print i
		for i in per_month_data:
			print i
	except Exception as e:
		print e
		raise e
	finally:
		pass

if __name__ == "__main__":
	url = 'http://esss.sz.pegatroncorp.com/ESSSNT/Modules/Class/ClassInfo.aspx'
	args = []
	for arg in sys.argv:
		args.append(arg)
	if 1 == len(args):
		help()
	elif '-o' == args[1]:
		userpassword = args[2] + ':' + args[3]
		result = myPycurl()
		print 'FINISH'
	elif '-s' == args[1]:
		arrange()
		print 'FINISH'
	elif '-h' == args[1] or '-H' == args[1] or '-help' == args[1] or '-HELP' == args[1]:
		help()
	else:
		help()





