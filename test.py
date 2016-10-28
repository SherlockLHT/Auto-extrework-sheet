#coding:utf-8
import xlwt
import os
import time
import random

print time.strftime('%Y')
li = [u'<td>S15037407</td>'
      u'<td>\u6797\u6d77\u6fe4</td>'
      u'<td>09/26(\u5468\u4e09)</td>'
      u'<td>09/26 09:00:00  </td>'
      u'<td>09/26 18:00:00  </td>'
      u'<td>09/26 08:29:13  </td>'
      u'<td>09/26 19:02:05  </td>']
row_item = li[0].split('</td><td>')
for i in row_item:
    print i.strip('</td> /').replace('/','')
try:
    row_result = []
    row_result.append(row_item[0].replace('<td>',''))
    temp = row_item[2].replace('/','')
    aa = temp.split('(')
    date = time.strftime('%Y') + row_item[2].replace('/','').split('(')[0]
    if row_item[2].split('(')[1].split(')')[0] in u'週一週二週三週四週五周一周二周三周四周五':
        type = 1
    else:
        type = 2
    temp = row_item[6].split(' ')[1].split(':')
    print temp
    if int(temp[1]) < 30:
        date = temp[0] + '00'
    else:
        date = temp[0] + '30'
    reason = [u'1', u'2',u'3',u'4',u'5']

    temp_1 = 18
    temp_2 = 19
    temp_3 = 00
    temp_4 = 00
    temp = (temp_2 - temp_1) + 0.5 * (temp_4 - temp_3) / 30
    print temp
    print random.sample(reason,1)
except Exception as e:
    print e
    raise e
finally:
    pass
