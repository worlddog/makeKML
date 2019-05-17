# -*- coding: utf-8 -*-
import openpyxl
import time

t0=time.time()

wb = openpyxl.load_workbook(filename='C:/Users/weiyi/Desktop/柳州LTE工程参数总表20190517V1.1.xlsx', read_only=True)
sheet = wb['FDD']

# 最大行数
max_row = 50
row_list = []
x_list = []
y_list = []
h_list = []

for i in range(2, max_row+1):
    row_list.append('H'+str(i))  # 小区名
for i in range(2, max_row+1):
    x_list.append('J'+str(i))
for i in range(2, max_row+1):
    y_list.append('K'+str(i))
# 处理数量
for w in range(0, 49):
    print(sheet[row_list[w]].value)


filename = 'C:/Users/weiyi/Desktop/result.kml'


# 写入KML头
with open(filename, 'a',encoding='utf-8') as make_kmlhead_object:
    make_kmlhead_object.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    make_kmlhead_object.write('<kml xmlns="http://www.opengis.net/kml/2.2"> \n')
    make_kmlhead_object.write('<Document>')

# 写入数据
with open(filename, 'a',encoding='utf-8') as write_point:
    for j in range(0, max_row-1):
        x = str(sheet[x_list[j]].value)
        y = str(sheet[y_list[j]].value)
        write_point.write('<Placemark>\n')
#  写入点名称
        write_point.write('<name>'+sheet[row_list[j]].value+'</name>\n')
        write_point.write('<Point>\n')
#  写入坐标
        write_point.write('<coordinates>'+x + ','+y+'</coordinates>\n')
        write_point.write('</Point>\n')
        write_point.write('</Placemark>\n')


with open(filename, 'a',encoding='utf-8') as write_end:
    write_end.write(' </Document>')
    write_end.write(' </kml>')

t1=time.time()
t2=t1-t0

print('DONE\n'+'已处理', str(j)+'\n'+'耗时' + str(t2))

#for i in range(2, sheet.max_row):
 #   print(sheet['Di'].value)
