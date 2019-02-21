#coding:utf-8

from openpyxl import Workbook
from openpyxl import load_workbook

wb =  load_workbook('a.xlsx')
ws = wb.active

data_a = [{'name':'echo_one','age':'56','rank':'A','Class':'43'},
          {'age':'44','rank':'B+','Class':'9','name':'echo_four','phone':'10086','add':'henan'},
          {'age':'47','rank':'A','Class':'9'}]
    keys = data[0].keys() #列表中第一个字典的所有key组成的集合
    for i in data:
        keys += i.keys()
    newkeys = sorted(list(set(keys)))   #合并列表中所有字典的所有key,并去重组成的列表，此行可为excel的首行
    vs = []
    for i in data:
        v = []
        for k in newkeys:
            try:
                v.append(i[k])  ###
            except:
                v.append(None)
        vs.append(v)  #列表所有各字典的value组成的列表
    return ([newkeys] + [[]] + vs)  #key 和 value组成的新列表
for i in get_v(data_a):
    print i
    ws.append(i)  #写入excel

wb.save('a.xlsx')
