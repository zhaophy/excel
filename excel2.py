#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
from xlutils.copy import copy
from openpyxl import Workbook

#list = [{'Name': 'Zara', 'Age': 7,'area':'shanghai'},{'Name': 'Herry', 'Age': 7,'area':'beijing'}]
#dict = {'Name': 'Zara', 'Age': 7, 'Class': 'First'};
data_a = [{'name':'echo_one','age':'56','rank':'A'},{'name':'echo_two','age':'47','rank':'A'},{'name':'echo_three','age':'44','rank':'B+'}]
data_b = [{'base':'sword','cat':'cat-5','level':'11'},{'base':'blade','cat':'cat-3','level':'15'},{'base':'shield','cat':'cat-11','level':'331'}]
dict = data_a[0]
print len(data_a)
print dict
rb = xlrd.open_workbook(r'ok.xls')
wb = copy(rb)
s = wb.get_sheet(0)
s.write(5,0,5)


wb.save('ok.xls')