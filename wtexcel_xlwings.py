#coding:utf-8
import xlwings as xw
wb = xw.Book('ok.xlsx')
sht =  wb.sheets('ok')

data_a = [{'name':'echo_one','age':'56','rank':'A','Class':'43'},
          {'age':'44','rank':'B+','Class':'9','name':'echo_four','phone':'10086','add':'henan'},
          {'age':'47','rank':'A','Class':'9'}]
data_b = [{'name':'echo_one','age':'54','rank':'A','Class':'33'},
          {'age':'43','rank':'B+','Class':'3','name':'echo_four','phone':'10086','add':'henan'},
          {'age':'4437','rank':'A','Class':'4'}]

keys = data_a[0].keys()
for i in data_a:
    keys += i.keys()
newkeys = sorted(list(set(keys)))
sht.range('A1').value = newkeys


def get_v():
    vs = []
    for i in data_a:
        v = []
        for k in newkeys:
            try:
                v.append(i[k])
            except:
                v.append(None)
        vs.append(v)
    return vs
sht.range('A2').value = get_v()
wb.save('ok.xlsx')
wb.close()
