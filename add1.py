#coding:utf-8

from openpyxl import Workbook
from openpyxl import load_workbook
import os



data_a = [{'name':'大家好','age':'56','rank':'A','Class':'43'},
          {'age':'44','rank':'B+','Class':'9','name':'echo_four','phone':'10086','add':'henan'},
          {'age':'47','rank':'A','Class':'9'},
          {'color':'grey'},
          {'language':'中国语'}]
data_b = [{u'上海电气电站设备有限公司汽轮机厂': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=14'}, {u'上海电气电站设备有限公司发电机厂': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=17'}, {u'上海电气电站设备有限公司电站辅机厂': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=21'}, {u'上海电气斯必克工程技术有限公司': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=30'}, {u'上海锅炉厂有限公司': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=24'}, {u'上海电气集团上海电机厂有限公司': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=12'}, {u'上海电气富士电机电气技术有限公司': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=84'}, {u'上海电气电站工程公司': 'http://www.shanghai-electric.com/PG/Pages/companies/EngineeringCompany.aspx?cid=19'}, {u'上海电气电站服务公司': 'http://www.shanghai-electric.com/PG/Pages/companies/servicecompany.aspx?cid=81'}, {u'上海电气风电设备有限公司': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=22'}, {u'上海电气电站环保工程有限公司': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=26'}, {u'上海电气海水淡化工程技术公司': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=27'}, {u'上海电气电站设备有限公司临港工厂': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=16'}, {u'上海人民电器厂': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=76'}, {u'上海电器陶瓷厂有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=77'}, {u'上海电气输配电工程成套有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=79'}, {u'上海电气输配电试验中心有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=80'}, {u'上海电气电力电子有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=78'}, {u'上海南桥变压器有限责任公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=71'}, {u'上海纳杰电气成套有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=75'}, {u'上海飞航电线电缆有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=68'}, {u'上海大华电器设备有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=73'}, {u'上海南华兰陵电气有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=72'}, {u'上海捷锦电力新材料有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=69'}, {u'吴江变压器有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=70'}, {u'上海华普电缆有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=67'}, {u'上海资文建设工程咨询有限公司': 'http://www.shanghai-electric.com/PTD/Pages/companies/company.aspx?cid=85'}, {u'上海电气核电设备有限公司': 'http://www.shanghai-electric.com/NP/Pages/CompanyIntro.aspx?cid=1'}, {u'上海第一机床厂有限公司': 'http://www.shanghai-electric.com/NP/Pages/CompanyIntro.aspx?cid=5'}, {u'上海电气凯士比核电泵阀有限公司': 'http://www.shanghai-electric.com/NP/Pages/CompanyIntro.aspx?cid=4'}, {u'上海核电技术装备有限公司': 'http://www.shanghai-electric.com/NP/Pages/CompanyIntro.aspx?cid=6'}, {u'上海凯士比泵有限公司': 'http://www.shanghai-electric.com/NP/Pages/CompanyIntro.aspx?cid=7'}, {u'上海重型机器厂有限公司': 'http://www.shanghai-electric.com/HIG/Pages/companies/company.aspx?cid=34'}, {u'上海船用曲轴有限公司': 'http://www.shanghai-electric.com/HIG/Pages/companies/company.aspx?cid=36'}, {u'上海市机械制造工艺研究所有限公司': 'http://www.shanghai-electric.com/HIG/Pages/companies/company.aspx?cid=44'}, {u'上海电气环保集团': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=43'}, {u'上海市机电设计研究院有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=43'}, {u'上海电气环保热电（南通）有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=43'}, {u'上海电气南通水处理有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=43'}, {u'上海轨道交通设备发展有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=41'}, {u'上海阿尔斯通交通设备有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=41'}, {u'上海阿尔斯通交通电气有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=41'}, {u'上海电气集团股份有限公司中央研究院': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=28'}, {u'上海电气自动化集团': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=46'}, {u'上海电气金融集团': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=87'}, {u'上海电气集团财务有限责任公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=32'}, {u'上海电气租赁有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=33'}, {u'上海电气保险经纪有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=47'}, {u'上海电气临港重型机械装备有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=45'}, {u'上海电气通讯技术有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=137'}, {u'上海电气风电集团': 'http://www.shanghai-electric.com/PG/Pages/companies/company.aspx?cid=22'}, {u'上海三菱电梯有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=5'}, {u'上海机床厂有限公司': 'http://www.shanghai-electric.com/Pages/pns/details.aspx?pid=13'}, {u'上海机床厂有限公司': 'http://www.shanghai-electric.com/Pages/pns/details.aspx?pid=13'}, {u'上海电气泰雷兹交通自动化系统有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=140'}, {u'上海电气国际经济贸易有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=35'}, {u'上海发那科机器人有限公司': 'http://www.shanghai-electric.com/Pages/companies/company.aspx?cid=51'}]

def get_v(data):
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
    return ([newkeys]  + vs)  #key 和 B
def insert(data,file):
    wb = load_workbook(file)
    ws = wb.active
    for i in get_v(data):
        print i
        ws.append(i)
    wb.save(file)
insert(data_a,'F:\\excel\\ab.xlsx')




