import pandas as pd
import openpyxl
from Constant import *


class Staff:
    def __init__(self, name):
        self.name = name
        self.job = 0
        self.is_formal_teacher = 0
        self.jiaowu_teacher = 0
        self.sale = 0
        self.polity = 0
        self.market = 0
        self.Base_salary = 0
        self.Commission = 0
        self.bonus = 0
        self.place = 0

    def cal(self):
        pass


def pre():
    p = pd.read_excel('tables/07月薪资表（结果数据）.xlsx', sheet_name='各校区员工名单一览表', header=2)
    s = []
    k = p.loc[:, ['姓名', '校区', '职务']]

    for i in k['姓名']:
        if not pd.isnull(i):
            s.append(Staff(i))

    for i in range(len(s)):
        if k['校区'][i] == '青浦':
            s[i].place = qingpu
        elif k['校区'][i] == '青浦万达茂':
            s[i].place = qingpuwandamao
        elif k['校区'][i] == '青浦二楼':
            s[i].place = qingpuerlou
        elif k['校区'][i] == '青浦三楼':
            s[i].place = qingpusanlou
        elif k['校区'][i] == '青浦吾悦':
            s[i].place = qingpuwuyue
        elif k['校区'][i] == '青浦宝龙校区':
            s[i].place = qingpubaolongxiaoqu
        elif k['校区'][i] == '茸北':
            s[i].place = rongbei
        elif k['校区'][i] == '松江万达':
            s[i].place = songjiangwanda
        elif k['校区'][i] == '大区':
            s[i].place = daqu
        elif k['校区'][i] == '文诚路':
            s[i].place = wenchenglu
        elif k['校区'][i] == '御上海':
            s[i].place = yushanghai
        elif k['校区'][i] == '颛桥':
            s[i].place = zhuanqiao

    for i in range(len(s)):
        if k['职务'][i].find('校长') != -1:
            s[i].job = xiaozhang
        elif k['职务'][i].find('市场') != -1:
            s[i].job = shichang
        elif k['职务'][i].find('顾问') != -1:
            s[i].job = guwen
        elif k['职务'][i].find('助理') != -1:
            s[i].job = zhuli
        elif k['职务'][i].find('保洁') != -1:
            s[i].job = baojie
        elif k['职务'][i].find('舞') != -1:
            s[i].job = wudaolaoshi
        elif k['职务'][i].find('模特') != -1:
            s[i].job = motelaoshi
        elif k['职务'][i].find('主持') != -1:
            s[i].job = zhuchilaoshi
        elif k['职务'][i].find('拉丁') != -1:
            s[i].job = ladinglaoshi
        elif k['职务'][i].find('负责人') != -1:
            s[i].job = fuzeren
        elif k['职务'][i].find('钢琴') != -1:
            s[i].job = gangqinlaoshi
        elif k['职务'][i].find('画') != -1:
            s[i].job = huihualaoshi
        elif k['职务'][i].find('前台') != -1:
            s[i].job = qiantai
        elif k['职务'][i].find('财务') != -1:
            s[i].job = caiwu
        elif k['职务'][i].find('HRM') != -1:
            s[i].job = HRM
        elif k['职务'][i].find('创始人') != -1:
            s[i].job = chaungshiren
        elif k['职务'][i].find('书法') != -1:
            s[i].job = shufalaoshi

    for i in range(len(s)):
        print(s[i].name, s[i].job, s[i].place, s[i].polity, s[i].market)


if __name__ == '__main__':
    pre()
