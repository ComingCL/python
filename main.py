import pandas as pd
import openpyxl
from Constant import *
from teacher import *


class Staff:
    def __init__(self, name):
        self.name = name
        self.job = 0
        self.is_formal_teacher = 0
        self.jiaowu_teacher = 0
        self.sale = 0
        self.polity = 0
        self.market = 0
        self.base_salary = 0
        self.Commission = 0
        self.bonus = 0
        self.place = 0
        self.yuexiaoke = 0
        self.tiyanke = 0
        self.tiyankekeshifei = 0
        self.zhengshike = 0
        self.zhengshikekeshifei = 0
        self.xinsheng_above24 = 0
        self.xinsheng_below24 = 0
        self.xufei = 0
        self.xiaokejine = 0
        self.jintie = 0
        self.ans = 0
        self.shitingke = 0


def pre():
    p = pd.read_excel('07月薪资表（结果数据）.xlsx', sheet_name='各校区员工名单一览表', header=2)
    xiaoke = pd.read_excel('7月消课量表.xls')
    s = []
    s2 = []
    # print(xiaoke['上课次数'])
    k = p.loc[:, ['姓名', '校区', '职务']]
    k2 = xiaoke.loc[:, ['教师', '学生课时']]
    for i in k['姓名']:
        if not pd.isnull(i):
            s.append(Staff(i))
    for i in s:
        for j in range(len(k2)):
            if(i.name.find(k2['教师'][j]) != -1 and
                    not pd.isnull(k2['教师'][j])):
                i.yuexiaoke += k2['学生课时'][j]
                break

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
        classes = 0
        if s[i].name.find('小雨') != -1:
            s[i].zhengshike += 278
            s[i].tiyanke += 25
        if s[i].name.find('高*爽') != -1:
            s[i].zhengshike += 86.5
            s[i].zhengshikekeshifei += 43250
        if s[i].name.find('董*理') != -1:
            s[i].zhengshike += 6
        if s[i].name.find('周*琳') != -1:
            s[i].zhengshike += 7.5
            s[i].tiyanke += 3
        if s[i].name.find('小羊') != -1:
            s[i].zhengshike += 85
        if s[i].name.find('wa') != -1:
            s[i].zhengshike += 80
            s[i].zhengshikekeshifei += 9600
            s[i].zhengshikekeshifei += 2280
        if s[i].name.find('孙*彬') != -1:
            s[i].zhengshike += 74
        if s[i].name.find('贾') != -1:
            s[i].zhengshike += 163.5
        if s[i].name.find('李*冉') != -1:
            s[i].zhengshike += 123
        if s[i].name.find('叶子') != -1:
            s[i].zhengshike += 61
        if s[i].name.find('文文') != -1:
            s[i].zhengshike += 39
        if s[i].name.find('宋*心') != -1:
            s[i].zhengshike += 4
        if s[i].name.find('胡*妮') != -1:
            s[i].zhengshike += 47
        if s[i].name.find('月月') != -1:
            s[i].zhengshike += 87
        if s[i].name.find('文文') != -1:
            s[i].zhengshike += 51
            s[i].zhengshike += 34
        if s[i].name.find('果果') != -1:
            s[i].zhengshike += 119
        if s[i].name.find('泡泡') != -1:
            s[i].zhengshike += 859 - 69
            s[i].zhengshike += 54
            s[i].zhengshike += 5
            s[i].bonus += 3 * 120
            s[i].bonus += 120 * 16
        if s[i].name.find('麦子') != -1:
            s[i].zhengshike += 54
        if s[i].name.find('沈') != -1 and s[i].job == wudaolaoshi:
            s[i].zhengshike += 108
            s[i].zhengshikekeshifei += 12960
        if s[i].name.find('孙') != -1 and s[i].job == wudaolaoshi:
            s[i].zhengshike += 8
        if s[i].name.find('梁') != -1:
            s[i].zhengshike += 130.5
        if s[i].name.find('韩') != -1 and s[i].job == wudaolaoshi:
            s[i].zhengshike += 5
        if s[i].name.find('颜') != -1:
            s[i].zhengshike += 163.5
        if (s[i].name.find('杨') != -1 and
                s[i].job == wudaolaoshi and s[i].place == wenchenglu):
            s[i].zhengshike += 15
        if s[i].name.find('wa') != -1:
            s[i].zhengshike += 40.5
            s[i].zhengshikekeshifei += 4860
        if s[i].name.find('毛') != -1:
            s[i].zhengshike += 28
        if s[i].name.find('周*雅') != -1:
            s[i].zhengshike += 10
        if s[i].name.find('饼干') != -1:
            s[i].zhengshike += 10
        if (s[i].name.find('黄') != -1 and
                s[i].job == wudaolaoshi and s[i].place == wenchenglu):
            s[i].zhengshike += 64
            s[i].zhengshikekeshifei += 7680
        if s[i].name.find('高') != -1 and s[i].job == huihualaoshi:
            s[i].zhengshike += 71
        if s[i].job == baojie:
            s[i].ans += 3170
        if s[i].name.find('韦*萍') != -1:
            s[i].zhengshike += 71
        if s[i].name.find('贝') != -1:
            s[i].zhengshike += 3
        if s[i].name.find('张*婷') != -1:
            s[i].zhengshike += 64
            s[i].shitingke += 9
            s[i].zhengshike += 7
        if s[i].name.find('高*爽') != -1:
            s[i].bonus += 2 * 120
            s[i].zhengshike += 2
        if s[i].name.find('朱') != -1:
            s[i].shitingke += 11
            s[i].zhengshike += 27
            s[i].bonus += 240
            s[i].zhengshikekeshifei += 3240
            s[i].bonus += 660
            s[i].bonus += 540
        if s[i].name.find('董') != -1 and s[i].job == wudaolaoshi:
            s[i].zhengshike += 7
        if s[i].name.find('熊') != -1:
            s[i].zhengshike += 31
        if s[i].name.find('玺') != -1:
            s[i].zhengshike += 5
        if s[i].name.find('潘') != -1:
            s[i].zhengshike += 12
        if s[i].name.find('高*爽') != -1:
            s[i].zhengshike += 9
            s[i].bonus += 4 * 120
        if s[i].name.find('朱') != -1:
            s[i].zhengshike += 6
        if s[i].name.find('项') != -1:
            s[i].zhengshike += 54
        if s[i].name.find('高*翔') != -1:
            s[i].zhengshike += 272
        if s[i].name.find('麦子') != -1:
            s[i].zhengshike += 166.5

    for i in range(len(s)):
        if s[i].job == huihualaoshi or s[i].job == shufalaoshi:
            s[i].base_salary += 3000
            s[i].jintie += 1000
            s[i].zhengshikekeshifei += 20
            s[i].tiyankekeshifei += 10
            s[i].base_salary += s[i].tiyankekeshifei * s[i].tiyanke
            s[i].base_salary += s[i].zhengshikekeshifei * s[i].zhengshike
            s[i].base_salary += 100 * s[i].xinsheng_above24
            s[i].base_salary += 50 * s[i].xinsheng_below24
            s[i].base_salary += 50 * s[i].xufei

        elif s[i].job == gangqinlaoshi:
            s[i].base_salary += 0
            s[i].Commission += 0.5 * s[i].xiaokejine
        elif s[i].job == wudaolaoshi:
            if s[i].is_formal_teacher:
                s[i].base_salary += 3000
            if s[i].jiaowu_teacher:
                s[i].base_salary += 1000
            # 某节课某班级人数 >= 4 以及刚开班前两个月 待定
                # if s[i].yuexiaoke > 1000:
                #     s[i].zhengshikekeshifei += 60
                #     s[i].bonus += 2000
                # elif s[i].yuexiaoke > 900:
                #     s[i].zhengshikekeshifei += 50
                #     s[i].bonus += 2000
                # elif s[i].yuexiaoke > 800:
                #     s[i].zhengshikekeshifei += 40
                #     s[i].bonus += 1500
                # elif s[i].yuexiaoke > 700:
                #     s[i].zhengshikekeshifei += 30
                #     s[i].bonus += 1000
                # elif s[i].yuexiaoke > 600:
                #     s[i].zhengshikekeshifei += 20
                #     s[i].bonus += 500
                # elif s[i].yuexiaoke > 500:
                #     s[i].zhengshikekeshifei += 10
            # 班级人数 < 4 待定
            s[i].base_salary += s[i].tiyanke * 60
            s[i].base_salary += s[i].xinsheng_above24 * 80
            s[i].base_salary += s[i].xinsheng_below24 * 40
            s[i].base_salary += s[i].xufei * 30

        elif s[i].job == zhuchilaoshi or s[i].job == motelaoshi:
            if s[i].is_formal_teacher:
                s[i].base_salary += 3000
            if s[i].jiaowu_teacher:
                s[i].base_salary += 1000
    for i in s:
        print(i.name, i.yuexiaoke)


if __name__ == '__main__':
    pre()
