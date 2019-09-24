import sqlite3
from openpyxl import Workbook
from datetime import date


conn = sqlite3.connect(r"Data\data.db")
cur = conn.cursor()


class Selesman():
    def __init__(self):
        self.month = None               # 考核月份
        self.ye_wu_yuan = None          # 业务员
        self.xing_ming = None           # 姓名
        self.gong_hao = None            # 工号
        self.ru_zhi_shi_jian = None     # 入职时间
        self.si_ling_gong_zi = None     # 司龄工资
        self.kao_he_lei_xing = None     # 考核类型
        self.he_tong_lei_xing = None    # 合同类型

    @property
    def zhong_zhi(self):
        return self._zhong_zhi

    @zhong_zhi.setter
    def zhong_zhi(self, value):
        sql_str = f"SELECT [中心支公司简称] \
            FROM [中心支公司] \
            WHERE [中心支公司] = '{value}'"
        cur.execute(sql_str)
        self._zhong_zhi = cur.fetchone()[0]

    @property
    def ji_gou(self):
        return self._ji_gou

    @ji_gou.setter
    def ji_gou(self, value):
        sql_str = f"SELECT [机构简称] \
            FROM [机构] \
            WHERE [机构] = '{value}'"
        cur.execute(sql_str)
        self._ji_gou = cur.fetchone()[0]

    @property
    def tuan_dui(self):
        return self._tuan_dui

    @tuan_dui.setter
    def tuan_dui(self, value):
        sql_str = f"SELECT [团队简称] \
            FROM [团队] \
            WHERE [团队] = '{value}'"
        cur.execute(sql_str)
        self._tuan_dui = cur.fetchone()[0]

    @property
    def zhi_ji(self):
        return self._zhi_ji

    @zhi_ji.setter
    def zhi_ji(self, value):
        if value == '(null)':
            self._zhi_ji = ''
        else:
            self._zhi_ji = value

    @property
    def ru_si_zhi_ji(self):
        return self._ru_si_zhi_ji

    @ru_si_zhi_ji.setter
    def ru_si_zhi_ji(self, value):
        if value == '(null)':
            self._ru_si_zhi_ji = ''
        else:
            self._ru_si_zhi_ji = value

    @property
    def ru_si_shi_jian(self):
        return self._ru_si_shi_jian

    @ru_si_shi_jian.setter
    def ru_si_shi_jian(self, value):
        pass


def write_excel(info):
    wb = Workbook()
    ws = wb.active
    ws.title = "前线人员司龄工资信息表"
    nrow = 1
    ws.cell(row=nrow, column=1).value = '序号'
    ws.cell(row=nrow, column=2).value = '机构'
    ws.cell(row=nrow, column=3).value = '团队'
    ws.cell(row=nrow, column=4).value = '姓名'
    ws.cell(row=nrow, column=5).value = '合同类型'
    ws.cell(row=nrow, column=6).value = '现任职级'
    ws.cell(row=nrow, column=7).value = '入司职级'
    ws.cell(row=nrow, column=8).value = '入司时间'
    ws.cell(row=nrow, column=9).value = '17年签单保费'
    ws.cell(row=nrow, column=10).value = '18年签单保费'
    ws.cell(row=nrow, column=11).value = '同比增长率'
    ws.cell(row=nrow, column=12).value = '现任司龄工资'
    ws.cell(row=nrow, column=13).value = '司龄工资变化'
    ws.cell(row=nrow, column=14).value = '考核后司龄工资'

    nrow += 1

    for v in info:
        ws.cell(row=nrow, column=1).value = nrow - 1
        ws.cell(row=nrow, column=2).value = v.zhong_zhi
        ws.cell(row=nrow, column=3).value = v.tuan_dui
        ws.cell(row=nrow, column=4).value = v.xing_ming
        ws.cell(row=nrow, column=5).value = v.he_tong_lei_xing
        ws.cell(row=nrow, column=6).value = v.zhi_ji
        ws.cell(row=nrow, column=7).value = v.ru_si_zhi_ji
        ws.cell(row=nrow, column=8).value = v.ru_si_shi_jian
        ws.cell(row=nrow, column=9).value = '17年签单保费'
        ws.cell(row=nrow, column=10).value = '18年签单保费'
        ws.cell(row=nrow, column=11).value = '同比增长率'
        ws.cell(row=nrow, column=12).value = '现任司龄工资'
        ws.cell(row=nrow, column=13).value = '司龄工资变化'
        ws.cell(row=nrow, column=14).value = '考核后司龄工资'
        nrow += 1

    wb.save('1.xlsx')


def main():
    month = input("请输入考核月份：")
    sql_str = f"SELECT [业务员], [姓名], [工号], [中心支公司], [机构], \
        [销售团队], [职级], [入司时间], [入职时间], [入司职级], \
        [考核类型], [合同类型] \
        FROM [销售人员] \
        WHERE [在职状态] = '在职' \
        ORDER BY [中心支公司], [销售团队], [考核类型], [业务员]"
    cur.execute(sql_str)
    datas = cur.fetchall()
    info = []
    for data in datas:
        ren = Selesman()
        ren.month = month
        ren.ye_wu_yuan = data[0]
        ren.xing_ming = data[1]
        ren.gong_hao = data[2]
        ren.zhong_zhi = data[3]
        ren.ji_gou = data[4]
        ren.tuan_dui = data[5]
        ren.zhi_ji = data[6]
        ren.ru_si_shi_jian = data[7]
        ren.ru_zhi_shi_jian = data[8]
        ren.ru_si_zhi_ji = data[9]
        ren.kao_he_lei_xing = data[10]
        ren.he_tong_lei_xing = data[11]
        info.append(ren)
    write_excel(info)


if __name__ == '__main__':
    main()
