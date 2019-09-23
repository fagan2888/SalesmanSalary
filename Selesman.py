import sqlite3
from openpyxl import Workbook

conn = sqlite3.connect(r"Data\data.db")
cur = conn.cursor()


class Selesman():
    def __init__(self):
        self.ye_wu_yuan = None          # 业务员
        self.xing_ming = None           # 姓名
        self.gong_hao = None            # 工号
        self.tuan_dui = None            # 团队
        self.ji_gou = None              # 机构
        self.zhi_ji = None              # 职级
        self.ru_si_shi_jian = None      # 入司时间
        self.ru_zhi_shi_jian = None     # 入职时间
        self.si_ling_gong_zi = None     # 司龄工资
        self.ru_si_zhi_ji = None        # 入司职级
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


def write_excel(info):
    wb = Workbook()
    ws = wb.active
    ws.title = "前线人员司龄工资信息表"
    nrow = 1
    for v in info:
        ws.cell(row=nrow, column=1).value = nrow
        ws.cell(row=nrow, column=2).value = v.zhong_zhi
        ws.cell(row=nrow, column=3).value = v.tuan_dui
        ws.cell(row=nrow, column=4).value = v.xing_ming
        ws.cell(row=nrow, column=5).value = v.he_tong_lei_xing
        nrow += 1

    wb.save('1.xlsx')


def main():
    sql_str = f"SELECT 业务员, 姓名, 工号, 中心支公司, 机构, 销售团队, 职级, \
        入司时间, 入职时间, 入司职级, 考核类型, 合同类型 \
        FROM [销售人员] WHERE 在职状态 = '在职'"
    cur.execute(sql_str)
    datas = cur.fetchall()
    info = []
    for data in datas:
        ren = Selesman()
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

