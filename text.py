import sqlite3

from datetime import date
from openpyxl import Workbook
from rens import Rens


conn = sqlite3.connect(r"Data\data.db")
cur = conn.cursor()

str_sql = "SELECT \
           业务员, \
           中支公司, \
           营销服务部, \
           职级, \
           人员分类, \
           入司时间 \
           FROM 前线人员信息表 \
           ORDER BY 中支公司, 营销服务部, 人员分类, 职级, 业务员"

cur.execute(str_sql)

rens = []

year = date.today().strftime("%Y")
month = date.today().strftime("%m")


for v in cur.fetchall():
    info = Rens()
    info.name = v[0][10:]
    info.zhong_zhi = v[1][7:]
    info.ji_gou = v[2][11:]
    if info.name == '马斌':
        info.zhi_ji = '中级渠道经理中2'
    else:
        info.zhi_ji = v[3][16:]
    info.he_tong = v[4][5:7]
    info.ru_si_shi_jian = v[5][:10]

    str_sql = f"SELECT \
               入司职级, \
               入司月份, \
               考核后司龄工资 \
               FROM 司龄工资核定表 \
               WHERE 姓名 = '{info.name}'"
    cur.execute(str_sql)
    si_ling_xin_xi = cur.fetchall()
    if si_ling_xin_xi == []:
        info.ru_si_zhi_ji = info.zhi_ji
        info.ru_si_yue_fen = info.ru_si_shi_jian[5:7]
        info.si_ling_gong_zi = '0'

    if si_ling_xin_xi is None:
        print('内容为None')

    for value in si_ling_xin_xi:
        info.ru_si_zhi_ji = value[0]
        info.ru_si_yue_fen = value[1]
        info.si_ling_gong_zi = value[2]

    rens.append(info)

for ren in rens:
    if ren.name == '江朝丽' or ren.name == '单红丽':
        print(ren)
