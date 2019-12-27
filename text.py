import sqlite3

from datetime import date
from openpyxl import Workbook
from openpyxl import Workbook
from openpyxl.styles import (Border,
                             Font,
                             NamedStyle,
                             PatternFill,
                             Side,
                             Alignment)

from rens import Rens
from write_excel import write_excel


conn = sqlite3.connect(r"Data\data.db")
cur = conn.cursor()

# 获取当前在职前线人员清单
str_sql = "SELECT \
        业务员, \
        中支公司, \
        营销服务部, \
        职级, \
        人员分类, \
        考核起始时间 \
        FROM 前线人员信息表 \
        ORDER BY 中支公司, 营销服务部, 人员分类, 职级, 业务员"

cur.execute(str_sql)

rens = []

# 人员信息中的基础信息进行赋值
for v in cur.fetchall():
    info = Rens(conn, cur)
    info.name = v[0][10:]           # 人员姓名
    info.zhong_zhi = v[1][7:]       # 人员中支名称
    info.ji_gou = v[2][11:]         # 人员机构名称
    if info.name == '马斌':         # 马斌在销管系统中无职级
        info.zhi_ji = '中级渠道经理中2'
    else:
        info.zhi_ji = v[3][16:]     # 人员职级信息
    info.he_tong = v[4][5:7]        # 人员的合同类型
    info.ru_si_shi_jian = v[5][:10]  # 人员的入司时间

    rens.append(info)

write_excel(rens)
