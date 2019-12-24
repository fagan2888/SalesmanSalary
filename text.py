import sqlite3

from datetime import date
from openpyxl import Workbook
from rens import Rens


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
    info = Rens()
    info.name = v[0][10:]           # 人员姓名
    info.zhong_zhi = v[1][7:]       # 人员中支名称
    info.ji_gou = v[2][11:]         # 人员机构名称
    if info.name == '马斌':         # 马斌在销管系统中无职级
        info.zhi_ji = '中级渠道经理中2'
    else:
        info.zhi_ji = v[3][16:]     # 人员职级信息
    info.he_tong = v[4][5:7]        # 人员的合同类型
    info.ru_si_shi_jian = v[5][:10]     # 人员的入司时间

    # 获取人员当前司龄工资的核定信息
    str_sql = f"SELECT \
               入司职级, \
               入司月份, \
               考核后司龄工资 \
               FROM 司龄工资核定表 \
               WHERE 姓名 = '{info.name}'"
    cur.execute(str_sql)
    si_ling_xin_xi = cur.fetchall()
    # 新入司人员的信息进行评定
    if si_ling_xin_xi == []:
        info.ru_si_zhi_ji = info.zhi_ji  # 当前职级等于入司职级
        info.ru_si_nian_fen = info.ru_si_shi_jian[:4]  # 入司年份
        info.ru_si_yue_fen = info.ru_si_shi_jian[5:7]  # 入司月份
        info.si_ling_gong_zi = '0'  # 新入司人员司龄工资为0

    for value in si_ling_xin_xi:
        info.ru_si_zhi_ji = value[0]
        info.ru_si_yue_fen = value[1]
        info.si_ling_gong_zi = value[2]

    rens.append(info)

for ren in rens:
    if ren.ru_si_yue_fen == str(int(month) - 1) \
      or ren.ru_si_nian_fen != year:
        str_sql = f"SELECT SUM(签单保费)/10000 \
                    FROM   [滚动12个月签单保费] \
                    WHERE  ([业务员] LIKE '%{ren.name}' \
                    OR [业务员] LIKE '%{ren.name}(%' \
                    OR [业务员] LIKE '%{ren.name}（%' \
                    OR [业务员] LIKE '%{ren.name}\\%' \
                    OR [业务员] LIKE '%{ren.name}/%')"
        cur.execute(str_sql)
        for v in cur.fetchone():
            if v is None:
                ren.bao_fei = 0
            else:
                ren.bao_fei = v

        str_sql = f"SELECT SUM(签单保费)/10000 \
                    FROM   [滚动12个月同期签单保费] \
                    WHERE  ([业务员] LIKE '%{ren.name}' \
                    OR [业务员] LIKE '%{ren.name}(%' \
                    OR [业务员] LIKE '%{ren.name}（%' \
                    OR [业务员] LIKE '%{ren.name}\\%' \
                    OR [业务员] LIKE '%{ren.name}/%')"
        cur.execute(str_sql)
        for v in cur.fetchone():
            if v is None:
                ren.tong_qi_bao_fei = 0
            else:
                ren.tong_qi_bao_fei = v

        if ren.tong_qi_bao_fei == 0:
            ren.zeng_zhang_lv = 1
        else:
            ren.zeng_zhang_lv = ren.bao_fei / ren.tong_qi_bao_fei - 1

        print(ren.name, ren.bao_fei, ren.tong_qi_bao_fei, ren.zeng_zhang_lv)
