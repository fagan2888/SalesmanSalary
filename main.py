import sqlite3
from Selesman import Selesman
from write_excel import write_excel

conn = sqlite3.connect(r"Data\data.db")
cur = conn.cursor()


def main():
    month = input("请输入考核月份：")
    sql_str = f"SELECT [业务员], [姓名], [工号], [中心支公司], [机构], \
        [销售团队], [职级], [入职时间], [入司职级], [考核类型],\
        [合同类型] ,[入司时间]\
        FROM [销售人员] \
        WHERE [在职状态] = '在职' \
        ORDER BY [中心支公司], [销售团队], [考核类型], [业务员]"
    cur.execute(sql_str)
    datas = cur.fetchall()
    info = []
    for data in datas:
        ren = Selesman(cur)
        ren.month = month
        ren.ye_wu_yuan = data[0]
        ren.xing_ming = data[1]
        ren.gong_hao = data[2]
        ren.zhong_zhi = data[3]
        ren.ji_gou = data[4]
        ren.tuan_dui = data[5]
        ren.zhi_ji = data[6]
        ren.ru_zhi_shi_jian = data[7]
        ren.ru_si_zhi_ji = data[8]
        ren.kao_he_lei_xing = data[9]
        ren.he_tong_lei_xing = data[10]
        ren.ru_si_shi_jian = data[11]

        info.append(ren)
    write_excel(info)

if __name__ == '__main__':
    main()
    cur.close()
    conn.close()
