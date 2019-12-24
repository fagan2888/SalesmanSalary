import sqlite3
from datetime import date
import logging

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG,
    format=" %(asctime)s | %(levelname)s | %(message)s")


class Rens():
    def __init__(self, conn, cur):
        self._name: str = None  # 姓名
        self._zhong_zhi: str = None  # 中支
        self._ji_gou: str = None  # 机构
        self._zhi_ji: str = None  # 当前职级
        self._he_tong: str = None  # 合同类型
        self._ru_si_zhi_ji: str = None  # 入司职级
        self._ru_si_shi_jian: str = None  # 入司时间
        self._ru_si_yue_fen: int = None  # 入司月份
        self._shi_fou_he_ding: str = None  # 是否重新核定
        self._bao_fei: float = None  # 滚动12个月签单保费
        self._tong_qi_bao_fei: float = None  # 滚动12个月同期签单保费
        self._zeng_zhang_lv: float = None  # 同比增长率
        self._si_ling_gong_zi: int = None  # 司龄工资
        self._bian_hua: int = None  # 司龄工资变化
        self._he_ding: int = None  # 核定后的司龄工资
        self._shuo_ming: str = None  # 核定说明

        # 获取当前年份及月份
        self._year = date.today().strftime("%Y")
        self._month = date.today().strftime("%m")

        self._conn = conn
        self._cur = cur

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @property
    def zhong_zhi(self):
        return self._zhong_zhi

    @zhong_zhi.setter
    def zhong_zhi(self, value):
        self._zhong_zhi = value

    @property
    def ji_gou(self):
        return self._ji_gou

    @ji_gou.setter
    def ji_gou(self, value):
        self._ji_gou = value

    @property
    def zhi_ji(self):
        return self._zhi_ji

    @zhi_ji.setter
    def zhi_ji(self, value):
        self._zhi_ji = value

    @property
    def he_tong(self):
        return self._he_tong

    @he_tong.setter
    def he_tong(self, value):
        self._he_tong = value

    @property
    def ru_si_zhi_ji(self):
        # 获取人员当前司龄工资的核定信息
        str_sql = f"SELECT 入司职级 \
                    FROM 司龄工资核定表 \
                    WHERE 姓名 = '{self._name}'"

        self._cur.execute(str_sql)
        value = self._cur.fetchone()

        logging.debug(f"{self._name}, {value}")

        if value is None:
            self._ru_si_zhi_ji = self._zhi_ji  # 当前职级等于入司职级
        else:
            for v in value:
                self._ru_si_zhi_ji = v

        return self._ru_si_zhi_ji

    @property
    def ru_si_shi_jian(self):
        return self._ru_si_shi_jian

    @ru_si_shi_jian.setter
    def ru_si_shi_jian(self, value):
        self._ru_si_shi_jian = value

    @property
    def ru_si_nian_fen(self):
        return self._ru_si_shi_jian[:4]

    @property
    def ru_si_yue_fen(self):
        return self._ru_si_shi_jian[5:7]

    @property
    def shi_fou_he_ding(self):
        if self.ru_si_nian_fen != self._year \
          and int(self.ru_si_yue_fen) == int(self._month) - 1:
            return "重新核定"
        else:
            return "不予重新定"

    @property
    def bao_fei(self):
        return self._bao_fei

    @bao_fei.setter
    def bao_fei(self, value):
        self._bao_fei = value

    @property
    def tong_qi_bao_fei(self):
        return self._tong_qi_bao_fei

    @tong_qi_bao_fei.setter
    def tong_qi_bao_fei(self, value):
        self._tong_qi_bao_fei = value

    @property
    def zeng_zhang_lv(self):
        return self._zeng_zhang_lv

    @zeng_zhang_lv.setter
    def zeng_zhang_lv(self, value):
        self._zeng_zhang_lv = value

    @property
    def si_ling_gong_zi(self):
        return self._si_ling_gong_zi

    @si_ling_gong_zi.setter
    def si_ling_gong_zi(self, value):
        self._si_ling_gong_zi = value

    @property
    def bian_hua(self):
        return self._bian_hua

    @bian_hua.setter
    def bian_hua(self, value):
        self._bian_hua = value

    @property
    def he_ding(self):
        return self._he_ding

    @he_ding.setter
    def he_ding(self, value):
        self._he_ding = value

    @property
    def shuo_ming(self):
        return self._shuo_ming

    @shuo_ming.setter
    def shuo_ming(self, value):
        self._shuo_ming = value


    # def si_ling_gong_zi_biao(self):

    #     # 获取人员当前司龄工资的核定信息
    #     str_sql = f"SELECT \
    #                 入司职级, \
    #                 FROM 司龄工资核定表 \
    #                 WHERE 姓名 = '{self._name}'"

    #     self._cur.execute(str_sql)
    #     value = self._cur.fetchone()

    #     # 新入司人员的信息进行评定
    #     if value == []:
    #         self._ru_si_nian_fen = self._ru_si_shi_jian[:4]  # 入司年份
    #         self._ru_si_yue_fen = self._ru_si_shi_jian[5:7]  # 入司月份
    #         self._si_ling_gong_zi = '0'  # 新入司人员司龄工资为 0

    #     for v in value:
    #         self._ru_si_yue_fen = v[1]
    #         self._si_ling_gong_zi = v[2]


def main():
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

    for ren in rens:
        print(
            f"{ren.name:<6}",
            f"{ren.ru_si_shi_jian:>12}",
            f"{ren.ru_si_nian_fen:>6}",
            f"{ren.ru_si_yue_fen:>4}",
            f"{ren.shi_fou_he_ding:>10}")


if __name__ == '__main__':
    main()
