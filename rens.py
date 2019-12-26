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
        self._ru_si_shi_jian: str = None  # 入司时间
        self._ru_si_yue_fen: int = None  # 入司月份
        self._shuo_ming: str = None  # 核定说明

        # 获取当前年份及月份
        self._year: int = int(date.today().strftime("%Y"))
        self._month: int = int(date.today().strftime("%m"))

        # 数据库连接
        self._conn = conn
        self._cur = cur

    @property
    def name(self):
        '''
        返回业务员姓名
        '''
        return self._name

    @name.setter
    def name(self, value):
        '''
        设置业务员姓名
        '''
        self._name = value

    @property
    def zhong_zhi(self):
        '''
        返回业务员所在中支
        '''
        return self._zhong_zhi

    @zhong_zhi.setter
    def zhong_zhi(self, value):
        '''
        设置业务员所在中支
        '''
        self._zhong_zhi = value

    @property
    def ji_gou(self):
        '''
        返回业务员所在机构
        '''
        return self._ji_gou

    @ji_gou.setter
    def ji_gou(self, value):
        '''
        设置业务员所在机构
        '''
        self._ji_gou = value

    @property
    def zhi_ji(self):
        '''
        返回业务员当前职级
        '''
        return self._zhi_ji

    @zhi_ji.setter
    def zhi_ji(self, value):
        '''
        设置业务员当前职级
        '''
        self._zhi_ji = value

    @property
    def he_tong(self):
        '''
        返回业务员的合同类型，“正编”或“劳务派遣”
        '''
        return self._he_tong

    @he_tong.setter
    def he_tong(self, value):
        '''
        设置业务员的合同类型，“正编”或“劳务派遣”
        '''
        self._he_tong = value

    @property
    def ru_si_zhi_ji(self):
        '''
        返回业务员的入司职级
        如果未能从司龄工资核定表中查询到相关人员则证明为新增人员，入司职级等于当前职级
        '''

        str_sql = f"SELECT 入司职级 \
                    FROM 司龄工资核定表 \
                    WHERE 姓名 = '{self._name}'"

        self._cur.execute(str_sql)
        value = self._cur.fetchone()

        logging.debug(f"{self._name}, {value}")

        for v in value:
            if v is None:
                return self._zhi_ji  # 当前职级等于入司职级
            else:
                return v

    @property
    def ru_si_shi_jian(self):
        '''
        返回入司时间
        '''
        return self._ru_si_shi_jian

    @ru_si_shi_jian.setter
    def ru_si_shi_jian(self, value):
        '''
        设置入司时间
        '''
        self._ru_si_shi_jian = value

    @property
    def ru_si_nian_fen(self):
        '''
        返回入司年份
        '''
        return int(self._ru_si_shi_jian[:4])

    @property
    def ru_si_yue_fen(self):
        '''
        返回入司月份
        '''
        return int(self._ru_si_shi_jian[5:7])

    @property
    def shi_fou_he_ding(self):
        '''
        返回是否重新核定
        评定依据为不是本年入司，并且入司月份不是上月
        '''
        if self.ru_si_nian_fen != self._year \
           and self.ru_si_yue_fen == self._month - 1:
            return "重新核定"
        else:
            return "不予重新核定"

    @property
    def bao_fei(self):
        '''
        返回滚动12个月签单保费
            如果该业务员符合重新核定标准则统计滚动12个月签单保费，否则统计保费为空
        '''
        if self.shi_fou_he_ding == '重新核定':
            str_sql = f"SELECT SUM(签单保费)/10000\
                        FROM   [滚动12个月签单保费]\
                        WHERE  ([业务员] LIKE '%{self.name}'\
                        OR [业务员] LIKE '%{self.name}(%'\
                        OR [业务员] LIKE '%{self.name}（%'\
                        OR [业务员] LIKE '%{self.name}\\%'\
                        OR [业务员] LIKE '%{self.name}/%')"

            self._cur.execute(str_sql)
            value = self._cur.fetchone()
            for v in value:
                if v is None:
                    return 0
                else:
                    return v
        else:
            return ''

    @property
    def tong_qi_bao_fei(self):
        '''
        返回滚动12个月同期签单保费
            如果该业务员符合重新核定标准则统计滚动12个月同期签单保费，否则统计保费为空
        '''
        if self.shi_fou_he_ding == '重新核定':
            str_sql = f"SELECT SUM(签单保费)/10000\
                        FROM   [滚动12个月同期签单保费]\
                        WHERE  ([业务员] LIKE '%{self.name}'\
                        OR [业务员] LIKE '%{self.name}(%'\
                        OR [业务员] LIKE '%{self.name}（%'\
                        OR [业务员] LIKE '%{self.name}\\%'\
                        OR [业务员] LIKE '%{self.name}/%')"

            self._cur.execute(str_sql)
            value = self._cur.fetchone()
            for v in value:
                if v is None:
                    return 0
                else:
                    return v
        else:
            return ''

    @property
    def zeng_zhang_lv(self):
        '''
        返回同比增长率
            如果该人员符合重新核定标准则进行计算同比增长率，否则同比增长率为空
            如果滚动12个月同期签单保费为0，则证明入司首次满一年，同比增长率为100%
        '''
        if self.shi_fou_he_ding == '重新核定':
            if self.tong_qi_bao_fei == 0:
                return 1
            else:
                return self.bao_fei / self.tong_qi_bao_fei - 1
        else:
            return ''

    @property
    def zheng_ti_bao_fei(self):
        '''
        返回分公司整体滚动12个月签单保费
        '''
        str_sql = f"SELECT SUM(签单保费)/10000\
                    FROM   [滚动12个月签单保费]"

        self._cur.execute(str_sql)
        for v in self._cur.fetchone():
            return v

    @property
    def zheng_ti_tong_qi_bao_fei(self):
        '''
        返回分公司整体滚动12个月同期签单保费
        '''
        str_sql = f"SELECT SUM(签单保费)/10000\
                    FROM [滚动12个月同期签单保费]"

        self._cur.execute(str_sql)
        for v in self._cur.fetchone():
            return v

    @property
    def zheng_ti_zeng_zhang_lv(self):
        '''
        返回分公司整体同比增长率
        '''
        return self.zheng_ti_bao_fei / self.zheng_ti_tong_qi_bao_fei - 1

    @property
    def si_ling_gong_zi(self):
        '''
        返回当前司龄工资表中该员工的司龄工资
        '''
        str_sql = f"SELECT 现任司龄工资 \
                    FROM 司龄工资核定表 \
                    WHERE 姓名 = '{self._name}'"

        self._cur.execute(str_sql)
        value = self._cur.fetchone()

        for v in value:
            if v is None:
                return 0
            else:
                return v

    @property
    def bian_hua(self):
        '''
        返回重新核定后司龄工资的变化
        '''

        if self.shi_fou_he_ding == '重新核定':
            if self.ru_si_nian_fen == self._year - 1:
                if '初级客户经理' in self.ru_si_zhi_ji:
                    # 符合重新核定规则、上一年度入司、入司职级为初级客户经理
                    # 司龄工资评定为 50
                    self.shuo_ming = '入职首次满一年且入司职级为初级客户经理'
                    return 50
                else:
                    self.shuo_ming = '入职首次满一年'
                    # 符合重新核定规则、上一年度入司、入司职级不是初级客户经理
                    # 司龄工资评定为 100
                    return 100
            elif self.si_ling_gong_zi >= '1000':
                if self.zeng_zhang_lv >= 0:
                    # 符合重新评定规则、入司超过一年，且司龄工资达1000或以上，同比正增长的
                    # 司龄工资已达上限，不在追加
                    self.shuo_ming = '司龄工资已达上限，不予增加'
                    return 0
                else:
                    self.shuo_ming = '同比负增长'
                    # 符合重新评定规则、入司超过一年，且司龄工资达1000或以上，同比副增长的
                    # 司龄工资减少50
                    return - 50
            elif self.zeng_zhang_lv >= self.zheng_ti_zeng_zhang_lv:
                # 符合重新评定规则，入司超过一年，同比增长率高于分公司整体
                # 司龄工资增加 100
                self.shuo_ming = '同比增长率高于分公司'
                return 100
            elif self.zeng_zhang_lv >= 0:
                # 符合重新评定规则，入司超过一年，同比正增长，但低于分公司整体
                # 司龄工资增加 50
                self.shuo_ming = '同比增长率大于0，但低于分公司'
                return 50
            else:
                # 符合重新评定规则，入司超过一年，同比负增长
                # 司龄工资减少 50
                self.shuo_ming = '同比负增长'
                return -50
        else:
            return 0

    @property
    def he_ding(self):
        '''
        返回最终的核定司龄工资
        '''
        return self.si_ling_gong_zi + self.bian_hua

    @property
    def shuo_ming(self):
        '''
        返回重新核定的说明
        '''
        return self._shuo_ming

    @shuo_ming.setter
    def shuo_ming(self, value):
        '''
        设置重新核定的说明
        '''
        self._shuo_ming = value


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
            f"{ren.shi_fou_he_ding:>10}",
            f"{ren.bao_fei:>10}",
            f"{ren.tong_qi_bao_fei:>10}",
            f"{ren.zeng_zhang_lv:>10}",
            f"{ren.bian_hua:>10}",
            f"{ren.shuo_ming}")


if __name__ == '__main__':
    main()
