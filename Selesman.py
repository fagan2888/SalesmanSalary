import sqlite3
from openpyxl import Workbook
from datetime import date
from datetime import timedelta
from datetime import datetime


class Selesman(object):
    def __init__(self, cur):
        self.cur = cur
        self.ye_wu_yuan = None              # 业务员
        self.ji_gou = None                  # 机构
        self.xing_ming = None               # 姓名
        self.gong_hao = None                # 工号
        self.ru_zhi_shi_jian = None         # 入职时间
        self.ru_si_shi_jian = None          # 入司时间
        self.kao_he_lei_xing = None         # 考核类型
        self.he_tong_lei_xing = None        # 合同类型
        self.he_ding_shuo_ming = None       # 核定说明

    @property
    def month(self):
        """获取考核月份"""
        return self._month

    @month.setter
    def month(self, value):
        """设置考核月份"""
        if int(value) > 0 and int(value) <= 12:
            self._month = value
        else:
            m = input("请输入正确的考核月份：")
            self.month = m

    @property
    def year(self):
        """获取今年的年份"""
        return datetime.now().year

    @property
    def zhong_zhi(self):
        """设置中心支公司"""
        return self._zhong_zhi

    @zhong_zhi.setter
    def zhong_zhi(self, value):
        """获取中心支公司简称"""
        sql_str = f"SELECT [中心支公司简称] \
            FROM [中心支公司] \
            WHERE [中心支公司] = '{value}'"
        self.cur.execute(sql_str)
        self._zhong_zhi = self.cur.fetchone()[0]

    @property
    def ji_gou_jian_cheng(self):
        """获取机构简称"""
        sql_str = f"SELECT [机构简称] \
            FROM [机构] \
            WHERE [机构] = '{self.ji_gou}'"
        self.cur.execute(sql_str)
        return self.cur.fetchone()[0]

    @property
    def tuan_dui(self):
        """获取销售团队简称"""
        return self._tuan_dui

    @tuan_dui.setter
    def tuan_dui(self, value):
        """获取销售团队全称并转化为销售团队简称"""
        sql_str = f"SELECT [团队简称] \
            FROM [团队] \
            WHERE [团队] = '{value}'"
        self.cur.execute(sql_str)
        self._tuan_dui = self.cur.fetchone()[0]

    @property
    def zhi_ji(self):
        """获取现任职级"""
        return self._zhi_ji

    @zhi_ji.setter
    def zhi_ji(self, value):
        """设置现任职级"""
        if value == '(null)':
            self._zhi_ji = ''
        else:
            self._zhi_ji = value

    @property
    def ru_si_zhi_ji(self):
        """设置入司职级"""
        return self._ru_si_zhi_ji

    @ru_si_zhi_ji.setter
    def ru_si_zhi_ji(self, value):
        """获取入司职级"""
        if value == '(null)':
            self._ru_si_zhi_ji = ''
        else:
            self._ru_si_zhi_ji = value

    @property
    def he_ding(self):
        """判断是否参与核定"""
        if int(self.ru_si_shi_jian[5:7]) == int(self.month) \
           and int(self.ru_si_shi_jian[:4]) != self.year:
            return True
        else:
            return False

    @property
    def shang_nian_bao_fei(self):
        """从数据库中获取上年保费"""
        # 入司月份与考核月份相同（入司每满12个月）
        # 同时不是本年入司的员工
        # 满足以上两点的人员才需要重新核定司龄工资
        if self.he_ding is True:
            start_date = f"{self.year-1}-{self.month:>02}"
            end_date = f"{self.year}-{(int(self.month)-1):>02}"

            sql_str = f"SELECT SUM ([签单保费/批改保费]) \
                FROM [司龄工资保费清单] \
                WHERE  [业务员] = '{self.ye_wu_yuan}' \
                AND [投保确认日期] >= '{start_date}' \
                AND [投保确认日期] <= '{end_date}'"
            self.cur.execute(sql_str)
            temp = self.cur.fetchone()

            if temp[0] is None:
                self._shang_nian_bao_fei = ''      # 如果没有查询到数据则返回空
            else:
                self._shang_nian_bao_fei = temp[0]
        else:
            self._shang_nian_bao_fei = ''      # 如果不需要重新核定司龄工资则返回空

        if self._shang_nian_bao_fei == '':
            return self._shang_nian_bao_fei
        else:
            return self._shang_nian_bao_fei / 10000

    @property
    def tong_qi_bao_fei(self):
        """从数据库中获取同期保费"""
        # 入司月份与考核月份相同（入司每满12个月）
        # 同时不是本年入司的员工
        # 满足以上两点的人员才需要重新核定司龄工资
        if self.he_ding is True:

            end_date = f"{self.year-1}-{(int(self.month)-1):>02}"
            start_date = f"{self.year-2}-{self.month:>02}"

            sql_str = f"SELECT SUM ([签单保费/批改保费]) \
                FROM [司龄工资保费清单] \
                WHERE  [业务员] = '{self.ye_wu_yuan}' \
                AND [投保确认日期] >= '{start_date}' \
                AND [投保确认日期] <= '{end_date}'"
            self.cur.execute(sql_str)
            temp = self.cur.fetchone()

            if temp[0] is None:
                self._tong_qi_bao_fei = ''      # 如果没有查询到数据则返回空
            else:
                self._tong_qi_bao_fei = temp[0]
        else:
            self._tong_qi_bao_fei = ''      # 如果不需要重新核定司龄工资则返回空

        if self._tong_qi_bao_fei == '':
            return self._tong_qi_bao_fei
        else:
            return self._tong_qi_bao_fei / 10000

    @property
    def tong_bi(self):
        """计算并返回保费同比"""
        if self.he_ding is True:
            print(self.xing_ming)
            if self.tong_qi_bao_fei == '':
                return ''
            else:
                self._tong_bi = float(self.shang_nian_bao_fei) \
                    / float(self.tong_qi_bao_fei) - 1
        else:
            self._tong_bi = ''
        return self._tong_bi

    @property
    def yuan_si_ling_gong_zi(self):
        """获取原司龄工资"""
        he_ding_date = f"{datetime.now().year}-{(int(self.month)-1):>02}"
        sql_str = f"SELECT [司龄工资] \
            FROM [司龄工资] \
            WHERE [业务员] = '{self.ye_wu_yuan}'\
            AND [核定时间] = '{he_ding_date}'"
        self.cur.execute(sql_str)
        return self.cur.fetchone()[0]

    @property
    def zheng_ti_shang_nian_bao_fei(self):
        """获取分公司整体的上年保费"""
        start_date = f"{self.year-1}-{self.month:>02}"
        end_date = f"{self.year}-{(int(self.month)-1):>02}"
        sql_str = f"SELECT SUM ([签单保费/批改保费]) \
            FROM [司龄工资保费清单] \
            WHERE [投保确认日期] >= '{start_date}' \
            AND [投保确认日期] <= '{end_date}'"
        self.cur.execute(sql_str)
        return self.cur.fetchone()[0]

    @property
    def zheng_ti_tong_qi_bao_fei(self):
        """获取分公司整体的同期保费"""
        start_date = f"{self.year-2}-{self.month:>02}"
        end_date = f"{self.year-1}-{(int(self.month)-1):>02}"
        sql_str = f"SELECT SUM ([签单保费/批改保费]) \
            FROM [司龄工资保费清单] \
            WHERE [投保确认日期] >= '{start_date}' \
            AND [投保确认日期] <= '{end_date}'"
        self.cur.execute(sql_str)
        return self.cur.fetchone()[0]

    @property
    def zheng_ti_tong_bi(self):
        """分公司整体同比增长率"""
        return (float(self.zheng_ti_shang_nian_bao_fei) /
                float(self.zheng_ti_tong_qi_bao_fei) - 1)

    @property
    def si_ling_gong_zi_bian_hua(self):
        """计算司龄工资的变化情况"""
        if self.he_ding is False:
            self.he_ding_shuo_ming = '距上次核定或入职未满一年，不予重新核定'
            return 0
        elif self.kao_he_lei_xing == '团队经理':
            if self.yuan_si_ling_gong_zi >= 1000:
                self.he_ding_shuo_ming = '司龄工资已达上限，不再增加'
                return 0
            else:
                self.he_ding_shuo_ming = '团队经理司龄工资按入职时间每年递增100元'
                return 100
        elif int(self.ru_si_shi_jian[:4]) == self.year - 1:
            sql_str = f"SELECT [职级级别] \
                FROM [销售人员] \
                JOIN [客户经理] \
                ON [销售人员].[入司职级] = [客户经理].[职级] \
                WHERE  [业务员] = '{self.ye_wu_yuan}'"
            self.cur.execute(sql_str)
            ji_bie = self.cur.fetchone()[0]
            if ji_bie <= 6:
                self.he_ding_shuo_ming = '入司首次满一年，且入司职级为初级客户经理，司龄工资为50元'
                return 50
            else:
                self.he_ding_shuo_ming = '入司首次满一年，且入司职级为中级或以上客户经理，司龄工资为100元'
                return 100
        else:
            if self.tong_bi > 0 and self.tong_bi >= self.zheng_ti_tong_bi:
                self.he_ding_shuo_ming = '入司超过一年，同比增长率大于分公司，司龄工资增加100元'
                return 100
            elif self.tong_bi > 0:
                self.he_ding_shuo_ming = '入司超过一年，同比增长率大于零但小于分公司，司龄工资增加50元'
                return 50
            else:
                self.he_ding_shuo_ming = '入司超过一年，同比增长率小于0，司龄工资减50元'
                return - 50

    @property
    def xin_si_ling_gong_zi(self):
        value = self.yuan_si_ling_gong_zi + self.si_ling_gong_zi_bian_hua
        if value < 0:
            value = 0
        return value
