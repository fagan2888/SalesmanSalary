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

    @property
    def month(self):
        """考核月份"""
        return self._month

    @month.setter
    def month(self, value):
        """考核月份"""
        if int(value) > 0 and int(value) <= 12:
            self._month = value
        else:
            m = input("请输入正确的考核月份：")
            self.month = m

    @property
    def year(self):
        """今年的年份"""
        return datetime.now().year

    @property
    def zhong_zhi(self):
        """中心支公司"""
        return self._zhong_zhi

    @zhong_zhi.setter
    def zhong_zhi(self, value):
        """中心支公司"""
        sql_str = f"SELECT [中心支公司简称] \
            FROM [中心支公司] \
            WHERE [中心支公司] = '{value}'"
        self.cur.execute(sql_str)
        self._zhong_zhi = self.cur.fetchone()[0]

    @property
    def ji_gou_jian_cheng(self):
        """机构"""
        sql_str = f"SELECT [机构简称] \
            FROM [机构] \
            WHERE [机构] = '{self.ji_gou}'"
        self.cur.execute(sql_str)
        return self.cur.fetchone()[0]

    @property
    def tuan_dui(self):
        """销售团队"""
        return self._tuan_dui

    @tuan_dui.setter
    def tuan_dui(self, value):
        """销售团队"""
        sql_str = f"SELECT [团队简称] \
            FROM [团队] \
            WHERE [团队] = '{value}'"
        self.cur.execute(sql_str)
        self._tuan_dui = self.cur.fetchone()[0]

    @property
    def zhi_ji(self):
        """现任职级"""
        return self._zhi_ji

    @zhi_ji.setter
    def zhi_ji(self, value):
        """现任职级"""
        if value == '(null)':
            self._zhi_ji = ''
        else:
            self._zhi_ji = value

    @property
    def ru_si_zhi_ji(self):
        """入司职级"""
        return self._ru_si_zhi_ji

    @ru_si_zhi_ji.setter
    def ru_si_zhi_ji(self, value):
        """入司职级"""
        if value == '(null)':
            self._ru_si_zhi_ji = ''
        else:
            self._ru_si_zhi_ji = value

    @property
    def he_ding(self):
        """是否参与核定"""
        if int(self.ru_si_shi_jian[5:7]) == int(self.month) \
           and int(self.ru_si_shi_jian[:4]) != self.year:
            return True
        else:
            return False

    @property
    def shang_nian_bao_fei(self):
        """上年保费"""
        # 入司月份与考核月份相同（入司每满12个月）
        # 同时不是本年入司的员工
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
                self._shang_nian_bao_fei = ''
            else:
                self._shang_nian_bao_fei = temp[0]
        else:
            self._shang_nian_bao_fei = ''

        if self._shang_nian_bao_fei == '':
            return self._shang_nian_bao_fei
        else:
            return self._shang_nian_bao_fei / 10000

    @property
    def tong_qi_bao_fei(self):
        """同期保费"""
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
                self._tong_qi_bao_fei = ''
            else:
                self._tong_qi_bao_fei = temp[0]
        else:
            self._tong_qi_bao_fei = ''

        if self._tong_qi_bao_fei == '':
            return self._tong_qi_bao_fei
        else:
            return self._tong_qi_bao_fei / 10000

    @property
    def tong_bi(self):
        """保费同比"""
        if self.he_ding is True:
            print(self.xing_ming)
            self._tong_bi = float(self.shang_nian_bao_fei) \
                / float(self.tong_qi_bao_fei) - 1
        else:
            self._tong_bi = ''
        return self._tong_bi

    @property
    def yuan_si_ling_gong_zi(self):
        """原司龄工资"""
        he_ding_date = f"{datetime.now().year}-{(int(self.month)):>02}"
        sql_str = f"SELECT [司龄工资] \
            FROM [司龄工资] \
            WHERE [业务员] = '{self.ye_wu_yuan}'\
            AND [核定时间] = '{he_ding_date}'"
        self.cur.execute(sql_str)
        return self.cur.fetchone()[0]

    @property
    def zheng_ti_tong_bi(self):
        """分公司整体同比增长率"""
        start_date = f"{self.year-1}-{self.month:>02}"
        end_date = f"{self.year}-{(int(self.month)-1):>02}"
        sql_str = f"SELECT SUM ([签单保费/批改保费]) \
            FROM [司龄工资保费清单] \
            WHERE [投保确认日期] >= '{start_date}' \
            AND [投保确认日期] <= '{end_date}'"
        self.cur.execute(sql_str)
        shang_nian_bao_fei = self.cur.fetchone()[0]

        start_date = f"{self.year-2}-{self.month:>02}"
        end_date = f"{self.year-1}-{(int(self.month)-1):>02}"
        sql_str = f"SELECT SUM ([签单保费/批改保费]) \
            FROM [司龄工资保费清单] \
            WHERE [投保确认日期] >= '{start_date}' \
            AND [投保确认日期] <= '{end_date}'"
        self.cur.execute(sql_str)
        tong_qi_bao_fei = self.cur.fetchone()[0]

        return shang_nian_bao_fei / tong_qi_bao_fei - 1

    @property
    def si_ling_gong_zi_bian_hua(self):
        sql_str = f""
        if he_ding is False:
            return 0
        elif int(self.ru_si_shi_jian[:4]) == self.year - 1:
            pass
