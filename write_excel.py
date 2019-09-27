from openpyxl import Workbook


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
    ws.cell(row=nrow, column=9).value = '18年签单保费'
    ws.cell(row=nrow, column=10).value = '17年签单保费'
    ws.cell(row=nrow, column=11).value = '同比增长率'
    ws.cell(row=nrow, column=12).value = '现任司龄工资'
    ws.cell(row=nrow, column=13).value = '司龄工资变化'
    ws.cell(row=nrow, column=14).value = '考核后司龄工资'

    nrow += 1

    for ren in info:
        ws.cell(row=nrow, column=1).value = nrow - 1
        ws.cell(row=nrow, column=2).value = ren.zhong_zhi
        ws.cell(row=nrow, column=3).value = ren.tuan_dui
        ws.cell(row=nrow, column=4).value = ren.xing_ming
        ws.cell(row=nrow, column=5).value = ren.he_tong_lei_xing
        ws.cell(row=nrow, column=6).value = ren.zhi_ji
        ws.cell(row=nrow, column=7).value = ren.ru_si_zhi_ji
        ws.cell(row=nrow, column=8).value = ren.ru_si_shi_jian
        ws.cell(row=nrow, column=9).value = ren.shang_nian_bao_fei
        ws.cell(row=nrow, column=10).value = ren.tong_qi_bao_fei
        ws.cell(row=nrow, column=11).value = ren.tong_bi
        ws.cell(row=nrow, column=12).value = ren.yuan_si_ling_gong_zi
        ws.cell(row=nrow, column=13).value = '司龄工资变化'
        ws.cell(row=nrow, column=14).value = '考核后司龄工资'
        nrow += 1

    wb.save('司龄工资考核表.xlsx')
