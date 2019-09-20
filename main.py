# _*_coding: utf-8 _*_

import tkinter as tk


class Root(tk.Tk):
    def __init__(self):
        super().__init__()

        self.lbl_name = tk.Label(self, text='姓名：')
        self.lbl_name.grid(row=0, column=0, padx=5, pady=5)
        self.ent_name = tk.Entry()
        self.ent_name.grid(row=0, column=1)
        self.lbl_job_id = tk.Label(self, text='工号：')
        self.lbl_job_id.grid(row=0, column=2, padx=5, pady=5)
        self.ent_job_id = tk.Entry()
        self.ent_job_id.grid(row=0, column=3, padx=5, pady=5)
        self.lbl_id_num = tk.Label(self, text='身份证号：')
        self.lbl_id_num.grid(row=0, column=4, padx=5, pady=5)
        self.ent_id_num = tk.Entry()
        self.ent_id_num.grid(row=0, column=5, padx=5, pady=5)

1	中心支公司	
2	机构		
3	销售团队				
7	在职状态		
8	合同类型		
10	考核类型		
11	职级		
12	入职时间		
13	入司职级		
14	入司时间
标准保费
司龄工资		
if __name__ == "__main__":
    root = Root()
    root.mainloop()
