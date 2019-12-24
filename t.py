import re
import sqlite3


conn = sqlite3.connect(r'Data\data.db')
cur = conn.cursor()

str_sql = "SELECT 业务员 FROM [滚动12个月同期签单保费]"

cur.execute(str_sql)

name_list = cur.fetchall()
for name in name_list:
    for n in name:
        print(n)

        