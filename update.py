from openpyxl import load_workbook
import sqlite3

wb = load_workbook(filename=r'司龄工资保费统计表.xlsx', read_only=True)
ws = wb['page']

conn = sqlite3.connect(r'Data\data.db')
cur = conn.cursor()

for row in ws.iter_rows(min_row=2, values_only=True):
    sql_str = f"INSERT INTO [司龄工资保费清单] VALUES ('{row[0]}', '{row[1]}', \
'{row[2]}', '{row[3]}', '{row[4]}', '{row[5]}', '{row[6]}', {row[7]})"
    cur.execute(sql_str)

conn.commit()

cur.close()
conn.close()
