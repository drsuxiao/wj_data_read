import xlrd
import pymysql
import re

conn = pymysql.connect(host='localhost', port=3306, user='root',
                       passwd='password', db='test1', charset='utf8mb4')
p = re.compile(r'\s')
data = xlrd.open_workbook('E:\\myproject\\wj_data\\全部原始问卷汇总.xls')
table = data.sheets()[0]
t = table.col_values(1)
col = table.row_values(0)
nrows = table.nrows
ops = []
for i in range(1, nrows):
    r1 = table.row_values(i)
    ops.append((r1[0], r1[1], r1[2], r1[3], r1[4], r1[5], r1[6], r1[7], r1[8], r1[9],
                r1[10], r1[11], r1[12], r1[13], r1[14], r1[15], r1[16], r1[17], r1[18], r1[19],
                r1[20], r1[21], r1[22], r1[23], r1[24], r1[25], r1[26], r1[27], r1[28], r1[29],
                r1[30], r1[31], r1[32], r1[33], r1[34], r1[35], r1[36], r1[37], r1[38], r1[39],
                r1[40], r1[41], r1[42], r1[43], r1[44], r1[45], r1[46], r1[47], r1[48], r1[49],
                r1[50], r1[51], r1[52], r1[53], r1[54], r1[55], r1[56]))
print(ops)

cur = conn.cursor()
cur.executemany('insert into `wj` (`col_1`, `col_2`, `col_3`, `col_4`, `col_5A`, `col_5B`,`col_5C`, `col_5D`, '
                '`col_6`, `col_7`,`col_8`, `col_9`, `col_10`, '
                '`col_11`, `col_12`, `col_13A`, `col_13B`, `col_13C`, `col_13D`, `col_13E`, `col_13F`,'
                '`col_14A`, `col_14B`, `col_14C`, `col_14D`, `col_14E`, `col_14F`, `col_14G`, `col_14H`, `col_14I`, '
                '`col_15`, `col_16`, `col_17`, `col_18`, `col_19`,'
                '`col_20A`, `col_20B`, `col_20C`, `col_20D`, '
                '`col_21A`, `col_21B`, `col_21C`, `col_21D`, `col_21E`, `col_21F`, `col_21G`, `col_21H`, `col_21I`, '
                '`col_22A`, `col_22B`, `col_22C`, `col_22D`, '
                '`col_23`,`col_24`, `col_25`, `col_26`, `col_27`) \
     values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, '
                '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,'
                '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,'
                '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,'
                '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,'
                '%s, %s, %s, %s, %s, %s, %s)', ops)
conn.commit()
cur.close()

conn.close()


