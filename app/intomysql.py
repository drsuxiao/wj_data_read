#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:lcj
import pymysql
from app.excel_write import writeExcel_tuple

def InsertData(tablename, dic):
    try:
        # 连接数据库
        conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='123456', db='test1')  # db：库名
        # 创建游标
        cur = conn.cursor()
        # 插入一条数据
        # reCount = cur.excute('insert into lcj(name,age) vaules(%s,%s)',('ff',18))
        # 向test库中的lcj表插入
        # ret = cur.executemany("insert into lcj(name,tel)values(%s,%s)", [("kk",13212344321),("kw",13245678906)])
        # 同时向数据库lcj表中插入多条数据
        # ret = cur.executemany("insert into lcj values(%s,%s,%s,%s,%s)", [(41, "xiaoluo41", 'man', 24, 13212344332), (42, "xiaoluo42", 'gril', 21, 13245678948)])

        COLstr = ''  # 列的字段
        ROWstr = ''  # 行字段
        fiedstr = ''


        i = 1
        for key in dic.keys():
            if i + 6 < len(dic):
                ColumnStyle = ' VARCHAR(10)'
            else:
                ColumnStyle = ' VARCHAR(2000)'
            COLstr = COLstr + ' ' + 'col_' + key + ColumnStyle + ' ' + 'null' + ','   # 字段不能数字开头
            ROWstr = (ROWstr + '"%s"' + ',') % (dic[key])    # 小括号不能少，会报错？？？
            fiedstr = fiedstr + ',' + 'col_' + key
            if i == 59:
                break
            i = i + 1

        create_sql = "CREATE TABLE %s (%s)" % (tablename, COLstr[:-1])
        insert_sql = "INSERT INTO %s(%s) VALUES (%s)" % (tablename, fiedstr[1:], ROWstr[:-1])
        # 推断表是否存在，存在运行try。不存在运行except新建表，再insert
        try:
            cur.execute("SELECT * FROM  %s" % tablename)
            cur.execute(insert_sql)
        except pymysql.Error as e:
            cur.execute(create_sql)
            cur.execute(insert_sql)
        conn.commit()
        cur.close()
        conn.close()
    except pymysql.Error as e:
        print("Mysql Error %d: %s" % (e.args[0], e.args[1]))

    #提交
    conn.commit()
    #关闭指针对象
    cur.close()
    #关闭连接对象
    conn.close()


def updatedata_number(tablename, start, end):
    try:
        # 连接数据库
        conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='123456', db='test1')  # db：库名
        # 创建游标
        cur = conn.cursor()

        query_sql = 'select * from %s' % tablename
        cur.execute(query_sql)
        rows = cur.fetchall()

        value_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
        for row in rows:
            for i in range(start, end+1):
                field = 'col_' + str(i)
                if start == 1:
                    n = i - 1
                elif start == 6:
                    n = i + 3 - 1  # 加上5题的偏移量
                elif start == 15:
                    n = i + 3 + 5 + 8 - 1  # 加上5,13,14题的偏移量
                value = row[n]
                index = -1
                if value in value_list:
                    index = value_list.index(value)
                if index != -1:
                    new_value = str(index + 1)
                    update_sql = 'update %s set %s="%s" where %s="%s"' % (tablename, field, new_value, field, value)
                    cur.execute(update_sql)
                    print(update_sql)
            #for i in range(6, 13):
             #   pass
            #for i in range(15, 20):
              #  pass
    except pymysql.Error as e:
        print(e)
    #提交
    conn.commit()
    #关闭指针对象
    cur.close()
    #关闭连接对象
    conn.close()


def exportdata_excel(filepath, filename):
    try:
        # 连接数据库
        conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='123456', db='test1')  # db：库名
        # 创建游标
        cur = conn.cursor()

        titles = []
        querys = []
        title = "1.您所在单位类别"
        querystr = "select (CASE col_1 " \
                  "WHEN '1' THEN 'A.行政单位' " \
                  "WHEN '2' THEN 'B.中小学校' " \
                  "WHEN '3' THEN 'C.高等学校' " \
                  "WHEN '4' THEN 'D.医院' " \
                  "WHEN '5' THEN 'E.基层医疗卫生机构' " \
                  "WHEN '6' THEN 'F.科学事业单位' " \
                  "WHEN '7' THEN 'G.测绘事业单位' " \
                  "WHEN '8' THEN 'H.地质勘查事业单位' " \
                  "WHEN '9' THEN 'I.国有林场和苗圃单位' " \
                  "WHEN '10' THEN 'J.彩票机构' " \
                  "WHEN '11' THEN 'K.除上述类别之外的其他事业单位' " \
                  "ELSE col_1 END) as '您所在单位类别', count(col_1) as '总数' from wj " \
                  "group by col_1"
        titles.append(title)
        querys.append(querystr)

        title = "2.您单位的级别"
        querystr = "select (CASE col_2 " \
                   "WHEN '1' THEN 'A.中央驻广西机构' " \
                   "WHEN '2' THEN 'B.自治区直属机构' " \
                   "WHEN '3' THEN 'C.市级直属机构' " \
                   "WHEN '4' THEN 'D.县、区及下属机构' " \
                   "ELSE col_2 END) as '您单位的级别', count(col_2) as '总数' from wj " \
                   "group by col_2"
        titles.append(title)
        querys.append(querystr)

        title = "3.您单位2019年以前执行的会计制度"
        querystr = "select (CASE col_3 " \
                   "WHEN '1' THEN 'A.《行政单位会计制度》' " \
                   "WHEN '2' THEN 'B.《国有林场和苗圃会计制度(暂行)》' " \
                   "WHEN '3' THEN 'C.《测绘事业单位会计制度》' " \
                   "WHEN '4' THEN 'D.《地质勘查单位会计制度》' " \
                   "WHEN '5' THEN 'E.《高等学校会计制度》' " \
                   "WHEN '6' THEN 'F.《中小学校会计制度》' " \
                   "WHEN '7' THEN 'G.《科学事业单位会计制度》' " \
                   "WHEN '8' THEN 'H.《医院会计制度》'" \
                   "WHEN '9' THEN 'I.《基层医疗卫生机构会计制度》' " \
                   "WHEN '10' THEN 'J.《彩票机构会计制度》' " \
                   "WHEN '11' THEN 'K.《事业单位会计制度》' " \
                   "ELSE col_3 END) as '您单位2019年以前执行的会计制度', count(col_3) as '总数' from wj " \
                   "group by col_3"
        titles.append(title)
        querys.append(querystr)

        title = "4.您在单位从事的工作"
        querystr = "select (CASE col_4 " \
                   "WHEN '1' THEN 'A.会计部门负责人' " \
                   "WHEN '2' THEN 'B.会计部门工作人员' " \
                   "WHEN '3' THEN 'C.其他部门人员' " \
                   "WHEN '4' THEN 'D.以上都不是' " \
                   "ELSE col_4 END) as '您在单位从事的工作', count(col_4) as '总数' from wj " \
                   "group by col_4"
        titles.append(title)
        querys.append(querystr)

        title = "6.您所在单位的会计人员是否集中参加《政府会计制度》相关培训"
        querystr = "select (CASE col_6 " \
                   "WHEN '1' THEN 'A.全部' " \
                   "WHEN '2' THEN 'B.部分' " \
                   "WHEN '3' THEN 'C.很少' " \
                   "WHEN '4' THEN 'D.没有' " \
                   "ELSE col_6 END) as '您所在单位的会计人员是否集中参加《政府会计制度》相关培训', " \
                   "count(col_6) as '总数' from wj " \
                   "group by col_6"
        titles.append(title)
        querys.append(querystr)

        title = "7.您对《政府会计制度》或相关衔接规定了解程度"
        querystr = "select (CASE col_7 " \
                   "WHEN '1' THEN 'A.非常了解' " \
                   "WHEN '2' THEN 'B.一般了解' " \
                   "WHEN '3' THEN 'C.了解较少' " \
                   "WHEN '4' THEN 'D.不了解' " \
                   "ELSE col_7 END) as '您对《政府会计制度》或相关衔接规定了解程度', " \
                   "count(col_7) as '总数' from wj " \
                   "group by col_7"
        titles.append(title)
        querys.append(querystr)

        title = "8.您单位自2019年1月1日起，是否已按照新制度的规定进行会计核算"
        querystr = "select (CASE col_8 " \
                   "WHEN '1' THEN 'A.完全按照' " \
                   "WHEN '2' THEN 'B.基本按照' " \
                   "WHEN '3' THEN 'C.稍微按照' " \
                   "WHEN '4' THEN 'D.完全不按照' " \
                   "WHEN '5' THEN 'E.不知道' " \
                   "ELSE col_8 END) as '您单位自2019年1月1日起，是否已按照新制度的规定进行会计核算', " \
                   "count(col_8) as '总数' from wj " \
                   "group by col_8"
        titles.append(title)
        querys.append(querystr)

        title = "9.2018年年末至今，您单位是否完全按规定进行了资产的清查并对清查结果进行了处理"
        querystr = "select (CASE col_9 " \
                   "WHEN '1' THEN 'A.完全' " \
                   "WHEN '2' THEN 'B.部分' " \
                   "WHEN '3' THEN 'C.很少' " \
                   "WHEN '4' THEN 'D.没有' " \
                   "ELSE col_9 END) as '2018年年末至今，您单位是否完全按规定进行了资产的清查并对清查结果进行了处理', " \
                   "count(col_9) as '总数' from wj " \
                   "group by col_9"
        titles.append(title)
        querys.append(querystr)

        title = "10.您单位是否全部将2018年原账科目余额转入2019年新账财务会计相应科目"
        querystr = "select (CASE col_10 " \
                   "WHEN '1' THEN 'A.完全' " \
                   "WHEN '2' THEN 'B.部分' " \
                   "WHEN '3' THEN 'C.很少' " \
                   "WHEN '4' THEN 'D.没有' " \
                   "WHEN '5' THEN 'E.不知道' " \
                   "ELSE col_10 END) as '您单位是否全部将2018年原账科目余额转入2019年新账财务会计相应科目', " \
                   "count(col_10) as '总数' from wj " \
                   "group by col_10"
        titles.append(title)
        querys.append(querystr)

        title = "11.您单位将2018年原账科目余额转入2019年新账相应科目的方法是"
        querystr = "select (CASE col_11 " \
                   "WHEN '1' THEN 'A.编制会计凭证记入新账' " \
                   "WHEN '2' THEN 'B.直接录入新账年初余额' " \
                   "WHEN '3' THEN 'C.其他' " \
                   "WHEN '4' THEN 'D.不知道' " \
                   "ELSE col_11 END) as '您单位将2018年原账科目余额转入2019年新账相应科目的方法是', " \
                   "count(col_11) as '总数' from wj " \
                   "group by col_11"
        titles.append(title)
        querys.append(querystr)

        title = "12.您单位是否进行了预算会计科目2019年年初余额的登记工作"
        querystr = "select (CASE col_12 " \
                   "WHEN '1' THEN 'A.全部完成' " \
                   "WHEN '2' THEN 'B.部分完成' " \
                   "WHEN '3' THEN 'C.尚未开始' " \
                   "WHEN '4' THEN 'D.不知道' " \
                   "ELSE col_12 END) as '您单位是否进行了预算会计科目2019年年初余额的登记工作', " \
                   "count(col_12) as '总数' from wj " \
                   "group by col_12"
        titles.append(title)
        querys.append(querystr)

        title = "15.您单位是否完全按相关衔接规定要求编制了新旧会计制度转账、登记新账科目对照表"
        querystr = "select (CASE col_15 " \
                   "WHEN '1' THEN 'A.完全' " \
                   "WHEN '2' THEN 'B.部分' " \
                   "WHEN '3' THEN 'C.很少' " \
                   "WHEN '4' THEN 'D.没有' " \
                   "WHEN '5' THEN 'E.不知道' " \
                   "ELSE col_15 END) as '您单位是否完全按相关衔接规定要求编制了新旧会计制度转账、登记新账科目对照表', " \
                   "count(col_15) as '总数' from wj " \
                   "group by col_15"
        titles.append(title)
        querys.append(querystr)

        title = "16.您单位是否编制了2019年1月1日完成新旧衔接后的资产负债表"
        querystr = "select (CASE col_16 " \
                   "WHEN '1' THEN 'A.已正确编制' " \
                   "WHEN '2' THEN 'B.已编制，但疑似有误' " \
                   "WHEN '3' THEN 'C.没有编制' " \
                   "WHEN '4' THEN 'D.不知道' " \
                   "ELSE col_16 END) as '您单位是否编制了2019年1月1日完成新旧衔接后的资产负债表', " \
                   "count(col_16) as '总数' from wj " \
                   "group by col_16"
        titles.append(title)
        querys.append(querystr)

        title = "17.迄今为止，您单位是否完成基建“并账”"
        querystr = "select (CASE col_17 " \
                   "WHEN '1' THEN 'A.2018年12月31日已不存在独立基建账' " \
                   "WHEN '2' THEN 'B.已完成对基建账的并账' " \
                   "WHEN '3' THEN 'C.尚未完成对基建账的并账' " \
                   "WHEN '4' THEN 'D.不打算并账，今后继续保留独立基建账' " \
                   "WHEN '5' THEN 'E.不知道' " \
                   "ELSE col_17 END) as '迄今为止，您单位是否完成基建“并账”', " \
                   "count(col_17) as '总数' from wj " \
                   "group by col_17"
        titles.append(title)
        querys.append(querystr)

        title = "18.您单位所用会计核算软件或系统是"
        querystr = "select (CASE col_18 " \
                   "WHEN '1' THEN 'A.用友政务大平台' " \
                   "WHEN '2' THEN 'B.用友单位独立版' " \
                   "WHEN '3' THEN 'C.金蝶软件' " \
                   "WHEN '4' THEN 'D.新中大软件' " \
                   "WHEN '5' THEN 'E.其他软件' " \
                   "WHEN '6' THEN 'F.手工核算，没有软件' " \
                   "WHEN '7' THEN 'G.不知道' " \
                   "ELSE col_18 END) as '您单位所用会计核算软件或系统是', " \
                   "count(col_18) as '总数' from wj " \
                   "group by col_18"
        titles.append(title)
        querys.append(querystr)

        title = "19.截至目前，您单位是否根据《政府会计制度》的要求进行了会计软件的更新或更换"
        querystr = "select (CASE col_19 " \
                   "WHEN '1' THEN 'A.软件或系统已升级到新的版本（识别双体系，支持双分录）' " \
                   "WHEN '2' THEN 'B.软件未升级，但会计科目表已更新' " \
                   "WHEN '3' THEN 'C.软件未升级，会计科目表也未更新' " \
                   "WHEN '4' THEN 'D.不知道' " \
                   "ELSE col_19 END) as '截至目前，您单位是否根据《政府会计制度》的要求进行了会计软件的更新或更换', " \
                   "count(col_19) as '总数' from wj " \
                   "group by col_19"
        titles.append(title)
        querys.append(querystr)

        title = "13.对原制度下2018年12月31日未入账的资产负债，按新制度已经补充确认入账的主要有"
        querystr = "select * from(" \
                   "select 'A.公共基础设施', sum(col_13A) as '总数' from wj " \
                   "union all " \
                   "select 'B.盘盈资产', sum(col_13B) as '总数' from wj " \
                   "union all " \
                   "select 'C.受托代理资产和负债', sum(col_13C) as '总数' from wj " \
                   "union all " \
                   "select 'D.预计负债', sum(col_13D) as '总数' from wj " \
                   "union all " \
                   "select 'E.其他', sum(col_13E) as '总数' from wj " \
                   "union all " \
                   "select 'F.无', sum(col_13F) as '总数' from wj) as t13 "
        titles.append(title)
        querys.append(querystr)

        title = "14.您单位对哪些原有的资产、负债账面余额进行了调整"
        querystr = "select * from(" \
                   "select 'A.应收账款的坏账准备', sum(col_14A) as '总数' from wj " \
                   "union all " \
                   "select 'B.固定资产的折旧', sum(col_14B) as '总数' from wj " \
                   "union all " \
                   "select 'C.公共基础设施的折旧及摊销', sum(col_14C) as '总数' from wj " \
                   "union all " \
                   "select 'D.保障性住房的折旧', sum(col_14D) as '总数' from wj " \
                   "union all " \
                   "select 'E.无形资产的摊销', sum(col_14E) as '总数' from wj " \
                   "union all " \
                   "select 'F.对外投资的应收利息或应收股利', sum(col_14F) as '总数' from wj " \
                   "union all " \
                   "select 'G.负债的应付利息', sum(col_14F) as '总数' from wj " \
                   "union all " \
                   "select 'H.其他', sum(col_14F) as '总数' from wj " \
                   "union all " \
                   "select 'I.无', sum(col_14F) as '总数' from wj) as t14 "
        titles.append(title)
        querys.append(querystr)

        title = "20. 您单位会计软件系统的网络体系结构是"
        querystr = "select * from(" \
                   "select 'A.客户端/远程网服务器（含同城网）', sum(col_20A) as '总数' from wj " \
                   "union all " \
                   "select 'B.客户端/局域网服务器（本院或本栋）', sum(col_20B) as '总数' from wj " \
                   "union all " \
                   "select 'C.单机版', sum(col_20C) as '总数' from wj " \
                   "union all " \
                   "select 'D.不知道', sum(col_20D) as '总数' from wj) as t20 "
        titles.append(title)
        querys.append(querystr)

        title = "21. 本单位的会计软件是否整合了下列功能"
        querystr = "select * from(" \
                   "select 'A．资产卡片管理系统', sum(col_21A) as '总数' from wj " \
                   "union all " \
                   "select 'B.客户端/局域网服务器（本院或本栋）', sum(col_21B) as '总数' from wj " \
                   "union all " \
                   "select 'C．国库收支结算系统', sum(col_21C) as '总数' from wj " \
                   "union all " \
                   "select 'D．银行结算系统', sum(col_21D) as '总数' from wj " \
                   "union all " \
                   "select 'E．工资计算与发放系统', sum(col_21E) as '总数' from wj " \
                   "union all " \
                   "select 'F．采购与合同管理系统', sum(col_21F) as '总数' from wj " \
                   "union all " \
                   "select 'G．前台报销系统', sum(col_21G) as '总数' from wj " \
                   "union all " \
                   "select 'H．其他系统', sum(col_21H) as '总数' from wj " \
                   "union all " \
                   "select 'I．不知道', sum(col_21I) as '总数' from wj" \
                   ") as t21"
        titles.append(title)
        querys.append(querystr)

        title = "22.软件厂商总体印象"
        querystr = "select (CASE col_22A" \
                   "WHEN '1' THEN '不知道' " \
                   "WHEN '2' THEN '不满意' " \
                   "WHEN '3' THEN '基本满意' " \
                   "WHEN '4' THEN '满意' " \
                   "WHEN '5' THEN '完美' " \
                   "ELSE col_22A END) as '软件厂商总体印象', " \
                   "count(col_22A) as '总数' from wj group by col_22A"
        titles.append(title)
        querys.append(querystr)

        title = "22.服务机构总体印象"
        querystr = "select (CASE col_22B" \
                   "WHEN '1' THEN '不知道' " \
                   "WHEN '2' THEN '不满意' " \
                   "WHEN '3' THEN '基本满意' " \
                   "WHEN '4' THEN '满意' " \
                   "WHEN '5' THEN '完美' " \
                   "ELSE col_22B END) as '服务机构总体印象', " \
                   "count(col_22B) as '总数' from wj group by col_22B"
        titles.append(title)
        querys.append(querystr)

        title = "22.服务人员总体印象"
        querystr = "select (CASE col_22C" \
                   "WHEN '1' THEN '不知道' " \
                   "WHEN '2' THEN '不满意' " \
                   "WHEN '3' THEN '基本满意' " \
                   "WHEN '4' THEN '满意' " \
                   "WHEN '5' THEN '完美' " \
                   "ELSE col_22C END) as '服务人员总体印象', " \
                   "count(col_22C) as '总数' from wj group by col_22C"
        titles.append(title)
        querys.append(querystr)

        title = "22.本次衔接工作支持服务"
        querystr = "select (CASE col_22D" \
                   "WHEN '1' THEN '不知道' " \
                   "WHEN '2' THEN '不满意' " \
                   "WHEN '3' THEN '基本满意' " \
                   "WHEN '4' THEN '满意' " \
                   "WHEN '5' THEN '完美' " \
                   "ELSE col_22D END) as '本次衔接工作支持服务', " \
                   "count(col_22D) as '总数' from wj group by col_22D"
        titles.append(title)
        querys.append(querystr)
        for i, val in enumerate(querys):
            cur.execute(val)
            rows = cur.fetchall()
            firstrow = []
            firstrow.append(titles[i])
            firstrow.append('总数')
            writeExcel_tuple(firstrow, filepath, filename)
            for row in rows:
                writeExcel_tuple(row, filepath, filename)
            writeExcel_tuple(('', ''), filepath, filename)
    except pymysql.Error as e:
        print(e)
    #提交
    conn.commit()
    #关闭指针对象
    cur.close()
    #关闭连接对象
    conn.close()


if __name__ == '__main__':
    # 将答案进行数字化
    #updatedata_number('wj', 1, 4)   # 1-4题  1-4
    #updatedata_number('wj', 6, 12)  # 6-12题  9-15
    #updatedata_number('wj', 15, 19)  # 15-19题  31-35

    filepath = "E:\\python\\wj_data\\"
    filename = "wj_data.xlsx"
    #exportdata_excel(filepath, filename)
