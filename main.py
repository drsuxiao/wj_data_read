import os
import datetime
from app.docx_read import each_file
from app.excel_write import writeExcel, writeExcel_ext
from app import intomysql as mydb


# 主程序入口
data_source_path = "D:\\datasource\\docx-516\\"
excel_file_path = "D:\\datasource\\excel\\"

file_list = os.listdir(data_source_path)
file_count = len(file_list)
file_max = 5000
m_count = 1  # 与文件数同步
data_list = []

if file_count > 0:
    print('总文件数：' + str(file_count))
    all_starttime = datetime.datetime.now()

    for file in file_list:
        # starttime = datetime.datetime.now()
        file_name = os.path.join(data_source_path, file)
        print(file_name)
        try:
            temp_dict = each_file(file_name, 0)   # 处理完删除文件，避免下次重复处理
            # 添加文件名到最后
            temp_dict['29'] = file
        except Exception as e:
            print(e)
        try:
            if file_count >= 100:
                data_list.append(temp_dict)
                a = file_count % 100  # 相除的余数
                n = file_count // 100  # 相除取整数
                m = m_count // 100
                if m <= n and len(data_list) == 100:
                    excel_starttime = datetime.datetime.now()
                    print('开始保存数据到excel#################################################')
                    writeExcel_ext(data_list, excel_file_path)
                    print('完成保存数据到excel#################################################')
                    excel_endtime = datetime.datetime.now()
                    print('数据处理耗时：' + str(excel_endtime - excel_starttime))
                    data_list.clear()
                elif m == n and len(data_list) < 100:
                    writeExcel(temp_dict, excel_file_path)
            else:
                writeExcel(temp_dict, excel_file_path)
        except Exception as e:
            print(e)
        # endtime = datetime.datetime.now()
        # print('第' + str(m_count) + '个文档数据处理时间：' + str(endtime - starttime))
        m_count = m_count + 1

    all_endtime = datetime.datetime.now()
    print(str(file_count) + '个文档数据处理时间：' + str(all_endtime - all_starttime))
else:
    print('当前路径下没有可处理的文件：' + data_source_path)





