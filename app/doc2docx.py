import os
import shutil
import datetime
from win32com import client as wc


# 复制文件到指定路径：重命名、归类
def mycopyfile(srcfile, dstfile):
    if not os.path.isfile(srcfile):
        print("%s not exist!" % (srcfile))
    else:
        fpath, fname = os.path.split(dstfile)  # 分离文件名和路径
        if not os.path.exists(fpath):
            os.makedirs(fpath)  # 创建路径
        try:
            shutil.copyfile(srcfile, dstfile)  # 复制文件
        except Exception as e:
            print(e)
        print("copy %s -> %s" % (srcfile, dstfile))


def doc2docx(from_path, to_path, ifmove=0):
    all_starttime = datetime.datetime.now()

    if not os.path.exists(from_path):
        os.mkdir(from_path)
    if not os.path.exists(to_path):
        os.mkdir(to_path)

    file_list = os.listdir(from_path)
    if len(file_list) == 0:
        print('当前路径下没有文件：' + from_path)
        return 0

    d_count = 1  # 从1开始才能与文件数同步
    m_count = len(os.listdir(to_path))
    move_path = os.path.join(from_path, "temp\\")

    word = wc.Dispatch("Word.Application")
    for file in file_list:
        name, ext = os.path.splitext(file)
        old_file = os.path.join(from_path, file)
        if old_file.endswith('doc') or old_file.endswith('docx'):    # 'Word 为文档指定的名称不能与已打开文档的名称相同
            try:
                new_file = os.path.join(to_path, (str(d_count + m_count) + '.docx'))
                doc = word.Documents.Open(old_file)
                doc.SaveAs(new_file, 12)
                doc.Close()
                d_count = d_count + 1
            except Exception as e:
                print(e)
            print("doc2docx %s -> %s" % (old_file, new_file))
        # elif old_file.endswith('docx'):
            # mycopyfile(old_file, new_file)
        else:
            continue
        if ifmove == 1:
            if not os.path.exists(move_path):
                os.mkdir(move_path)
            try:                                # 存在相同名称的文件 解决方案：重命名文件名，统一编号
                shutil.move(old_file, move_path)
            except Exception as e:
                print(e)
        # os.remove(old_file)  # word.Quit()之前执行remove会报错提示：[WinError 5] 拒绝访问。

    word.Quit()
    #sleep(2)

    all_endtime = datetime.datetime.now()
    print(str(d_count) + '个文档doc2docx所需时间：' + str(all_endtime - all_starttime))


if __name__ == '__main__':
    data_from_path = "D:\\datafrom\\第三部分 6月\\"
    data_to_path = "D:\\datasource\\docx\\第三部分 6月\\"
    doc2docx(data_from_path, data_to_path, 1)

