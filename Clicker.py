# coding=utf-8 
import xlwt
import os
from pip._vendor.distlib.compat import raw_input
import codecs
import re
import zipfile
import chardet

 
Root_Url = "O:\\workspace\\src\\"


def txt2excel(file):
    if os.path.exists(file):
        wb = xlwt.Workbook()
        sheet1 = wb.add_sheet('sheet1', cell_overwrite_ok=True)
        f = codecs.open(file, 'r' ,'gbk')
#         f = codecs.open(file, 'r')
        lines = f.readlines()
        sheet1.write(0, 0, u'地市代码')
        sheet1.write(0, 1, u'地市名称')
        sheet1.write(0, 2, u'笔数')

        i = 1
        for line in lines: 
            j = 0
            for item in line.split(';'):
                sheet1.write(i, j, item)
                j += 1
            i += 1
        f.close()
        bname = os.path.basename(str(file[:-4]) + '.xls')
        savefile = os.path.join(Root_Url, bname)
#         print chardet.detect(savefile);
        wb.save(savefile)

def delete_file(src):
    if os.path.isfile(src):
        try:
            print "正在删除文件"+os.path.basename(src)
            os.remove(src)
        except:
            print "删除文件失败"+os.path.basename(src)
    else:
        pass
  
def getfiles():
    f_list = []
    for file in os.listdir(Root_Url):
        if  str(file).lower().endswith(".txt"):
            f_list.append(os.path.join(Root_Url,file))
    return f_list            

def zip_files():
    zip_name = u"申报情况.zip"
    full_name = os.path.join(Root_Url, zip_name)
    f = zipfile.ZipFile(full_name, 'w', zipfile.ZIP_DEFLATED)
    for file in os.listdir(Root_Url):
        if str(file).lower().endswith('.xls'):
            f.write(os.path.join(Root_Url, file))
    f.close()


if __name__ == '__main__':
    for file in getfiles():
        txt2excel(file);

    zip_files()
    print "处理完毕，按回车键结束"
    raw_input()
