#######################################################
#filename:test_xlwt.py
#author:defias
#date:xxxx-xx-xx
#function：新建excel文件并写入数据
#######################################################
import xlwt
#创建workbook和sheet对象
workbook = xlwt.Workbook() #注意Workbook的开头W要大写
sheet1 = workbook.add_sheet('sheet1',cell_overwrite_ok=True)
sheet2 = workbook.add_sheet('sheet2',cell_overwrite_ok=True)
#向sheet页中写入数据
sheet1.write(0,0,'地市代码')
sheet1.write(0,1,'地市名称')
sheet1.write(0,2,'笔数') 
"""
#-----------使用样式-----------------------------------
#初始化样式
style = xlwt.XFStyle() 
#为样式创建字体
font = xlwt.Font()
font.name = 'Times New Roman'
font.bold = True
#设置样式的字体
style.font = font
#使用样式
sheet.write(0,1,'some bold Times text',style)
"""
#保存该excel文件,有同名文件时直接覆盖
workbook.save('O:\\workspace\\dst\\test3.xls')
print ('创建excel文件完成！')
