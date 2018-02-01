#!/usr/bin/env python
#coding=utf-8
import time
import xlwt
from tlapi import *
from xlwt import *
'''
设置单元格样式
'''
class writeExcel:
	row0 = [u'编号', u'名称', u'类别', u'子类', u'优先级', u'测试内容', u'需求编号', u'适用产品', u'自动化', u'备注']
	
	def __init__(self,row0=None):
		if row0:self.row0 = row0
		
	def set_style(self, name, height, bold=False):
		style = XFStyle() # 初始化样式
		
		font = Font() # 为样式创建字体
		font.name = name # 'Times New Roman'
		font.bold = bold
		font.color_index = 4
		font.height = height
		
		borders= xlwt.Borders()
		borders.left= 1
		borders.right= 1
		borders.top= 1
		borders.bottom= 1
		borders.bottom_colour = 0x3A
		
		pattern = Pattern()
		pattern.pattern = Pattern.SOLID_PATTERN  # 设置其模式为实型
		pattern.pattern_fore_colour = 40
		
		style.font = font
		style.borders = borders
		style.pattern = pattern
		return style
	
	def set_style2(self, name, height, bold=False):
		style = XFStyle()  # 初始化样式
		
		font = Font()  # 为样式创建字体
		font.name = name  # 'Times New Roman'
		# font.bold = bold
		font.height = height
		
		borders = xlwt.Borders()
		borders.left = 1
		borders.right = 1
		borders.top = 1
		borders.bottom = 1
		borders.bottom_colour = 0x3A
		
		style.font = font
		style.borders = borders
		return style
	
	#写excel
	def write_excel(self, row):
		f = xlwt.Workbook() #创建工作簿
		sheetname = readtxt()[2]
		sheet1 = f.add_sheet(unicode(sheetname),cell_overwrite_ok=True)
		col0 = sheet1.col(0)
		col0.width = 256 * 10
		col1 = sheet1.col(1)
		col1.width = 256 * 20
		col2 = sheet1.col(2)
		col2.width = 256 * 10
		col3 = sheet1.col(3)
		col3.width = 256 * 10
		col4 = sheet1.col(4)
		col4.width = 256 * 10
		col5 = sheet1.col(5)
		col5.width = 256 * 15
		col6 = sheet1.col(6)
		col6.width = 256 * 15
		col7 = sheet1.col(7)
		col7.width = 256 * 15
		col8 = sheet1.col(8)
		col8.width = 256 * 10
		col9 = sheet1.col(9)
		col9.width = 256 * 10
		style1 = self.set_style(u'宋体',220,True)
		style2 = self.set_style2(u'宋体',220,True)
		
		for i in range(0,len(self.row0)):
			sheet1.write(0,i,self.row0[i],self.set_style(u'宋体',220,True))
		for i in range(len(row)):
			for j in range(len(row[i])):
				sheet1.write(i+1, j, row[i][j],style2)
		fname = readtxt()[3]
		filename = time.strftime("%Y%m%d%H%M%S", time.localtime())+'_'+unicode(fname)+'.xls'
		f.save(filename)
	
		
if __name__ == '__main__':
	# count = 0
	# for i in range(3):
	# 	projectName = raw_input("请输入项目名称：".decode('utf-8').encode('gbk'))
	# 	check = TLDealer().getTestProjectInfo(projectName)
	# 	if check:
	# 		print '请稍后...'.decode('utf-8').encode('gbk')
	# 		row = TLDealer()._getcaseinfo(projectName)
	# 		writeExcel().write_excel(row)
	# 		break
	# 	else:
	# 		print '请输入正确的项目名称...'.decode('utf-8').encode('gbk')
	# 		count += 1
	# if count == 3:
	# 	print '输入项目名称三次均错误，程序结束....'.decode('utf-8').encode('gbk')
	# 	time.sleep(2)
	a = u'测试目的:'+'\n'+'1111'+'\n'\
				      +u'预置条件:'+'\n'+'22222'+'\n'\
				      +u'测试步骤:'+'\n'+'333333'+'\n'\
				      +u'预期结果:'+'\n'+'44444'
	row=[[1,1,1,1,1,a,1,1,1,1]]
	writeExcel().write_excel(row)


