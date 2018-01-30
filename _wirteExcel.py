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
		sheet1 = f.add_sheet(u'全部用例',cell_overwrite_ok=True)
		
		style2 = self.set_style2(u'宋体',220,True)
		
		for i in range(0,len(self.row0)):
			sheet1.write(0,i,self.row0[i],self.set_style(u'宋体',220,True))
		for i in range(len(row)):
			for j in range(len(row[i])):
				sheet1.write(i+1, j, row[i][j],style2)
		f.save(u'V1R2用例.xls')
	
		
if __name__ == '__main__':
	row = TLDealer()._getcaseinfo('S2PJ')
	writeExcel().write_excel(row)

