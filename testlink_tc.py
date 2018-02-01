#!/usr/bin/python
# coding=utf-8

from testlink import TestlinkAPIClient, TestLinkHelper
from platform import python_version
import os
import sys
import time
import re
import testlink
import collections
import xlrd
import xlwt
import getopt
import logging

reload(sys)
sys.setdefaultencoding('utf-8')

if os.path.exists('testlink.log'):
	os.remove('testlink.log')
logging.basicConfig(level = logging.INFO, filename='testlink.log',
					format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s')

def readconfig():
	try:
		f = open('config.txt','r')
		r = f.read()
		f.close()
	except Exception,e:
		print 'Config.txt file open fails'
		sys.exit(3)
	testlink_url = re.findall('testlink_url=(.*)\\n', r)[0]
	testlink_devkey = re.findall('testlink_devkey=(.*)\\n', r)[0]
	project_name = re.findall('project_name=(.*)\\n', r)[0]
	excel_name = re.findall('excel_name=(.*)', r)[0]
	sheet_name = re.findall('sheet_name=(.*)\\n', r)[0]
	author_login=re.findall('author_login=(.*)\\n', r)[0]
	configs = {'url':testlink_url,'key':testlink_devkey,
			'project':project_name,'excel':excel_name,
			'sheet':sheet_name,'author':author_login}
	for v in configs.values():
		if not v:
			print 'Config file content incomplete'
			sys.exit(4)
	return configs

def argumentcheck():
	if len(sys.argv)<2:
		usage()
	if (sys.argv[1] != '-i' and sys.argv[1] != '-e'):
		usage()
	configs = readconfig()
	opts, args = getopt.getopt(sys.argv[2:], "hu:k:p:x:s")
	for op, value in opts:
		if op == "-u":
			if re.search(r'http.*xmlrpc\.php',value):
				configs['url'] = value 
			elif re.search(r'\d+\.\d+\.\d+\.\d+:*\d*',value):
				configs['url'] = 'http://'+value+'/testlink/lib/api/xmlrpc/v1/xmlrpc.php'
			else:
				print 'url parameter incorrect, use data from config file'
		elif op == "-k":
			configs['key'] = value
		elif op == "-p":
			configs['project'] = value
		elif op == "-x":
			if '.xls' not in value:
				print 'excel parameter incorrect, use data from config file'
			else: 
				configs['excel'] = value
		elif op == "-s":
			configs['sheet'] = value
		elif op == "-h":
			usage()
	configs['op']=sys.argv[1][1]
	return configs

def usage():
	print 'Usage: %s <-i or -e> -u [url] -k [devkey] -p [project] -x [excel] -s [sheet] -h'% os.path.basename(sys.argv[0])
	#print '\t-u [url] -k [devkey] -p [project] -e [excel] -s [sheet] -h\n'
	print '\t-i or -e is mandatory. indicate Import or Export'
	print '\t-u option: testlink url'
	print '\t-k option: devkey'
	print '\t-p option: project name'
	print '\t-x option: excel file name'
	print '\t-s option: excel sheet name'
	print '\t-h option: print this help information'
	sys.exit(10)

class TL_API:

	def __init__(self,url=None, devkey=None):
		if url:	self.testlink_url = url
		if devkey: self.testlink_devkey = devkey
		self.mylink = testlink.TestlinkAPIClient(self.testlink_url, self.testlink_devkey)

	'''
	def getTestPlanInfo(self, projectid = 1):
		plans = self.mylink.getProjectTestPlans(projectid)
		for plan in plans:
			print 'Test Plan Id:',plan['id']
			for key, value in plan.items():
				print '\t',key, '-->', value

	def getTestCaseList_inTP(self, tpid=924):
		tcs = self.mylink.getTestCasesForTestPlan(tpid,details='simple')
		for tcid,cont in tcs.items():
			print 'TCID:',tcid, '-->', cont

	def getTCIDbyName(self,name='RRC建立请求次数',project='BST',suite='RRC'):
		ret = self.mylink.getTestCaseIDByName(name,testprojectname=project,testsuitename=suite)
		print str(ret)

	def getTestCaseInfo(self, tcid=894):
		ret = self.mylink.getTestCaseIDByName('erab建立成功次数',testprojectname='BST')	#Useful
		print '\nTestCase ID By Name:', ret[0]['id'],'\n'
		ret = self.mylink.getTestCase(tcid)
		print 'TC_Info:==>',str(ret[0]['testcase_id']),'\n'
		for key, value in ret[0].items():
			print key,'-->', repr(value).decode('unicode-escape')
		ret = self.mylink.getTestCase(testcaseexternalid='BS-212')
		print 'Ext-TC_Info:==>',str(ret[0]['testcase_id']),'\n'
		for key, value in ret[0].items():
			print key,'-->', repr(value).decode('unicode-escape')

	def updateTestCaseInfo(self, tcid=894):
		ret = self.mylink.updateTestCase('BS-212',summary='<p>Changed by API44</p><p>Changed by API55</p>') 
		print ret

	def getProjectKeywords(self, projectid=1):
		ret = self.mylink.getProjectKeywords(projectid)
		print ret

	def addTestCaseKeywords(self):
		args = {'BS-212':['Qualcomm']}
		ret = self.mylink.addTestCaseKeywords(args)
		print ret

	def removeTestCaseKeywords(self):
		args = {'BS-212':['Common']}
		ret = self.mylink.removeTestCaseKeywords(args)
		print ret
	'''

	def getTestProjectInfo(self, projectname):
		projectID = self.mylink.getProjectIDByName(projectname)
		projects = self.mylink.getProjects()
		for pj in projects:
			if pj['id'] == projectID:
				return  pj

	def getTestSuiteInfo(self, projectid):
		suites = self.mylink.getFirstLevelTestSuitesForTestProject(projectid)
		s = collections.OrderedDict()
		for suite in suites:
			testsuites =  self.mylink.getTestSuitesForTestSuite(suite['id'])
			s[suite['name']] = testsuites
		return s

	def getTestCaseList_inTS(self, tsid):
		tcs = self.mylink.getTestCasesForTestSuite(testsuiteid=tsid, deep=True, details='full',getkeywords=True)
		return tcs

	def getFullPath(self, nodeid):
		return self.mylink.getFullPath(nodeid)

	def getTestCaseCustomFieldDesignValue(self,ext_id, version, testprojectid, field='Remark'):
		version = int(version)
		ret = self.mylink.getTestCaseCustomFieldDesignValue(testcaseexternalid=ext_id,version=version,
				customfieldname=field,testprojectid=testprojectid)
		return ret

	def getPJInfo(self,pjname='BST'):
		try:
			return self.mylink.getTestProjectByName(pjname)
		except Exception,e:
			print e
			sys.exit(1)

	def getTCbyName(self,name,project,suite1,suite2):
		try:
			tcs = self.mylink.getTestCaseIDByName(name,testprojectname=project)
			#print 'TCS:',tcs,'\n'
			for tc in tcs:
				fullpath = self.getFullPath(int(tc['id']))[tc['id']]
				if fullpath[1]==suite1 and fullpath[2]==suite2: #Found the correct TC
					return self.mylink.getTestCase(int(tc['id']))[0]
			return {}
		except Exception,e:
			#print e
			return {}

	def getTCbyID(self, extid, suite1, suite2):
		try:
			tcs = self.mylink.getTestCase(testcaseexternalid=extid)
			for tc in tcs:
				fullpath = self.getFullPath(int(tc['id']))[tc['id']]
				if fullpath[1]==suite1 and fullpath[2]==suite2: #Found the correct TC
					return tc
			return {}
		except Exception,e:
			#print e
			return {}

	def getTestCaseKeywords(self, ext_id='BS-212'):
		try:
			ret = self.mylink.getTestCaseKeywords(testcaseexternalid=ext_id)
			return ret[ext_id]
		except Exception,e:
			print e
			return {}

	def checkTestSuites(self,suite1,suite2,project):
		try:
			tss1 = self.mylink.getTestSuite(suite1, project['prefix'])
		except Exception,e:	#tss1 empty
			ts1 = self.mylink.createTestSuite(project['id'],suite1,'create by API')[0]
			ts2 = self.mylink.createTestSuite(project['id'],suite2,'create by API',parentid=ts1['id'])[0]
			return ts2['id']

		try:
			for ts1 in tss1:
				if ts1['parent_id']==project['id']:
					break
			else:
				ts1 = self.mylink.createTestSuite(project['id'],suite1,'create by API')[0]
				ts2 = self.mylink.createTestSuite(project['id'],suite2,'create by API',parentid=ts1['id'])[0]
				return ts2['id']
			try:
				tss2 = self.mylink.getTestSuite(suite2, project['prefix'])
			except Exception,e: 	#tss2 empty
				ts2 = self.mylink.createTestSuite(project['id'],suite2,'create by API',parentid=ts1['id'])[0]
				return ts2['id']

			for ts2 in tss2:
				if ts2['parent_id']==ts1['id']:
					break
			else:
				ts2 = self.mylink.createTestSuite(project['id'],suite2,'create by API',parentid=ts1['id'])[0]
				return ts2['id']
			return ts2['id']
		except Exception,e:
			print e
			print 'Check Test Suite %s | %s Error'%(suite1, suite2)
			return None

	def getAllTCInfo(self, projectname):
		allinfo = []
		suitesid = []
		p = self.getTestProjectInfo(projectname)
		s = self.getTestSuiteInfo(p['id'])
		for key, value in s.iteritems():
			for k,v in value.iteritems():
				if type(v) == dict:
					suitesid.append(k)
				else:
					suitesid.append(value['id'])
					break
		count = 0.0
		for suiteid in suitesid:
			tcs = self.getTestCaseList_inTS(suiteid)
			count += len(tcs)
		c = 0.0
		for suiteid in suitesid:
			tcs = self.getTestCaseList_inTS(suiteid)
			for tc in tcs:
				msg =  u'测试目的:'+'\n'+str(tc['summary'].replace('</p>','\n').replace('<p>',''))\
				      +u'预置条件:'+'\n'+str(tc['preconditions'].replace('</p>','\n').replace('<p>',''))\
				      +u'测试步骤:'+'\n'+str(tc['steps'][0]['actions'].replace('</p>','\n').replace('<p>',''))\
				      +u'预期结果:'+'\n'+str(tc['steps'][0]['expected_results'].replace('</p>','\n').replace('<p>',''))
				keywords = []
				if tc.has_key('keywords'):
					for k,v in tc['keywords'].iteritems():
						keywords.append(v['keyword'])
				else:
					keywords.append(' ')
				if len(keywords)>1:
					for i in range(len(keywords)-1):
						keywords[i+1] = ','+str(keywords[i+1])
				suite2name = self.mylink.getTestSuiteByID(suiteid)['name']
				for key, value in s.iteritems():
					for k, v in value.iteritems():
						if type(v) == dict:
							if k == suiteid:
								suite1name = key
								break
						else:
							if value['id'] == suiteid:
								suite1name = key
								break
				importance = ''
				if tc['importance'] == '3':
					importance = 'H'
				elif tc['importance'] == '2':
					importance = 'M'
				else:
					importance = 'L'
				execution_type = ''
				if tc['execution_type'] == '2':
					execution_type = 'A'
				else:
					execution_type = 'M'
				customfield = self.getTestCaseCustomFieldDesignValue(tc['external_id'],tc['version'],p['id'])
				info = [tc['external_id'], tc['name'], suite1name, suite2name, importance, msg, ' ', keywords, execution_type , customfield]
				c += 1
				rate = c / count
				rate_num = round(rate * 100,1)
				r = '\rExport progress: %s%%' %rate_num
				sys.stdout.write(r)
				sys.stdout.flush()
				allinfo.append(info)
		return allinfo

class TL_Export:
	row0 = [u'编号', u'名称', u'类别', u'子类', u'优先级', u'测试内容', u'需求编号', u'适用产品', u'自动化', u'备注']

	def __init__(self,excel, sheet, row0=None):
		if row0:self.row0 = row0
		self.excelname = excel
		self.sheetname = sheet

	def set_style(self, name, height, bold=False):
		style = xlwt.XFStyle() # 初始化样式
		font = xlwt.Font() # 为样式创建字体
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
		pattern = xlwt.Pattern()
		pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # 设置其模式为实型
		pattern.pattern_fore_colour = 40
		style.font = font
		style.borders = borders
		style.pattern = pattern
		return style
	
	def set_style2(self, name, height, bold=False):
		style = xlwt.XFStyle()  # 初始化样式
		font = xlwt.Font()  # 为样式创建字体
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
	def write_excel(self, tcsinfo):
		f = xlwt.Workbook() #创建工作簿
		sheet1 = f.add_sheet(unicode(self.sheetname),cell_overwrite_ok=True)
		style2 = self.set_style2(u'宋体',220,True)
		col0 = sheet1.col(0)
		col0.width = 256 * 10
		col1 = sheet1.col(1)
		col1.width = 256 * 40
		col2 = sheet1.col(2)
		col2.width = 256 * 15
		col3 = sheet1.col(3)
		col3.width = 256 * 15
		col4 = sheet1.col(4)
		col4.width = 256 * 10
		col5 = sheet1.col(5)
		col5.width = 256 * 10
		col6 = sheet1.col(6)
		col6.width = 256 * 10
		col7 = sheet1.col(7)
		col7.width = 256 * 15
		col8 = sheet1.col(8)
		col8.width = 256 * 10
		col9 = sheet1.col(9)
		col9.width = 256 * 15
		
		for i in range(0,len(self.row0)):
			sheet1.write(0,i,self.row0[i],self.set_style(u'宋体',220,True))
		for i in range(len(tcsinfo)):
			for j in range(len(tcsinfo[i])):
				sheet1.write(i+1, j, tcsinfo[i][j],style2)
		filename = time.strftime("%Y%m%d%H%M%S", time.localtime())+'_'+unicode(self.excelname)
		f.save(filename)

class TL_Import:
	table = None
	rowIndex = 1
	project = None
	author  = None
	tlapi= None
	logger = logging.getLogger('import')
	
	def __init__(self, tlapi, pjname, excelname,author='admin'):
		self.tlapi = tlapi
		self.project = self.tlapi.getPJInfo(pjname)
		data = xlrd.open_workbook(excelname)
		self.table = data.sheets()[0]
		self.author = author
		print 'Records in excel file: %d' % (self.table.nrows-1)

	def getProgress(self):
		return round(float(self.rowIndex-1)*100/float(self.table.nrows-1),1)

	def readline(self):
		while True:
			if self.rowIndex >= self.table.nrows: 
				return []
			rowvalue = self.table.row_values(self.rowIndex)
			if not rowvalue or len(rowvalue)<3 or  not rowvalue[1] or not rowvalue[2]: 
				self.logger.warning('excel file line: %s content incorrect.' % self.rowIndex)
				self.rowIndex += 1
				continue
			self.rowIndex += 1
			return rowvalue

	def preprocess(self,rowvalue):
		record = {}
		record['extid']  = rowvalue[0].strip()
		record['name']   = rowvalue[1].strip()
		record['suite1'] = rowvalue[2].strip()
		record['suite2'] = rowvalue[3].strip() if len(rowvalue)>=4 else ''
		record['prio']   = rowvalue[4].strip() if len(rowvalue)>=5 else ''
		record['req']    = rowvalue[6].strip() if len(rowvalue)>=7 else ''
		record['type']   = rowvalue[8].strip() if len(rowvalue)>=9 else ''
		record['remark'] = rowvalue[9].strip() if len(rowvalue)>=10 else ''
		if not record['suite2']: record['suite2'] = record['suite1']
		if not record['prio'] or record['prio'] not in 'HML': record['prio'] = 'M'	#Middle
		if record['prio']=='H':
			record['prio']='3'
		elif record['prio']=='L':
			record['prio']='1'
		else:
			record['prio']='2'
		if not record['type'] or record['type'] not in 'AM':  record['type'] = 'M'	#Manual
		record['type'] = '2' if record['type'] == 'A' else '1'
		#Keyword -> List
		kwstemp = re.split('[:,;，；：]',rowvalue[7].strip())
		keywords = []
		for kw in kwstemp:
			if kw.strip():
				keywords.append(kw.strip())
		record['keywords'] = keywords
		#Main content
		cont = rowvalue[5].strip() if len(rowvalue)>=6 else ''
		cont = cont.replace(u'输出数据要求及预期结果',u'预期结果')
		maindata = {}
		items = ["测试目的","预置条件","测试步骤","预期结果"]
		for item in items:
			maindata[item] = ''
		maindata["测试步骤"] = cont
		for i in range(len(items)):
			item1 = items[i]
			if not re.findall(ur'%s[:：\s]*[\r\n]+'%item1, cont):
				continue
			for j in range(len(items)):
				if i==j: continue
				item2 = items[j]
				matchobj = re.findall(ur'%s[:：\s]*[\r\n]+(.*?)\n%s[:：\s]*[\r\n]+'%(item1,item2),cont,re.S)
				if matchobj:
					maindata[item1] = matchobj[0]
					break
			else:
				matchobj = re.findall(ur'%s[:：\s]*[\r\n]+(.*)'%item1,cont,re.S)
				if matchobj: 
					maindata[item1] = matchobj[0]
		record['summary'] = self.addP(maindata["测试目的"])
		record['precond'] = self.addP(maindata["预置条件"])
		record['steps']   = self.addP(maindata["测试步骤"])
		record['expect']  = self.addP(maindata["预期结果"])
		return record

	def addP(self, data):
		lines = data.strip().split('\n')
		result = ''
		for line in lines:
			result=result+"<p>"+line+"</p>"
		return result

	def process(self,record):
		if record['extid']:
			tc = self.tlapi.getTCbyID(record['extid'],record['suite1'],record['suite2'])
			if not tc:
				self.logger.error('Cannot find TC: %s with suites %s | %s'%(record['extid'],record['suite1'],record['suite2']))
			else:
				self.update(record, tc)
		else:
			tc = self.tlapi.getTCbyName(record['name'],self.project['name'],record['suite1'],record['suite2'])
			if not tc:
				self.logger.info('Test Case: %s not found, create it.'%record['name'])
				self.create(record)
			else:
				self.update(record, tc)

	def update(self, record, tc):
		#Testcase update
		if(record['name'] != tc['name'] or record['summary'] != tc['summary']
			or record['precond'] != tc['preconditions']
			or record['prio'] != tc['importance']
			or record['type'] != tc['execution_type']
			or record['steps'] != tc['steps'][0]['actions']
			or record['expect'] != tc['steps'][0]['expected_results'] ):
			'''
			print record['name'] ,'VS', tc['name'] 
			print record['summary'] ,'VS', tc['summary'] 
			print record['precond'] ,'VS', tc['preconditions'] 
			print record['prio'] ,'VS', tc['importance'] 
			print record['type'] ,'VS', tc['execution_type'] 
			print record['steps'] ,'VS', tc['steps'][0]['actions'] 
			print record['expect'] ,'VS', tc['steps'][0]['expected_results']
			'''
			try:
				ret = self.tlapi.mylink.updateTestCase(tc['full_tc_external_id'],
					testcasename=record['name'],
					summary=record['summary'],
					preconditions=record['precond'],
					importance=record['prio'],
					executiontype=record['type'],
					steps=[{'step_number' : 1, 'actions' : record['steps'] ,
						'expected_results' : record['expect'], 'execution_type' : 0}])
				if ret:
					self.logger.info('updateTestCase success: %s'%record['name'])
			except Exception,e:
				self.logger.error('Update test case: %s fails'%record['name'])
		#Keyword update
		kw_t = self.tlapi.getTestCaseKeywords(tc['full_tc_external_id'])
		curr_kws = kw_t.values() if kw_t else []
		addlist = []
		rmvlist = []
		for kw in record['keywords']:
			if kw not in curr_kws:
				addlist.append(kw)
		for kw in curr_kws:
			if kw not in record['keywords']:
				rmvlist.append(kw)
		if rmvlist:
			self.tlapi.mylink.removeTestCaseKeywords({tc['full_tc_external_id']:rmvlist})
		if addlist:
			self.tlapi.mylink.addTestCaseKeywords({tc['full_tc_external_id']:addlist})
		if rmvlist or addlist:
			self.logger.info('update Keywords %s to TC: %s'%(str(record['keywords']),record['name']))
		#Custom field update
		cf_t = self.tlapi.getTestCaseCustomFieldDesignValue(
			tc['full_tc_external_id'],tc['version'],self.project['id'])
		if(cf_t != record['remark']):
			ret = self.tlapi.mylink.updateTestCaseCustomFieldDesignValue(
				tc['full_tc_external_id'], int(tc['version']),self.project['id'],
				{'Remark':record['remark']})
			self.logger.info('update custome %s to TC: %s'%(record['remark'],record['name']))

	def create(self, record):
		ts2_id = self.tlapi.checkTestSuites(record['suite1'],record['suite2'],self.project)
		if not ts2_id:
			self.logger.critical('Test suites %s | %s check fails'%(record['suite1'],record['suite2']))
			return
		try:
			tc_t = self.tlapi.mylink.createTestCase(record['name'],ts2_id,
					self.project['id'],self.author,record['summary'],
					steps=[{'step_number' : 1, 'actions' : record['steps'] ,
							'expected_results' : record['expect'], 'execution_type' : 0}],
					preconditions=record['precond'],importance=record['prio'],
					executiontype=record['type'])[0]
			tc = self.tlapi.mylink.getTestCase(tc_t['id'])[0]
			try:
				if record['keywords']:
					kw_t = self.tlapi.mylink.addTestCaseKeywords(
						{tc['full_tc_external_id']:record['keywords']})
			except Exception,e:
				self.logger.warning('TC %s,Add Keywords:%s Fails' % (tc['full_tc_external_id'],str(record['keywords'])))
			try:
				if record['remark']:
					cf_t = self.tlapi.mylink.updateTestCaseCustomFieldDesignValue(
						tc['full_tc_external_id'], int(tc['version']),self.project['id'],
						{'Remark':record['remark']})
			except Exception,e:
				self.logger.warning('TC %s,Add Custom:%s Fails' % (tc['full_tc_external_id'],record['remark']))
		except Exception,e:
			#print e
			self.logger.warning("Create test case:%s error" % record['name'])

if __name__ == '__main__':
	configs = argumentcheck()
	tlapi = TL_API(configs['url'],configs['key'])
	if not tlapi.getTestProjectInfo(configs['project']):
		print '请确保项目名称正确.'.decode('utf-8').encode('gbk')
		sys.exit(1)
	#time1 = time.strftime("%y-%m-%d %H:%M:%S", time.localtime()) 
	time1 = time.strftime("%H:%M:%S", time.localtime()) 
	if(configs['op']=='e'):		#Export: testlink -> excel
		#print '请稍后...'.decode('utf-8').encode('gbk')
		tcsinfo = tlapi.getAllTCInfo(configs['project'])
		TL_Export(configs['excel'],configs['sheet']).write_excel(tcsinfo)
	else:						#Import: excel -> testlink
		if not os.path.exists(configs['excel']):
			print 'Excel文件不存在.'.decode('utf-8').encode('gbk')
			sys.exit(2)
		importer = TL_Import(tlapi,configs['project'],configs['excel'],configs['author'])
		while True:
			row = importer.readline()
			if not row: break
			record = importer.preprocess(row)
			importer.process(record)
			sys.stdout.write('\rImport progress: %s%%' % importer.getProgress())
			sys.stdout.flush()
			#break
	time2 = time.strftime("%H:%M:%S", time.localtime()) 
	print '\nTime Duration: %s - %s' % (time1, time2)
