#!/usr/bin/python
# coding=utf-8

from testlink import TestlinkAPIClient, TestLinkHelper
from platform import python_version
import os
import testlink
import collections

class TLDealer:
	testlink_url='http://192.168.103.114/testlink/lib/api/xmlrpc/v1/xmlrpc.php'
	testlink_devkey='b3042cb26bec47083ddf3f5693571c05'
	def __init__(self,url=None, devkey=None):
		if url:	self.testlink_url = url
		if devkey: self.testlink_devkey = devkey
		self.mylink = testlink.TestlinkAPIClient(self.testlink_url, self.testlink_devkey)
	
	def reportResult(self):
		#ret = self.mylink.reportTCResult(test_case_ID, test_plan_id , buidname,test_result, test_note, bugid='',user=username)
		pass

	def getTestProjectInfo(self, projectname):
		projectID = self.mylink.getProjectIDByName(projectname)
		projects = self.mylink.getProjects()
		for pj in projects:
			if pj['id'] == projectID:
				# print 'Project Id:',pj['id']
				# for key, value in pj.items():
				# 	print '\t',key, '-->', value
				return  pj

	def getTestPlanInfo(self, projectid = 1):
		plans = self.mylink.getProjectTestPlans(projectid)
		for plan in plans:
			print 'Test Plan Id:',plan['id']
			for key, value in plan.items():
				print '\t',key, '-->', value

	def getTestSuiteInfo(self, projectid):
		suites = self.mylink.getFirstLevelTestSuitesForTestProject(projectid)
		s = collections.OrderedDict()
		for suite in suites:
			testsuites =  self.mylink.getTestSuitesForTestSuite(suite['id'])
			s[suite['name']] = testsuites
		# for key, value in s.items():
		# 	print '\t',key, '-->', value
		return s
			
	def getTestCaseList_inTS(self, tsid):
		tcs = self.mylink.getTestCasesForTestSuite(testsuiteid=tsid, deep=True, details='full',getkeywords=True)
		# for tc in tcs:
		# 	print 'Test case name:', tc['name']
		# 	for key, value in tc.items():
		# 		print '\t',key, '-->', value
		return tcs

	def getTestCaseList_inTP(self, tpid=924):
		tcs = self.mylink.getTestCasesForTestPlan(tpid,details='simple')
		for tcid,cont in tcs.items():
			print 'TCID:',tcid, '-->', cont

	def getTCIDbyName(self,name='RRC建立请求次数',project='BST',suite='RRC'):
		ret = self.mylink.getTestCaseIDByName(name,testprojectname=project,testsuitename=suite)
		print str(ret)

	def getTestCaseInfo(self, tcid=894):
		ret = self.mylink.getTestCaseIDByName('erab建立成功次数',testprojectname='BST')	#Useful
		#print ret
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
		
	def afun(self, ProjectID):
		response = self.mylink.getTestProjectByName('BST')
		print("getTestProjectByName", str(response).decode('unicode-escape'))
		response = self.mylink.getProjectTestPlans(ProjectID)
		print("getProjectTestPlans", str(response).decode('unicode-escape'))
		response = self.mylink.getFirstLevelTestSuitesForTestProject(ProjectID)
		print("getFirstLevelTestSuitesForTestProject", str(response).decode('unicode-escape'))
		response = self.mylink.getProjectPlatforms(ProjectID)
		print("getProjectPlatforms", str(response).decode('unicode-escape'))
		response = self.mylink.getProjectKeywords(ProjectID)
		print("getProjectKeywords", str(response).decode('unicode-escape'))

		# get information - testPlan
		response = self.mylink.getTestPlanByName('BST', 'QC_TestPlan') #projectname, planname
		print("getTestPlanByName", response)
		response = self.mylink.getTotalsForTestPlan('23') #plan id
		print("getTotalsForTestPlan", response)
		response = self.mylink.getBuildsForTestPlan('23')
		print("getBuildsForTestPlan", response)
		response = self.mylink.getLatestBuildForTestPlan('23')
		print("getLatestBuildForTestPlan", response)
		response = self.mylink.getTestPlanPlatforms('23')
		print("getTestPlanPlatforms", response)
		response = self.mylink.getTestSuitesForTestPlan('23')
		print("getTestSuitesForTestPlan", response)
		# get failed Testcases
		response = self.mylink.getTestCasesForTestPlan('23', executestatus='f')
		print("getTestCasesForTestPlan  failed ", response)
		# get Testcases for Plattform SmallBird
		response = self.mylink.getTestCasesForTestPlan('23', platformid=2)
		print("getTestCasesForTestPlan", response)
		# get information - TestSuite
		response = self.mylink.getTestSuiteByID('814')
		print("getTestSuiteByID", response)
		print 'name:', response['name']
		response = self.mylink.getTestSuitesForTestSuite('814')
		print("getTestSuitesForTestSuite", response)
		print 'name:', response['name']
		response = self.mylink.getTestCasesForTestSuite('814', True, 'full')
		for tc in response:
			print("Test case: " + str(tc).decode('unicode-escape'))
		response = self.mylink.getTestCasesForTestSuite('814', True, 'only_id')
		print("getTestCasesForTestSuite", response)
		response = self.mylink.updateTestSuite('814', prefix='BS',
				testsuitename='功能',details="auto updated Details") #suitid  prefix必填
		print("updateTestSuite", response)

	def getFullPath(self, nodeid=894):
		ret = self.mylink.getFullPath(nodeid)
		print ret

	def getProjectKeywords(self, projectid=1):
		ret = self.mylink.getProjectKeywords(projectid)
		print ret

	def getTestCaseKeywords(self, ext_id='BS-212'):
		ret = self.mylink.getTestCaseKeywords(testcaseexternalid=ext_id)
		print ret

	def addTestCaseKeywords(self):
		args = {'BS-212':['Qualcomm']}
		ret = self.mylink.addTestCaseKeywords(args)
		print ret

	def removeTestCaseKeywords(self):
		args = {'BS-212':['Common']}
		ret = self.mylink.removeTestCaseKeywords(args)
		print ret

	def getTestCaseCustomFieldDesignValue(self,ext_id, version, testprojectid):
		version = int(version)
		ret = self.mylink.getTestCaseCustomFieldDesignValue(testcaseexternalid=ext_id,version=version,
				customfieldname='Remark',testprojectid=testprojectid)
		return ret

	def _getcaseinfo(self, projectname):
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
		for suiteid in suitesid:
			tcs = self.getTestCaseList_inTS(suiteid)
			for tc in tcs:
				msg = u'测试目的:'+str(tc['summary'].replace('</p>','\n').replace('<p>',''))+'\n'\
				      +u'预置条件:'+str(tc['preconditions'].replace('</p>','\n').replace('<p>',''))+'\n'\
				      +u'测试步骤:'+str(tc['steps'][0]['actions'].replace('</p>','\n').replace('<p>',''))+'\n'\
				      +u'预期结果:'+str(tc['steps'][0]['expected_results'].replace('</p>','\n').replace('<p>',''))
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
				print info
				allinfo.append(info)
		return allinfo
		
if __name__ == '__main__':
	mytl = TLDealer()
	#mytl.getTestProjectInfo()
	#mytl.getTestPlanInfo()
	#mytl.getTestSuiteInfo()
	#mytl.getTestCaseList_inTS()
	#mytl.getTestCaseList_inTP()
	#mytl.getTestCaseInfo()
	#mytl.updateTestCaseInfo()
	#mytl.afun(1)
	#mytl.getTCIDbyName()
	#mytl.getFullPath()
	#mytl.getProjectKeywords()
	#mytl.getTestCaseKeywords()
	#mytl.addTestCaseKeywords()
	#mytl.removeTestCaseKeywords()
	# mytl.getTestProjectInfo('S2PJ')
	# mytl.getTestSuiteInfo('932')
	# mytl._getcaseinfo('S2PJ')
	# mytl.getTestCaseList_inTS('1094')


	

