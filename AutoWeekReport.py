# -*- coding: utf-8 -*-
"""
Created on Fri Nov 27 17:48:11 2020

@author: LQS
"""
import xlrd
import copy
import datetime

class AutoWeekReport(object):

	def __init__(self, path: str, month: str, startDate: str, endDate: str):
		self._path = path
		self._month = month
		self._startDate = datetime.date(*map(int, startDate.split('-')))
		self._endDate = datetime.date(*map(int, endDate.split('-')))
		self._d = {}
		self._container = []

	def __getdata__(self) -> list:
		data = xlrd.open_workbook(self._path)
		table = data.sheet_by_name(self._month)

		nrows = table.nrows
		self._container = [None]*nrows

		for i in range(nrows):
			cur_row = list(filter(None, table.row_values(i)))
			if '监督抽检' in cur_row or '大检查' in cur_row:
				cur_date = xlrd.xldate.xldate_as_datetime(table.row_values(i)[1], 0)
				if cur_date.date().__ge__(self._startDate) and cur_date.date().__le__(self._endDate):

					self._container[i] = cur_row
		self._container = list(filter(None, self._container))
		return self._container
	
	#输出数量前五的工厂项目名称
	def __products__(self, l: list) -> list:
		pro_map = {}
		for prod in l:
			if prod[6] not in pro_map:
				pro_map[prod[6]] = 1
			else:
				pro_map[prod[6]] += 1
		prod_sorted = sorted(pro_map.items(), key = lambda x : x[1])
		res = ''
		for i in range(1, 6):
			res += prod_sorted[-i][0] + '，'
		print(res)
		del res
		return prod_sorted

	def __getMap__(self):
		allList = self.__getdata__()
		for i in allList:
			if i:
				if i[3].strip() not in self._d:
					self._d[i[3].strip()] = i[4]
				else:
					self._d[i[3].strip()] += i[4]

	def __dataCleaning__(self):
		self.__getMap__()
		#将各种集料汇集成 集料
		self._d['集料'] = self._d.get('粗集料', 0.0) + self._d.get('细集料', 0.0) + \
		self._d.get('沥青路面粗集料', 0.0) + self._d.get('沥青路面细集料', 0.0) + \
		self._d.get('砂', 0.0) + self._d.get('碎石', 0.0) + self._d.get('沥青碎石', 0.0) + \
		self._d.get('沥青细集料', 0.0) + self._d.get('沥青粗集料', 0.0)

		aggregate = ['粗集料', '细集料', '沥青路面粗集料', '沥青路面细集料',
		'沥青粗集料', '沥青细集料', '砂', '碎石', '沥青碎石']

		#删除剩余的各种集料
		for i in aggregate:
			if i in self._d:
				del self._d[i]

		#将减水剂和外加剂合并成外加剂
		if '减水剂' in self._d:
			if '外加剂' in self._d:
				self._d['外加剂'] += self._d['减水剂']
				del self._d['减水剂']

		#将钢筋原材加反向弯曲加入到钢筋原材里面
		if '钢筋原材加反向弯曲' in self._d:
			if '钢筋原材' in self._d:
				self._d['钢筋原材'] += self._d['钢筋原材加反向弯曲']
				del self._d['钢筋原材加反向弯曲']

		#将钢绞线力学加松弛加入到钢绞线力学性能中
		if '钢绞线力学加松弛' in self._d:
			if '钢绞线力学性能' in self._d:
				self._d['钢绞线力学性能'] += self._d['钢绞线力学加松弛']
				del self._d['钢绞线力学加松弛']

		#将混凝土试块更改为混凝土力学
		if '混凝土试块' in self._d:
			self._d['混凝土力学'] = self._d['混凝土试块']
			del self._d['混凝土试块']

	def getRes(self) -> dict:
		self.__dataCleaning__()
		allNums = int(sum(self._d.values()))
		print('分类：{0} 类'.format(len(self._d)))
		print('总数： {0} 组'.format(allNums))

		products = self.__products__(self._container)
		
		res = copy.deepcopy(self._d)
		self._d.clear()

		#增加写入文本的功能
		with open(r'.\材料检测部周报数据汇总.txt', 'a') as f:
			f.write('\n------------------------------------------------\n')
			f.write(self._startDate.strftime('%Y%m%d') + '                ' +\
			self._endDate.strftime('%Y%m%d') + '\n')
			nums = 1
			for i in res.keys():
				f.write(str(nums) + '\t' + i + '\t' + str(int(res[i])) + '\n')
				nums += 1
			f.write('**一共 {0} 类**'.format(i) + '\n')
			f.write('**一共 {0} 组**'.format(allNums) + '\n\n')
			f.write(r'**项目名称：数量**' + '\n')
			for name, pro_num in products:
				f.write(name + '\t' + str(pro_num) + '\n')
			f.write('-----------------------------------------------\n')
		return res