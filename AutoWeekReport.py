import xlrd

import datetime

class AutoWeekReport(object):

	def __init__(self, path: str, month: str, startDate: str, endDate: str):
		self._path = path
		self._month = month
		self._startDate = datetime.date(*map(int, startDate.split('-')))
		self._endDate = datetime.date(*map(int, endDate.split('-')))
		self._d = {}

	def __getdata__(self) -> list:
		data = xlrd.open_workbook(self._path)
		table = data.sheet_by_name(self._month)

		nrows = table.nrows
		container = [None]*nrows

		for i in range(nrows):
			cur_row = list(filter(None, table.row_values(i)))
			if '监督抽检' in cur_row or '大检查' in cur_row:
				cur_date = xlrd.xldate.xldate_as_datetime(table.row_values(i)[1], 0)
				if cur_date.date().__ge__(self._startDate) and cur_date.date().__le__(self._endDate):

					container[i] = cur_row
		container = list(filter(None, container))
		return container

	def __getMap__(self):
		res = self.__getdata__()
		for i in res:
			if i:
				if i[3] not in self._d:
					self._d[i[3]] = i[4]
				else:
					self._d[i[3]] += i[4]

	def __dataCleaning__(self):
		self.__getMap__()
		#将各种集料汇集成 集料
		self._d['集料'] = self._d.get('粗集料', 0.0) + self._d.get('细集料', 0.0) + self._d.get('沥青路面粗集料', 0.0) + self._d.get('沥青路面细集料', 0.0) + self._d.get('砂', 0.0) + self._d.get('碎石', 0.0)

		aggregate = ['粗集料', '细集料', '沥青路面粗集料', '沥青路面细集料',
		'沥青路面粗集料', '沥青路面细集料', '砂', '碎石',]

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

	def getRes(self):
		self.__dataCleaning__()
		allNums = sum(self._d.values())
		print('总数： {0} 组'.format(allNums))
		return self._d

