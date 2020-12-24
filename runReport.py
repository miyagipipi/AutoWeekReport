# -*- coding: utf-8 -*-
"""
Created on Fri Nov 27 14:04:13 2020

@author: 李启晟
"""

import AutoWeekReport as awr

if __name__ == '__main__':
    path = input('请输入产值表excel的路径（注意这里只输入路径，不要输入excel表格名）: ')
    month = input('请输入月份： ')

    startData = input('请输入起始日期（注意此处的日期格式必须为xxxx-xx-xx）: \n')
    endDate = input('请输入结束日期（注意此处的日期格式必须为xxxx-xx-xx）: \n')

    a = awr.AutoWeekReport(path, month, startData, endDate)
    b = a.getRes()