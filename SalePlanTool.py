#!/usr/bin/env python
# # -*- coding: utf-8 -*-
# # @Time : 2021/2/3 09:36
# # @Author : qilin.wang
# # @Site :
# # @File : SalePlanTool.py
# # @Software: PyCharm
import os
import time
import pandas as pd
import numpy as np
# from colorama import init
# init(autoreset=True)
os.system('')
#获取指定文件夹下文件的路径
#pathname:文件夹名
#str_path:需拼接的路径
#list_path:返回文件路径
def file_path(pathname, str_path):
    list_path = []
    #获取指定文件夹下的文件名
    filename = os.listdir(pathname)
    for f_name in filename:
        f_name = str_path + f_name
        list_path.append(f_name)
    return list_path

if __name__ == '__main__':
    try:
        print("\033[36m正在加载销售计划转换前文件中的excel文件...\033[0m")
        # 读取转化之前的文件路径
        list_file_before = file_path("销售计划转换前文件", "./销售计划转换前文件/")
        print("\033[36m正在进行表格行列转置...\033[0m")
        print('\033[31m*******************************************************************\033[0m')
        #excel_path：excel路径
        for excel_path in list_file_before:
            source_excel = pd.read_excel(excel_path, index_col=None)
            source_excel.fillna(0, inplace=True)
            #删除合计行
            source_excel_without_zero = source_excel[~source_excel['部门'].isin([0])]

            list_month = source_excel_without_zero.columns[13:21]
            #print(list_month)
            #print(source_excel_without_zero)
            #数量
            target_qty = pd.melt(source_excel_without_zero,
                                 id_vars=['部门', '事务所', '客户编码', '项目编码', '客户名称', '项目名称', 'u9料号', '大类', '统计口径', '型号', '核算币种', '汇率', '未税单价', '产品性质'],
                                 value_vars=[list_month[0], list_month[1], list_month[2], list_month[3]],
                                 var_name='月份',
                                 value_name='数量')
            #删除年月字段最后两个字符
            target_qty['月份'] = target_qty['月份'].str[:-2].str[3:]
            #金额
            target_mny = pd.melt(source_excel_without_zero,
                                 id_vars=['部门', '事务所', '客户编码', '项目编码', '客户名称', '项目名称', 'u9料号', '大类',
                                          '统计口径', '型号', '核算币种', '汇率', '未税单价', '产品性质'],
                                 value_vars=[list_month[4], list_month[5], list_month[6], list_month[7]],
                                 var_name='月份',
                                 value_name='金额'
                                 )
            target_mny['月份'] = target_mny['月份'].str[:-2].str[3:]
            target_with_adjust = pd.concat([target_qty, target_mny['金额']], axis=1)
            target_excel = target_with_adjust[['部门', '事务所', '客户编码', '项目编码', '客户名称', '项目名称', 'u9料号', '大类', '统计口径', '型号',
                                               '核算币种', '汇率', '未税单价', '数量', '金额', '产品性质', '月份']]
            target_excel.rename(columns={'部门':'事业部'}, inplace=True)
            # 文件名拆分
            excel_after = "#" + excel_path.split('/')[2].split('.')[0] + ".xlsx"
            file_after = './销售计划转换后文件/' + excel_after
            print("正在写入" + excel_after + "...")
            # 写入excel
            target_excel.to_excel(file_after, index=False, encoding='utf-8')
            print(excel_after + "转化完毕！")
            print('\033[31m*******************************************************************\033[0m')
        print('\033[1;32m所有文件已转化成功！\033[0m')
        input('请按<Enter>退出~')
    except Exception as msg:
        print('\033[1;31m发生错误：%s\033[0m' %msg)
        input('\033[5m请按<Enter>退出~\033[0m')


