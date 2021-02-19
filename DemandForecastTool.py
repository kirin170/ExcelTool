#!/usr/bin/env python
# # -*- coding: utf-8 -*-
# # @Time : 2021/2/19 09:36
# # @Author : qilin.wang
# # @Site :
# # @File : DemandForecastTool.py
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

        print("\033[36m正在加载需求预测计划转换前文件中的excel文件...\033[0m")
        # 读取转化之前的文件路径
        list_file_before = file_path("需求预测计划转换前文件", "./需求预测计划转换前文件/")
        print("\033[36m正在进行表格行列转置...\033[0m")
        print('\033[31m*******************************************************************\033[0m')
        #excel_path：excel路径
        for excel_path in list_file_before:
            SourceExcel = pd.read_excel(excel_path, index_col=None)
            #处理Nan值
            SourceExcel.fillna(0, inplace=True)
            #设置列索引方便动态取值
            NewCol = ['生产事业部', '部门', '事务所', '业务员', '客户编码', '客户名称', '项目编码', '项目名称', '产品性质', '大类名称', '小类名称', '采购周期', 'u9料号',
                      '型号', 'a1', 'a2', 'a3', 'b1', 'b2', 'b3', 'c1', 'c2',
                      'c3', 'd1', 'd2', 'd3', 'e1', 'e2', 'e3', 'f1', 'f2', 'f3', '备注1']
            SourceExcel.columns = NewCol
            #取月份
            Months = [SourceExcel.loc[0, 'a1'], SourceExcel.loc[0, 'b1'], SourceExcel.loc[0, 'c1'],
                      SourceExcel.loc[0, 'd1'], SourceExcel.loc[0, 'e1'], SourceExcel.loc[0, 'f1']]
            #print(Months)
            ValueCol = ['要货数', '消耗库存数', '生产数']
            #剔除第0行和第1行数据
            SourceExcel_1 = SourceExcel.loc[2:]
            #剔除合计行数据
            SourceExcel_2 = SourceExcel_1[~SourceExcel_1['生产事业部'].isin([0])]

            #数据清洗后再次重新设置列索引方便转置
            FinalCol = ['生产事业部', '部门', '事务所', '业务员', '客户编码', '客户名称', '项目编码', '项目名称', '产品性质', '大类名称', '小类名称', '采购周期', 'u9料号',
                        '型号',
                        Months[0] + ValueCol[0], Months[0] + ValueCol[1], Months[0] + ValueCol[2],
                        Months[1] + ValueCol[0], Months[1] + ValueCol[1], Months[1] + ValueCol[2],
                        Months[2] + ValueCol[0], Months[2] + ValueCol[1], Months[2] + ValueCol[2],
                        Months[3] + ValueCol[0], Months[3] + ValueCol[1], Months[3] + ValueCol[2],
                        Months[4] + ValueCol[0], Months[4] + ValueCol[1], Months[4] + ValueCol[2],
                        Months[5] + ValueCol[0], Months[5] + ValueCol[1], Months[5] + ValueCol[2],
                        '备注1']
            SourceExcel_2.columns = FinalCol

            #需求数
            DemandQty = pd.melt(
                SourceExcel_2,
                id_vars=['生产事业部', '部门', '事务所', '业务员', '客户编码', '客户名称', '项目编码', '项目名称', '产品性质', '大类名称', '小类名称', '采购周期',
                         'u9料号', '型号', '备注1'],
                value_vars=[Months[0] + ValueCol[0], Months[1] + ValueCol[0], Months[2] + ValueCol[0],
                            Months[3] + ValueCol[0], Months[4] + ValueCol[0], Months[5] + ValueCol[0]],
                var_name='月份',
                value_name=ValueCol[0])
            DemandQty['月份'] = DemandQty['月份'].str[:-3]
            #print(DemandQty)

            #库存消耗数
            ConsumeQty = pd.melt(
                SourceExcel_2,
                id_vars=['生产事业部', '部门', '事务所', '业务员', '客户编码', '客户名称', '项目编码', '项目名称', '产品性质', '大类名称', '小类名称', '采购周期',
                         'u9料号', '型号', '备注1'],
                value_vars=[Months[0] + ValueCol[1], Months[1] + ValueCol[1], Months[2] + ValueCol[1],
                            Months[3] + ValueCol[1], Months[4] + ValueCol[1], Months[5] + ValueCol[1]],
                var_name='月份',
                value_name=ValueCol[1])
            ConsumeQty['月份'] = ConsumeQty['月份'].str[:-5]

            #生产数
            ProduceQty = pd.melt(
                SourceExcel_2,
                id_vars=['生产事业部', '部门', '事务所', '业务员', '客户编码', '客户名称', '项目编码', '项目名称', '产品性质', '大类名称', '小类名称', '采购周期',
                         'u9料号', '型号', '备注1'],
                value_vars=[Months[0] + ValueCol[2], Months[1] + ValueCol[2], Months[2] + ValueCol[2],
                            Months[3] + ValueCol[2], Months[4] + ValueCol[2], Months[5] + ValueCol[2]],
                var_name='月份',
                value_name=ValueCol[2])
            ProduceQty['月份'] = ProduceQty['月份'].str[:-3]

            #合并连接
            TargetExcel = pd.concat([DemandQty, ConsumeQty[ValueCol[1]], ProduceQty[ValueCol[2]]], axis=1)
            TargetExcel = TargetExcel[['生产事业部', '部门', '事务所', '业务员', '客户编码', '客户名称', '项目编码', '项目名称',
                                        '产品性质', '大类名称', '小类名称', '采购周期', 'u9料号', '型号',
                                        '月份', '要货数', '消耗库存数', '生产数', '备注1']]

            # 文件名拆分
            ExcelAfter = "#" + excel_path.split('/')[2].split('.')[0] + ".xlsx"
            FileAfter = './需求预测计划转换后文件/' + ExcelAfter
            print("正在写入" + ExcelAfter + "...")

            # 写入excel
            TargetExcel.to_excel(FileAfter, index=False, encoding='utf-8')
            print(ExcelAfter + "转化完毕！")
            print('\033[31m*******************************************************************\033[0m')
        print('\033[1;32m所有文件已转化成功！\033[0m')
        input('\033[5m请按<Enter>退出~\033[0m')
    except Exception as msg:
        print('\033[1;31m发生错误：%s\033[0m' %msg)
        input('\033[5m请按<Enter>退出~\033[0m')



