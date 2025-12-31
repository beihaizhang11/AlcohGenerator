#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建示例Excel文件用于测试Excel合并工具
"""

import pandas as pd

# 示例1：学生信息表1
data1 = {
    '序号': [1, 2, 3],
    '姓名': ['张三', '李四', '王五'],
    '年龄': [20, 21, 19],
    '成绩': [85, 90, 88]
}
df1 = pd.DataFrame(data1)
df1.to_excel('example1.xlsx', index=False, engine='openpyxl')
print("✓ 已创建 example1.xlsx (3行数据)")

# 示例2：学生信息表2
data2 = {
    '序号': [1, 2],
    '姓名': ['赵六', '孙七'],
    '年龄': [22, 20],
    '成绩': [92, 87]
}
df2 = pd.DataFrame(data2)
df2.to_excel('example2.xlsx', index=False, engine='openpyxl')
print("✓ 已创建 example2.xlsx (2行数据)")

# 示例3：学生信息表3
data3 = {
    '序号': [1, 2, 3, 4],
    '姓名': ['周八', '吴九', '郑十', '钱一'],
    '年龄': [21, 19, 20, 22],
    '成绩': [89, 91, 86, 93]
}
df3 = pd.DataFrame(data3)
df3.to_excel('example3.xlsx', index=False, engine='openpyxl')
print("✓ 已创建 example3.xlsx (4行数据)")

print("\n✅ 所有示例文件创建完成！")
print("\n可以使用以下命令测试合并功能：")
print("python excel_merger.py result.xlsx example1.xlsx example2.xlsx example3.xlsx")
