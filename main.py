#!/usr/bin/env python
# -*- coding:utf-8 -*-

from collections import OrderedDict
from pyexcel_xls import save_data
from pyexcel_xls import get_data


def read_xls_file(filname, sheet):
    data = get_data(filname)
    return data[sheet]


def save_xls_file(filname, sheet_1):
    data = OrderedDict()
    data.update({"Sheet1": sheet_1})
    save_data(filname, data)


if __name__ == "__main__":
    # 读取文件中的‘输出结果’sheet
    s2 = read_xls_file(r"数据2.xls", '输出结果')
    # 将要提取数据的‘科目代码’字段，存入 code_list
    code_list = []
    for row in s2:
        print(row[0])
        code_list.append(row[0])
    # 读取文件中的‘原始数据’sheet
    s1 = read_xls_file(r"数据2.xls", '原始数据')
    new_sheet = []
    for row in s1:
        try:
            if row[0] in code_list:
                # 是要提取的数据
                print('要提取的数据', row)
                # 逐条添加数据，到 new_sheet
                new_sheet.append(row)
        except IndexError:
            print("List index out of range! 该列表为空列表")
    # 将 new_sheet 写入到新的 excel 文件
    save_xls_file("writefile.xls", new_sheet)
