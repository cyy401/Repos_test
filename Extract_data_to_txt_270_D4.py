'''
@Descripttion:
@version:
@Author:
@Date: 2020-03-17 06:45:22
@LastEditors: Please set LastEditors
@LastEditTime: 2020-03-28 20:58:44
'''
import re
import xlwt
import xlrd
import codecs
from os.path import os
from xlutils.copy import copy


def write_excel_xls_append(path, row, col, value):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    # rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    # print(rows_old)
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    new_worksheet.write(row, col, value)
    new_workbook.save(path)  # 保存工作簿
    # print("xls格式表格【追加】写入数据成功！")


def Write_to_txt(Testcase):
    f_temp = open(fileName+"_"+Testcase+".txt", 'a+')
    flag_power = re.search(r'(Power of board is :)(\d{2})', line)
    if flag_power:
        f_temp.write(flag_power.group(2))
        f_temp.write('\n')
    f_temp.close()


if __name__ == "__main__":
    outputExcel = "result.xls"
    fileName = "111912100081_log_0328.txt"
    f = open(fileName, encoding='utf-8')
    line = f.readline()
    check = False
    keyword_index = 0

    filter_list = [
        'Run Resnet50 in 270 package',

        'Run 300 loops Resnet50 1-batch in SDK', 'Run 300 loops Resnet50 4-batch in SDK', 'Run 300 loops Resnet50 16-batch in SDK',

        'Run 300 loops mobilenet 1-batch in SDK', 'Run 300 loops mobilenet 4-batch in SDK', 'Run 300 loops mobilenet 16-batch in SDK',

        'Run 100 loops ssd 1-batch in SDK', 'Run 100 loops ssd 4-batch in SDK', 'Run 100 loops ssd 16-batch in SDK',

        'Run 50 loops fastercnn 1-batch in SDK', 'Run 50 loops fastercnn 4-batch in SDK', 'Run 50 loops fastercnn 16-batch in SDK',

        'Run 50 loops yolov3 1-bacth in SDK', 'Run 50 loops yolov3 4-bacth in SDK', 'Run 50 loops yolov3 16-bacth in SDK',

        'Run 100 loops yolov2 1-batch in SDK', 'Run 100 loops yolov2 4-batch in SDK', 'Run 100 loops yolov2 16-batch in SDK',

        'Run 50 loops mtcnn in SDK']

    while line:
        line_str = "".join(line)
        if keyword_index < len(filter_list):
            flag_keyword = re.search(filter_list[keyword_index], line_str)
        elif keyword_index == len(filter_list):
            keyword_index = len(filter_list)-1
            flag_keyword = re.search(filter_list[keyword_index], line_str)

        if flag_keyword:
            check_flag = [False] * len(filter_list)
            check_flag[keyword_index] = True
            check = check_flag[keyword_index]
            keyword_index += 1

        if flag_keyword or check:
            if flag_keyword:
                print("Processing %s" % filter_list[keyword_index-1])
            re_result = re.search(r'(Power of board is :)(\d{2})', line)
            if re_result:
                Write_to_txt(filter_list[keyword_index-1])

        line = f.readline()
    f.close()
