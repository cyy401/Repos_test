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
    flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
    if flag_power:
        f_temp.write(flag_power.group(2))
        f_temp.write('\n')
    f_temp.close()


if __name__ == "__main__":
    outputExcel = "result.xls"
    fileName = "332001200002_full_log.txt"
    f = open(fileName)
    line = f.readline()
    check1 = False
    check2 = False
    check3 = False
    check4 = False
    check5 = False
    check6 = False
    check7 = False
    check8 = False
    check9 = False
    keyword_index = 0
    filter_list = ['Run Resnet50 in 220 package',
                   'Run 300 loops Resnet50 1-batch in SDK', 'Run 300 loops Resnet50 4-batch in SDK', 'Run 300 loops Resnet50 16-batch in SDK',
                   'Run 300 loops mobilenet 1-batch in SDK', 'Run 300 loops mobilenet 4-batch in SDK', 'Run 300 loops mobilenet 16-batch in SDK',
                   'Run 100 loops ssd 1-batch in SDK', 'Run 100 loops ssd 4-batch in SDK', 'Run 100 loops ssd 16-batch in SDK',
                   'Run 50 loops fastercnn 1-batch in SDK', 'Run 50 loops fastercnn 4-batch in SDK', 'Run 50 loops fastercnn 16-batch in SDK',
                   'Run 50 loops yolov3 1-bacth in SDK', 'Run 50 loops yolov3 4-bacth in SDK', 'Run 50 loops yolov3 16-bacth in SDK',
                   'Run 100 loops yolov2 1-batch in SDK', 'Run 100 loops yolov2 4-batch in SDK', 'Run 100 loops yolov2 16-batch in SDK',
                   'Run 50 loops mtcnn in SDK']
    check_flag = [False] * len(filter_list)

    while line:
        line_str = "".join(line)
        str_Res220 = "Run Resnet50 in 220 package"
        str_ResSDk = 'Run 300 loops Resnet50 4-batch in SDK'
        str_mobile = 'Run 300 loops mobilenet 4-batch in SDK'
        str_ssd = 'Run 100 loops ssd 4-batch in SDK'
        str_fastercnn = 'Run 50 loops fastercnn 4-batch in SDK'
        str_yolov3 = 'Run 50 loops yolov3 4-bacth in SDK'
        str_yolov2 = "Run 100 loops yolov2 4-batch in SDK"
        str_mtcnn = 'Run 50 loops mtcnn in SDK'

        flag_Res220 = re.search(str_Res220, line_str)
        flag_ResSDK = re.search(str_ResSDk, line_str)
        flag_mobile = re.search(str_mobile, line_str)
        flag_ssd = re.search(str_ssd, line_str)
        flag_fastercnn = re.search(str_fastercnn, line_str)
        flag_yolov3 = re.search(str_yolov3, line_str)
        flag_yolov2 = re.search(str_yolov2, line_str)
        flag_mtcnn = re.search(str_mtcnn, line_str)
        flag_keyword = re.search(filter_list[keyword_index], line_str)

        if flag_keyword:
            check_flag = [False] * len(filter_list)
            check_flag[keyword_index] = True

        if flag_keyword or check_flag[keyword_index]:
            if flag_keyword:
                print("Processing %s" % filter_list[keyword_index])
            re_result = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_keyword:
                Write_to_txt(filter_list[keyword_index])

            # f_temp = open(fileName+"_fastercnn.txt", 'a+')
            # flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            # if flag_power:
            #     f_temp.write(flag_power.group(2))
            #     f_temp.write('\n')
            #     print("mean power is %s" % flag_power.group(2))
            # f_temp.close()

        if flag_Res220:
            row = 2
            check1 = True
            check2 = False
            check3 = False
            check4 = False
            check5 = False
            check6 = False
            check7 = False
            check8 = False
        elif flag_ResSDK:
            row = 2
            check1 = False
            check2 = True
            check3 = False
            check4 = False
            check5 = False
            check6 = False
            check7 = False
            check8 = False

        elif flag_mobile:
            row = 2
            check1 = False
            check2 = False
            check3 = True
            check4 = False
            check5 = False
            check6 = False
            check7 = False
            check8 = False
        elif flag_ssd:
            row = 2
            check1 = False
            check2 = False
            check3 = False
            check4 = True
            check5 = False
            check6 = False
            check7 = False
            check8 = False
        elif flag_fastercnn:
            row = 2
            check1 = False
            check2 = False
            check3 = False
            check4 = False
            check5 = True
            check6 = False
            check7 = False
            check8 = False
        elif flag_yolov3:
            row = 2
            check1 = False
            check2 = False
            check3 = False
            check4 = False
            check5 = False
            check6 = True
            check7 = False
            check8 = False
        elif flag_yolov2:
            row = 2
            check1 = False
            check2 = False
            check3 = False
            check4 = False
            check5 = False
            check6 = False
            check7 = True
            check8 = False
        elif flag_mtcnn:
            row = 2
            check1 = False
            check2 = False
            check3 = False
            check4 = False
            check5 = False
            check6 = False
            check7 = False
            check8 = True

        if flag_Res220 or check1:
            if flag_Res220:
                print("Start Resnet50 220\n")
            col = 1
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1
            # Write_to_txt("Resnet50_package")

        elif flag_ResSDK or check2:
            if flag_ResSDK:
                print("Start Resnet50 SDK\n")
            col = 2
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1
            # Write_to_txt("Resnet50_SDK")

        elif flag_mobile or check3:
            if flag_mobile:
                print("Start MobileNet\n")
            col = 3
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1
            # Write_to_txt("MobileNet")

        elif flag_ssd or check4:
            if flag_ssd:
                print("Start ssd\n")
            col = 4
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1
            # Write_to_txt("MobileNet")

        elif flag_fastercnn or check5:
            if flag_fastercnn:
                print("Start fastercnn\n")
            col = 5
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1

            # f_temp = open(fileName+"_fastercnn.txt", 'a+')
            # flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            # if flag_power:
            #     f_temp.write(flag_power.group(2))
            #     f_temp.write('\n')
            #     print("mean power is %s" % flag_power.group(2))
            # f_temp.close()

        elif flag_yolov3 or check6:
            if flag_yolov3:
                print("Start yolov3\n")
            col = 6
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1

            # f_temp = open(fileName+"_yolov3.txt", 'a+')
            # flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            # if flag_power:
            #     f_temp.write(flag_power.group(2))
            #     f_temp.write('\n')
            #     print("mean power is %s" % flag_power.group(2))
            # f_temp.close()

        elif flag_yolov2 or check7:
            if flag_yolov2:
                print("Start yolov2\n")

            col = 7
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1

            # f_temp = open(fileName+"_yolov2.txt", 'a+')
            # flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            # if flag_power:
            #     f_temp.write(flag_power.group(2))
            #     f_temp.write('\n')
            #     print("mean power is %s" % flag_power.group(2))
            # f_temp.close()

        elif flag_mtcnn or check8:
            if flag_mtcnn:
                print("Start mtcnn\n")
            col = 8
            flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            if flag_power:
                write_excel_xls_append(
                    outputExcel, row, col, flag_power.group(2))
                row += 1

            # f_temp = open(fileName+"_mtcnn.txt", 'a+')
            # flag_power = re.search(r'([a-z]{4} )(\d.\d\d)', line)
            # if flag_power:
            #     f_temp.write(flag_power.group(2))
            #     f_temp.write('\n')
            #     print("mean power is %s" % flag_power.group(2))
            # f_temp.close()

        line = f.readline()
    f.close()
