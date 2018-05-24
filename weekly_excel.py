# -*- coding: gbk -*-

import argparse
import time
import xlrd
import xlwt


def str_replace(input):
    output = input
    if "北活" in input:
        output = input[2:]
    elif "第二教学楼" in input:
        output = input[5:]
    elif "第一教学楼" in input:
        output = input[5:]
    elif "北区活动中心" in input:
        output = input[6:]
    elif "江南大学就业创业指导服务中心" in input:
        output = '就业指导中心'
    elif "物联网工程学院" in input:
        output = '物联网学院'
    elif "化学与材料工程学院" in input:
        output = '化工学院'
    elif "纺织与服装学院" in input:
        output = '纺服学院'
    elif "环境与土木工程学院" in input:
        output = '环土学院'

    return output


def main(args):
    begin_date = time.strptime(args.begin_date, "%Y-%m-%d")
    end_date = time.strptime(args.end_date, "%Y-%m-%d")

    title = ['日期', '时间', '公司名称', '宣讲地址',
             '联系人', '电话', '手机', '笔试时间',
             '笔试地址', '面试时间', '面试地址', '下发学院'
             ]

    if begin_date > end_date:
        print('Check the date.')
        exit(0)
    file = xlrd.open_workbook('宣讲会信息导出.xlsx')
    tabel = file.sheets()[0]
    nrows = tabel.nrows

    datas = []
    datas_title = tabel.row_values(0)
    for i in range(1, nrows):
        data = tabel.row_values(i)
        date = data[4].split( )
        temp_date = time.strptime(date[0], "%Y-%m-%d")
        if (temp_date >= begin_date) and (temp_date <= end_date):
            datas.append(data)

    # 格式化
    need = []
    for i in range(len(datas)):
        one_data = [None] * (len(title) - 1)
        # 日期时间
        one_data[0] = datas[i][4]
        # 公司名称
        one_data[1] = datas[i][0]
        # 宣讲地址
        one_data[2] = str_replace(datas[i][5])
        # 联系人
        one_data[3] = datas[i][14]
        # 电话
        one_data[4] = datas[i][15]
        # 手机
        one_data[5] = datas[i][16]
        # 笔试时间
        one_data[6] = datas[i][6]
        # 笔试地址
        one_data[7] = str_replace(datas[i][7])
        # 面试时间
        one_data[8] = datas[i][8]
        # 面试地址
        one_data[9] = str_replace(datas[i][9])
        # 下发学院
        one_data[10] = str_replace(datas[i][18])

        need.append(one_data)

    need.sort()
    for i in range(len(need)):
        tdate, ttime = need[i][0].split( )
        need[i][0] = tdate
        need[i].insert(1, ttime)

    # write excel
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('sheet1')
    for i in range(len(need)+1):
        for j in range(len(need[0])):
            if i == 0:
                sheet.write(0, j, title[j])
            else:
                sheet.write(i, j, need[i-1][j])
    outname = args.begin_date[5:7] + args.begin_date[8:10] + '-' + args.end_date[5:7] + args.end_date[8:10]
    workbook.save(outname+'.xls')
    print('Save the result at', outname+'.xls')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="weekly excel")
    # set date
    parser.add_argument('-begin_date', type=str, default='2018-05-12')
    parser.add_argument('-end_date', type=str, default='2018-05-25')

    main(parser.parse_args())
