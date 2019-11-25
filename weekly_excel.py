import argparse
import time
import xlrd
import xlwt
from collections import Counter

# 字段替换和简写
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

    title = ['日期',
             '时间',
             '公司名称',
             '宣讲地址',
             '公司联系人', 
             '电话', 
             '手机', 
             '宣讲联系人',
             '联系电话', 
             '笔试时间', 
             '笔试地址', 
             '面试时间', 
             '面试地址', 
             '下发学院'
             ]

    col_width = [10, 10, 35, 10, 10, 15, 12, 10, 15, 10, 10, 10, 10, 12]
    style1_col = [1,3,4,6,7,9,10,11,12,13]
    style2_col = [2,5,8,]

    if begin_date > end_date:
        print('Check the date.')
        exit(0)
    file = xlrd.open_workbook('excel.xlsx')
    tabel = file.sheets()[0]
    nrows = tabel.nrows

    datas = []
    datas_title = tabel.row_values(0)
    for i in range(1, nrows):
        data = tabel.row_values(i)
        state = data[-5]
        date = data[4].split()
        if date[0] == '活动已取消': continue
        temp_date = time.strptime(date[0], "%Y-%m-%d")
        if (temp_date >= begin_date) and (temp_date <= end_date) and state=='审核成功':
            # print(data[0])
            datas.append(data)
    # exit(0)
    # 新表列
    need = []
    for i in range(len(datas)):
        one_data = [None] * (len(title) - 1)
        # 日期时间
        one_data[0] = datas[i][4]
        # 公司名称
        one_data[1] = datas[i][0]
        # 宣讲地址
        one_data[2] = str_replace(datas[i][5])
        # 公司联系人
        one_data[3] = datas[i][14]
        # 电话
        one_data[4] = datas[i][15]
        # 手机
        one_data[5] = datas[i][16]
        # 宣讲联系人
        one_data[6] = datas[i][-3]
        # 联系电话
        one_data[7] = datas[i][-2]  
        # 笔试时间
        one_data[8] = datas[i][6]
        # 笔试地址
        one_data[9] = str_replace(datas[i][7])
        # 面试时间
        one_data[10] = datas[i][8]
        # 面试地址
        one_data[11] = str_replace(datas[i][9])
        # 下发学院
        one_data[12] = str_replace(datas[i][18])

        need.append(one_data)

    need.sort()
    for i in range(len(need)):
        tdate, ttime = need[i][0].split( )
        need[i][0] = tdate
        need[i].insert(1, ttime)

    # 设置居中
    alignment1 = xlwt.Alignment()
    alignment1.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    alignment1.vert = xlwt.Alignment.VERT_CENTER  # 垂直方向
    alignment1.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行

    alignment2 = xlwt.Alignment()
    alignment2.vert = xlwt.Alignment.VERT_CENTER  # 垂直方向
    alignment2.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行

    # 设置边框
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 设置字体
    font = xlwt.Font()
    font.bold = True

    # 定义不同的excel style
    # style1 居中有边框
    style1 = xlwt.XFStyle()
    style1.borders = borders
    style1.alignment = alignment1
    # style2 靠左有边框
    style2 = xlwt.XFStyle()
    style2.borders = borders
    style2.alignment = alignment2
    # title style
    style_title = xlwt.XFStyle()
    style_title.borders = borders
    style_title.alignment = alignment1
    style_title.font = font

    # write excel
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('sheet1')
    sheet_date = []
    # title
    for j in range(len(need[0])):
        sheet.write(0, j, title[j], style=style_title)
        sheet.col(j).width = (col_width[j]+1) * 256
    for i in range(len(need)):
        sheet_date.append(need[i][0])
        for j in range(len(need[0])):
            if j in style1_col:
                sheet.write(i+1, j, need[i][j], style=style1)
            if j in style2_col:
                sheet.write(i+1, j, need[i][j], style=style2)
    # 合并单元格
    count_date = list(Counter(sheet_date).items())
    trow = 1
    for i in range(len(count_date)):
        sheet.write_merge(trow, trow+count_date[i][1]-1, 0, 0, label=count_date[i][0], style=style1)
        trow += count_date[i][1]

    # save file
    outname = args.begin_date[5:7] + args.begin_date[8:10] + '-' + args.end_date[5:7] + args.end_date[8:10]
    workbook.save('./output_excel/'+outname+'.xls')
    print('Save the result at', outname+'.xls')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="weekly excel")
    # set date
    parser.add_argument('-b', '--begin_date', type=str, default='2019-11-24')
    parser.add_argument('-e', '--end_date', type=str, default='2019-11-30')

    main(parser.parse_args())
