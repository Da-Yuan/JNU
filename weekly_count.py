import argparse
import time
import xlrd
import xlwt


def main():
    begin_date = time.strptime('2019-09-01', "%Y-%m-%d")
    # end_date = time.localtime()
    # end_date = time.strftime("%Y-%m-%d", time.localtime())
    end_date = time.strptime('2019-11-18', "%Y-%m-%d")
    # print(end_date)
    # exit(0)

    if begin_date > end_date:
        print('Check the date.')
        exit(0)
    # 读入文件
    file = xlrd.open_workbook('excel.xlsx')
    ## 计算宣讲会数
    # 读入子表1
    tabel = file.sheets()[0]
    nrows = tabel.nrows

    company = []
    for i in range(1, nrows):
        data = tabel.row_values(i)
        state = data[-5]
        date = data[4].split( )
        if date[0]=='活动已取消':
            continue
        temp_date = time.strptime(date[0], "%Y-%m-%d")
        if (temp_date >= begin_date) and (temp_date <= end_date) and state == '审核成功':
            company.append(data[0])
    print('已经举办的宣讲会数：', len(company))
    # print(company)

    # 计算宣讲会的岗位数
    tabel_jobs = file.sheets()[1]
    nrows = tabel_jobs.nrows
    datas = []
    for i in range(1, nrows):
        data = tabel_jobs.row_values(i)
        state = data[-2]
        date = data[9].split()
        temp_date = time.strptime(date[0], "%Y-%m-%d")
        if (temp_date >= begin_date) and (temp_date <= end_date) and state == '审核成功':
            if data[0] in company:
                datas.append(int(data[2]))
    print('宣讲会提供的岗位数：', sum(datas))
    exit(0)


if __name__ == '__main__':
    main()
