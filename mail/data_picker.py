import os,sys,io
import time
import datetime
import traceback

import crawler_util as util

# 输入参数

# 输入收件人地址:
email_target_addr_list = ["625259906@qq.com","277693035@qq.com"]#,"625259906@qq.com","277693035@qq.com"]
email_debug_addr_list = ["342875289@qq.com","625259906@qq.com"]


# 输入SMTP服务器地址:
smtp_server = "smtp.sina.com"
# 输入Email地址和口令:
email_source_addr = "mon_dayuan@sina.com"
email_password = "labcat127"
# 调试参数
# 是否发送邮件
isSendEmail = 1
isWriteFile = 0
isWriteClipboard = 0
# 是否打开调试模式
isDebug = 0
delta_time = 1


# 程序区 请勿改动
# 获得当天日期
today = datetime.date.today()
# 计算目标日期
tagert_date = today + datetime.timedelta(days=delta_time)
print("今天是"+today.isoformat())
print("尝试获取"+tagert_date.isoformat()+"的信息")

try:
    # 获取指定某个月的招聘会/宣讲会信息
    data_one_day_a_slice = util.get_one_month_data(tagert_date.year,tagert_date.month,isDebug)

    # 从一个月的数据中整理出某一天的详细数据
    list_zhaopin, list_xuanjiang = util.get_one_day_data(data_one_day_a_slice,tagert_date,isDebug)

    # 分析各场次所在时间段,生成文本描述
    str_email = util.get_email_content(list_zhaopin,list_xuanjiang,tagert_date)
except Exception as e:# failed
    str_email  = '数据获取失败通知:\n'
    str_email += '尝试获取的数据日期:' + tagert_date.isoformat() + '\n'
    str_email += '错误信息:\n' + traceback.format_exc() + '\n'
    print(str_email)
    # 发送出错后的通知邮件 
    util.send_email(smtp_server, email_source_addr, email_password, email_debug_addr_list, tagert_date.isoformat(),str_email,sub_add='数据获取失败')

else:# sucess
    print('数据获取正常')
    print(str_email)

    if isSendEmail:
        util.send_email(smtp_server, email_source_addr, email_password, email_target_addr_list, tagert_date.isoformat(),str_email)

    if isWriteFile:
        txt_file = open(os.path.dirname(os.path.realpath(__file__))+'\\'+tagert_date.isoformat()+'的招聘会宣讲会信息.txt','w',encoding='utf-8')
        txt_file.write(str_email)
        txt_file.close()
        print("宣讲会的详细内容已保存到txt")

    if isWriteClipboard:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, str_email)
        win32clipboard.CloseClipboard()
        print("宣讲会的详细内容已复制到剪切板")

print('程序会在三秒后关闭')
time.sleep(3)
