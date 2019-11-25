import urllib.request
import re,os,sys
import smtplib,time
from email.mime.text import MIMEText
from email.header import Header
import datetime,gzip


# 获取指定某个月的招聘会/宣讲会信息
def get_one_month_data(year,month,isDebug):

    url='http://jiangnan.91job.org.cn/default/date'
    #创建Request对象
    request = urllib.request.Request(url)
    #添加数据
    #添加http header
    request.add_header('Host', 'jiangnan.91job.org.cn')
    request.add_header('User-Agent','Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.4.8.2000 Chrome/30.0.1599.101 Safari/537.36')
    request.add_header('Connection', 'keep-alive')
    # request.add_header('Content-Length','18')
    request.add_header('Accept', '*/*')
    request.add_header('Origin','http://jiangnan.91job.org.cn')
    request.add_header('Content-Type','application/x-www-form-urlencoded; charset=UTF-8')#
    request.add_header('DNT','1')
    request.add_header('Referer','http://jiangnan.91job.org.cn/')
    request.add_header('X-Requested-With','XMLHttpRequest')
    request.add_header('Accept-Encoding','gzip,deflate')
    request.add_header('Accept-Language','zh-CN')
    request.add_header('Cookie', 'PHPSESSID2=1tbfvr7r60nbgcu19spale0g82; __jsluid=0cfb92f2b0a229fd890c619271a30d66; UM_distinctid=165be61c97d15c-097a760eba4651-9393265-1fa400-165be61c98610; Hm_lvt_d7682ab43891c68a00de46e9ce5b76aa=1536497458; Hm_lpvt_d7682ab43891c68a00de46e9ce5b76aa=1536497575; CNZZDATA1259394577=1060619681-1536494425-%7C1536549820')

    # POST_Data
    post_data = {'year' : str(year),
    'month' : str(month)}
    post_data_code= urllib.parse.urlencode(post_data).encode(encoding='UTF8')

    # 发送请求
    response = urllib.request.urlopen(request,data=post_data_code)
    data = gzip.decompress(response.read()).decode('unicode_escape')#UTF-8
    # 去掉大括号
    data = data[1:len(data) - 1]
    data = data.replace('\\', '')
    # 构造用于按照时间切分
    p_getOneDay = re.compile(r'","')

    #把数据按时间天数切分
    data_one_day_a_slice=p_getOneDay.split(data)
    num_split = len(data_one_day_a_slice)
    if(isDebug==1):
        print("收集到"+str(num_split-1)+"天的信息，分别是:")
        for i in range(num_split):
            print(data_one_day_a_slice[i])

    return data_one_day_a_slice




# 从一个月的数据中整理出某一天的详细数据
def get_one_day_data(data_one_day_a_slice,tagert_date,isDebug):
    list_zhaopin, list_xuanjiang = [],[]
    count_zhaopin,count_xuanjiang = 0,0
    # 筛选日期参数
    str_date = str(tagert_date.year) + '-' + str(tagert_date.month) + '-' + str(tagert_date.day)


    # 用于提取招聘会/宣讲会的链接
    p_getItemDetail = re.compile(r'href="(.+)">(.+)</a')
    # 用于提取招聘会/宣讲会的详细信息
    # p_item_detail = re.compile(r'([^【]+)【(.+)】[^<]*<')
    p_item_detail = re.compile(r'"date-subtitle"> *([^【]+)【(.+)】[^<]*<')
    # >　　机械学院a202【18:30-21:00】</li>
    # 提取招聘会信息
    # p_item_detail_time1 = re.compile(r'时间[^<]*<span>(\d{4}-\d{2}-\d{1,2}) (.{5,11})</span>')
    # p_item_detail_location1 = re.compile(r'举办地址[^<]*<span>(.+)</span>')
    # 提取宣讲会信息
    # p_item_detail_time2 = re.compile(r'宣讲时间[^<]*<span>(.{10}) (.{11})')
    # p_item_detail_location2 = re.compile(r'宣讲地址[^<]*<span>(.+)</span>')

    # < li >　　北区活动中心F105【14: 00 - 16:00】 < / li >

    num_split=len(data_one_day_a_slice)
    for i in range(num_split-1):#减一为了屏蔽最后的时间日期数据
        if(isDebug):
            print(i)
            print(data_one_day_a_slice[i])
        str_data = data_one_day_a_slice[i]
        if(str_data.find(str_date+'"')!=-1):
            # 切分目标日期的每个宣讲会/招聘会信息
            # 一天中的会议场次
            items_in_oneday = str_data.split('<li>')
            num_a_day_split = len(items_in_oneday)
            # print('items_in_oneday:{}'.format(items_in_oneday))
            # print('num_a_day_split:{}'.format(num_a_day_split))
            if(isDebug==1):
                for i in range(num_a_day_split):
                    print('a_day_data_after_split:'+items_in_oneday[i])
            #for j in range(1,num_a_day_split,2):
            for j in range(1, num_a_day_split):
                item_link_name = items_in_oneday[j]
                item_location_time = items_in_oneday[j]
                #item_location_time = items_in_oneday[j+1]
                [link,enterpriseName] = p_getItemDetail.search(item_link_name).groups()

                # print(item_link_name)
                # print(item_location_time)

                [location, time_t]= p_item_detail.search(item_location_time).groups()
                location = location.replace('　','').replace(' ','')
                # print('item_location_time:{}'.format(item_location_time))
                # print('location:{}'.format(location))
                # print('time_t:{}'.format(time_t))

                if(link.find('jobfair') != -1 ):#是招聘会
                    count_zhaopin=count_zhaopin+1
                    # print(the_page)
                    # print(p_item_detail_time1.search(the_page).groups())
                    # time_t=p_item_detail_time1.search(the_page).group(2)
                    # location=p_item_detail_location1.search(the_page).group(1)
                    #print(time_t)
                    #print(location)
                    list_zhaopin.append([enterpriseName,time_t,location,0,'http://jiangnan.91job.org.cn'+link])
                elif(link.find('teachin') != -1 ):#宣讲会
                    count_xuanjiang=count_xuanjiang+1
                    # time_t=p_item_detail_time2.search(the_page).group(2)
                    # location=p_item_detail_location2.search(the_page).group(1)
                    #print(time_t)
                    #print(location)
                    list_xuanjiang.append([enterpriseName,time_t,location,0,'http://jiangnan.91job.org.cn'+link])

    def _getKey( x ):
        temp = int(x[1][3:5])+int(x[1][0:2])*100
        return int(temp)
    list_zhaopin.sort(key=_getKey)
    list_xuanjiang.sort(key=_getKey)

    return list_zhaopin,list_xuanjiang




# 分析各场次所在时间段,生成文本描述
def get_email_content(list_zhaopin,list_xuanjiang,tagert_date):
    def time_division( item ):
        time_division_num=[0,0,0]
        for i in range(len(item)):
            if(item[i][1][0:2]<="12"):
                time_division_num[0]=time_division_num[0]+1
                item[i][3]=1
            elif(item[i][1][0:2]>"12" and item[i][1][0:2]<"18"):
                time_division_num[1]=time_division_num[1]+1
                item[i][3]=2
            else:
                time_division_num[2]=time_division_num[2]+1
                item[i][3]=3
        return time_division_num

    count_zhaopin=len(list_zhaopin)
    count_xuanjiang=len(list_xuanjiang)
    # 邮件开头
    str_date = str(tagert_date.year) + '-' + str(tagert_date.month) + '-' + str(tagert_date.day)
    str_email ="日期:"+str_date+"\n"
    time_division_zhaopin=time_division(list_zhaopin)
    time_division_xuanjiang=time_division(list_xuanjiang)

    if(count_zhaopin==0 and count_xuanjiang==0):
        str_email+="今天没有招聘、宣讲会"
    else:
        str_email+="各学院:今天共有"
        if(count_zhaopin!=0):
            str_email+=str(count_zhaopin)+"场招聘会"
        if(count_zhaopin!=0 and count_xuanjiang!=0):
            str_email+="和"
        if(count_xuanjiang!=0):
            str_email+=str(count_xuanjiang)+"场宣讲会"
        str_email+="，请通知学生查看信息、积极参加:\n"
        time_division = 0
        time_division_i=1
        if(count_zhaopin!=0):
            str_email +="招聘会"+"\n"
        for i in range(count_zhaopin):
            if(list_zhaopin[i][3]!=time_division):
                if(list_zhaopin[i][3]==1):
                    str_email +="上午"+str(time_division_zhaopin[0])+"场"+"\n"
                elif(list_zhaopin[i][3]==2):
                    str_email +="下午"+str(time_division_zhaopin[1])+"场"+"\n"
                else:
                    str_email +="晚上"+str(time_division_zhaopin[2])+"场"+"\n"
                time_division = list_zhaopin[i][3]
                time_division_i=1
            #print(str(time_division_i)+'.'+list_zhaopin[i][0]+","+list_zhaopin[i][1]+','+list_zhaopin[i][2]+',详情请查看:'+list_zhaopin[i][4])
            str_email +=str(time_division_i)+'.'+list_zhaopin[i][0]+","+list_zhaopin[i][1]+','+list_zhaopin[i][2]+',详情请点击:'+list_zhaopin[i][4]+"\n"
            time_division_i=time_division_i+1
        time_division = 0
        time_division_i=1
        if(count_xuanjiang!=0):
            str_email +="宣讲会"+"\n"
        for i in range(count_xuanjiang):
            if(list_xuanjiang[i][3]!=time_division):
                if(list_xuanjiang[i][3]==1):
                    str_email +="上午"+str(time_division_xuanjiang[0])+"场"+"\n"
                elif(list_xuanjiang[i][3]==2):
                    str_email +="下午"+str(time_division_xuanjiang[1])+"场"+"\n"
                else:
                    str_email +="晚上"+str(time_division_xuanjiang[2])+"场"+"\n"
                time_division = list_xuanjiang[i][3]
                time_division_i=1

            str_email +=str(time_division_i)+'.'+list_xuanjiang[i][0]+","+list_xuanjiang[i][1]+','+list_xuanjiang[i][2]+',详情请点击:'+list_xuanjiang[i][4]+"\n"
            time_division_i=time_division_i+1

    return str_email

def send_email(smtp_server,email_source_addr,email_password,email_target_addr_list,str_date,str_email,sub_add='',smtp_server_port=465):
    server = smtplib.SMTP_SSL(smtp_server, smtp_server_port)
    #server = smtplib.SMTP(smtp_server)
    server.set_debuglevel(0)
    server.login(email_source_addr,email_password)
    msg = MIMEText(str_email, 'plain', 'utf-8')
    msg['From'] = Header("Xiao_Ming<mon_dayuan@sina.com>")
    msg['Subject'] = Header(sub_add + str_date + "招聘宣讲信息", 'utf-8')
    for target_addr in email_target_addr_list:
        msg['To'] = Header('<' + target_addr + '>', 'utf-8')
        server.sendmail(email_source_addr,target_addr, msg.as_string())
    server.quit()
    print("宣讲会的详细内容已通过邮件发送")

