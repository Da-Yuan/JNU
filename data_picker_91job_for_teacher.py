import urllib.request
import re,os
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import datetime
#输入参数
#输入Email地址和口令:
from_addr = "mon_dayuan@sina.com"
password = "labcat127"
# 输入收件人地址:
to_addr = "342875289@qq.com"
# 输入SMTP服务器地址:
smtp_server = "smtp.sina.com"
#调试参数
#是否发送邮件
isSendEmail = 0
#是否打开调试模式
isDebug = 0


#计算当天日期
today = datetime.date.today()
#计算第二天的日期
tomorrow = today + datetime.timedelta(days=1) 
print("今天是"+today.isoformat())
print("尝试获取明天"+tomorrow.isoformat()+"的信息")


txt_file = open(os.path.dirname(os.path.realpath(__file__))+'\\'+tomorrow.isoformat()+'的招聘会宣讲会信息.txt','w')


#HTTPCookiesProcessor
#HTTPRedirectHandler
url='http://jiangnan.91job.gov.cn/default/date'
#创建Request对象
request = urllib.request.Request(url)
#添加数据
#添加http header
request.add_header('Host', 'jiangnan.91job.gov.cn')
request.add_header('User-Agent','Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.4.8.2000 Chrome/30.0.1599.101 Safari/537.36')
request.add_header('Connection', 'keep-alive')
request.add_header('Content-Length','17')
request.add_header('Accept', '*/*')
request.add_header('Origin','http://jiangnan.91job.gov.cn')
request.add_header('Content-Type','application/x-www-form-urlencoded')
request.add_header('DNT','1')
request.add_header('Referer','http://jiangnan.91job.gov.cn/')
request.add_header('X-Requested-With','XMLHttpRequest')
#request.add_header('Accept-Encoding','gzip,deflate')
request.add_header('Accept-Language','zh-CN')
request.add_header('Cookie', 'PHPSESSID=4adcb2a5571b6b6724600; TS_think_language=zh-CN;TS_LOGGED_USER=fQkWi20ie94JrnxRpBXxIIX88;Hm_lvt_dd3ea352543392a029ccf9da1be54a50=1461215406,1461300191,1461414536,1461417954;Hm_lpvt_dd3ea352543392a029ccf9da1b')

#POST_Data
post_data = {'year' : str(tomorrow.year),
'month' : str(tomorrow.month)}
post_data_code= urllib.parse.urlencode(post_data).encode(encoding='UTF8')

#发送请求
response = urllib.request.urlopen(request,data=post_data_code)
the_page = response.read()
data = the_page.decode('unicode_escape')
print(data)

str_email=""
#筛选日期参数
str_date = str(tomorrow.year)+'-'+str(tomorrow.month)+'-'+str(tomorrow.day)
#编辑邮件内容信息
str_email +="日期:"+str_date+"\n"
#去掉大括号
data=data[1:len(data)-1]
#print(data)
data=data.replace('\\','')
#print(data)
list_zhaopin =[]
list_xuanjiang=[]
#用于按照时间切分
p_getOneDay = re.compile(r'","')
#用于切分每一个项目
p_getOneItem = re.compile(r'<li>')
#用于提取项目的详细信息
p_getItemDetail = re.compile(r'href="((.)+)">((.)+)</a')
#提取招聘会信息
p_item_detail_time1 = re.compile(r'时间.<span>(\d){4}-(\d){2}-(\d){1,2} ((.){5,11})</span>')
p_item_detail_location1 = re.compile(r'举办地址.<span>((.)+)</span>')
#提取宣讲会信息
p_item_detail_time2 = re.compile(r'宣讲时间.<span>(.){10} ((.){11})')
p_item_detail_location2 = re.compile(r'宣讲地址.<span>((.)+)</span>')
#把数据按时间天数切分
data_after_split=p_getOneDay.split(data)
num_split = len(data_after_split)
if(isDebug==1):
    print("收集到"+str(num_split-1)+"天的信息，分别是:")
    for i in range(num_split):
        print(data_after_split[i])
num_zhaopin=0
num_xuanjiang=0

for i in range(num_split-1):#减一为了屏蔽最后的时间日期数据
    str_data = data_after_split[i]
    if(str_data.find(str_date+'"')!=-1):
        #切分出每个项目
        a_day_data_after_split=p_getOneItem.split(str_data)
        num_a_day_split = len(a_day_data_after_split)
        if(isDebug==1):
            for i in range(num_a_day_split):
                print(a_day_data_after_split[i])
        for j in range(1,num_a_day_split):
            a_day_a_piece = a_day_data_after_split[j]
            link = "http://jiangnan.91job.gov.cn"+p_getItemDetail.search(a_day_a_piece).group(1)
            enterpriseName = p_getItemDetail.search(a_day_a_piece).group(3)
    #        print(link)
    #        print(enterpriseName)
            url=link
            #创建Request对象
            request = urllib.request.Request(url)
            response = urllib.request.urlopen(request)
            the_page = response.read().decode('utf8')
            #print(the_page)
            if(a_day_a_piece[0:3]=="招聘会"):
                num_zhaopin=num_zhaopin+1
                time=p_item_detail_time1.search(the_page).group(4)
                location=p_item_detail_location1.search(the_page).group(1)
                print(time)
    #            print(location)
                list_zhaopin.append([enterpriseName,time,location,0,link])
            elif(a_day_a_piece[0:3]=="宣讲会"):
                num_xuanjiang=num_xuanjiang+1
                time=p_item_detail_time2.search(the_page).group(2)
                location=p_item_detail_location2.search(the_page).group(1)
     #          print(time)
    #           print(location)
                list_xuanjiang.append([enterpriseName,time,location,0,link])
    else :
      #  print(str(i)+" NO")
        continue
def getKey( x ):  
    temp = int(x[1][3:5])+int(x[1][0:2])*100
    return int(temp)
list_zhaopin.sort(key=getKey)
list_xuanjiang.sort(key=getKey)
#print(list_zhaopin)
#print(list_xuanjiang)

#分析各场次所在时间段
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

time_division_zhaopin=time_division(list_zhaopin)
time_division_xuanjiang=time_division(list_xuanjiang)
if(num_zhaopin==0 and num_xuanjiang==0):
    print("今天没有招聘、宣讲会")
    str_email+="今天没有招聘、宣讲会"
else:
    print("各学院:今天共有", end='')
    str_email+="各学院:今天共有"
    if(num_zhaopin!=0):
        print(str(num_zhaopin)+"场招聘会", end='')
        str_email+=str(num_zhaopin)+"场招聘会"
    if(num_zhaopin!=0 and num_xuanjiang!=0):
        print("和",end='')
        str_email+="和"
    if(num_xuanjiang!=0):
        print(str(num_xuanjiang)+"场宣讲会", end='')
        str_email+=str(num_xuanjiang)+"场宣讲会"
    print("，请通知学生查看信息、积极参加:")
    str_email+="，请通知学生查看信息、积极参加:\n"
    time_division = 0
    time_division_i=1
    if(num_zhaopin!=0):
        print("招聘会")
        str_email +="招聘会"+"\n"
    for i in range(num_zhaopin):
        if(list_zhaopin[i][3]!=time_division):
            if(list_zhaopin[i][3]==1):
                print("上午"+str(time_division_zhaopin[0])+"场")
                str_email +="上午"+str(time_division_zhaopin[0])+"场"+"\n"
            elif(list_zhaopin[i][3]==2):
                print("下午"+str(time_division_zhaopin[1])+"场")
                str_email +="下午"+str(time_division_zhaopin[1])+"场"+"\n"
            else:
                print("晚上"+str(time_division_zhaopin[2])+"场")
                str_email +="晚上"+str(time_division_zhaopin[2])+"场"+"\n"
            time_division = list_zhaopin[i][3]
            time_division_i=1
        print(str(time_division_i)+'.'+list_zhaopin[i][0]+","+list_zhaopin[i][1]+','+list_zhaopin[i][2]+',详情请查看:'+list_zhaopin[i][4])
        str_email +=str(time_division_i)+'.'+list_zhaopin[i][0]+","+list_zhaopin[i][1]+','+list_zhaopin[i][2]+',详情请查看:'+list_zhaopin[i][4]+"\n"
        time_division_i=time_division_i+1
    time_division = 0
    time_division_i=1
    if(num_xuanjiang!=0):
        print("宣讲会")
        str_email +="宣讲会"+"\n"
    for i in range(num_xuanjiang):
        if(list_xuanjiang[i][3]!=time_division):
            if(list_xuanjiang[i][3]==1):
                print("上午"+str(time_division_xuanjiang[0])+"场")
                str_email +="上午"+str(time_division_xuanjiang[0])+"场"+"\n"
            elif(list_xuanjiang[i][3]==2):
                print("下午"+str(time_division_xuanjiang[1])+"场")
                str_email +="下午"+str(time_division_xuanjiang[1])+"场"+"\n"
            else:
                print("晚上"+str(time_division_xuanjiang[2])+"场")
                str_email +="晚上"+str(time_division_xuanjiang[2])+"场"+"\n"
            time_division = list_xuanjiang[i][3]
            time_division_i=1
        print(str(time_division_i)+'.'+list_xuanjiang[i][0]+","+list_xuanjiang[i][1]+','+list_xuanjiang[i][2]+',详情请查看:'+list_xuanjiang[i][4])
        str_email +=str(time_division_i)+'.'+list_xuanjiang[i][0]+","+list_xuanjiang[i][1]+','+list_xuanjiang[i][2]+',详情请查看:'+list_xuanjiang[i][4]+"\n"
        time_division_i=time_division_i+1
#在这里插入检测时间安排是否的程序   
list_all = list_zhaopin + list_xuanjiang
list_all.sort(key=getKey)
print(list_zhaopin)
print(list_xuanjiang)
print(list_all)      
#发送邮件
if(isSendEmail==1):
    msg = MIMEText(str_email, 'plain', 'utf-8')
    msg['To'] = Header('<'+to_addr+'>','utf-8')
    msg['From'] = Header("Xiao_Ming<mon_dayuan@sina.com>")
    msg['Subject'] = Header(str_date+"招聘宣讲信息", 'utf-8')
    server = smtplib.SMTP_SSL(smtp_server, 465)
    server.set_debuglevel(1)
    server.login(from_addr,password)
    server.sendmail(from_addr,to_addr, msg.as_string())
    server.quit()
    print("邮件发送完成")
else:
    print("按照选择没有发送邮件")
txt_file.write(str_email)
txt_file.close()