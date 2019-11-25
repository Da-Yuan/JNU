import urllib.request
import http.cookiejar
import json
import os


# 加入对cookies的支持
cookie_file = 'cookies'
cookie = http.cookiejar.MozillaCookieJar(cookie_file)
# 首先尝试读取上次退出时保存的cookeis
if os.path.isfile(cookie_file):
    cookie.load(cookie_file, ignore_discard=True, ignore_expires=True)

handler = urllib.request.HTTPCookieProcessor(cookie)
opener = urllib.request.build_opener(handler)
urllib.request.install_opener(opener)

headers_dict = {
    'Connection': 'keep-alive',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64, AppleWebKit/537.36 (KHTML, like Gecko, Maxthon/4.4.8.2000 Chrome/30.0.1599.101 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
    'DNT': '1',
    'Host': 'jiangnan.91job.org.cn',
    'Origin': 'http://jiangnan.91job.org.cn',
    'Referer': 'http://jiangnan.91job.org.cn/admin/default/login',
    'Accept-Language': 'zh-CN',
    }

# 用于测试是否需要验证码登录
url_index = 'http://jiangnan.91job.org.cn/admin/default/index'
# 创建Request对象
request = urllib.request.Request(url_index)
# 添加http header
[request.add_header(key, value) for key, value in headers_dict.items()]
# 发送请求
response = urllib.request.urlopen(request, timeout=60)
# 保存网页内容
context = response.read().decode('UTF8')
# 输出网页内容
# print(context)
if context.find('当前登录用户') == -1:  # cookies无效,重新登录
    print('未读取到有效的登录记录,请输入收到的验证码,按回车结束')
    # 获取验证码
    url_sendcode = 'http://jiangnan.91job.org.cn/admin/default/sendcode'
    request = urllib.request.Request(url_sendcode)
    [request.add_header(key, value) for key, value in headers_dict.items()]
    # POST_Data
    post_data = {'username': 'jyzdzx'}
    post_data_code = urllib.parse.urlencode(post_data).encode(encoding='UTF8')
    # 发送请求
    response = urllib.request.urlopen(request, data=post_data_code, timeout=60)
    context = response.read().decode('UTF8')
    # 输出网页内容
    response_json = json.loads(context)
    print(response_json['message'])
    cookie.save(ignore_discard=True, ignore_expires=True)

    code = str(input())
    # 完成登陆
    url_login = 'http://jiangnan.91job.org.cn/admin/default/login'
    # 创建Request对象
    request = urllib.request.Request(url_login)
    # 添加http header
    [request.add_header(key, value) for key, value in headers_dict.items()]
    # POST_Data
    post_data = {'UserULoginForm[username]': 'jyzdzx',
                 'UserULoginForm[password]': 'Password123',
                 'UserULoginForm[smsCode]': code,
                 'UserULoginForm[rememberMe]': '1',
                 'yt1': '登录',
                 }
    post_data_code = urllib.parse.urlencode(post_data).encode(encoding='UTF8')
    # 发送请求
    response = urllib.request.urlopen(request, data=post_data_code, timeout=60)
    context = response.read().decode('UTF8')
    cookie.save(ignore_discard=True, ignore_expires=True)
else:
    print('读取到有效的登录记录,准备开始下载')

# 下载excel
url_index = 'http://jiangnan.91job.org.cn/admin/teachin/export?x=1&page=1&limit=50&target=_blank'
# 创建Request对象
request = urllib.request.Request(url_index)
[request.add_header(key, value) for key, value in headers_dict.items()]
print('正在下载....')
response = urllib.request.urlopen(request, timeout=120)
# 保存execl
excel_data = response.read()
with open('excel.xlsx', mode='wb') as f:
    f.write(excel_data)
print('下载完成')
