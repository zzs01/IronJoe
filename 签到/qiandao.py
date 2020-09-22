#_Author:Iron Joe
#date:2020/9/17

'''
已实现：
1.自动登录
2.自动签到，在于formhash动态命令，先登录后在抓取
未实现：
1.判断是否已登录
2.返回是否登录成功
3.定时启动
4.用户名，密码改成变量登录
'''
import requests
import time
import re

forum_session = requests.session()

all_headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.116 Safari/537.36',
    'Origin': 'http://120.77.14.141',
    'Host': '120.77.14.141'}

def response_get_text(url_,reg_rule):
    response_get = forum_session.get(url_, verify=False)
    ret = re.findall(reg_rule, response_get.content.decode('utf-8'), re.DOTALL)[0]# re.DOTALL代表全局搜索，忽略\n，并入搜索
    return ret

def login_post():
    login_url = 'http://120.77.14.141/member.php?mod=logging&action=login&loginsubmit=yes&loginhash=LtOdf&inajax=1'
    login_data = {'username':"",
                  'password':""}
    forum_session.post(url=login_url, data=login_data, headers=all_headers)

def sign_in(form_hash_,homepage_url_):
    sign_in_url = 'http://120.77.14.141/plugin.php?id=dc_signin:sign&inajax=1'
    sign_in_data = {'formhash':form_hash_,
                    'signsubmit': 'yes',
                    'handlekey': 'signin',
                    'emotid': '7',
                    'referer':homepage_url_,
                    'comment':'0000'}
    ret_text=forum_session.post(url=sign_in_url,data=sign_in_data,headers=all_headers)

    return ret_text

def sign_in_judge(text_):
    reg1 = r".*,\s'(.*?)'.*" #签到成功的匹配规则
    ret1 = re.findall(reg1, text_.content.decode('utf-8'), re.DOTALL)# re.DOTALL代表全局搜索，忽略\n，并入搜索

    reg2 = r".*CDATA\[(.*)<script.*"#已经签到的匹配规则
    ret2 = re.findall(reg2, text_.content.decode('utf-8'), re.DOTALL)# re.DOTALL代表全局搜索，忽略\n，并入搜索
    if ret1:
        ret = ret1[0]
    elif ret2:
        ret = ret2[0]
    else:
        ret = '签到失败~~'
    return ret



def main_operation():
    # 登录
    login_post()

    # 登录后进入个人主页,提取formhash动态参数
    homepage_url = 'http://120.77.14.141/?129'
    reg_rule1 = r".*formhash=(.*)&amp;searchsubmit.*"
    form_hash = response_get_text(homepage_url,reg_rule1)

    # 自动签到
    sign_in_text = sign_in(form_hash, homepage_url)
    title=sign_in_judge(sign_in_text)
    time.sleep(1)

    # 字段
    money_url='http://120.77.14.141/plugin.php?id=dc_signin&action=index'
    all_reg = r'.*?<div class="mytips">(.*?)</div>\r\n</div>.*?'#提取的内容
    all_text=response_get_text(money_url, all_reg)
    main_comment = all_text.replace(r"\r\n",'').replace(r"<b>",'**').replace(r"</b>",'**').replace(r"</p>",'``` ```\n').replace(r"<p>",'')
    content ='''
-----------------
%s
    '''%(main_comment)
    print(content)

    SCKEY = ""
    api = "https://sc.ftqq.com/" + SCKEY + ".send"
    data = {
       "text":title,
       "desp":content
    }
    req = requests.post(api,data = data)
    
if __name__ == '__main__':
    main_operation()
