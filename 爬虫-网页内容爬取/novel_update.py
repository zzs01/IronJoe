# -*- coding:utf8 -*-

import requests
import re
import sys
from time import strftime   
from datetime import datetime
import sqlite3


# 查询数据    
def query_sql(name,query_key):
    conn = sqlite3.connect('./novel.db')
    cur = conn.cursor()
    que_sql = '''SELECT %s from NOVEL where NAME = "%s"''' % (query_key,name)
    
    tb_sql = '''CREATE TABLE NOVEL 
               (NAME   CHAR(50)  PRIMARY KEY  UNIQUE  NOT NULL,
               LINK             CHAR(50),
               LOCATION         INT,
               CHAPTER          CHAR(50));'''
    de_sql = '''DROP TABLE NOVEL;'''
    try:
        cursor = cur.execute(que_sql)
    except Exception as er:
        try:
            cur.execute(tb_sql)
        except Exception as err:
            cur.execute(de_sql)
            cur.execute(tb_sql)
        cursor = cur.execute(que_sql)
    data_tupple_list = cursor.fetchall()
    data_list = [str(i[0]) for i in data_tupple_list if i[0] != '']
    data = "".join(data_list)
    conn.close()
    return data

# 添加或更新数据库    
def create_sql(novel_name, link, location=-1, chapter=""):
    conn = sqlite3.connect('./novel.db')
    cur = conn.cursor()
    ins_sql = '''INSERT INTO NOVEL (NAME,LINK,LOCATION,CHAPTER) \
        VALUES ("%s", "%s", %s, "%s") \
        ''' % (novel_name, link, location, chapter)
    up_sql = '''UPDATE NOVEL set LINK = "%s",LOCATION = %s,CHAPTER = "%s"  where NAME = "%s"''' % (link, location, chapter, novel_name)
    
    if query_sql(novel_name,"name"):
        cur.execute(up_sql)
    else:
        cur.execute(ins_sql)
        
    conn.commit()
    conn.close()

#爬取并推送小说最新章节
#来源，笔趣阁
#首次运行只发送最新一章
headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
'Accept-Encoding': 'gzip, deflate',
'Accept-Language': 'zh-CN,zh;q=0.9',
'Cache-Control': 'no-cache',
'Connection': 'keep-alive',
'Pragma': 'no-cache',
'Upgrade-Insecure-Requests': '1',
'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.116 Safari/537.36}'
}
# 获取小说链接
def get_novel_link(novel_name, searchkey):
    home_page = "http://www.biquge.info/"
#第二个源    https://www.shujy.com/ 章节名字不一样
    html = requests.get(home_page,headers=headers)
    
    home_page_link = "http://www.biquge.info/modules/article/search.php?searchkey=" + searchkey
    home_page_response = requests.get(home_page_link,headers=headers)
    home_page_response.encoding = home_page_response.apparent_encoding
    home_page_text = home_page_response.text

    match_key = r'>%s<[\s.|\S.]+?href="(.+?/)?\d.*'%searchkey
    # 小说不存在
    try:
        red = re.findall(match_key,home_page_text,re.DOTALL)[0]
    except Exception as err:
        dd_send("找不到书籍信息", "书籍不存在", "请检查书名" + searchkey + "是否正确")
        return ""
    novel_link = home_page+red
    create_sql(novel_name, novel_link)
    return novel_link
    
# 钉钉推送
def dd_send(novel_name, novel_title, novel_content):
    from dingtalkchatbot.chatbot import DingtalkChatbot, is_not_null_and_blank_str, ActionCard, FeedLink, CardItem

    webhook_token = ''
    webhook = 'https://oapi.dingtalk.com/robot/send?access_token=' + webhook_token
    secret = ''  # 可选：创建机器人勾选“加签”选项时使用
    if not webhook_token or not secret:
        print("请在程序填写钉钉推送token和secret")
        return
    xiaoding = DingtalkChatbot(webhook, secret=secret)
    now = datetime.now()
    now_time = now.strftime("%Y年%m月%d日 %H:%M:%S")
    dd_content = '''%s \n
%s \n
%s \n
%s
''' % (novel_name,novel_title,novel_content,now_time)
    send_title = novel_name + " " + novel_title
    xiaoding.send_markdown(title=send_title, text=dd_content)
    
# 获取小说章节内容并推送        
def get_novel_content(chapter_link, chapter_title, novel_link, novel_name):  
    req = requests.get(novel_link + chapter_link + ".html",headers=headers)
    req.encoding = req.apparent_encoding
    content = req.text
    re_novel_content = r'<div id="content"><!--go-->(.*?)</div>'
    chapter_content = re.findall(re_novel_content,content,re.DOTALL)[0]
    chapter_content = chapter_content.replace(r"<br/>",'\n')
    drop_content = "正在手打中，请稍等片刻，内容更新后"
    if re.search(drop_content,chapter_content) == None:
        dd_send(novel_name, chapter_title,chapter_content)
    else:
        return "failed"
    
# 获取小说更新状态    
def get_novel_status(novel_name, newest_chapter_location):    
    search_key = novel_name
    
    novel_link = query_sql(novel_name,"link")
    if not novel_link:
        novel_link = get_novel_link(novel_name, search_key)
    
    # 小说不存在
    if not novel_link:
        return
    
    local_path="."
    # 获取当前工作目录, 使用绝对路径是可用（软路由上可用）
    # path = sys.argv[0].split("/")
    # local_path = "/".join(path[0:-1])
        
    res = requests.get(novel_link,headers=headers)
    res.encoding = res.apparent_encoding
    content = res.text
    re_novel_content = r'<dd><a href="(.*?).html" title="'
    novel_content = re.findall(re_novel_content,content,re.DOTALL)
    re_novel_title = r' title="(.*?)">'
    novel_title = re.findall(re_novel_title,content,re.DOTALL)

    newest = newest_chapter_location
    latest_chapter_title = novel_title[newest]
        
    # 查询更新章节数
    def update_chapters(saved_content,last_newest):
        this_newest = last_newest + 1
        if saved_content == "":
            return this_newest
        while saved_content != novel_title[last_newest]:
            last_newest = last_newest -1
            if newest - last_newest > 10:
                get_novel_link(novel_name,search_key)
                get_novel_status(novel_name, newest_chapter_location)
                return
        else:
            return last_newest + 1

    #查询是否最新一章
    content = query_sql(novel_name,"chapter")
        
    if content == latest_chapter_title:
        return
    else:
        for i in range(update_chapters(content, newest - 1), newest + 1):
            if_update = get_novel_content(novel_content[i],novel_title[i], novel_link, search_key)
            if if_update == "failed":
                create_sql(novel_name, novel_link, newest_chapter_location, novel_title[i])
                return
        create_sql(novel_name, novel_link, newest_chapter_location, latest_chapter_title)
    
    #给另一个文件传参
    global status
    status = "success"

    
if __name__ == '__main__':
    novel_name_cn = sys.argv[1]
    #如修罗武神最新一章在倒数第13章，需要指定第二个参数
    try:
        newest_chapter_location = int("-" + str(sys.argv[2]))
    except Exception as er:
        newest_chapter_location = -1
        
    get_novel_status(novel_name_cn, newest_chapter_location)
