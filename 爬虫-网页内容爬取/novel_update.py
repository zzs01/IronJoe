import requests
import re
import sys
from time import strftime   
from datetime import datetime

#爬取并推送小说最新章节
#来源，笔趣阁
#首次运行只发送最新一章

# 获取小说链接
def get_novel_link(searchkey):
    home_page = "http://www.biquge.info/"
    html = requests.get(home_page)

    home_page_link = "http://www.biquge.info/modules/article/search.php?searchkey=" + searchkey
    home_page_response = requests.get(home_page_link)
    home_page_response.encoding = home_page_response.apparent_encoding
    home_page_text = home_page_response.text

    match_key = r'>%s<[\s.|\S.]+?href="(.+?/)?\d.*'%searchkey
    red = re.findall(match_key,home_page_text,re.DOTALL)[0]
    novel_link = home_page+red
    return novel_link
    
# 钉钉推送
def dd_send(novel_title, novel_content):
    from dingtalkchatbot.chatbot import DingtalkChatbot, is_not_null_and_blank_str, ActionCard, FeedLink, CardItem

    webhook_token = ''
    webhook = 'https://oapi.dingtalk.com/robot/send?access_token=' + webhook_token
    secret = ''  # 可选：创建机器人勾选“加签”选项时使用
    if not webhook_token or not secret:
        print("请在程序填写钉钉推送token和secret")
        return
    xiaoding = DingtalkChatbot(webhook, secret=secret)  # 方式二：勾选“加签”选项时使用（v1.5以上新功能）
    now = datetime.now()
    now_time = now.strftime("%Y年%m月%d日 %H:%M:%S")
    search_key = sys.argv[1]
    dd_content = '''%s \n
%s \n
%s \n
%s
''' % (search_key,now_time,novel_title,novel_content)
    send_title = search_key + " " + novel_title
    xiaoding.send_markdown(title=send_title, text=dd_content)
    
# 获取小说章节内容并推送        
def get_novel_content(chapter_link, chapter_title, novel_link):  
    req = requests.get(novel_link + chapter_link + ".html")
    req.encoding = req.apparent_encoding
    content = req.text
    re_novel_content = r'<div id="content"><!--go-->(.*?)</div>'
    chapter_content = re.findall(re_novel_content,content,re.DOTALL)[0]
    chapter_content = chapter_content.replace(r"<br/>",'\n')
    dd_send(chapter_title,chapter_content)
    
# 获取小说更新状态    
def get_novel_status():    
    search_key = sys.argv[1]
    novel_link = get_novel_link(search_key)
    res = requests.get(novel_link)
    res.encoding = res.apparent_encoding
    content = res.text
    re_novel_content = r'<dd><a href="(.*?).html" title="'
    novel_content = re.findall(re_novel_content,content,re.DOTALL)
    re_novel_title = r' title="(.*?)">'
    novel_title = re.findall(re_novel_title,content,re.DOTALL)

    #如修罗武神最新一章在倒数第13章，需要指定第二个参数
    try:
        newest = int("-" + str(sys.argv[2]))
    except Exception as er:
        newest = -1
    latest_chapter_title = novel_title[newest]
        
    # 查询更新章节数
    def update_chapters(saved_content,last_newest):
        this_newest = last_newest + 1
        if saved_content == "":
            return this_newest
        while saved_content != novel_title[last_newest]:
            last_newest = last_newest -1
        else:
            return last_newest + 1
            
    #查询是否最新一章
    try:
        with open("./" + search_key + ".txt", 'r', encoding='utf-8') as fh:
            content = fh.read()
    except Exception as er:
        content = ""        
        
    if content == latest_chapter_title:
        sys.exit()
    else:
        for i in range(update_chapters(content, newest - 1), newest + 1):
            get_novel_content(novel_content[i],novel_title[i],novel_link)
            
        with open("./" + search_key + ".txt", 'w', encoding='utf-8') as fs:
            fs.write(latest_chapter_title)
            
if __name__ == '__main__':
    get_novel_status()
