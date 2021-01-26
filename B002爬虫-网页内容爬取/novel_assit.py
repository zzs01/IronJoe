# -*- coding:utf8 -*-

import novel_update as novel
import re
import sys
from datetime import datetime

def start_check_novel(novel_file):
    with open(novel_file, 'r', encoding='utf-8') as fh:
        novel_all = []
        for i in fh:
            novel_all.append(i.strip().split(" "))
            
    novel_all_not_blank = [i for i in novel_all if i[0] != ''] #清除空元素
    for novel_name in novel_all_not_blank:
        try:
            newest_chapter_location =  int("-" + str(novel_name[1]))
        except Exception as er:
            newest_chapter_location = -1

        novel.get_novel_status(novel_name[0], newest_chapter_location)
                
if __name__ == '__main__':
    start = datetime.now()
    try:
        novel_file =  sys.argv[1]
    except Exception as er:
        novel_file = "novel_list"
    
    start_check_novel(novel_file)
    
    try:
        novel.status == "success"
        status = "finish"
    except Exception as er:
        status = "unfinished"
    if status == "finish":
        end = datetime.now()
        novel.dd_send("总共用时", str(end-start), "程序运行时间：" + str(end-start))
