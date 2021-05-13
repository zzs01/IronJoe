#!/usr/bin/env python
# -*- coding:utf-8 -*-
#@Time      :2021/5/5 17:40
#@Author    :IRONJOE
#@File      :001-彩票数据爬取.py

import requests
import re
import sqlite3
import matplotlib.pyplot as plt

def lottery_web(history_lottery_url):
    """
    历史数据爬取
    :param history_lottery_url:
    :return: 期号，开奖日期，开奖结果
    """
    # 获取之前的彩票的网址
    # lottery_url = 'https://www.lottery.gov.cn/kj/kjlb.html?qxc'
    # 全部开奖结果的首页
    # history_lottery_url = 'https://webapi.sporttery.cn/gateway/lottery/getHistoryPageListV1.qry?gameNo=04&provinceId=0&pageSize=30&isVerify=1&pageNo=1'
    # 最近100期结果
    # history_lottery_url = 'https://webapi.sporttery.cn/gateway/lottery/getHistoryPageListV1.qry?gameNo=04&provinceId=0&pageSize=100&isVerify=1&pageNo=1&termLimits=100'

    # 请求头
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36'}

    # 响应值处理
    url_response = requests.get(url=history_lottery_url,headers=headers)
    url_response.encoding = url_response.apparent_encoding
    response_html = url_response.text

    # 期号
    rule_lotteryDrawNum = r'"lotteryDrawNum":"(\w*)"'
    lotteryDrawNum_list = re.findall(rule_lotteryDrawNum,response_html,re.DOTALL)

    # 开奖日期
    rule_lotteryDrawTime = r'"lotteryDrawTime":"(.{10})"'
    lotteryDrawTime_list = re.findall(rule_lotteryDrawTime,response_html,re.DOTALL)
    lotteryDrawTime_list = [i.replace('-', '') for i in lotteryDrawTime_list]

    # 开奖结果
    rule_lotteryDrawResult = r'"lotteryDrawResult":"(.{13})"?'
    lotteryDrawResult_list = re.findall(rule_lotteryDrawResult,response_html,re.DOTALL)
    lotteryDrawResult_list = [i.replace(' ','') for i in lotteryDrawResult_list] #清除空格
    # print(lotteryDrawResult_list)
    # lotteryDrawResult_list = [i.replace(' ','')[:-3] for i in lotteryDrawResult_list] #清除空格，余下四合彩

    return lotteryDrawNum_list,lotteryDrawTime_list,lotteryDrawResult_list

def lottery_database_create(history_lottery_url,table_name):
    """
    数据库管理,新表创建
    :return:
    """
    # 连接数据库
    conn = sqlite3.connect('lottery.db')
    form = conn.cursor()
    print('打开数据库成功...,创建新表')

    # 数据写入命令
    lottery_data_str = "INSERT INTO %s (lotteryDrawNum,lotteryDrawTime,lotteryDrawResult) VALUES ({}, {}, '{}')" % table_name

    # 判断数据是否存在
    form.execute("SELECT tbl_name FROM sqlite_master WHERE type='table'")
    table_name_list = [tb_name[0] for tb_name in form.fetchall()]
    data_list = lottery_web(history_lottery_url)
    # 判断数据库表格是否存在
    if table_name not in table_name_list:
        print("数据库不存在乐透表，开始创建乐透表 %s ...."%table_name)
        form.execute('''CREATE TABLE {}
                        (lotteryDrawNum   INT PRIMARY KEY   NOT NULL,
                        lotteryDrawTime   INT               NOT NULL,
                        lotteryDrawResult TEXT              NOT NULL);'''.format(table_name))
        print("乐透表 %s 创建完毕...."%table_name)
        for i in range(len(data_list[0])):
            # form.execute(lottery_data_str.format(int(data_list[0][i]),int(data_list[1][i]),int(data_list[2][i])))
            form.execute(lottery_data_str.format(data_list[0][i], data_list[1][i], data_list[2][i]))
        print("数据写入完毕....")
    else:
        print("数据库已存在乐透表 %s ，开始写入最新数据..." % table_name)
        for i in range(len(data_list[0])):
            cursor = form.execute("SELECT lotteryDrawNum from {} where lotteryDrawNum={}".format(table_name,data_list[0][-i]))
            lotteryDrawNum = [row[0] for row in cursor]
            # 判断数据是否存在于数据库表中
            if not lotteryDrawNum:
                form.execute(lottery_data_str.format(data_list[0][-i], data_list[1][-i], data_list[2][-i]))
            else:
                break
    conn.commit()
    conn.close()

def data_visualization(table_name):
    """
    数据可视化
    :return:
    """
    # 准备数据
    conn = sqlite3.connect('lottery.db')
    form = conn.cursor()
    print('打开数据库成功...')
    cursor = form.execute("SELECT lotteryDrawNum, lotteryDrawResult from {}".format(table_name))
    lo_data_dic = {}
    for row in cursor:
        lo_data_dic.update({row[0]:row[1]})
        # print('lotteryDrawNum = ',row[0])
        # print('lotteryDrawResult = ', row[1],'\n')
    # 按照key排序
    lo_data_dic = sorted(lo_data_dic.items(), key=lambda d: d[0])
    lotteryDrawNum = sorted([row[0] for row in lo_data_dic])
    # print(lotteryDrawNum)
    lotteryDrawResult = [int(row[1][0]) for row in lo_data_dic]
    # print(lotteryDrawResult)
    conn.close()

    # 创建画布
    plt.figure()

    # 绘画折线图
    plt.plot(lotteryDrawNum,  # x轴数据
             lotteryDrawResult,  # y轴数据
             linestyle='-',  # 折线类型
             linewidth=2,  # 折线宽度
             color='steelblue',  # 折线颜色
             marker='o',  # 折线图中添加圆点
             markersize=6,  # 点的大小
             markeredgecolor='black',  # 点的边框色
             markerfacecolor='brown')  # 点的填充色
    # 获取图的坐标信息
    # ax = plt.gca()
    # 设置日期的显示格式
    # date_format = mpl.dates.DateFormatter("%m-%d")
    # ax.xaxis.set_major_formatter(date_format)
    # # 设置x轴每个刻度的间隔天数
    # xlocator = mpl.ticker.MultipleLocator(7)
    # ax.xaxis.set_major_locator(xlocator)
    # 添加x，y轴标签以及图表名称
    plt.ylabel('开奖结果', fontsize=18)  # 横坐标轴的标题
    plt.xlabel('期号', fontsize=18)  # 纵坐标轴的标题
    plt.title('七星彩图表', fontsize=18)

    plt.grid() # 显示网格
    # 添加图形标题

    plt.show()

def main():
    # try:
    #     lottery_database_create()
    # except:
    #     print('网络有误，未更新数据...')
    #     pass
    # #
    history_lottery_url = "https://webapi.sporttery.cn/gateway/lottery/getHistoryPageListV1.qry?gameNo=04&provinceId=0&pageSize=30&isVerify=1&pageNo={}&startTerm=20001&endTerm=20133"
    for i in range(1,6):
        print('第 %d 页数据写入。。。'%i)
        lottery_database_create(history_lottery_url.format(i),'LOTTERY2020')
        print(history_lottery_url.format(i))

    data_visualization('LOTTERY2020')

if __name__ == '__main__':
    main()