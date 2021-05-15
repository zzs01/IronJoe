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

def data_read(table_name):
    # 准备数据
    conn = sqlite3.connect('lottery.db')
    form = conn.cursor()
    print('打开数据库成功...')
    cursor = form.execute("SELECT lotteryDrawNum, lotteryDrawResult from {}".format(table_name))
    lo_data_dic = {}
    for row in cursor:
        lo_data_dic.update({row[0]:row[1]})

    # 按照key排序
    lo_data_dic = sorted(lo_data_dic.items(), key=lambda d: d[0])
    lotteryDrawNum = sorted([row[0] for row in lo_data_dic])
    lotteryDrawResult = [int(row[1][2]) for row in lo_data_dic]
    conn.close()
    return lotteryDrawNum, lotteryDrawResult

def data_visualization_plt(table_name1,table_name2):
    """
    数据可视化-折线图
    :return:
    """
    # 创建画布
    plt.figure(figsize=(15,8),dpi=100)

    y_list1 = data_read(table_name1)[1]
    y_list2 = data_read(table_name2)[1]
    x_list1 = [i+1 for i in range(len(y_list1))]
    x_list2 = [i+1 for i in range(len(y_list2))]
    # 绘画折线图
    plt.plot(x_list1,  # x轴数据
             y_list1,  # y轴数据
             linestyle='-',  # 折线类型
             linewidth=2,  # 折线宽度
             color='steelblue',  # 折线颜色
             marker='o',  # 折线图中添加圆点
             markersize=3,  # 点的大小
             # markeredgecolor='black',  # 点的边框色
             # markerfacecolor='brown', # 点的填充色
             label= table_name1)

    plt.plot(x_list2,  # x轴数据
             y_list2,  # y轴数据
             linestyle='-',  # 折线类型
             linewidth=2,  # 折线宽度
             color='red',  # 折线颜色
             marker='o',  # 折线图中添加圆点
             markersize=3,  # 点的大小
             # markeredgecolor='black',  # 点的边框色
             # markerfacecolor='brown', # 点的填充色
             label= table_name2)

    # 添加图例
    plt.legend()

    # 添加x，y轴刻度
    yticks = range(0,10,1)
    plt.yticks(yticks)

    # 设置x，y轴名称以及标题
    plt.ylabel('开奖结果', fontsize=18)
    plt.xlabel('期号', fontsize=18)
    plt.title('七星彩图表', fontsize=18)
    # 显示网格
    plt.grid(True, linestyle = '--', alpha = 0.5)
    plt.show()

def data_visualization_bar(table_name1,table_name2):
    """
    数据可视化-柱状图
    :return:
    """
    # 创建画布
    plt.figure(figsize=(15, 8), dpi=100)

    x_list = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    y_list1 = data_read(table_name1)[1]
    y_sum1 = len(y_list1)
    print('%s的平均数为：%s'%(table_name1,sum(y_list1)/y_sum1))
    # print(y_sum1)
    y_list1 = [(y_list1.count(i))/y_sum1 for i in x_list]

    y_list2 = data_read(table_name2)[1]
    y_sum2 = len(y_list2)
    print('%s的平均数为：%s' % (table_name2, sum(y_list2) / y_sum2))
    # print(y_sum2)
    y_list2 = [(y_list2.count(i))/y_sum2 for i in x_list]

    total_width, n = 0.6, 2
    width = total_width / n

    x_list1 = [x_list[i] - width/2 for i in x_list]
    plt.bar(x_list1, y_list1, width=width, color="green", label = table_name1)
    for a, b in zip(x_list, y_list1):  # 柱子上的数字显示
        plt.text(a-width/2, b, '%.3f' % b, ha='center', va='bottom', fontsize=10)
    # for i in range(len(x_list)):
    #     x_list[i] = x_list[i] + width/2
    x_list2 = [x_list[i] + width / 2 for i in x_list]
    plt.bar(x_list2, y_list2, width=width, color="blue", label =table_name2)
    for a, b in zip(x_list, y_list2):  # 柱子上的数字显示
        plt.text(a+width/2, b, '%.3f' % b, ha='center', va='bottom', fontsize=10)

    # 添加图例
    plt.legend(fontsize = 10)

    # 添加x，y轴刻度
    xticks = range(0,10,1)
    plt.xticks(xticks)

    # 设置x，y轴名称以及标题
    plt.ylabel('开奖结果', fontsize=18)
    plt.xlabel('期号', fontsize=18)
    plt.title('七星彩图表', fontsize=18)
    # 显示网格
    plt.grid(True, linestyle = '--', alpha = 0.5)
    plt.show()

def main():
    # try:
    #     lottery_database_create()
    # except:
    #     print('网络有误，未更新数据...')
    #     pass
    # #
    # history_lottery_url = "https://webapi.sporttery.cn/gateway/lottery/getHistoryPageListV1.qry?gameNo=04&provinceId=0&pageSize=30&isVerify=1&pageNo={}&startTerm=20001&endTerm=20133"
    # for i in range(1,7):
    #     print('第 %d 页数据写入。。。'%i)
    #     lottery_database_create(history_lottery_url.format(i),'LOTTERY2020')
    #     print(history_lottery_url.format(i))
    #
    table_name1 = 'LOTTERY2019'
    table_name2 = 'LOTTERY2020'
    data_visualization_bar(table_name1,table_name2)
    # data_visualization_plt(table_name1,table_name2)

if __name__ == '__main__':
    main()