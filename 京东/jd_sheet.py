#_Author:Iron Joe
#date:2020/10/8

import re
import copy
#读取有效数据
def useful_text(text_list,start_text,end_text):
    start_flag = False
    jd_game_use=[]
    for i in range(len(text_list)):
        if text_list[i] == start_text:
            start_flag=True
        elif text_list[i] == end_text:
            start_flag= False
            pass
        if start_flag:
            if 20<len(text_list[i])<45:
                jd_game_use.append(text_list[i])
    return jd_game_use

# 有效数据处理
# sacrifice_value 致力于助力前面的号,总人数小于8时建议为0，人数为11时建议小于3
def text_process(jd_game,available_value,sacrifice_value):
    jd_game_all_list = []
    jd_game_copy = copy.deepcopy(jd_game)
    len_list = len(jd_game)
    if not sacrifice_value:
        sacrifice_value = 0
    available_sacrifice_value = len_list-(4*len_list)//5
    if sacrifice_value > available_sacrifice_value:
        sacrifice_value = available_sacrifice_value
        jd_game_copy = jd_game_copy[0:len_list-sacrifice_value]
    for i in range(len(jd_game_copy)):
        if i > 0:
            item = jd_game_copy.pop(0)
            jd_game_copy.append(item)
        jd_game2 = jd_game_copy[0:available_value]
        game_all = '@'.join(jd_game2)
        jd_game_all_list.append(game_all)
    for i in range(sacrifice_value):
        k = len_list - sacrifice_value + i
        jd_game3 = jd_game[4*i:4*(i+1)]
        jd_game3.insert(0,jd_game[k])
        game_all = '@'.join(jd_game3)
        jd_game_all_list.append(game_all)
        
    return jd_game_all_list

text_list = []
with open('jd_cookies','r',encoding='utf-8') as fh,open('jd_cookies','r',encoding='utf-8') as fh2:
    for i in fh:
        text_list.append(i.strip())
    text_all = fh2.read()

# cookies
re_pt_key=r'(?m)^pt_key=?.*.;$'
pt_key=re.findall(re_pt_key,text_all)#6
cookies_num = len(pt_key)
cookies='&'.join(pt_key) + '&'

# 种豆得豆互助码：
jd_bean = useful_text(text_list,'###bean','###bean_end')
bean_all_list = text_process(jd_bean,4,0)
bean_all_text = '&'.join(bean_all_list)+'&'

# 东东农场
jd_farm=useful_text(text_list,'###farm','###farm_end')
farm_all_list = text_process(jd_farm,5,1)
farm_all_text = '&'.join(farm_all_list)+'&'

# 京东萌宠互助码：
jd_pet = useful_text(text_list,'###pet','###pet_end')
pet_all_list = text_process(jd_pet,None,0)
for i in range(len(pet_all_list)):
    pet_all_list[i] = pet_all_list[i] + '='
    pet_all_list[i] = pet_all_list[i].replace('=@','==@')
    re_pro = r'[.*@]{1}(.*)'#匹配删除第一个助力码
    pet_all_list[i] = re.findall(re_pro,pet_all_list[i])[0]
pet_all_text = '&'.join(pet_all_list)+'&'

# # 京小超互助码：
# jd_market = useful_text(text_list,'###market','###market_end')
# market_all_list = text_process(jd_market,None,0)
# market_all_text = '&'.join(market_all_list)+'&'

with open('jd_sheet', 'w', encoding='utf-8') as fh:
    content = '''
京东cookies
JD_COOKIE:
%s

种豆得豆：
PLANT_BEAN_SHARECODES:
%s

京东农场：
FruitShareCodes:
%s

京东萌宠：   
PETSHARECODES:
%s

京小超：
SUPERMARKET_SHARECODES	:
%s
    ''' % (cookies, bean_all_text, farm_all_text, pet_all_text)
    # (cookies, bean_all_text, farm_all_text, pet_all_text,market_all_text)
    fh.write(content)
# 京小超：
# SUPERMARKET_SHARECODES	:
# %s
