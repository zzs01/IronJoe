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
def text_process(jd_game,available_value):
    jd_game_all_list = []
    jd_game_copy = copy.deepcopy(jd_game)
    len_list = len(jd_game)
    for i in range(len_list):
        if i > 0:
            item = jd_game_copy.pop(0)
            jd_game_copy.append(item)
        jd_game2 = jd_game_copy[0:available_value]
        game_all = '@'.join(jd_game2)
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
bean_all_list = text_process(jd_bean,None)
bean_all_text = '&'.join(bean_all_list)+'&'

# 东东农场
jd_farm=useful_text(text_list,'###farm','###farm_end')
farm_all_list = text_process(jd_farm,6)
for i in farm_all_list:
    print(i)
farm_all_text = '&'.join(farm_all_list)+'&'

# 京东萌宠互助码：
jd_pet = useful_text(text_list,'###pet','###pet_end')
pet_all_list = text_process(jd_pet,None)
for i in range(len(pet_all_list)):
    pet_all_list[i] = pet_all_list[i] + '='
    pet_all_list[i] = pet_all_list[i].replace('=@','==@')
    re_pro = r'[.*@]{1}(.*)'#匹配删除第一个助力码
    pet_all_list[i] = re.findall(re_pro,pet_all_list[i])[0]
pet_all_text = '&'.join(pet_all_list)+'&'

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
    ''' % (cookies, bean_all_text, farm_all_text, pet_all_text)
    fh.write(content)
