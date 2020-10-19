## 京东薅羊毛 互助码生成

### 使用
* 直接下载程序，并按照```jd_cookies```填写信息
* 程序使用python3编写，命令```python3 jd_sheet.py```

### 说明
* 每人得到的助力机会是均等的
* 圣人为无私奉献者，即只助力，无被助力（暂时只支持农场）

### 函数text_process(jd_game,available_value,sacrifice_value)的说明
* 种豆得豆：每人每天只有三次助力机会
```
text_process(jd_bean,4,0)
```
* 京东农场：每人每天有四次助力机会，一个圣人
```
text_process(jd_farm,5,1)
```
* 京东萌宠：每人每天有五次获取助力机会，助力别人的机会未知，不过程序限制了五个成功的互助码
```
text_process(jd_pet,None,0)
```
