# coding=gbk
import random
import pandas as pd

with open('百家姓.txt','r',encoding='utf-8') as fp:
    txt=fp.read().split('|')

info=[]
for i in range(100):
    name=random.choice(txt)+'**'
    num=random.randint(0,9)
    str_end=''.join(random.sample('0123456789', 2))
    phone=f'13{num}******{str_end}'
    info.append({'name':name,'phone':phone})

df=pd.DataFrame(info)
df.to_csv('虚拟数据.csv', index=False, encoding='gbk')
print(df)