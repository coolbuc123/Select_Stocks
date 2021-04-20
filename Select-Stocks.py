#!/usr/bin/env python
# coding: utf-8

# In[1]:


import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np
import my_function as my


# ### 실적기준월

# In[2]:


yyyymm = '202006'


# ### Data - fanace : pkl활용

# ### 현 보유 포트(+순위) 정보 

# In[3]:


# finance_file = 'data/재무_'+yyyymm+'.xlsx'
# my.get_finance(finance_file).to_pickle('pkl/finance.pkl')


# In[4]:


port_except = ['TIGER 소프트웨어','KODEX 바이오','현대차','태림포장']
port = my.get_port( port_except, yyyymm, 500)
port = port.sort_values('투자기준월').reset_index(drop=True)
print(port.shape)
port # 현재가 표기 안하기로


# ### [매도-기간기준] 투자기준월 가장 오래된 2종목

# In[5]:


closing = port.sort_values('투자기준월') [['투자기준월','종목코드','종목명','최종순위','평가손익','수익률','현재가','체결잔고','매수금액']]
closing.head(2)


# ### [매도-순위기준] : 24위 순위 밖

# In[6]:


sell = port[port['최종순위'] > 24] [['투자기준월','종목코드','종목명','최종순위','평가손익','수익률','현재가','체결잔고','평가금액']]
sell.sort_values('최종순위')


# In[7]:


base_amt = 650000


# ### [추가매수] 증분만 더 사기

# In[8]:


stock_nm = '조선선재'
print("추가매수 : ", stock_nm)
round((base_amt - port[port['종목명'] == stock_nm]['매수금액']) / port[port['종목명'] == stock_nm]['현재가'],1)


# ### [매수] : 미보유 종목 중 topn(24종목 기준, 매도대상 고려)

# In[9]:


import math
sell_cnt = 24 - (len(port) - sell.shape[0])
sell_cnt = 5
buy = my.get_stock_topn(sell_cnt, yyyymm, 500, 1)
buy['매수가능'] = round(base_amt / buy['현재가'],1)
buy['매수금액'] = buy['현재가'] * buy['매수가능'].apply(lambda x : math.floor(x))
buy[['종목코드','종목명','최종순위','현재가','매수가능','매수금액']].reset_index(drop=True)


# ### (참고) 전체 종목 기준으로 top24 살펴보기

# In[10]:


my.get_stock_topn(24, yyyymm, 500, 0)

