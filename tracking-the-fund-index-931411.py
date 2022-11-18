#!/usr/bin/env python
# coding: utf-8


import json
import requests
from pathlib import Path
from datetime import date
from datetime import timedelta
import pandas as pd
import datacompy

# Import Webex Teams room_id and token from the local file "./myTokens/MyTokens.py"
from myTokens.MyTokens import *

folder = './xls'
Path(folder).mkdir(parents=True, exist_ok=True)

today = date.today()
yesterday = today - timedelta(days=1)

path_today = folder + '/' + str(today) + '.xls'
path_yesterday = folder + '/' + str(yesterday) + '.xls'

xls_url = 'https://csi-web-dev.oss-cn-shanghai-finance-1-pub.aliyuncs.com/static/html/csindex/public/uploads/file/autofile/closeweight/931411closeweight.xls'
r = requests.get(xls_url, allow_redirects=True)
open(path_today, 'wb').write(r.content)

df_today = pd.read_excel(path_today, engine='xlrd')
df_today = df_today.iloc[:, 6:]
df_today.columns = ['code_shh', 'name_shh', 'code_szh', 'name_szh', 'col_1', 'col_2', 'weight_%']
df_today.pop('col_1')
df_today.pop('col_2')
df_today['code'] = pd.concat([df_today.iloc[:, 0], df_today.iloc[:, 2]]).dropna()
df_today['name'] = pd.concat([df_today.iloc[:, 1], df_today.iloc[:, 3]]).dropna()
df_today = df_today.filter(['code', 'name', 'weight_%'], axis=1)

# 首次运行没有旧文件，复制xls充当旧文件，避免运行失败
if not Path(path_yesterday).exists():
	xls_today = Path(path_today)
	xls_yesterday = Path(path_yesterday)
	xls_yesterday.write_bytes(xls_today.read_bytes())

df_yesterday = pd.read_excel(path_yesterday, engine='xlrd')
df_yesterday = df_yesterday.iloc[:, 6:]
df_yesterday.columns = ['code_shh', 'name_shh', 'code_szh', 'name_szh', 'col_1', 'col_2', 'weight_%']
df_yesterday.pop('col_1')
df_yesterday.pop('col_2')
df_yesterday['code'] = pd.concat([df_yesterday.iloc[:, 0], df_yesterday.iloc[:, 2]]).dropna()
df_yesterday['name'] = pd.concat([df_yesterday.iloc[:, 1], df_yesterday.iloc[:, 3]]).dropna()
df_yesterday = df_yesterday.filter(['code', 'name', 'weight_%'], axis=1)

compare = datacompy.Compare(df_yesterday, df_today, join_columns='code')
# 需要卖出的可转债
short_list = compare.df1_unq_rows
# 需要买入的可转债
long_list = compare.df2_unq_rows

if short_list.empty and long_list.empty:
	message = '可转债指数-931411-持仓无变化'
else:
	# 在df_yesterday中列出没有被卖出的可转债，及其权重
	df_yesterday = df_yesterday.drop(short_list.index, errors='ignore', axis=0)
	df_yesterday.reset_index(drop=True, inplace=True)
	
	# 在df_today中列出没有被卖出的可转债，及其权重
	df_today = df_today.drop(long_list.index, errors='ignore', axis=0)
	df_today.reset_index(drop=True, inplace=True)
    
	# 计算没有被卖出的可转债的持仓权重变化
	df_diff = df_yesterday.filter(['code', 'name'], axis=1)
	df_diff['weight_%'] = df_today['weight_%'] - df_yesterday['weight_%']
	df_diff = df_diff.loc[~(df_diff['weight_%'] == 0)]
	
	# 生成调仓通知
	short_list = short_list.to_string(index=False)
	long_list =	long_list.to_string(index=False)
	df_diff = df_diff.to_string(index=False)
	message = '可转债卖出: \n' + short_list + '\n\n买入: \n' + long_list + '\n\n权重变化: \n' + df_diff

def sendMessage(token, room_id, message):
	header = {"Authorization": "Bearer %s" % token,
			  "Content-Type": "application/json"}
	data = {"roomId": room_id,
			"text": message}
	res = requests.post("https://api.ciscospark.com/v1/messages/",
						 headers=header, data=json.dumps(data), verify=True)
	if res.status_code == 200:
		print("消息已经发送至 Webex Teams")
	else:
		print("failed with statusCode: %d" % res.status_code)
		if res.status_code == 404:
			print("please check the bot is in the room you're attempting to post to...")
		elif res.status_code == 400:
			print("please check the identifier of the room you're attempting to post to...")
		elif res.status_code == 401:
			print("please check if the access token is correct...")

sendMessage(token=webex_token, room_id=webex_room_id, message=message)

# 删除前一日的xls表格
Path(path_yesterday).unlink()