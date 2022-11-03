#!/usr/bin/env python
# coding: utf-8


# get_ipython().system('pip install --upgrade pip')
# get_ipython().system('pip install pandas')
# get_ipython().system('pip install requests')
# get_ipython().system('pip install xlrd')
# get_ipython().system('pip install datacompy')


import json
import requests
from pathlib import Path
from datetime import date
from datetime import timedelta
import pandas as pd
import datacompy

# Import Webex Teams room_id and token
from myTokens.MyTokens import *


folder = './xls'
Path(folder).mkdir(parents=True, exist_ok=True)

today = date.today()
yesterday = today - timedelta(days=1)

path_today = folder + '/' + str(today) + '.xls'
path_yesterday = folder + '/' + str(yesterday) + '.xls'

xls_url = 'https://csi-web-dev.oss-cn-shanghai-finance-1-pub.aliyuncs.com/static/html/csindex/public/uploads/file/autofile/cons/930994cons.xls'
r = requests.get(xls_url, allow_redirects=True)
open(path_today, 'wb').write(r.content)


df_today = pd.read_excel(path_today, engine='xlrd')
df_today = df_today.iloc[:, 4:6]
df_today.columns = ['Code', 'Name']

# 首次运行没有旧文件，复制xls充当旧文件，避免运行失败
if not Path(path_yesterday).exists():
	xls_today = Path(path_today)
	xls_yesterday = Path(path_yesterday)
	xls_yesterday.write_bytes(xls_today.read_bytes())

df_yesterday = pd.read_excel(path_yesterday, engine='xlrd')
df_yesterday = df_yesterday.iloc[:, 4:6]
df_yesterday.columns = ['Code', 'Name']


compare = datacompy.Compare(df_yesterday, df_today, join_columns='Code')
# 需要卖出的基金
short_list = compare.df1_unq_rows
# 需要买入的基金
long_list = compare.df2_unq_rows

if short_list.empty and long_list.empty:
	message = '工银股混持仓无变化'
else:
	short_list = short_list.to_string(index=False)
	long_list =	long_list.to_string(index=False)
	message = '卖出: \n' + short_list + '\n\n买入：\n' + long_list

# Send messages to Webex Teams
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

# Delete old xls file
Path(path_yesterday).unlink()