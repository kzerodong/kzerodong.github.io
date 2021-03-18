#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

import pandas as pd

import time
import datetime

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def extractDataframeFromHTML(page_src, table_id):
	table = page_src.find(id = table_id)
	table_rows = table.find_all('tr')

	# HTML table to dataframe
	l = []
	for tr in table_rows:
		td = tr.find_all('td')
		row = [tr.text for tr in td]
		l.append(row)

	df = pd.DataFrame(l)

	return df

class isProtoWinLose():
	def __call__(self, driver):
		soup = BeautifulSoup(driver.page_source, "lxml")
		df = extractDataframeFromHTML(soup, "grd_closedGmList")
		df.dropna(axis = 0, how ='all', thresh=None, subset=None, inplace=True)
		temp_list = df.iloc[0,1].split()
		if (temp_list[0] == u'프로토') & (temp_list[1] == u'승부식'):
			temp = temp_list[-1]
			temp = temp.replace(u'회차', '')
			return int(temp)
		else:
	 		return False

def getRecentGameNumberFromURL(url):

	options = webdriver.ChromeOptions()
	options.add_argument('headless')
	options.add_argument('window-size=1920x1080')
	options.add_argument("disable-gpu")

	# FIXME - install overhead 
	driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
	driver.get(url) # 마감게임 보기
	
	el = driver.find_element_by_xpath("//button[@title='게임유형 선택']") # open dropdown box 1 - 게임유형 전체
	el.click()

	ppt = driver.find_element_by_xpath("//div[text() = '프로토']") # select proto category
	ppt.click()

	el2 = driver.find_element_by_xpath("//button[@title='게임종류 선택']") # open dropdown box - 프로토 전체
	el2.click()

	pptwl = driver.find_element_by_xpath("//div[text() = '-프로토 승부식']") # select proto category
	pptwl.click()

	btn_sch = driver.find_element_by_xpath("//button[@id='btn_sch']")
	btn_sch.click()

	g_num = WebDriverWait(driver, 10).until(isProtoWinLose()) # last game number
	g_num += 1 # current game number

	return g_num

def updateURL():
	t_url = 'http://www.betman.co.kr/main/mainPage/gamebuy/closedGameList.do'
	num = getRecentGameNumberFromURL(t_url)

	year = datetime.datetime.now().year
	year = year%100

	g_num = year*10000 + num
	
	new_url = 'http://www.betman.co.kr/main/mainPage/gamebuy/closedGameSlip.do?frameType=typeA&gmId=G101&gmTs='
	new_url += str(g_num)

	return new_url, g_num





def writeInputToExcel(ec_writer, df, sheetname):
	new_sheetname = sheetname
	workbook = ec_writer.book
	worksheet = workbook.add_worksheet(new_sheetname)
	ec_writer.sheets[new_sheetname] = worksheet

	# Time format
	time_cell_format = workbook.add_format({'align':'center',
																					'num_format':'hh:mm'})
	time_col = df.columns.get_loc(u'시간')
	worksheet.set_column(time_col, time_col, None, time_cell_format)

	df.to_excel(ec_writer, sheet_name = new_sheetname, index=False)

def extractTableColumnNameFromHTML(page_src, table_id):
	table = page_src.find(id = table_id)
	table_rows = table.find_all('tr')

	tr = table_rows[0]
	td = tr.find_all('th')
	row = [];
	for tr in td:
		tr = tr.text
		tr = tr.replace(u'오름차순', '')
		tr = tr.replace('\n', '')
		row.append(tr)

	return row 

def preprocessTime(x):
	try: 
		# conversion to serial time format (excel time format) 
		ret = x.hour
		ret *= 60
		ret += x.minute
		ret *= 60
		ret += x.second
		ret = float(ret / 86400)
	except:
		ret = ''

	return ret

def getGameDataframeFromURL(url):

	options = webdriver.ChromeOptions()
	options.add_argument('headless')
	options.add_argument('window-size=1920x1080')
	options.add_argument("disable-gpu")

	# TODO - reduce install overhead 
	driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
	driver.get(url)
	
	# 'btn_gmBuySlipDt0' is the first dynamic tag ID 
	try:
		element = WebDriverWait(driver, 10).until(
						EC.presence_of_element_located((By.ID, "btn_gmBuySlipDt0")) 
						)
	except:
		print('error while crawling')

	# move to 발매마감 table 
	if (driver.find_element_by_id("buyPsblStTab_3").is_enabled()):
		print("ENALBLED : buyPsblStTab_3")
		driver.execute_script('document.getElementById("buyPsblStTab_3").click();') 
		soup = BeautifulSoup(driver.page_source, "lxml")
		crawled_df_closed = extractDataframeFromHTML(soup, "tbl_gmBuySlipList")
	else:
		print("DISABLED : buyPsblStTab_3")
		crawled_df_closed = pd.DataFrame()

	# move to 발매중 table
	if (driver.find_element_by_id("buyPsblStTab_2").is_enabled()):
		print("ENALBLED : buyPsblStTab_2")
		driver.execute_script('document.getElementById("buyPsblStTab_2").click();') 
		soup = BeautifulSoup(driver.page_source, "lxml")
		crawled_df_open = extractDataframeFromHTML(soup, "tbl_gmBuySlipList")
	else:
		print("DISABLED : buyPsblStTab_2")
		crawled_df_open = pd.DataFrame()


	# move to 발매전 table
	if (driver.find_element_by_id("buyPsblStTab_1").is_enabled()):
		print("ENALBLED : buyPsblStTab_1")
		driver.execute_script('document.getElementById("buyPsblStTab_1").click();') 
		soup = BeautifulSoup(driver.page_source, "lxml")
		crawled_df_pending = extractDataframeFromHTML(soup, "tbl_gmBuySlipList")
	else:
		print("DISABLED : buyPsblStTab_1")
		crawled_df_pending = pd.DataFrame()

	table_columns = extractTableColumnNameFromHTML(soup, "thd_gmBuySlipList")

	driver.close()

	# concat 발매전/중/마감 tables
	temp_df = pd.concat([crawled_df_closed, crawled_df_open], ignore_index=True)
	crawled_df = pd.concat([temp_df, crawled_df_pending], ignore_index=True)

	# reset the columns
	crawled_df.columns = table_columns

	# TODO - isolate this code by function
	# drop empty rows
	crawled_df.dropna(axis = 0, how ='all', thresh=None, subset=None, inplace=True)

	# drop '긴급 공지닫기' string
	crawled_df[u'번호'] = crawled_df[u'번호'].str.replace(u'긴급 공지닫기', '')
	crawled_df[u'번호'] = pd.to_numeric(crawled_df[u'번호'])

	# sort dataframe and reset index
	crawled_df = crawled_df.sort_values([u'번호'])
	crawled_df = crawled_df.reset_index(drop=True)

	# FIXME test dataframe
	test_df = crawled_df.copy()

	# drop cols
	test_df = test_df.drop(u'마감일시', axis=1)
	test_df = test_df.drop(u'장소', axis=1)
	test_df = test_df.drop(u'정보', axis=1)

	# rename column names
	test_df = test_df.rename(columns={u'종목/대회':u'대회'})

	# create '종목' column
	test_df[u'종목'] = test_df[u'대회'].str[0:2]

	# refine row data in '대회' column
	test_df[u'대회'] = test_df[u'대회'].str.replace(u'축구', '')
	test_df[u'대회'] = test_df[u'대회'].str.replace(u'농구', '')
	test_df[u'대회'] = test_df[u'대회'].str.replace(u'배구', '')

	# create new cols
	test_df[u'분류'] = ''
	test_df[u'날짜'] = ''
	test_df[u'시간'] = ''
	test_df[u'홈'] = ''
	test_df[u'원정'] = ''
	test_df[u'승'] = ''
	test_df[u'무'] = ''
	test_df[u'패'] = ''

	# loop over rows
	for index, row in test_df.iterrows():
		game_type = row[u'게임유형']
		info = row[u'홈팀 vs 원정팀']
		info = info.replace('H ', ' H ')
		info = info.replace('U/O', ' U/O')
		info = info.replace(':', ' : ')

		# get game type and number
		if (game_type == u'일반'):
			test_df.at[index, u'분류'] = row[u'대회']

		elif (game_type == u'핸디캡'):
			temp = info
			temp = temp.replace(u'사전조건 변경', '')
			temp = temp.replace(':', '')

			temp_list = temp.split()
			h_idx = temp_list.index('H')
			handi = temp_list[h_idx+1]
			handi = handi.replace('+', '+ ')
			handi = handi.replace('-', '- ')
			handi = 'H ' + handi
			test_df.at[index, u'분류'] = handi

		elif (game_type == u'언더오버'):
			temp = info
			temp = temp.replace(u'사전조건 변경', '')
			temp = temp.replace(':', '')

			temp_list = temp.split()
			uo_index = temp_list.index('U/O')
			uo = temp_list[uo_index+1]
			uo = 'U/O ' + uo

			test_df.at[index, u'분류'] = uo

		# get date
		temp = row[u'경기일시']
		if (temp == u'미정'):
			test_df.at[index, u'날짜'] = temp
		else:
			date_dt = datetime.datetime.strptime(temp.split(' ')[0], "%m.%d")
			date_dt = date_dt.replace(year=datetime.datetime.now().year) # FIXME - year from now
			date = date_dt.strftime("%Y/%m/%d")
			test_df.at[index, u'날짜'] = pd.Timestamp(date)

		# get time
		temp = row[u'경기일시']

		if (temp !=  u'미정'):
			temp = temp.replace(')', ') ')
			time_dt = datetime.datetime.strptime(temp.split(' ')[-1], "%H:%M")
			time = preprocessTime(time_dt) # change datetime to serial time format(excel time format)
			test_df.at[index, u'시간'] = float(time) # excel time format corresponds to float 

		# get home team
		home = info.split(' ')[0]
		test_df.at[index, u'홈'] = home

		# get away team
		away = info.split(' ')[-1]
		test_df.at[index, u'원정'] = away

		# get rate
		temp = row[u'배당률선택']
		if (temp != '---'):
			temp = temp.replace(u'배당률 하락', '')
			temp = temp.replace(u'배당률 상승', '')
			temp = temp.replace(u'발매차단', '') # FIXME
			temp = temp.replace(u'승', ' 승 ')
			temp = temp.replace(u'패', ' 패 ')
			temp = temp.replace(u'무', ' 무 ')
			temp = temp.replace('U', ' U ')
			temp = temp.replace('O', ' O ')
			temp = temp.replace('-', ' - ')
			temp_list = temp.split(' ')

			if (game_type == u'일반' or game_type == u'핸디캡'):
				win_rate = temp_list[temp_list.index(u'승') + 1]
				test_df.at[index, u'승'] = float(win_rate)

				lose_rate = temp_list[temp_list.index(u'패') + 1]
				test_df.at[index, u'패'] = float(lose_rate)

				try:
					draw_idx = temp_list.index(u'무')
					test_df.at[index, u'무'] = float(temp_list[draw_idx + 1])
				except:
					pass

			elif (game_type == u'언더오버'):
				win_rate = temp_list[temp_list.index('U') + 1]
				test_df.at[index, u'승'] = float(win_rate)

				lose_rate = temp_list[temp_list.index('O') + 1]
				test_df.at[index, u'패'] = float(lose_rate)

	# drop cols
	#test_df = test_df.drop(u'대회', axis=1)
	#test_df = test_df.drop(u'게임유형', axis=1)
	test_df = test_df.drop(u'홈팀 vs 원정팀', axis=1)
	test_df = test_df.drop(u'배당률선택', axis=1)
	test_df = test_df.drop(u'경기일시', axis=1)

	return test_df, crawled_df


def getDataFromURL(url, filename):

	test_df, crawled_df = getGameDataframeFromURL(url)

	excel_writer = pd.ExcelWriter(filename,
																engine='xlsxwriter',
																datetime_format='yyyy/mm/dd')

	# FIXME to funciton
	new_sheetname = 'crawled_input'
	writeInputToExcel(excel_writer, test_df, new_sheetname) # to set time_col cell formatting on '시간' col

	# FIXME - change to Omod function
	#crawled_df.to_excel(excel_writer, sheet_name = 'Temp', index=False)

	excel_writer.save()

def createPost(g_num):

	post_text = ''
	post_text += '---' + '\n'
	post_text += 'layout: post' + '\n'
	title = '프로토 승부식 ' + str(int(g_num%10000))  + '회차 초기배당'
	post_text += 'title:  \"' + title + '\"' + '\n'
	now = datetime.datetime.now()
	cur_time = now.strftime('%Y-%m-%d %H:%M:%S')
	post_text += 'date:   ' + cur_time + ' +0900' + '\n'
	post_text += '---' + '\n'
	post_text += '' + '\n' # empty contents
	post_text += '' + '\n'
	post_text += 'Excel file : [' + str(g_num) + '][' + str(g_num) + ']' + '\n'
	post_text += '' + '\n'
	post_text += '[' + str(g_num) + ']: {{ site.url }}/crawler/output/' + str(g_num) + '.xlsx' + '\n'

	date = now.strftime('%Y-%m-%d-')
	f = open('./_posts/' + date + title.replace(' ', '-') + '.markdown', 'w')
	f.write(post_text)
	f.close()

if __name__ == "__main__":

	url, g_num = updateURL()

	o_filename = "crawler/output/"
	o_filename += str(g_num) + ".xlsx"

	getDataFromURL(url, o_filename)

	createPost(g_num)
	print (url)

