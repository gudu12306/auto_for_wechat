import win32clipboard as w
import win32con
import win32api
import win32gui
import ctypes
import time
import requests
from urllib.request import urlopen
from bs4 import BeautifulSoup
from selenium import webdriver

#把文字放入剪贴板
def setText(aString):
	w.OpenClipboard()
	w.EmptyClipboard()
	w.SetClipboardData(win32con.CF_UNICODETEXT,aString)
	w.CloseClipboard()
	
#模拟ctrl+V
def ctrlV():
	win32api.keybd_event(17,0,0,0) #ctrl
	win32api.keybd_event(86,0,0,0) #V
	win32api.keybd_event(86,0,win32con.KEYEVENTF_KEYUP,0)#释放按键
	win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)
	
#模拟alt+s
def altS():
	win32api.keybd_event(18,0,0,0)
	win32api.keybd_event(83,0,0,0)
	win32api.keybd_event(83,0,win32con.KEYEVENTF_KEYUP,0)
	win32api.keybd_event(18,0,win32con.KEYEVENTF_KEYUP,0)
# 模拟enter
def enter():
	win32api.keybd_event(13,0,0,0)
	win32api.keybd_event(13,0,win32con.KEYEVENTF_KEYUP,0)
#模拟单击
def click():
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
#移动鼠标的位置
def movePos(x,y):
	win32api.SetCursorPos((x,y))

def getZZWeatherAndSendMsg():
	HTML1='http://www.weather.com.cn/weather1dn/101190201.shtml'
	driver=webdriver.Edge()
	driver.get(HTML1)
	soup=BeautifulSoup(driver.page_source,'html5lib')
	
	#获取实时温度
	tem=soup.find('span',class_='temp').string
	#获取最高温度与最低温度
	maxtem=soup.find('span',id='maxTemp').string
	mintem=soup.find('span',id='minTemp').string
	#获取污染状况
	poll=soup.find('a',href='http://www.weather.com.cn/air/?city=101190201').string
	#获取风向和湿度
	win=soup.find('span',id='wind').string
	humidity=soup.find('span',id='humidity').string
	#获取紫外线状况
	sun=soup.find('div',class_='lv').find('em').string
	#获取穿衣指南
	cloth=soup.find('dl',id='cy').find('dd').string

	HTML2='http://www.weather.com.cn/weathern/101190201.shtml'
	driver.get(HTML2)
	soup=BeautifulSoup(driver.page_source,'html5lib')
	#获取天气情况
	wea=soup.find_all('p',class_='weather-info')[1].string
	weatherContent='实时温度：'+tem+'℃'+'\n'+'今日温度变化：'+mintem+'~'+maxtem+'\n'+'今日天气：'+wea+'\n'+'当前风向：'+win+'\n'+'相对湿度：'+humidity+'\n'+'紫外线：'+sun+'\n'+'污染指数：'+poll+'\n'+'穿衣指南:'+cloth+'\n'+'注意天气变化！！'
	driver.quit()
	return weatherContent

def getWeibo():
	url='https://s.weibo.com/top/summary'
	headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.72 Safari/537.36 Edg/90.0.818.41'}
	r=requests.get(url,headers=headers)
	r.raise_for_status()
	r.encoding = r.apparent_encoding
	soup = BeautifulSoup(r.text, "html.parser")
	tr=soup.find_all('tr')
	weiboContent='今日微博热榜：'+'\n'
	for i in range(2,52):
		text=tr[i].find('td',class_='td-02').find('a').string
		weiboContent=weiboContent+str(i-1)+'"'+text+'"'+'\n'
	return weiboContent

if __name__=="__main__":
	target_a=['06:55','11:55','19:53']
	target_b=['07:00','12:00','19:54']
	name_list=['Squirrel B','Squirrel B']
	while True:
		now=time.strftime("%m月%d日%H:%M",time.localtime())
		print(now)
		if now[-5:] in target_a:
			base_weatherContent=getZZWeatherAndSendMsg()
			weiboContent=getWeibo()
		if now[-5:] in target_b:
			hwnd=win32gui.FindWindow("WeChatMainWndForPC", '微信')
			win32gui.ShowWindow(hwnd,win32con.SW_SHOW)
			win32gui.MoveWindow(hwnd,0,0,1000,700,True)
			time.sleep(1)
			for name in name_list:
				movePos(28,147)
				click()
				#2.移动鼠标到搜索框，单击，输入要搜索的名字
				movePos(148,35)
				click()
				time.sleep(1)
				setText(name)
				ctrlV()
				time.sleep(1)  # 等待联系人搜索成功
				enter()
				time.sleep(1)
				now=time.strftime("%m月%d日%H:%M",time.localtime())
				weatherContent='现在是'+now+'\n'+base_weatherContent
				setText(weatherContent)
				ctrlV()
				time.sleep(1)
				altS()
				time.sleep(1)
				setText(weiboContent)
				ctrlV()
				time.sleep(1)
				altS()
				time.sleep(1)
			win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
		time.sleep(60)
