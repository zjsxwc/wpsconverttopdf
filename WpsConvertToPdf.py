# -*- coding: utf-8 -*

# 该脚本 依赖pywin32，需要保证已经执行了  `pip install pywin32`

# 需要设置wps为xlsx、docx、pptx等拓展名文件的默认打开程序！

import win32gui
import win32api
import win32con
from ctypes import *

import os
import time
import sys
import platform

class WpsConvertToPdf:
	def __init__(self):
		pass

	def keyboardWithAlt(self, k):
		# k 必须的大写的字符 比如 'F'
		win32api.keybd_event(win32con.VK_MENU, 0,0,0)
		time.sleep(0.1)
		win32api.keybd_event(ord(k), 0,0,0)
		time.sleep(0.1)
		win32api.keybd_event(ord(k), 0,win32con.KEYEVENTF_KEYUP,0)
		time.sleep(0.1)
		win32api.keybd_event(win32con.VK_MENU, 0,win32con.KEYEVENTF_KEYUP,0)

	def keyboardPressKey(self, k):
		# k 必须的大写的字符 比如 'F'
		time.sleep(0.1)
		win32api.keybd_event(ord(k), 0,0,0)
		time.sleep(0.1)
		win32api.keybd_event(ord(k), 0,win32con.KEYEVENTF_KEYUP,0)
		time.sleep(0.1)

	def keyboardPressEnter(self):
		time.sleep(0.1)
		win32api.keybd_event(win32con.VK_RETURN, 0,0,0)
		time.sleep(0.1)
		win32api.keybd_event(win32con.VK_RETURN, 0,win32con.KEYEVENTF_KEYUP,0)
		time.sleep(0.1)

	def convert(self, srcFilePath):

		fileBaseName = os.path.splitext(os.path.basename(srcFilePath))[0]
		fileOrigExtName = os.path.splitext(os.path.basename(srcFilePath))[1]
		fileExtName = fileOrigExtName.lower()

		#如果已经存在导出的pdf文件先删除这个pdf文件
		mayPdfFilePath = os.path.dirname(os.path.realpath(srcFilePath)) + "\\" +fileBaseName + ".pdf"
		if (os.path.exists(mayPdfFilePath)):
			print("remove" + mayPdfFilePath)
			os.remove(mayPdfFilePath)

		
		# 需要设置wps为xlsx、docx、pptx等拓展名文件的默认打开程序！
		os.startfile(srcFilePath, "edit")

		# 这里wps进程后缀不同系统语言等原因可能不一样，需要修改
		wndMainTitleSuffix = " - WPS 表格"
		if (fileExtName == ".docx"):
			wndMainTitleSuffix = " - WPS 文字"
		if (fileExtName == ".pptx"):
			wndMainTitleSuffix = " - WPS 演示"
		if (platform.platform().startswith("Windows-10")):
			wndMainTitleSuffix = " - WPS Office"

		wndMainTitle = fileBaseName + fileOrigExtName + wndMainTitleSuffix

		wndMain = None
		while not wndMain:
			time.sleep(1)
			wndMain = win32gui.FindWindow(None, wndMainTitle)

		wndMainRect = win32gui.GetWindowRect(wndMain)
		win32gui.SetForegroundWindow(wndMain)
		time.sleep(1)
		self.keyboardWithAlt('F')
		self.keyboardPressKey('F')
		time.sleep(0.4)
		self.keyboardPressEnter()

		processingWnd = None
		while not processingWnd:
			time.sleep(1)
			processingWnd = win32gui.FindWindow(None, "输出 PDF 文件")

		couldRenamePdfFile = False
		pdfFilePath = mayPdfFilePath
		tryRenamePdfFileNamePath = pdfFilePath + "~"
		while (not couldRenamePdfFile):
			if (os.path.exists(mayPdfFilePath)):
				try:
					os.rename(pdfFilePath, tryRenamePdfFileNamePath)
					couldRenamePdfFile = True
					break
				except Exception as e:
					print(e)
					pass
			time.sleep(0.5)
		time.sleep(1)
		self.keyboardPressEnter()
		self.keyboardWithAlt('F')
		self.keyboardPressKey('Q')

		os.rename(tryRenamePdfFileNamePath, pdfFilePath)

# 例子
# x = WpsConvertToPdf()
# x.convert(r"C:\Users\zjsxwc\Desktop\WpsConvertToPdf\test.xlsx")
