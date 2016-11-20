# -*- coding: utf-8 -*-
#修改EXCEL库，支持excel打开的时候写入和存储

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import unittest, time, re
import xlrd
import xlwt
from xlutils.copy import copy
import pyperclip
import pyscreeze
from pywinauto import application
import pyautogui
from win32com.client import Dispatch
import win32com.client
from lxml import etree
import autoit
import os

UrlGitHub = "https://github.com/dennishud/pyautoit/tree/master"
FolderToCheck = "E:\Desktop\PyTrail"

def LoginGitHub(driver,URL):

    driver.implicitly_wait(30)
    driver.maximize_window()

    driver.get(URL)

    xPath = "//div[@class='site-header-actions']/a[contains(text(),'Sign in')]"
    driver.find_element_by_xpath(xPath).click()

    xPath = "//input[@id='login_field']"
    driver.find_element_by_xpath(xPath).send_keys("dennishud@outlook.com")

    xPath = "//input[@id='password']"
    driver.find_element_by_xpath(xPath).send_keys("Aa7788250")

    xPath = "//input[@type='submit']"
    driver.find_element_by_xpath(xPath).click()

    time.sleep(5)

    return

def paste(ContentToPaste):
    pyperclip.copy(ContentToPaste)
    pyautogui.hotkey('ctrl', 'v')

def CreateFilesInGitHub(driver,FileNameToGive,FileContent):

    #点击创建新文件
    xPath = "//button[contains(text(),'Create new file')]"
    driver.find_element_by_xpath(xPath).click()
    #输入文件名
    xPath = "//input[@name='filename']"
    driver.find_element_by_xpath(xPath).send_keys(FileNameToGive)
    #输入文件内容 (暂时不能定位到输入框，用autoit方案替代)
    # xPath = "//textarea[@id='blob_contents_']"
    xPath = "//div[@class='commit-create']/textarea"
    # xPath = "//div[@id='js-repo-pjax-container']/div[2]/div[1]/div/form[2]/div[3]/div[2]/div/div[5]/div[1]/div/div/div"
    # driver.execute_script("document.getElementById('blob_contents_').style.top = 0;")
    # driver.find_element_by_css_selector('#blob_contents_').send_keys(FileContent)
    # # driver.find_element_by_xpath(xPath).send_keys(FileContent)

    autoit.win_activate(u"New File - Google Chrome")
    time.sleep(2)
    autoit.mouse_click("left",840,696)
    paste(FileContent)

    #输入commit summury
    xPath = "//input[@id='commit-summary-input']"
    driver.find_element_by_xpath(xPath).send_keys(FileNameToGive)
    #点击commit
    xPath = "//button[contains(text(),'Commit new file')]"
    driver.find_element_by_xpath(xPath).click()
    time.sleep(3)

    return

def ReadFileAndReplaceKeywords(FilePath,FileName):

    FileToRead = FilePath+FileName
    # FileToRead_ = u"E:\Desktop\PyTrail\YourProjectNameScenarioTest_Main_1117.py"

    file_object = open(FileToRead, 'r')
    all_the_text = file_object.read()
    all_the_text = all_the_text.decode('utf-8')

    xlApp = win32com.client.Dispatch('Excel.Application')
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlBook = xlApp.Workbooks.Open(u'E:\Desktop\PyTrail\StringReplaceTable.xlsx')
    table_ContentToReplace = xlBook.Worksheets(u'ContentToReplace')

    myRange = table_ContentToReplace.Range("B1:B20000")
    ColCount= xlApp.Application.WorksheetFunction.CountA(myRange)
    ColCount=int(ColCount)

    Cell_OldString_Col =1
    Cell_NewString_Col =2
    Cell_OldString =""
    Cell_NewString =""

    for i in range(2, ColCount+1, 1):
        Cell_OldString=table_ContentToReplace.Cells(i, Cell_OldString_Col).Value
        Cell_NewString=table_ContentToReplace.Cells(i, Cell_NewString_Col).Value

        all_the_text = all_the_text.replace(Cell_OldString, Cell_NewString)

    FileContent=all_the_text
    xlBook.Close()
    xlApp.Quit()

    return FileContent

def ProcessFilename(OldFileName):

    xlApp = win32com.client.Dispatch('Excel.Application')
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlBook = xlApp.Workbooks.Open(u'E:\Desktop\PyTrail\StringReplaceTable.xlsx')
    table_FilenameToReplace = xlBook.Worksheets(u'FilenameToReplace')

    myRange = table_FilenameToReplace.Range("B1:B20000")
    ColCount= xlApp.Application.WorksheetFunction.CountA(myRange)
    ColCount=int(ColCount)

    Cell_OldString_Col =1
    Cell_NewString_Col =2
    Cell_OldString =""
    Cell_NewString =""

    for i in range(2, ColCount+1, 1):
        Cell_OldString=table_FilenameToReplace.Cells(i, Cell_OldString_Col).Value
        Cell_NewString=table_FilenameToReplace.Cells(i, Cell_NewString_Col).Value

        OldFileName = OldFileName.replace(Cell_OldString, Cell_NewString)

    NewFileName=OldFileName

    return NewFileName

def GetAllPyfileUnderFoler(Folder):

    PyfileList=os.listdir(Folder)
    # for i in range(0,len(PyfileList),1):
    #     if PyfileList[i][:3]!=".py":
    #         PyfileList.remove()

    FileListLen = len(PyfileList)
    print "Orignal length: "+str(FileListLen)

    for file in PyfileList:
        # print file[-3:]
        if file[-3:] != ".py":
            # print file
            PyfileList.remove(file)
            FileListLen=len(PyfileList)

    print "Last length: "+str(FileListLen)

    for file in PyfileList:
        print file

    return PyfileList

if __name__ == '__main__':

    driver = webdriver.Chrome()
    LoginGitHub(driver, UrlGitHub)
    FilePath=u'''E:\Desktop\PyTrail\\'''

    for TempFile in GetAllPyfileUnderFoler(FilePath):
        if TempFile[-3:] == ".py":
            NewFilename = ProcessFilename(TempFile)

            FileContentForNewFile = ReadFileAndReplaceKeywords(FilePath, TempFile)

            CreateFilesInGitHub(driver, NewFilename, FileContentForNewFile)



