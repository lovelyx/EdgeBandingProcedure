import re
# Author: LiYiXiao
# 处理xml,得到相关文件

from xml.dom import minidom
import pymsgbox as mb
import logging
import traceback
import os
import pandas as pd
import datetime

from afterEnd.page.My_sheet import My_sheet
from afterEnd.page.styleT import style
import lib.Sql as sqlUnit



def rth(ebl,panel,eb):
    if ebl.split('*')[0]=='':
        return
    elif float(ebl.split('*')[0])>=1.0:
        pattern = r'([\u4e00-\u9fa5]+)'
        result = re.findall(pattern, ebl)
        for i in result:
            i = '▲▲▲▲' + i
            panel.setAttribute(eb, i)
    elif float(ebl.split('*')[0])<1.0:
        pattern = r'([\u4e00-\u9fa5]+)'
        result = re.findall(pattern, ebl)
        for i in result:
            i = '△△△△' + i
            panel.setAttribute(eb, i)

        # panel.setAttribute(ebl, "")

def changeWwidth(Folderpath,Folderpath3):
    # 打开文件夹得到文件夹下的所有文件
    listPacth=[]
    print(Folderpath)
    xml_list = os.listdir(Folderpath)
    for j in xml_list:  # 遍历所有xml文件
        try:
            fullPath = Folderpath + "/" + j  # 完整路径
            listPacth.append(j)
            # 打开文件
            xmldoc = minidom.parse(fullPath)
            #找到Machining节点（标签）
            Panelnodes = xmldoc.getElementsByTagName("Panel")
            PanelList=[]
            for panel in Panelnodes:
                EBL1 = panel.getAttribute('EBL1')
                rth(EBL1,panel,'EBL1')
                EBL2 = panel.getAttribute('EBL2')
                rth(EBL2, panel,'EBL2')
                EBW1 = panel.getAttribute('EBW1')
                rth(EBW1, panel,'EBW1')
                EBW2 = panel.getAttribute('EBW2')
                rth(EBW2, panel,'EBW2')

            with open(Folderpath3 +"/"+ fullPath.split('/')[-1], "w", encoding="UTF-8") as fs:
                fs.write(xmldoc.toxml())
                fs.close()
            #     插入数据库
            # SqlUnit.InserUnit(PanelList)
        except:
            logging.basicConfig(filename='log.txt', level=logging.DEBUG,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            # 方案一，自己定义一个文件，自己把错误堆栈信息写入文件。
            errorFile = open('log.txt', 'a')
            errorFile.write(traceback.format_exc())
            errorFile.close()
            # tkinter.messagebox.showerror(title='错误',message=f"{j.split('.')[0]}文件处理失败，请联系管理员")
            mb.alert(f"{j}文件处理失败，请联系管理员",'报错')
            # win32api.MessageBox(0,f"{j.split('.')[0]}文件处理失败，请联系管理员","错误",win32con.MB_OK)
    # return IdList
if __name__ == '__main__':
    Folderpath='D:/Desktop/1'
    Folderpath3='D:/Desktop/3'
    print('sdfa')
    changeWwidth(Folderpath,Folderpath3)



