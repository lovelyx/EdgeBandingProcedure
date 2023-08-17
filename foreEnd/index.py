import time

from PyQt5.QtWidgets import QAction, qApp
from PySide2 import QtCore
from PySide2.QtAxContainer import QAxWidget
from PySide2.QtCore import QSettings
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QMessageBox
from PySide2 import QtGui


import foreEnd.PatchSetting as PachIndex
import afterEnd.banding2 as banding
import afterEnd.Shield as shield

from lib.share import SI
import os
from threading import Thread
import lib.Sql as sqlUnit

class Win_Main :

    def __init__(self):
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('index.py'))
        filename = os.path.join(CURRENT_DIRECTORY, "foreEnd\\index.ui")
        self.ui = QUiLoader().load(filename)
        self.ui.PatchActionc.triggered.connect(self.PatchSignIn)
        self.ui.ExitAction.triggered.connect(self.ExitAction)
        self.ui.Buttonbanding.clicked.connect(self.BandingExt)
        self.ui.sqlActionc.triggered.connect(self.sqlActionc)
        # self.ui.ButtonAvailability.clicked.connect(self.Availability)
        self.ui.VersionAction.triggered.connect(self.about)

    def sqlActionc(self):
        QMessageBox.warning(
                self.ui,
                '提示',
                "用户无权限")


    def about(self):
        QMessageBox.about(
            self.ui,
            '版本信息',
            '后台封边程序工具V3.3.0\n2023.8.5')

    def PatchSignIn(self):
        PachIndex.main()


    def ExitAction(self):
        # s = requests.Session()
        # url = "http://192.168.1.105/api/sign"
        #
        # res = s.post(url,json={
        #     "action" : "signout",
        # })
        #
        # resObj = res.json()
        # if resObj['ret'] != 0:
        #     QMessageBox.warning(
        #         self.ui,
        #         '登出失败',
        #         resObj['msg'])
        #     return
        # 返回登录界面
        SI.loginWin.ui.show()
        self.ui.hide()


    def BandingExt(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        XmlInput= settings.value("XmlInput")
        MprInput = settings.value("MprInput")
        CsvOutput= settings.value("CsvOutput")
        XmlOutput = settings.value("XmlOutput")
        MprOutput = settings.value("MprOutput")
        SawCsvOutput = settings.value("SawCsvOutput")
        IdList = []
        # 按钮禁用
        self.ui.Buttonbanding.setEnabled(False)
        self.ui.textBrowser.append('开始处理')
        def run():
            lsi,SqlUnit = banding.main(XmlInput,CsvOutput,XmlOutput,SawCsvOutput,IdList)
            num = 0
            for i in lsi:
                num+=1
                self.ui.textBrowser.append(i)
                # 定位鼠标到最后一行中
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
            self.ui.textBrowser.append(f'xml处理完成,共处理{num}个文件')
            self.ui.textBrowser.append(f'正在处理MPR文件,请稍后.........')
            self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
            # print(IdList)
            shield.main(MprInput,MprOutput,IdList)
            print("mpr处理完成")
            self.ui.textBrowser.append(f'MPR文件处理完成')
            self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
            self.ui.Buttonbanding.setEnabled(True)


            SqlUnit.conn.close()
            print('关闭成功')
        # 多线程运行
        t = Thread(target=run)
        t.setDaemon(True)
        t.start()





        # def mpr(MprInput,IdList):
            # shield.main(MprInput,MprOutput,IdList)
            # List =[]
            # # mpr_list = os.listdir(MprInput)
            # for j in IdList:
            #     for p in j:
            #         List.append(p)
            #
            # for root, dirs, files in os.walk(MprInput):
            #     # for file in files:
            #     #     path = os.path.join(root, file)
            #     #     print(path)
            #     for i in files:
            #         print(i)
            #         Mprfile = i.split(".")[0]
            #         if Mprfile in List and (i.split(".")[1]=='MPR' or i.split(".")[1]=='mpr'):
            #             path = os.path.join(root, i)
            #             # print(path)
            #             shield.main(MprOutput, path)
            #         else:



    # def Availability(self):
    #     # 按钮禁用
    #     self.ui.Buttonbanding.setEnabled(False)
    #
    #     def run():
    #         lsi = avali.main()
    #         for i in lsi:
    #             self.ui.textBrowser.append(i)
    #         self.ui.Buttonbanding.setEnabled(True)
    #
    #     #     多线程运行
    #     t = Thread(target=run)
    #     t.start()







def index():
    SI.mainWin = Win_Main()
    SI.mainWin.ui.show()