import shutil
import time
import cryptocode
from PySide2.QtCore import QSettings
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QMessageBox
import foreEnd.PatchSetting as PatchIndex
import afterEnd.banding2 as banding
import afterEnd.bandingFive as bandingFive
import afterEnd.Shield as Shield
import afterEnd.compress as zipfile
import foreEnd.SqlSet as SqlSet
import afterEnd.MprProcessing as Mpr
import afterEnd.statistics as statistics
import afterEnd.FourSawed as FourSawed
from lib.share import SI
import os
import datetime
from threading import Thread


class Win_Main:

    def __init__(self, username, authority):
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('index.py'))
        filename = os.path.join(CURRENT_DIRECTORY, "foreEnd\\index.ui")
        self.username = username
        self.authority = authority
        self.ui = QUiLoader().load(filename)
        self.ui.PatchActionc.triggered.connect(self.PatchSignIn)
        self.ui.ExitAction.triggered.connect(self.ExitAction)
        self.ui.Buttonbanding.clicked.connect(self.BandingExt)
        self.ui.Buttonbanding2.clicked.connect(self.BandingExtFive)
        self.ui.MprProcessing.clicked.connect(self.MprProcessing)
        self.ui.pushButton.clicked.connect(self.statistics)
        # self.ui.FourSawed.clicked.connect(self.FourSawed)
        self.ui.sqlActionc.triggered.connect(self.sqlAction)
        self.ui.VersionAction.triggered.connect(self.about)

    def sqlAction(self):
        if self.authority == '1':
            SqlSet.main()
        else:
            QMessageBox.warning(
                    self.ui,
                    '提示',
                    "用户无权限")

    def about(self):
        QMessageBox.about(
            self.ui,
            '版本信息',
            '加工文件处理程序V4.2.8\n2023.1.12')

    @staticmethod
    def PatchSignIn():
        PatchIndex.main()

    def ExitAction(self):
        # 返回登录界面
        SI.loginWin.ui.show()
        self.ui.hide()

    def BandingExt(self):
        if int(self.authority) < 3:
            settings = QSettings("config.ini", QSettings.IniFormat)
            XmlInput = settings.value("XmlInput")
            MprInput = settings.value("MprInput")
            CsvOutput = settings.value("CsvOutput")
            XmlOutput = settings.value("XmlOutput")
            MprOutput = settings.value("MprOutput")
            SawCsvOutput = settings.value("SawCsvOutput")
            CheckBox = settings.value("CheckBox")
            sqlIp = settings.value("sqlIp")
            edt_username = settings.value("edt_username")
            edt_password = settings.value("edt_password")
            sqlIp = cryptocode.decrypt(sqlIp, "kfht.")
            edt_username = cryptocode.decrypt(edt_username, "kfht.")
            edt_password = cryptocode.decrypt(edt_password, "kfht.")

            IdList = []
            IdDictNo = {}
            # 按钮禁用
            self.ui.Buttonbanding.setEnabled(False)
            self.ui.Buttonbanding2.setEnabled(False)
            self.ui.MprProcessing.setEnabled(False)
            self.ui.pushButton.setEnabled(False)
            # self.ui.FourSawed.setEnabled(False)
            self.ui.textBrowser.append('开始处理')

            def run():
                self.DeleteFile(XmlOutput)
                self.DeleteFile(MprOutput)
                lsi, SqlUnit = banding.main(XmlInput, CsvOutput, XmlOutput, SawCsvOutput, IdList, CheckBox,
                                            self.username, IdDictNo, sqlIp, edt_username, edt_password)
                num = 0
                for i in lsi:
                    num += 1
                    self.ui.textBrowser.append(i)
                    time.sleep(0.2)
                    # 定位鼠标到最后一行中
                    self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                self.ui.textBrowser.append(f'xml处理完成,共处理{num}个文件')
                self.ui.textBrowser.append(f'正在处理MPR文件,请稍后.........')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                print(f'IdDictNo{IdDictNo}')
                Shield.main(MprInput, MprOutput, IdList, IdDictNo)
                print("mpr处理完成")
                self.ui.textBrowser.append(f'MPR文件处理完成')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                self.ui.Buttonbanding.setEnabled(True)
                self.ui.Buttonbanding2.setEnabled(True)
                self.ui.MprProcessing.setEnabled(True)
                self.ui.pushButton.setEnabled(True)
                # self.ui.FourSawed.setEnabled(True)
                SqlUnit.conn.close()
                print('关闭成功')

                self.zip(XmlOutput, CsvOutput, "先达加工")
                self.zip(MprOutput, CsvOutput, "通过式加工")
            # 多线程运行
            t = Thread(target=run)
            t.setDaemon(True)
            t.start()
        else:
            QMessageBox.warning(
                self.ui,
                '提示',
                "用户无权限")

    @staticmethod
    def DeleteFile(patch):
        # 遍历文件夹中的所有文件，判断是文件则删除
        # 遍历文件夹中的所有文件，一一删除
        for file_name in os.listdir(patch):
            file_path = os.path.join(patch, file_name)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # 删除文件
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # 删除目录
            except Exception as e:
                print("删除 %s 时出错,原因：%s" % (file_path, e))

    @staticmethod
    def zip(InputPatch, zipPath, file):
        filetime = datetime.datetime.now().strftime("%Y%m%d%H%M")
        CsvOutputZip = zipPath + "/" + file + filetime + ".zip"
        zipfile.zip_folder(InputPatch, CsvOutputZip)

    def BandingExtFive(self):
        if int(self.authority) < 3:
            settings = QSettings("config.ini", QSettings.IniFormat)
            XmlInput = settings.value("XmlInput")
            CsvOutput = settings.value("CsvOutput")
            XmlOutput = settings.value("XmlOutput")
            MprOutput = settings.value("MprOutput")
            SawCsvOutput = settings.value("SawCsvOutput")
            # CheckBox = settings.value("CheckBox")
            sqlIp = settings.value("sqlIp")
            edt_username = settings.value("edt_username")
            edt_password = settings.value("edt_password")
            sqlIp = cryptocode.decrypt(sqlIp, "kfht.")
            edt_username = cryptocode.decrypt(edt_username, "kfht.")
            edt_password = cryptocode.decrypt(edt_password, "kfht.")
            CheckBox = True
            IdList = []
            # 按钮禁用
            self.ui.Buttonbanding.setEnabled(False)
            self.ui.Buttonbanding2.setEnabled(False)
            self.ui.MprProcessing.setEnabled(False)
            self.ui.pushButton.setEnabled(False)
            # self.ui.FourSawed.setEnabled(False)
            self.ui.textBrowser.append('开始处理')

            def run():
                self.DeleteFile(XmlOutput)
                self.DeleteFile(MprOutput)
                lsi, SqlUnit = bandingFive.main(XmlInput, CsvOutput, XmlOutput, SawCsvOutput, IdList, CheckBox,
                                                self.username, sqlIp, edt_username, edt_password)
                num = 0
                for i in lsi:
                    num += 1
                    self.ui.textBrowser.append(i)
                    time.sleep(0.2)
                    # 定位鼠标到最后一行中
                    self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                self.ui.textBrowser.append(f'xml处理完成,共处理{num}个文件')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                self.ui.Buttonbanding.setEnabled(True)
                self.ui.Buttonbanding2.setEnabled(True)
                self.ui.MprProcessing.setEnabled(True)
                self.ui.pushButton.setEnabled(True)
                # self.ui.FourSawed.setEnabled(True)

                SqlUnit.conn.close()
                print('关闭成功')

            # 多线程运行
            t = Thread(target=run)
            t.setDaemon(True)
            t.start()
        else:
            QMessageBox.warning(
                self.ui,
                '提示',
                "用户无权限")

    def MprProcessing(self):
        if int(self.authority) < 4:
            settings = QSettings("config.ini", QSettings.IniFormat)
            MprInput = settings.value("MprInput")
            MprOutput = settings.value("MprOutput")
            # 按钮禁用
            self.ui.Buttonbanding.setEnabled(False)
            self.ui.Buttonbanding2.setEnabled(False)
            self.ui.MprProcessing.setEnabled(False)
            self.ui.pushButton.setEnabled(False)
            # self.ui.FourSawed.setEnabled(False)
            self.ui.textBrowser.append('开始处理')

            def run():
                self.DeleteFile(MprOutput)
                self.ui.textBrowser.append(f'正在处理MPR文件,请稍后.........')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                Mpr.main(MprInput, MprOutput)
                print("mpr处理完成")
                self.ui.textBrowser.append(f'MPR文件处理完成')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                self.ui.Buttonbanding.setEnabled(True)
                self.ui.Buttonbanding2.setEnabled(True)
                self.ui.MprProcessing.setEnabled(True)
                self.ui.pushButton.setEnabled(True)
                # self.ui.FourSawed.setEnabled(True)

            # 多线程运行
            t = Thread(target=run)
            t.setDaemon(True)
            t.start()
        else:
            QMessageBox.warning(
                self.ui,
                '提示',
                "用户无权限")

    def statistics(self):
        if int(self.authority) < 3:
            settings = QSettings("config.ini", QSettings.IniFormat)
            XmlOutput = settings.value("XmlOutput")
            # 按钮禁用
            self.ui.Buttonbanding.setEnabled(False)
            self.ui.Buttonbanding2.setEnabled(False)
            self.ui.MprProcessing.setEnabled(False)
            self.ui.pushButton.setEnabled(False)
            # self.ui.FourSawed.setEnabled(False)
            self.ui.textBrowser.append('开始处理')

            def run():
                self.ui.textBrowser.append(f'正在统计,请稍后.........')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                statistics.main(XmlOutput)
                self.ui.textBrowser.append(f'统计完成')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                self.ui.Buttonbanding.setEnabled(True)
                self.ui.Buttonbanding2.setEnabled(True)
                self.ui.MprProcessing.setEnabled(True)
                self.ui.pushButton.setEnabled(True)
                # self.ui.FourSawed.setEnabled(True)

            # 多线程运行
            t = Thread(target=run)
            t.setDaemon(True)
            t.start()
        else:
            QMessageBox.warning(
                self.ui,
                '提示',
                "用户无权限")

    def FourSawed(self):
        if int(self.authority) < 3:
            settings = QSettings("config.ini", QSettings.IniFormat)
            XmlOutput = settings.value("XmlOutput")
            # 按钮禁用
            self.ui.Buttonbanding.setEnabled(False)
            self.ui.Buttonbanding2.setEnabled(False)
            self.ui.MprProcessing.setEnabled(False)
            self.ui.pushButton.setEnabled(False)
            self.ui.FourSawed.setEnabled(False)
            self.ui.textBrowser.append('开始处理')

            def run():
                self.ui.textBrowser.append(f'正在处理四锯边,请稍后.........')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                FourSawed.main(XmlOutput)
                self.ui.textBrowser.append(f'四锯边处理完成')
                self.ui.textBrowser.moveCursor(self.ui.textBrowser.textCursor().End)
                self.ui.Buttonbanding.setEnabled(True)
                self.ui.Buttonbanding2.setEnabled(True)
                self.ui.MprProcessing.setEnabled(True)
                self.ui.pushButton.setEnabled(True)
                self.ui.FourSawed.setEnabled(True)

            # 多线程运行
            t = Thread(target=run)
            t.setDaemon(True)
            t.start()
        else:
            QMessageBox.warning(
                self.ui,
                '提示',
                "用户无权限")


def index(username, authority):
    SI.mainWin = Win_Main(username, authority)
    SI.mainWin.ui.show()
