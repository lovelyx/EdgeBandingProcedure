from PySide2.QtCore import QSettings
from PySide2.QtWidgets import QFileDialog, QMessageBox
from PySide2.QtUiTools import QUiLoader
import os

from lib.share import SI


class Patch_Main:
    def __init__(self):
        self.traceNum = True
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('PatchSetting.py'))
        filename = os.path.join(CURRENT_DIRECTORY, "foreEnd\\PatchSetting.ui")
        self.ui = QUiLoader().load(filename)
        self.init_login_info()

        self.ui.ButtonInput.clicked.connect(self.XmlInputPath)
        self.ui.ButtonInput_2.clicked.connect(self.MprInputPath)
        self.ui.ButtonOut.clicked.connect(self.CsvOutputPath)
        self.ui.ButtonOut_2.clicked.connect(self.XmlOutputPath)
        self.ui.ButtonOut_3.clicked.connect(self.MprOutputPath)
        self.ui.ButtonOut_4.clicked.connect(self.SawCsvOutputPath)
        self.ui.checkBox.stateChanged.connect(self.checkF)
        self.ui.save.clicked.connect(self.save_login_info)

    def XmlInputPath(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择Xml输入文件夹")
        self.ui.lineEditInput.setText(FileDirectory)
        print(FileDirectory)

    def MprInputPath(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择MPR输入文件夹")
        self.ui.lineEditInput_2.setText(FileDirectory)
        print(FileDirectory)

    def CsvOutputPath(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择Excel输出文件夹")
        self.ui.lineEditOut.setText(FileDirectory)
        print(FileDirectory)

    def XmlOutputPath(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择Xml输出文件夹")
        self.ui.lineEditOut_2.setText(FileDirectory)
        print(FileDirectory)

    def MprOutputPath(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择MPR输出文件夹")
        self.ui.lineEditOut_3.setText(FileDirectory)
        print(FileDirectory)

    def SawCsvOutputPath(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择存放sawCsv文件夹")
        self.ui.lineEditOut_4.setText(FileDirectory)
        print(FileDirectory)

    def checkF(self):
        if self.ui.checkBox.isChecked():
            self.ui.checkBox.setChecked(False)
        else:
            self.ui.checkBox.setChecked(True)

    # 保存lineEdit中内容
    def save_login_info(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        settings.setValue("XmlInput", self.ui.lineEditInput.text())
        settings.setValue("MprInput", self.ui.lineEditInput_2.text())
        settings.setValue("CsvOutput", self.ui.lineEditOut.text())
        settings.setValue("XmlOutput", self.ui.lineEditOut_2.text())
        settings.setValue("MprOutput", self.ui.lineEditOut_3.text())
        settings.setValue("SawCsvOutput", self.ui.lineEditOut_4.text())
        if self.ui.checkBox.isChecked():
            g = 1
        else:
            g = 0
        settings.setValue("CheckBox", g)
        QMessageBox.information(self.ui, '保存', '保存成功')
        self.ui.close()

    # 显示上次保存的值
    def init_login_info(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        XmlInput = settings.value("XmlInput")
        MprInput = settings.value("MprInput")
        CsvOutput = settings.value("CsvOutput")
        XmlOutput = settings.value("XmlOutput")
        MprOutput = settings.value("MprOutput")
        SawCsvOutput = settings.value("SawCsvOutput")
        CheckBox = settings.value("CheckBox")

        self.ui.lineEditInput.setText(XmlInput)
        self.ui.lineEditInput_2.setText(MprInput)
        self.ui.lineEditOut.setText(CsvOutput)
        self.ui.lineEditOut_2.setText(XmlOutput)
        self.ui.lineEditOut_3.setText(MprOutput)
        self.ui.lineEditOut_4.setText(SawCsvOutput)
        y = True
        if int(CheckBox) == 1:
            y = True
        elif int(CheckBox) == 0:
            y = False
        self.ui.checkBox.setChecked(y)


def main():
    SI.PatchWin = Patch_Main()
    SI.PatchWin.ui.show()
