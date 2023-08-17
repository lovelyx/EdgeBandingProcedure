from PySide2.QtCore import QSettings
from PySide2.QtWidgets import QFileDialog,QMessageBox
from PySide2.QtUiTools import QUiLoader
import os

from lib.share import SI


class Patch_Main:
    def __init__(self):
        self.traceNum=True
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('PatchSetting.py'))
        filename = os.path.join(CURRENT_DIRECTORY, "foreEnd\\PatchSetting.ui")
        self.ui = QUiLoader().load(filename)
        self.init_login_info()

        self.ui.ButtonInput.clicked.connect(self.XmlInputPacth)
        self.ui.ButtonInput_2.clicked.connect(self.MprInputPacth)
        self.ui.ButtonOut.clicked.connect(self.CsvoutputPacth)
        self.ui.ButtonOut_2.clicked.connect(self.XmloutputPacth)
        self.ui.ButtonOut_3.clicked.connect(self.MproutputPacth)
        self.ui.ButtonOut_4.clicked.connect(self.SawCsvoutputPacth)
        self.ui.save.clicked.connect(self.save_login_info)



    def XmlInputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择Xml输入文件夹")
        self.ui.lineEditInput.setText(FileDirectory)
        print(FileDirectory)

    def MprInputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择MPR输入文件夹")
        self.ui.lineEditInput_2.setText(FileDirectory)
        print(FileDirectory)

    def CsvoutputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择Excel输出文件夹")
        self.ui.lineEditOut.setText(FileDirectory)
        print(FileDirectory)

    def XmloutputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择Xml输出文件夹")
        self.ui.lineEditOut_2.setText(FileDirectory)
        print(FileDirectory)

    def MproutputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择MPR输出文件夹")
        self.ui.lineEditOut_3.setText(FileDirectory)
        print(FileDirectory)

    def SawCsvoutputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择存放sawCsv文件夹")
        self.ui.lineEditOut_4.setText(FileDirectory)
        print(FileDirectory)


    ##保存lineEdit中内容
    def save_login_info(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        settings.setValue("XmlInput", self.ui.lineEditInput.text())
        settings.setValue("MprInput", self.ui.lineEditInput_2.text())
        settings.setValue("CsvOutput", self.ui.lineEditOut.text())
        settings.setValue("XmlOutput",self.ui.lineEditOut_2.text())
        settings.setValue("MprOutput", self.ui.lineEditOut_3.text())
        settings.setValue("SawCsvOutput", self.ui.lineEditOut_4.text())
        QMessageBox.information(self.ui,'保存','保存成功')
        self.ui.close()

    ##显示上次保存的值
    def init_login_info(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        XmlInput= settings.value("XmlInput")
        MprInput = settings.value("MprInput")
        CsvOutput = settings.value("CsvOutput")
        XmlOutput= settings.value("XmlOutput")
        MprOutput = settings.value("MprOutput")
        SawCsvOutput = settings.value("SawCsvOutput")



        self.ui.lineEditInput.setText(XmlInput)
        self.ui.lineEditInput_2.setText(MprInput)
        self.ui.lineEditOut.setText(CsvOutput)
        self.ui.lineEditOut_2.setText(XmlOutput)
        self.ui.lineEditOut_3.setText(MprOutput)
        self.ui.lineEditOut_4.setText(SawCsvOutput)




def main():
    SI.PatchWin = Patch_Main()
    SI.PatchWin.ui.show()



