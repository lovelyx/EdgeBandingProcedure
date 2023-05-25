from PySide2.QtCore import QSettings
from PySide2.QtWidgets import QFileDialog,QMessageBox
from PySide2.QtUiTools import QUiLoader
import os

from lib.share import SI


class Patch_Main:
    def __init__(self):
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('PatchSetting.py'))
        filename = os.path.join(CURRENT_DIRECTORY, "foreEnd\\PatchSetting.ui")
        self.ui = QUiLoader().load(filename)
        self.init_login_info()

        self.ui.ButtonInput.clicked.connect(self.InputPacth)
        self.ui.ButtonOut.clicked.connect(self.outputPacth)
        self.ui.save.clicked.connect(self.save_login_info)



    def InputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择输入文件夹")
        self.ui.lineEditInput.setText(FileDirectory)
        print(FileDirectory)

    def outputPacth(self):
        FileDirectory = QFileDialog.getExistingDirectory(self.ui, "选择输出文件夹")
        self.ui.lineEditOut.setText(FileDirectory)
        print(FileDirectory)

    ##保存lineEdit中内容
    def save_login_info(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        settings.setValue("Input", self.ui.lineEditInput.text())
        settings.setValue("output", self.ui.lineEditOut.text())
        QMessageBox.information(self.ui,'保存','保存成功')
        self.ui.close()

    ##显示上次保存的值
    def init_login_info(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        the_Input= settings.value("Input")
        the_output= settings.value("output")

        self.ui.lineEditInput.setText(the_Input)
        self.ui.lineEditOut.setText(the_output)




def main():
    SI.PatchWin = Patch_Main()
    SI.PatchWin.ui.show()



