import cryptocode
from PySide2.QtCore import QSettings
from PySide2.QtWidgets import QMessageBox
from PySide2.QtUiTools import QUiLoader
import os

from lib.share import SI


class Patch_Main:
    def __init__(self):
        self.traceNum = True
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('SqlSetting.ui'))
        filename = os.path.join(CURRENT_DIRECTORY, "foreEnd\\SqlSetting.ui")
        self.ui = QUiLoader().load(filename)
        self.init_login_info()
        self.ui.save.clicked.connect(self.save_login_info)

    # 保存lineEdit中内容
    def save_login_info(self):
        sqlIP = self.ui.sqlIp.text()
        edt_username = self.ui.edt_username.text()
        edt_password = self.ui.edt_password.text()

        sqlIP = cryptocode.encrypt(sqlIP, "kfht.")
        edt_username = cryptocode.encrypt(edt_username, "kfht.")
        edt_password = cryptocode.encrypt(edt_password, "kfht.")

        settings = QSettings("config.ini", QSettings.IniFormat)
        settings.setValue("sqlIp", sqlIP)
        settings.setValue("edt_username", edt_username)
        settings.setValue("edt_password", edt_password)
        QMessageBox.information(self.ui, '保存', '保存成功')
        self.ui.close()

    # 显示上次保存的值
    def init_login_info(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        sqlIp = cryptocode.decrypt(settings.value("sqlIp"), 'kfht.')
        edt_username = cryptocode.decrypt(settings.value("edt_username"), 'kfht.')
        edt_password = cryptocode.decrypt(settings.value("edt_password"), 'kfht.')

        self.ui.sqlIp.setText(sqlIp)
        self.ui.edt_username.setText(edt_username)
        self.ui.edt_password.setText(edt_password)


def main():
    SI.PatchWin = Patch_Main()
    SI.PatchWin.ui.show()
