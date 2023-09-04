import base64
import hashlib
import sys

from PySide2.QtWidgets import QApplication, QMessageBox
from PySide2.QtUiTools import QUiLoader
import foreEnd.index2 as index2
import foreEnd.index as index

from lib.share import SI
import lib.Sql as sqlUnit

import os

class Win_Login:

    def __init__(self):
        # 从文件中加载UI定义

        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('login.py'))
        filename = os.path.join(CURRENT_DIRECTORY, "Login\\login.ui")
        self.ui = QUiLoader().load(filename)

        self.ui.btn_login.clicked.connect(self.onSignIn)
        self.ui.edt_password.returnPressed.connect(self.onSignIn)

    def onSignIn(self):
        username = self.ui.edt_username.text().strip()
        password = self.ui.edt_password.text().strip()

        # re=sqlUnit.SqlUnit.selectUnit(sqlUnit.SqlUnit, username, password)
        # 加密 base64
        # password2 = 'kfht'+password
        # password2 = password2.encode("utf-8")
        # password2 = base64.b64encode(password2)
        # password2 = base64.b64encode(password2)
        # password2=password2.decode("utf-8")
        # 加密md5
        md5=hashlib.md5()
        password='kfht.'+password
        password=str(password).encode(encoding="utf-8")
        md5.update(password)

        password2=md5.hexdigest()
        print(password2)
        SqlUnit = sqlUnit.main()
        re=SqlUnit.selectUnit(username,password2)
        # re=True
        # re = sqlUnit.selectUnit(username, password)
        if re:
            index.index(username)
        else:
            QMessageBox.warning(
                self.ui,
                 '登录失败','请检查密码')
            return


        # s = requests.Session()
        # url = "http://192.168.1.105/api/sign"
        #
        # res = s.post(url,json={
        #     "action" : "signin",
        #     "username" : username,
        #     "password" : password
        # })
        #
        # resObj = res.json()
        # if resObj['ret'] != 0:
        #     QMessageBox.warning(
        #         self.ui,
        #         '登录失败',
        #         resObj['msg'])
        #     return
        # if username != '123' or password !='123':
        #     QMessageBox.warning(
        #                 self.ui,
        #                  '登录失败','请检查密码')
        #     return



        self.ui.edt_password.setText('')
        self.ui.hide()

def main():
    app = QApplication([])
    SI.loginWin = Win_Login()
    SI.loginWin.ui.show()
    sys.exit(app.exec_())