import os
import re
import subprocess
import tkinter as tk
from tkinter import messagebox

import requests
from PySide2.QtCore import QSettings

import Login.login as login
import sys

sys.path.append("path")

def LocalVersion():
    url = 'http://192.168.10.16:220/Version.txt'
    file = 'Version.txt'
    response = requests.get(url)
    with open(file, 'wb') as f:
        f.write(response.content)
    # 打开文本文件
    file = open(f"Version.txt", "r")
    # 读取第一行
    first_line = file.readline()
    # 关闭文件
    file.close()
    os.remove('Version.txt')
    # 输出第一行内容
    return first_line

def ResponseVersion():
    settings = QSettings("config.ini", QSettings.IniFormat)
    version = settings.value("version")
    return version

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    response = LocalVersion()
    response =response.replace("\n", "")
    local = ResponseVersion()
    if response == local:
        login.main()
    else:
        result = messagebox.askyesno("更新", "检测到新版本是否更新？")
        if result:
            subprocess.Popen("main\\main.exe")
            sys.exit()
        else:
            login.main()

