import ftplib
import os
import re
import shutil
import subprocess
import sys
import textwrap
import tkinter.messagebox
import zipfile
from time import sleep

from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QMessageBox

from lib.share import SI


class FtpUtil:
    ftp = ftplib.FTP()

    def __init__(self, host, user, password):
        self.ftp = ftplib.FTP(host,user, password)      # 连接ftp服务器

    def close(self):
        self.ftp.quit()                                                    # 关闭服务器

    # 下载版本文件
    def ObtainVersion(self):
        file_list = self.ftp.nlst()

        return file_list[1]

    def download_file(self, local_file,remote_file):
        with open(local_file, 'wb') as f:
            self.ftp.retrbinary('RETR ' + remote_file, f.write)
        return True

class Update:
    def __init__(self):
        CURRENT_DIRECTORY = os.path.dirname(os.path.realpath('UpdateVersion.py'))
        filename = os.path.join(CURRENT_DIRECTORY, "foreEnd\\update.ui")
        self.ui = QUiLoader().load(filename)

        self.ui.progressBar.setRange(0, 4)
        self.ui.progressBar.setValue(0)

        self.ui.button.clicked.connect(self.updateFile)
        self.ui.textBrowser.append(f'{self.dowVersion()}')

    def updateFile(self):
        # 重置倒退进度条的进度
        self.ui.progressBar.reset()
        # 当代码执行至此，进读条为：0%
        self.ui.textBrowser.append("正在下载压缩包")
        ftp = FtpUtil('192.168.10.16', 'HT', 'admin@325')
        ftp.download_file('./EdgeBanding.zip', 'EdgeBanding.zip')
        ftp.close()
        # setValue(0):表示完成了 0/4
        self.ui.progressBar.setValue(0)
        # 睡眠一秒，方便看到进度条加载样式
        sleep(1)

        # 当代码执行至此，进读条为：25%
        self.ui.textBrowser.append("下载完成")

        # setValue(1):表示完成了 2/4
        self.ui.progressBar.setValue(1)
        # 睡眠一秒，方便看到进度条加载样式
        sleep(1)

        # 当代码执行至此，进读条为：50%
        self.ui.textBrowser.append("正在解压文件")
        self.extract_files('./EdgeBanding.zip')
        # setValue(2):表示完成了 2/4
        self.ui.progressBar.setValue(2)
        # 睡眠一秒，方便看到进度条加载样式
        sleep(1)
        self.ui.textBrowser.append("解压完成")

        # 当代码执行至此，进读条为：75%
        self.ui.textBrowser.append("正在编写bat脚本信息")
        self.make_updater_bat()
        # setValue(3):表示完成了 3/4
        self.ui.progressBar.setValue(3)
        self.ui.textBrowser.append("bat脚本编写完成")
        # 睡眠一秒，方便看到进度条加载样式
        sleep(3)


        # 当代码执行至此，进读条为：100%
        self.ui.textBrowser.append("正在替换文件")
        self.do_replace_files()
        # setValue(4):表示全部完成了 4/4
        self.ui.progressBar.setValue(4)
        self.ui.textBrowser.append("替换文件完成")
        # 睡眠一秒，方便看到进度条加载样式
        sleep(1)
        self.ui.textBrowser.append("更新完成")
        choice = QMessageBox.question(self.ui, '更新完成', '是否重启？')
        if choice == QMessageBox.Yes:
            subprocess.call("updater.bat")
            sleep(6)
            sys.exit()
        elif choice == QMessageBox.No:
            sys.exit()




    def dowVersion(self):
        ftp = FtpUtil('192.168.10.16', 'HT', 'admin@325')
        FtpVersion = ftp.ObtainVersion()

        pattern = re.compile("Version_(.*?).txt")
        FtpVersion = pattern.findall(FtpVersion)[0]
        ftp.download_file(f'Version_{FtpVersion}.txt',f'Version_{FtpVersion}.txt')
        ftp.close()
        with open(f'Version_{FtpVersion}.txt','r') as file:
            list =file.read()
        print(list)
        return str(list)





    def extract_files(self,pacth):
        file = pacth.split('.')[1]
        file = file.split('/')[1]
        print(file)
        if not os.path.exists(file):
            os.mkdir(file)
        with zipfile.ZipFile(pacth) as zf:
            zf.extractall(path=file)  # 解压目录
        print(f'extract file: {file}')

    def make_updater_bat(self):
        pattern = re.compile("Version_(.*?).txt")
        files = os.listdir('.')
        LocalVersion = [file for file in files if 'Version' in file][0]
        LocalVersion = pattern.findall(LocalVersion)[0]
        app_name = 'EdgeBanding.exe'
        app_file = os.path.join('EdgeBanding', app_name)
        print(app_file)
        with open('updater.bat', 'w', encoding='utf-8') as updater:
            updater.write(textwrap.dedent(f'''\
            @echo off
            echo 正在更新[EdgeBanding.exe]最新版本，请勿关闭窗口...
            ping -n 2 127.0.0.1 > nul
            echo 正在复制[./EdgeBanding\EdgeBanding.exe]，请勿关闭窗口...
            del EdgeBanding.exe
            copy EdgeBanding\EdgeBanding.exe . /Y
            del EdgeBanding.zip
            rd /S /Q EdgeBanding
            del Version_{LocalVersion}.txt
            echo 更新完成，等待自动启动EdgeBanding.exe...
            ping -n 3 127.0.0.1 > nul
            start EdgeBanding.exe
            pause
                '''))
            updater.flush()

    def do_replace_files(self):
        extract_dir = 'EdgeBanding'
        for file in os.listdir(extract_dir):
            if not file.endswith('EdgeBanding.exe') and not file.endswith('EdgeBanding') and not file.endswith('Cryptodome') and not file.endswith('numpy')\
                    and not file.endswith('pandas') and not file.endswith('.dll') and not file.endswith('pyexpat.pyd') and not file.endswith('pymssql') and not file.endswith('PyQt5')\
                    and not file.endswith('PySide2') and not file.endswith('select.pyd') and not file.endswith('shiboken2') and not file.endswith('unicodedata.pyd') and not file.endswith('_bz2.pyd'):
                try:
                    print(f'替换文件：{file}')
                    self.copy_files(os.path.join(extract_dir, file), '.')
                except BaseException as e:
                    if os.path.isdir(file):
                        tkinter.messagebox.showwarning(title='错误',
                                                       message=f'文件夹[{file}]有文件正在使用,更新失败,请关闭文件后重试')
                    else:
                        tkinter.messagebox.showwarning(title='错误',
                                                       message=f'文件[{file}]正在使用,更新失败,请关闭文件后重试')
                    print(f'替换文件错误{e}')
                    raise e

    def copy_files(self,src_file, dest_dir):
        file_name = src_file.split(os.sep)[-1]
        if os.path.isfile(src_file):
            shutil.copyfile(src_file, os.path.join(dest_dir, file_name))
        else:
            dest_path = os.path.join(dest_dir, file_name)
            if os.path.exists(dest_path):
                shutil.rmtree(dest_path)
            shutil.copytree(src_file, os.path.join(dest_dir, file_name))


def ObtainVersion():
    ftp = FtpUtil('192.168.10.16','HT','admin@325')
    FtpVersion = ftp.ObtainVersion()

    pattern = re.compile("Version_(.*?).txt")
    FtpVersion = pattern.findall(FtpVersion)[0]
    ftp.close()

    files = os.listdir('./')
    LocalVersion = [file for file in files if 'Version' in file][0]
    LocalVersion = pattern.findall(LocalVersion)[0]

    if FtpVersion == LocalVersion:
        return False
    else:
        return True

def main():
    SI.PatchWin = Update()
    SI.PatchWin.ui.show()





