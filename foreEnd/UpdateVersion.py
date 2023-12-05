import ftplib
import os
import re


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

def ObtainVersion():
# if __name__ == '__main__':
    ftp = FtpUtil('192.168.10.16','HT','admin@325')
    FtpVersion = ftp.ObtainVersion()

    pattern = re.compile("Version_(.*?).txt")
    FtpVersion = pattern.findall(FtpVersion)[0]
    ftp.close()

    files = os.listdir('./')
    LocalVersion = [file for file in files if 'Version' in file][0]
    LocalVersion = pattern.findall(LocalVersion)[0]

    print(FtpVersion)
    print(LocalVersion)

    if FtpVersion == LocalVersion:
        return False
    else:
        return True

