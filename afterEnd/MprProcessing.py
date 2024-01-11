import os
import pathlib
import logging
import traceback
import pymsgbox as mb


def main(MprInput, MprOutput):
    # noinspection PyBroadException
    try:
        for root, dirs, files in os.walk(MprInput):
            for i in files:
                path = os.path.join(root, i)
                with open(path) as fr:
                    FilePath = MprOutput + "/" + path.split('\\')[-2] + "/" + path.split('\\')[-1]
                    FilePath2 = MprOutput + "/" + path.split('\\')[-2]
                    pathlib.Path(FilePath2).mkdir(parents=True, exist_ok=True)
                    with open(FilePath, 'w') as fw:
                        # 循环读取文件内容，逐行修改
                        for line in fr:
                            if line.startswith('<131 \\UfluBohr\\'):
                                line = line.replace('<131 \\UfluBohr\\', '<102\\ Vertical Drilling\\')

                            if line.startswith("<105 \\Konturfraesen\\"):
                                line = line + 'EN="0"\n'
                            fw.write(line)
        return True

    except:
        logging.basicConfig(filename='log.txt', level=logging.DEBUG,
                            format='%(ascTime)s - %(levelName)s - %(message)s')
        # 方案一，自己定义一个文件，自己把错误堆栈信息写入文件。
        errorFile = open('log.txt', 'a')
        errorFile.write(traceback.format_exc())
        errorFile.close()
        # yield f"{j}文件处理失败，请联系管理员"
        # tkinter.messagebox.showerror(title='错误',message=f"{j.split('.')[0]}文件处理失败，请联系管理员")
        mb.alert(f"mpr文件处理失败，请联系管理员", '报错')
        # win32api.MessageBox(0,f"{j.split('.')[0]}文件处理失败，请联系管理员","错误",win32con.MB_OK)
