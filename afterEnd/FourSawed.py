import os
import pandas as pd
import logging
import traceback
import pymsgbox as mb


def main(SawCsvOutput):
    try:
        CsvList = os.listdir(SawCsvOutput)

        for j in CsvList:
            fullPath = SawCsvOutput + "/" + j  # 完整路径
            csvData = pd.read_csv(fullPath, encoding='gb18030')
            df = pd.DataFrame(csvData)
            # print(csvData)
            name = df.loc[:, '输出名称']
            for index, i in enumerate(name):
                if '拼色' not in i and df.loc[index, '厚'] == 5:
                    df.loc[index, '加工程序'] = f'{df.loc[index, "长"]- 10} * {df.loc[index, "宽"]- 10}'
            df.to_csv(fullPath, index=False, encoding='gb18030')
            yield j

    except:
        logging.basicConfig(filename='log.txt', level=logging.DEBUG,
                            format='%(ascTime)s - %(levelName)s - %(message)s')
        # 方案一，自己定义一个文件，自己把错误堆栈信息写入文件。
        errorFile = open('log.txt', 'a')
        errorFile.write(traceback.format_exc())
        errorFile.close()
        yield f"文件处理失败，请联系管理员"
        # tkinter.messagebox.showerror(title='错误',message=f"{j.split('.')[0]}文件处理失败，请联系管理员")
        mb.alert(f"文件处理失败，请联系管理员", '报错')
