import pandas as pd
import os
import tkinter.messagebox
from tkinter import *


# 指定文件夹路径
folder_path = r'D:\Desktop\2'  # Excle是我自己创建的文件夹，将需要合并的Excle文件放到这里
# 获取文件夹中所有的Excel文件名
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xls')]  # 注意要合并的文件格式，填写你需要的
print(excel_files)
# 创建一个空的DataFrame，用于保存所有数据
combined_df = pd.DataFrame()

# 循环读取每个Excel文件，并将数据合并到总DataFrame中
i=0
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    print(file)
    i+=1
    df = pd.read_excel(file_path)  # 跳过前三行，

    col_name = df.columns.tolist()
    col_name.insert(19, '文件名')  # 设置新增列的位置和名称
    wb = df.reindex(columns=col_name)
    df['18'] = file  # 计算方式，根据自己设定
    combined_df = pd.concat([combined_df, df], ignore_index=True)
    os.remove(file_path)
    print(i)



# 将合并后的DataFrame写入新的Excel文件
output_path = r'D:\Desktop\1\combined.csv'
combined_df.to_csv(output_path, index=False,encoding='gbk')

tkinter.Tk().withdraw();
tkinter.messagebox.showinfo(title= '成功',message='处理成功');