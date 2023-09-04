'''
1.读取时新增读取板材材质用数组保存，同时工件信息中的长宽也要读取可以用二维数组保存【【长，宽】,【长，宽】】，这两者都要按顺序保存，便于后面按照显示的数值读取,读取使用的板件位置和工件位置时，用字典保存
2.添加字段实际发板数，利用率为工件总面积除以实际发板数的总面积(dist)

'''
import sys
import tkinter as tk
import os
import re
import datetime as dt
from afterEnd.page.sheetAvailability import My_sheet
from afterEnd.page.style import style
from afterEnd.page.Logger import Logger

# 判断为文件还是文件夹
def panduan(fullPath,listPath2):
    # 首先遍历当前目录所有文件及文件夹
    file_list = os.listdir(fullPath)
    # 循环判断每个元素是否是文件夹还是文件，是文件夹的话，递归
    for file in file_list:
        # 利用os.path.join()方法取得路径全名，并存入cur_path变量，否则每次只能遍历一层目录
        cur_path = os.path.join(fullPath, file)
        # 判断是否是文件夹
        if os.path.isdir(cur_path):
            panduan(cur_path)
        else:
            listPath2.append(cur_path)

# 读取数据
def readData(Folderpath,listPath2,listPath3,numListDay,datalist2,dist2,numList):
    panduan(Folderpath)
    com = re.compile('.*.saw')
    for i in listPath2:
        if com.match(i.split('\\')[-1]):
            listPath3.append(i)
    for i in listPath3:  # 遍历所有saw文件
        print("处理的文件："+i)
        numListDay.append(i.split("\\")[2])
        # 打开文件
        dist = {}
        datalist = []
        dataMaterial = {}
        dataMaterialNum = {}
        workpieceList = []
        sumAreDict={}
        ls=[]
        with open(i, "r",  encoding='gb18030', errors='ignore') as f:
            for line in f:
                list2 = []
                list2.append(i.split('\\')[-1])
                # 按照逗号切割
                stringTOOL = line.split(",")
                # 判断每一行第一个，是什么
                if stringTOOL[0] == "BRD2":
                    # 当存在余料时，材料不唯一，所以键用材料名称作为键
                    dataMaterialList = []
                    primaryMaterialName = stringTOOL[1]
                    Material =stringTOOL[7]
                    MaterialLength = stringTOOL[2]
                    MaterialWidth = stringTOOL[3]
                    dataMaterialList.append(Material)
                    dataMaterialList.append(MaterialLength)
                    dataMaterialList.append(MaterialWidth)
                    dataMaterialList.append(primaryMaterialName)
                    dataMaterial[primaryMaterialName.upper()] = dataMaterialList
                    # print(dataMaterial)
                elif stringTOOL[0] == "PNL2":
                    workpieceListOne=[]
                    workpieceLenth = stringTOOL[3]
                    workpieceWidth = stringTOOL[4]
                    workpieceListOne.append(workpieceLenth)
                    workpieceListOne.append(workpieceWidth)
                    workpieceList.append(workpieceListOne)
                elif stringTOOL[0] == "MAT2":
                    # 板材名称
                    MaterialName = stringTOOL[1]
                    list2.append(MaterialName)
                    # 工件个数
                    workpiecesNum = stringTOOL[47]
                    list2.append(int(workpiecesNum))
                    # 工件面积
                    workpiecesArea = stringTOOL[48]
                    list2.append(float(workpiecesArea))
                    # 板件数
                    SurplusNum = stringTOOL[50]
                    # print(SurplusNum)
                    list2.append(int(SurplusNum))
                    # 板件总面积
                    SurplusArea = stringTOOL[51]
                    list2.append(float(SurplusArea))
                    # 这个批次这种板材的利用率
                    Utilization = (float(workpiecesArea) / float(SurplusArea))
                    result = '{:.2%}'.format(Utilization)
                    list2.append(result)
                    # 保存的都是同一个文件，会清空
                    datalist.append(list2)
                    # 汇总，为全局变量，不会清空
                    datalist2.append(list2)
                elif stringTOOL[0] =="PTN2":
                    # print(stringTOOL[1])
                    # 根据PTN2中的数字得到对应的原材料名称以及长宽
                    key = list(dataMaterial.keys())[int(stringTOOL[1]) - 1]
                    areaList = dataMaterial[key]
                    # print(dataMaterial)
                    # print(areaList)
                    if areaList[3].upper() in dataMaterialNum:
                        dataMaterialNum[areaList[3].upper()] +=1
                    else:
                        dataMaterialNum[areaList[3].upper()] =1

                elif stringTOOL[0] == "PTNR":
                    # num = stringTOOL[1]
                    num =stringTOOL[1:len(stringTOOL)]
                    # print(stringTOOL[1:len(stringTOOL)])
                    # 使用推导式，直接循环删除会造成下标改变，漏筛元素
                    num = [x for x in num if not x.startswith('X')]
                    # 将数组转变为字符串形式
                    s = ",".join(num)
                    # 正则表达式得到各各工件的下标
                    pattern = re.compile(r'\d+')
                    m=pattern.findall(s)

                    # if not areaList[3].startswith('X'):
                    # 计算利用率,工件下标需要从0开始，工件显示为1，工件在数组中的储存位置为0
                    Are = float(areaList[1]) * float(areaList[2])
                    workAre = 0.0
                    for p in m:
                        workAre = workAre + (
                                float(workpieceList[int(p) - 1][0]) * float(workpieceList[int(p) - 1][1]))
                    # print(workAre)
                    ActualUtilization = (float(workAre) / float(Are))
                    ActualResult = '{:.2%}'.format(ActualUtilization)

            dist[i.split('\\')[-1]] = datalist



        # 添加到汇总字典中
        dist2.update(dist)
        numList.append(len(datalist))
    yield i


# 写入数据到单元格
def writeCell(fileName,test,styleCell,datalist2):
    # 设置冻结窗口
    test.set_panes_frozen(True,1,2)
    col = ('排产时间','批次号','材料名称', '按花色工件数', '按花色工件总面积', '按花色使用板件数', '按花色使用板件总面积','批次按花色板件利用率','批次利用率','日利用率')
    test.write_row(0, 0, col,styleCell.style1())
    test.write_rows(1,1,datalist2,styleCell.style0())
    # test.save("D:/{}余料利用率.xls".format(time))


# 合并单元格
def mergeCell(dist2,test):
    z=0
    i=0
    for row1,data in enumerate(dist2.values()):
        listKey=list(dist2.keys())
        num = len(data)
        # print(num)
        # print(data)
        for row2,data2 in enumerate(data):
            i = i + 1
        # test.write_merge(i - num + 1, i, 1, 1, listKey[z])
        test.write_merge(i - num + 1, i, 8, 8,data[len(data)-1][7])
        test.write_merge(i - num + 1, i, 8, 8,data[len(data) - 1][6])
        # test.write_col(i - num + 1, 13, data[len(data) - 1][10])
        # test.write_merge(1,i,0,0,time)
        # test.write_merge(i - num + 1, i, 8, 8, resultSum, style.style2(self=""))
        # test.write_merge(1,i,8, 8, test.MaterialresultSum, style.style2(self=""))
        z+=1
     # msgbox.showinfo("运行失败", "运行失败第{0}个文件错误".format())

# 计算批次利用率
def write_pici(dist2,distGongJian,distGongJianSum):
    disNum=[]
    sumNum=0
    for i in dist2:
        gongjian = 0.0
        gongjianSum = 0
        for j in dist2[i]:
            gongjian = gongjian+j[3]
            gongjianSum = gongjianSum + j[5]
            Utilization = (float(gongjian)/float(gongjianSum))
            sumNum = '{:.2%}'.format(Utilization)
        j.append(sumNum)

            # print(j)

        # distGongJian.append(gongjian)
        # distGongJianSum.append(gongjianSum)



    # print(dist)
    # a= np.array(distGongJian)
    # b = np.array(distGongJianSum)
    # disNum = (a/b)*100+'%'

    # return disNum
    # print(distGongJian)
    # print(distGongJianSum)

# 计算实际花色利用率,
def ActualUtilization(dist2):
    disNum = []
    PiUtilization = 0
    for i in dist2.values():
        for n in i:
            if float(n[8]) ==0.0:
                n.append("100%")
            else:
                sumNum = '{:.2%}'.format(float(n[3])/float(n[8]))
                n.append(sumNum)

    for i in dist2:
        # print(dist[i])
        gongjian = 0.0
        gongjianSum = 0
        for j in dist2[i]:
            gongjian = gongjian + j[3]
            gongjianSum = gongjianSum + j[8]
        Utilization = (float(gongjian) / float(gongjianSum))
        PiUtilization = '{:.2%}'.format(Utilization)
        j.append(PiUtilization)

def main():

    sys.stderr = Logger('a.log_file', sys.stderr)
    # 1. 读取文件
    # 设置弹窗
    FileNum = 1
    root = tk.Tk()
    root.withdraw()
    # 创建表格对象
    test = My_sheet()
    # 创建样式对象
    styleCell = style()
    # 设置弹窗标题
    # 输入路径
    Folderpath = "D:/处理文件"
    print("处理文件存放路径"+Folderpath)
    # 保存数据到字典，用于单元格合并时使用
    dist2 = {}
    # 保存数据到数组，用于写入数据
    datalist2=[]
    listPath2 = []
    listPath3 = []
    numList = []
    numListDay = []
    distGongJian=[]
    distGongJianSum = []
    # 获取当前时间
    # time = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    time = dt.datetime.now().strftime("%m")
    fileName= Folderpath+"/"+time+"月余料利用率.xls"

    # 读取数据
    readData(Folderpath,listPath2,listPath3,numListDay,datalist2,dist2,numList)

    # 计算批次利用率，所得结果保存在dist2中的每一个key对应的value的最后一个数组中
    write_pici(dist2,distGongJian,distGongJianSum)
    # ActualUtilization(dist2)
    print(datalist2)
    writeCell(fileName,test,styleCell,datalist2)
    print(dist2)
    # test.write_col(1, 10,listPath2, styleCell.style0())
    print("写入数据成功")
    # print(dist2)
    # 合并单元格
    mergeCell(dist2,test)
    print("合并单元格成功")
    test.save(fileName)
    print(numListDay)
    test.Day(numList, numListDay, fileName)
    print("统计每天利用率成功")
    # 保存数据
    save = test.save(fileName)