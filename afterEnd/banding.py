
# Author: Nonove. nonove[at]msn[dot]com
# XML simple operation Examples and functions
# encoding = gbk
# @刘益孝 益孝xml里，如果工具是，2.5T型刀型刀（暂定这个名字），的加工信息就删除掉，做一个这样的工具。
import sys

import tkinter.messagebox as msgbox


from xml.dom import minidom
import pymsgbox as mb
import logging
import traceback
import os


from PySide2.QtWidgets import QMessageBox

from afterEnd.page.My_sheet import My_sheet
from afterEnd.page.styleT import style
import lib.Sql as sqlUnit

class banding:
    def __init__(self):
        self.changeWwidth



    def changeWwidth(self,Folderpath,Folderpath2,test,styleCell,SqlUnit):
        # 1. 读取xml文件
        # 设置弹窗
        # root = tk.Tk()
        # root.withdraw()
        # 设置弹窗标题
        # 输入路径
        # Folderpath = filedialog.askdirectory(title="选择输入目录")
        # Folderpath = "D:/原文件"
        # 输出路径
        # Folderpath2 = filedialog.askdirectory(title="选择输出目录")
        # Folderpath2 = "D:/处理后的文件"
        BandingProcessing,BandingCode = SqlUnit.selectBandingProcessing()
        ProcessingDict={}
        for t in BandingProcessing:
            ProcessingDict[t[1]]=t[0]
        BandingCodingDict={}
        # BandingCode= SqlUnit.selectBandingCode()
        for t in BandingCode:
            BandingCodingDict[t[1]]=t[2]
        print(BandingCodingDict)
        # 打开文件夹得到文件夹下的所有文件
        list=[]
        # window = tkinter.Tk()
        # window.withdraw()
        xml_list = os.listdir(Folderpath)
        for j in xml_list:  # 遍历所有xml文件
            try:
                fullPath = Folderpath + "/" + j  # 完整路径
                list.append(j)
                # list.append(i)
                # 打开文件
                xmldoc = minidom.parse(fullPath)
                #找到Machining节点（标签）
                Panelnodes = xmldoc.getElementsByTagName("Panel")
                PanelList=[]
                for panel in Panelnodes:
                    WorkpieceRotation=0
                    for i in range(1,3):
                        PanelBasics = []
                        # 二维码
                        QRCode = panel.getAttribute('ID')
                        PanelBasics.append(QRCode)
                        # 备注1
                        RemarksOne = panel.getAttribute('BatchNo')
                        PanelBasics.append(RemarksOne)
                        # 备注2
                        RemarksTwo = panel.getAttribute('Info4')
                        PanelBasics.append(RemarksTwo)
                        # 备注3
                        RemarksThree = panel.getAttribute('ContractNo')
                        PanelBasics.append(RemarksThree)
                        # 左机预铣
                        # 右机预铣
                        # 完工长度
                        FinishedLength = float(panel.getAttribute('Length'))
                        PanelBasics.append(FinishedLength)
                        # 完工宽度
                        FinishedWidth = float(panel.getAttribute('Width'))
                        PanelBasics.append(FinishedWidth)
                        # 完工厚度
                        FinishedThickness = int(panel.getAttribute('Thickness'))
                        PanelBasics.append(FinishedThickness)
                        # 数量
                        Quantity = int(panel.getAttribute('Qty'))
                        PanelBasics.append(Quantity)
                        # 进给次序
                        if i == 1:
                            Order = 1
                            PanelBasics.append(Order)
                        elif i == 2:
                            Order = 2
                            PanelBasics.append(Order)
                        # 工件旋转
                        if i == 1:
                            if FinishedLength>=FinishedWidth:
                                WorkpieceRotation = 0
                                PanelBasics.append(WorkpieceRotation)
                            elif FinishedLength<FinishedWidth:
                                WorkpieceRotation = 1
                                PanelBasics.append(WorkpieceRotation)
                        elif i== 2:
                            if WorkpieceRotation == 0:
                                WorkpieceRotation = 1
                                PanelBasics.append(WorkpieceRotation)
                            elif WorkpieceRotation == 1:
                                WorkpieceRotation = 0
                                PanelBasics.append(WorkpieceRotation)

                        # 浮动铣刀
                        FloatingCutter = 0
                        PanelBasics.append(FloatingCutter)
                        # 速度
                        Speed = 25
                        PanelBasics.append(Speed)
                        # 左机封边带编码
                        if i ==1:
                            LeftBandingCodeOne = panel.getAttribute('EBL1')
                            LeftBandingCodeOne=BandingCodingDict[LeftBandingCodeOne.strip()]
                            print(LeftBandingCodeOne)
                            PanelBasics.append(LeftBandingCodeOne)
                        elif i ==2:
                            LeftBandingCodeTwo = panel.getAttribute('EBW1')
                            LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
                            print(LeftBandingCodeTwo)
                            PanelBasics.append(LeftBandingCodeTwo)
                         # 左机加工编码
                        if i==1:
                            EBL1 = panel.getAttribute('EBL1')
                            str=''
                            if EBL1 == "":
                                str='无封边'
                            else:
                                if '▲▲▲▲' in EBL1:
                                    str = '1.0mm封边'
                                elif '△△△△' in EBL1:
                                    str = '0.5mm封边'
                                Length = float(panel.getAttribute('Length'))
                                Width = float(panel.getAttribute('Width'))
                                Machines = panel.getElementsByTagName("Machines")
                                for Machining in Machines:
                                    macin = Machining.getElementsByTagName("Machining")
                                    for m in macin:
                                        Type = int(m.getAttribute('Type'))
                                        Face = int(m.getAttribute('Face'))
                                        X = float(m.getAttribute('X'))
                                        Y = float(m.getAttribute('Y'))

                                        if Type == 4:
                                            EndX = float(m.getAttribute('EndX'))
                                            EndY = float(m.getAttribute('EndY'))
                                            widthTool = float(m.getAttribute('Width'))
                                            ToolOffset = m.getAttribute('ToolOffset')
                                            # 判断槽宽
                                            if widthTool == 13:
                                                # 判断通槽
                                                if (abs(EndX - X) == Length and abs(EndY - Y) == 0):
                                                    if Face == 5:
                                                        # 判断走刀方向
                                                        if X < EndX or Y < EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode',"0")
                                                            elif ToolOffset == '右' and (X == 7 or Length - X - widthTool== 7 or Y - widthTool == 7 or Width - Y== 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X - widthTool== 7 or Length - X== 7 or Y== 7 or Width - Y - widthTool == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                        elif X>EndX or Y>EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")

                                                    elif Face == 6:
                                                        # 判断走刀方向
                                                        if X < EndX or Y < EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                        elif X > EndX or Y > EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")



                                                else:
                                                    continue

                                            else:
                                                continue
                                        else:
                                            continue
                                str +='+跟踪'
                            if '6面槽+5面槽' in str :
                                str = str.replace('6面槽+5面槽','5&6面槽')
                            elif '5面槽+6面槽' in str :
                                str = str.replace('5面槽+6面槽', '5&6面槽')
                            print(str)


                        elif i==2:
                            EBW1 = panel.getAttribute('EBW1')
                            str = ''
                            if EBW1 == "":
                                str = '无封边'
                            else:
                                if '▲▲▲▲' in EBW1:
                                    str = '1.0mm封边'
                                elif '△△△△' in EBW1:
                                    str = '0.5mm封边'
                                Length = float(panel.getAttribute('Length'))
                                Width = float(panel.getAttribute('Width'))
                                Machines = panel.getElementsByTagName("Machines")
                                for Machining in Machines:
                                    macin = Machining.getElementsByTagName("Machining")
                                    for m in macin:
                                        Type = int(m.getAttribute('Type'))
                                        Face = int(m.getAttribute('Face'))
                                        X = float(m.getAttribute('X'))
                                        Y = float(m.getAttribute('Y'))

                                        if Type == 4:
                                            EndX = float(m.getAttribute('EndX'))
                                            EndY = float(m.getAttribute('EndY'))
                                            widthTool = float(m.getAttribute('Width'))
                                            ToolOffset = m.getAttribute('ToolOffset')
                                            # 判断槽宽
                                            if widthTool == 13:
                                                # 判断通槽
                                                if (abs(EndX - X) == 0 and abs(EndY - Y) == Width):
                                                    if Face == 5:
                                                        # 判断走刀方向
                                                        if X < EndX or Y < EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                        elif X > EndX or Y > EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                str = str + "+5面槽"
                                                                m.setAttribute('IsGenCode', "0")

                                                    elif Face == 6:
                                                        # 判断走刀方向
                                                        if X < EndX or Y < EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                        elif X > EndX or Y > EndY:
                                                            if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")
                                                            elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                str = str + "+6面槽"
                                                                m.setAttribute('IsGenCode', "0")



                                                else:
                                                    continue

                                            else:
                                                continue
                                        else:
                                            continue
                                str += '+跟踪'
                            if '6面槽+5面槽' in str:
                                str = str.replace('6面槽+5面槽', '5&6面槽')
                            elif '5面槽+6面槽' in str:
                                str = str.replace('5面槽+6面槽', '5&6面槽')
                            print(str)



                        LeftProcessingCode=ProcessingDict[str]
                        print(LeftProcessingCode)
                        PanelBasics.append(LeftProcessingCode)

                        # 右机封边带编码
                        if i == 1:
                            RightBandingCodeOne = panel.getAttribute('EBL2')
                            RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
                            print(RightBandingCodeOne)
                            PanelBasics.append(RightBandingCodeOne)
                        elif i == 2:
                            RigthBandingCodeTwo = panel.getAttribute('EBW2')
                            RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
                            print(RigthBandingCodeTwo)
                            PanelBasics.append(RigthBandingCodeTwo)
                        # 右机加工编码
                        if i == 1:
                            EBL2 = panel.getAttribute('EBL2')
                            str = ''
                            if EBL2 == "":
                                str = '无封边'
                            else:
                                if '▲▲▲▲' in EBL2:
                                    str = '1.0mm封边'
                                elif '△△△△' in EBL2:
                                    str = '0.5mm封边'
                                str += '+跟踪'
                            print(str)
                        elif i == 2:
                            EBW2 = panel.getAttribute('EBW2')
                            str = ''
                            if EBW2 == "":
                                str = '无封边'
                            else:
                                if '▲▲▲▲' in EBW2:
                                    str = '1.0mm封边'
                                elif '△△△△' in EBW2:
                                    str = '0.5mm封边'

                                str += '+跟踪'
                            print(str)
                        RightProcessingCode =ProcessingDict[str]
                        print(RightProcessingCode)
                        PanelBasics.append(RightProcessingCode)



                        #左右机预铣
                        EdgeGroups = panel.getElementsByTagName("EdgeGroup")
                        for EdgeGroup in EdgeGroups:
                            Edges = EdgeGroup.getElementsByTagName("Edge")
                            for Edge in Edges:
                                Faces = Edge.getAttribute('Face')
                                if i == 1:
                                    if int(Faces) == 1:
                                        Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(4,Pre_MillingLeft)
                                    if int(Faces) == 2:
                                        Pre_MillingRigth =float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(5,Pre_MillingRigth)
                                elif i ==2:
                                    if int(Faces) == 3:
                                        Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(4,Pre_MillingLeft)
                                    if int(Faces) == 4:
                                        Pre_MillingRigth =float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(5,Pre_MillingRigth)
                        PanelList.append(PanelBasics)
                col = ('二维码', '备注1', '备注2', '备注3', '左机预铣', '右机预铣', '完工长度', '完工宽度', '完工厚度', '数量','进给次序','工件旋转','浮动铣刀','速度','左机封边带编码','左机加工编码','右机封边带编码','右机加工编码')
                # 写入标题
                test.write_row(0, 0, col,styleCell.style1())
                # 写入数据
                test.write_rows(1, 0, PanelList,styleCell.style0())
                # SqlUnit.InserUnit(PanelList)
                # 保存按原文件命名  "D:/{}月余料利用率.xls".format(time)
                path = Folderpath2+'/'+j.split('.')[0]+'.xls'
                test.save(path)
                with open(Folderpath2 + "/" + fullPath.split('/')[-1]+"封边后", "w", encoding="UTF-8") as fs:
                    fs.write(xmldoc.toxml())
                    fs.close()
                yield j
            except:
                logging.basicConfig(filename='log.txt', level=logging.DEBUG,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
                # 方案一，自己定义一个文件，自己把错误堆栈信息写入文件。
                errorFile = open('log.txt', 'a')
                errorFile.write(traceback.format_exc())
                errorFile.close()
                yield f"{j}文件处理失败，请联系管理员"
                # tkinter.messagebox.showerror(title='错误',message=f"{j.split('.')[0]}文件处理失败，请联系管理员")
                mb.alert(f"{j}文件处理失败，请联系管理员",'报错')
                # win32api.MessageBox(0,f"{j.split('.')[0]}文件处理失败，请联系管理员","错误",win32con.MB_OK)
def main(Folderpath,Folderpath2):
    # 创建表格对象
    test = My_sheet()
    # 创建样式对象
    styleCell = style()
    #
    ban=banding()
    SqlUnit = sqlUnit.main()
    file = ban.changeWwidth(Folderpath,Folderpath2,test,styleCell,SqlUnit)
    # sqlUnit.SqlUnit.InserUnit(PaneListNUm)
    return file,SqlUnit


