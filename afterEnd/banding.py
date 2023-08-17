
# Author: Nonove. nonove[at]msn[dot]com
# XML simple operation Examples and functions
# encoding = gbk

from xml.dom import minidom
import pymsgbox as mb
import logging
import traceback
import os
import pandas as pd

from afterEnd.page.My_sheet import My_sheet
from afterEnd.page.styleT import style
import lib.Sql as sqlUnit

class banding:
    def __init__(self):
        self.changeWwidth

    # 修改csv堆垛
    def SawCsv(self,CsvIdFile,path,file):
        fullPath = path + "/" + file.split(".")[0] +'.csv'  # 完整路径
        print(fullPath)
        csvData = pd.read_csv(fullPath,encoding='gb18030')
        df = pd.DataFrame(csvData)
        # print(csvData)
        code = df.loc[:,'条形码']
        for index,i in enumerate(code):
            if i.split('N')[1] in CsvIdFile:
                num = df.loc[index,'堆垛']
                if num == 'AC':
                    # print(f'堆垛为{num,index}')
                    df.loc[index,'堆垛'] = 'A'
                    df.loc[index,'Info6'] ='A'
        df.to_csv(fullPath, index=False,encoding='gb18030')

    def changeWwidth(self,Folderpath,Folderpath2,Folderpath3,Folderpath4,test,styleCell,SqlUnit,IdList):
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
        # Folderpath3="./ProcessingFile"
        BandingProcessing,BandingCode = SqlUnit.selectBandingProcessing()
        ProcessingDict={}
        for t in BandingProcessing:
            ProcessingDict[t[1]]=t[0]
        BandingCodingDict={}
        # BandingCode= SqlUnit.selectBandingCode()
        for t in BandingCode:
            BandingCodingDict[t[1]]=t[2]
        # 打开文件夹得到文件夹下的所有文件
        listPacth=[]
        xml_list = os.listdir(Folderpath)
        for j in xml_list:  # 遍历所有xml文件
            try:
                fullPath = Folderpath + "/" + j  # 完整路径
                listPacth.append(j)
                # 打开文件
                xmldoc = minidom.parse(fullPath)
                #找到Machining节点（标签）
                Panelnodes = xmldoc.getElementsByTagName("Panel")
                PanelList=[]
                IdFile=[]
                CsvIdFile = []
                for panel in Panelnodes:
                    WorkpieceRotation=0
                    for i in range(1,3):
                        PanelBasics = []

                        # 1.二维码，参考
                        # QRCode=self.QRCode(panel,PanelBasics)
                        QRCode = panel.getAttribute('ID')
                        PanelBasics.append(QRCode)

                        # 2.备注1，装饰
                        RemarksOne = panel.getAttribute('BatchNo')
                        PanelBasics.append(RemarksOne)

                        # 3.备注2，物料流
                        # RemarksTwo = panel.getAttribute('Info4')
                        PanelBasics.append('')

                        # 4.备注3，批号
                        RemarksTwo = panel.getAttribute('Info4')
                        PanelBasics.append(RemarksTwo)

                        # 5.备注4，客户编号
                        RemarksThree = panel.getAttribute('ContractNo')
                        PanelBasics.append(RemarksThree)


                        # 6.完工长度
                        FinishedLength = float(panel.getAttribute('Length'))
                        if FinishedLength<250:
                            break
                        PanelBasics.append(FinishedLength)

                        # 7.完工宽度
                        FinishedWidth = float(panel.getAttribute('Width'))
                        if FinishedWidth<250:
                            break
                        PanelBasics.append(FinishedWidth)

                        # 8.完工厚度
                        FinishedThickness = float(panel.getAttribute('Thickness'))
                        if FinishedThickness !=18:
                            break
                        PanelBasics.append(FinishedThickness)

                        # 9.数量
                        Quantity = int(panel.getAttribute('Qty'))
                        PanelBasics.append(Quantity)

                        # 10.进给次序
                        if i == 1:
                            Order = 1
                            PanelBasics.append(int(Order))
                        elif i == 2:
                            Order = 2
                            PanelBasics.append(int(Order))



                        # 浮动铣刀
                        # FloatingCutter = 0
                        # PanelBasics.append(FloatingCutter)

                        # 11.速度
                        Speed = 25
                        PanelBasics.append(float(Speed))

                        # 12.工件旋转
                        # 设置变量宽大于长时判断标准槽时，长宽，坐标也要对应变化
                        bianlian =0;
                        if i == 1:
                            # if FinishedLength >= FinishedWidth:
                            #     WorkpieceRotation = 1
                            #
                            #     PanelBasics.append(int(WorkpieceRotation))
                            # elif FinishedLength < FinishedWidth:
                            #     WorkpieceRotation = 0
                            #     bianlian = 1
                            WorkpieceRotation=1
                            bianlian = 1
                            PanelBasics.append(int(WorkpieceRotation))
                        elif i == 2:
                            if WorkpieceRotation == 1:
                                WorkpieceRotation = 0

                                PanelBasics.append(int(WorkpieceRotation))
                            elif WorkpieceRotation == 0:
                                WorkpieceRotation = 1
                                # bianlian = 1
                                PanelBasics.append(int(WorkpieceRotation))
                        # if i == 1:
                        #     Order = 0
                        #     PanelBasics.append(int(Order))
                        # elif i == 2:
                        #     Order = 1
                        #     PanelBasics.append(int(Order))


                        # 13.左机封边带编码
                        if i ==2:
                            # if bianlian ==1:
                            #     LeftBandingCodeOne = panel.getAttribute('EBW1')
                            # else:
                            #     LeftBandingCodeOne = panel.getAttribute('EBL1')
                            LeftBandingCodeOne = panel.getAttribute('EBL1')
                            if LeftBandingCodeOne =="":
                                LeftBandingCodeOne="无封边"
                                LeftBandingCodeOne = BandingCodingDict[LeftBandingCodeOne.strip()]
                            else:
                                LeftBandingCodeOne=BandingCodingDict[LeftBandingCodeOne.strip()]
                            PanelBasics.append(LeftBandingCodeOne)
                        elif i ==1:
                            # if bianlian ==1:
                            #     LeftBandingCodeTwo = panel.getAttribute('EBL1')
                            # else:
                            #     LeftBandingCodeTwo = panel.getAttribute('EBW1')
                            LeftBandingCodeTwo = panel.getAttribute('EBW1')
                            if LeftBandingCodeTwo == "":
                                LeftBandingCodeTwo = "无封边"
                                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
                            else:
                                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
                            # LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
                            PanelBasics.append(LeftBandingCodeTwo)

                        # 14.左机加工编码
                        # 当开槽时
                        # 控制厚边右机封
                        t = 0
                        if i == 2:
                            # if bianlian==1:
                            #     EBL1 = panel.getAttribute('EBW1')
                            #     EBL2 = panel.getAttribute('EBW2')
                            # elif bianlian==0:
                            #     EBL1 = panel.getAttribute('EBL1')
                            #     EBL2 = panel.getAttribute('EBL2')
                            EBL1 = panel.getAttribute('EBL1')
                            EBL2 = panel.getAttribute('EBL2')
                            strname=''
                            if EBL1 == "":
                                strname='无封边'
                            else:
                                if '▲▲▲▲' in EBL1:
                                    strname = '1.0mm封边'
                                elif '△△△△' in EBL1:
                                    strname = '0.5mm封边'
                                # Length = float(panel.getAttribute('Length'))
                                # Width = float(panel.getAttribute('Width'))
                                Length = FinishedLength
                                Width = FinishedWidth
                                ID = panel.getAttribute('ID')
                                Machines = panel.getElementsByTagName("Machines")

                                for Machining in Machines:
                                    macin = Machining.getElementsByTagName("Machining")
                                    for m in macin:
                                        Type = int(m.getAttribute('Type'))
                                        Face = int(m.getAttribute('Face'))
                                        X = float(m.getAttribute('X'))
                                        Y = float(m.getAttribute('Y'))
                                        if Type == 3:
                                            strname = strname.split('+')[0]
                                            break
                                        elif Type == 4:
                                            EndX = float(m.getAttribute('EndX'))
                                            EndY = float(m.getAttribute('EndY'))
                                            widthTool = float(m.getAttribute('Width'))
                                            ToolOffset = m.getAttribute('ToolOffset')

                                            # 判断槽宽
                                            if widthTool == 13:
                                                # if bianlian == 0:
                                                # 判断通槽
                                                    if (abs(EndX - X) == Length and abs(EndY - Y) == 0):
                                                        if Face == 5:
                                                            # 判断走刀方向
                                                            if X < EndX or Y < EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t=1
                                                                        else:
                                                                            strname=''
                                                                            t=1

                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode',"0")
                                                                    ids = '1' +QRCode+str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool== 7 or Y - widthTool == 7 or Width - Y== 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X - widthTool== 7 or Length - X== 7 or Y== 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                            elif X>EndX or Y>EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)

                                                        elif Face == 6:
                                                            # 判断走刀方向
                                                            if X < EndX or Y < EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                            elif X > EndX or Y > EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+6面槽"
                                                                    hl =panel.getAttribute('Info6')
                                                                    # print(f'sdfsdf{hl}')
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBL2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBL2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1

                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)

                                                    else:
                                                        continue
                                                # elif bianlian ==1:
                                                #     if (abs(EndX - X) == 0 and abs(EndY - Y) == Width):
                                                #         if Face == 5:
                                                #             # 判断走刀方向
                                                #             if X < EndX or Y < EndY:
                                                #                 if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #             elif X > EndX or Y > EndY:
                                                #                 if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #
                                                #         elif Face == 6:
                                                #             # 判断走刀方向
                                                #             if X < EndX or Y < EndY:
                                                #                 if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #             elif X > EndX or Y > EndY:
                                                #                 if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname=='1.0mm封边':
                                                #                         if '▲▲▲▲' in EBL2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBL2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12]=(BandingCodingDict[EBL2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #     else:
                                                #         continue
                                                #

                                            else:
                                                continue
                                        else:
                                            continue
                                strname +='+跟踪'
                            if '6面槽+5面槽' in strname :
                                strname = strname.replace('6面槽+5面槽','5&6面槽')
                            elif '5面槽+6面槽' in strname :
                                strname = strname.replace('5面槽+6面槽', '5&6面槽')
                        elif i==1:
                            t = 0
                            # if bianlian==1:
                            #     EBW1 = panel.getAttribute('EBL1')
                            #     EBW2 = panel.getAttribute('EBL2')
                            # elif bianlian==0:
                            EBW1 = panel.getAttribute('EBW1')
                            EBW2 = panel.getAttribute('EBW2')
                            # EBW1 = panel.getAttribute('EBW1')
                            # EBW2 = panel.getAttribute('EBW2')
                            strname = ''
                            if EBW1 == "":
                                strname = '无封边'
                            else:
                                if '▲▲▲▲' in EBW1:
                                    strname = '1.0mm封边'
                                elif '△△△△' in EBW1:
                                    strname = '0.5mm封边'
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
                                        if Type ==3:
                                            strname=strname.split('+')[0]
                                            break
                                        elif Type == 4:
                                            EndX = float(m.getAttribute('EndX'))
                                            EndY = float(m.getAttribute('EndY'))
                                            widthTool = float(m.getAttribute('Width'))
                                            ToolOffset = m.getAttribute('ToolOffset')
                                            # 判断槽宽
                                            if widthTool == 13:
                                                # 判断通槽
                                                # if bianlian ==0:
                                                    if (abs(EndX - X) == 0 and abs(EndY - Y) == Width):
                                                        if Face == 5:
                                                            # 判断走刀方向
                                                            if X < EndX or Y < EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                            elif X > EndX or Y > EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6',"A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    m.setAttribute()
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+5面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)

                                                        elif Face == 6:
                                                            # 判断走刀方向
                                                            if X < EndX or Y < EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                            elif X > EndX or Y > EndY:
                                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                                    if strname=='1.0mm封边':
                                                                        if '▲▲▲▲' in EBW2:
                                                                            strname = '1.0mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        elif '△△△△' in EBW2:
                                                                            strname = '0.5mm封边'
                                                                            PanelBasics[12]=(BandingCodingDict[EBW2.strip()])
                                                                            t = 1
                                                                        else:
                                                                            strname = ''
                                                                            t = 1
                                                                    strname = strname + "+6面槽"
                                                                    if panel.getAttribute('Info6') == 'AC':
                                                                        panel.setAttribute('Info6', "A")
                                                                    m.setAttribute('IsGenCode', "0")
                                                                    ids = '1' + QRCode + str(Face)
                                                                    IdFile.append(ids)
                                                                    CsvIds = QRCode
                                                                    CsvIdFile.append(CsvIds)
                                                    else:
                                                        continue
                                                # elif bianlian == 1:
                                                #     if (abs(EndX - X) == Length and abs(EndY - Y) == 0):
                                                #         if Face == 5:
                                                #             # 判断走刀方向
                                                #             if X < EndX or Y < EndY:
                                                #                 if ToolOffset == '中' and (
                                                #                         X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (
                                                #                         X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (
                                                #                         X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #             elif X > EndX or Y > EndY:
                                                #                 if ToolOffset == '中' and (
                                                #                         X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (
                                                #                         X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (
                                                #                         X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+5面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #
                                                #         elif Face == 6:
                                                #             # 判断走刀方向
                                                #             if X < EndX or Y < EndY:
                                                #                 if ToolOffset == '中' and (
                                                #                         X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (
                                                #                         X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (
                                                #                         X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #             elif X > EndX or Y > EndY:
                                                #                 if ToolOffset == '中' and (
                                                #                         X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '右' and (
                                                #                         X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #                 elif ToolOffset == '左' and (
                                                #                         X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                #                     if strname == '1.0mm封边':
                                                #                         if '▲▲▲▲' in EBW2:
                                                #                             strname = '1.0mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         elif '△△△△' in EBW2:
                                                #                             strname = '0.5mm封边'
                                                #                             PanelBasics[12] = (
                                                #                             BandingCodingDict[EBW2.strip()])
                                                #                             t = 1
                                                #                         else:
                                                #                             strname = ''
                                                #                             t = 1
                                                #
                                                #                     strname = strname + "+6面槽"
                                                #                     m.setAttribute('IsGenCode', "0")
                                                #                     ids = '1' + QRCode + str(Face)
                                                #                     IdFile.append(ids)
                                                #                     CsvIds = QRCode
                                                #                     CsvIdFile.append(CsvIds)
                                                #
                                                #     else:
                                                #         continue
                                                # else:
                                                #     continue

                                            else:
                                                continue
                                        else:
                                            continue
                                strname += '+跟踪'
                            if '6面槽+5面槽' in strname:
                                strname = strname.replace('6面槽+5面槽', '5&6面槽')
                            elif '5面槽+6面槽' in strname:
                                strname = strname.replace('5面槽+6面槽', '5&6面槽')
                        LeftProcessingCode=ProcessingDict[strname]
                        PanelBasics.append(LeftProcessingCode)

                        # 15.右机封边带编码
                        if i == 2:
                            # if bianlian == 1:
                            #     RightBandingCodeOne = panel.getAttribute('EBW2')
                            # else:
                            #     RightBandingCodeOne = panel.getAttribute('EBL2')
                            RightBandingCodeOne = panel.getAttribute('EBL2')
                            if RightBandingCodeOne == "":
                                RightBandingCodeOne = "无封边"
                                RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
                            else:
                                RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
                            # RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
                            PanelBasics.append(RightBandingCodeOne)
                        elif i == 1:
                            # if bianlian == 1:
                            #     RigthBandingCodeTwo = panel.getAttribute('EBL2')
                            # else:
                            #     RigthBandingCodeTwo = panel.getAttribute('EBW2')
                            RigthBandingCodeTwo = panel.getAttribute('EBW2')
                            if RigthBandingCodeTwo == "":
                                RigthBandingCodeTwo = "无封边"
                                RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
                            else:
                                RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
                            # RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
                            PanelBasics.append(RigthBandingCodeTwo)

                        # 16.右机加工编码
                        if i == 2:
                            # if bianlian == 1:
                            #     if t == 0:
                            #         EBL2 = panel.getAttribute('EBW2')
                            #     elif t == 1:
                            #         EBL2 = panel.getAttribute('EBW1')
                            #         PanelBasics[14] = BandingCodingDict[EBL2.strip()]
                            # else:
                            #     if t == 0:
                            #         EBL2 = panel.getAttribute('EBL2')
                            #     elif t == 1:
                            #         EBL2 = panel.getAttribute('EBL1')
                            #         PanelBasics[14]=BandingCodingDict[EBL2.strip()]
                            if t == 0:
                                EBL2 = panel.getAttribute('EBL2')
                            elif t == 1:
                                EBL2 = panel.getAttribute('EBL1')
                                PanelBasics[14]=BandingCodingDict[EBL2.strip()]
                            strname = ''
                            if EBL2 == "":
                                strname = '无封边'
                            else:
                                if '▲▲▲▲' in EBL2:
                                    strname = '1.0mm封边'
                                elif '△△△△' in EBL2:
                                    strname = '0.5mm封边'
                                strname += '+跟踪'
                        elif i == 1:
                            # if bianlian == 1:
                            #     if t == 0:
                            #         EBW2 = panel.getAttribute('EBL2')
                            #     elif t == 1:
                            #         EBW2 = panel.getAttribute('EBL1')
                            #         PanelBasics[14] = BandingCodingDict[EBW2.strip()]
                            # else:
                            #     if t == 0:
                            #         EBW2 = panel.getAttribute('EBW2')
                            #     elif t == 1:
                            #         EBW2 = panel.getAttribute('EBW1')
                            #         PanelBasics[14] = BandingCodingDict[EBW2.strip()]
                            if t == 0:
                                EBW2 = panel.getAttribute('EBW2')
                            elif t == 1:
                                EBW2 = panel.getAttribute('EBW1')
                                PanelBasics[14] = BandingCodingDict[EBW2.strip()]

                            strname = ''
                            if EBW2 == "":
                                strname = '无封边'
                            else:
                                if '▲▲▲▲' in EBW2:
                                    strname = '1.0mm封边'
                                elif '△△△△' in EBW2:
                                    strname = '0.5mm封边'

                                strname += '+跟踪'
                        RightProcessingCode =ProcessingDict[strname]
                        PanelBasics.append(RightProcessingCode)

                        #17.左右机预铣
                        EdgeGroups = panel.getElementsByTagName("EdgeGroup")
                        for EdgeGroup in EdgeGroups:
                            Edges = EdgeGroup.getElementsByTagName("Edge")
                            for Edge in Edges:
                                Faces = Edge.getAttribute('Face')
                                if i == 1:
                                    if int(Faces) == 1:
                                        Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(12,Pre_MillingLeft)
                                    if int(Faces) == 2:
                                        Pre_MillingRigth =float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(15,Pre_MillingRigth)
                                elif i ==2:
                                    if int(Faces) == 3:
                                        Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(12,Pre_MillingLeft)
                                    if int(Faces) == 4:
                                        Pre_MillingRigth =float(Edge.getAttribute('Pre_Milling'))
                                        PanelBasics.insert(15,Pre_MillingRigth)
                        PanelList.append(PanelBasics)
                IdList.append(IdFile)
                # print(len(CsvIdFile))
                # 去重
                CsvIdFile = list(set(CsvIdFile))
                if Folderpath4!='':
                    self.SawCsv(CsvIdFile,Folderpath4,j)


                col = ('二维码', '备注1', '备注2', '备注3', '左机预铣', '右机预铣', '完工长度', '完工宽度', '完工厚度', '数量','进给次序','工件旋转','浮动铣刀','速度','左机封边带编码','左机加工编码','右机封边带编码','右机加工编码')
                # Excel
                # 写入标题
                # test.write_row(0, 0, col,styleCell.style1())
                # 写入数据
                # test.write_rows(1, 0, PanelList,styleCell.style0())
                # SqlUnit.InserUnit(PanelList)
                # 保存按原文件命名  "D:/{}月余料利用率.xls".format(time)
                # Csv
                df = pd.DataFrame(PanelList)
                # df.columns = ['参考', '装饰', '物料流', '批号', '客户编号', '长', '宽', '厚', '数量','进给次序，通过值','速度','方向(工件旋转)','左机预铣','左机封边','左机加工','浮动铣刀','右机预铣','右机封边','右机加工']
                df.columns = ['Reference', 'Decor', 'Materialflow', 'BatchNumber', 'CustomerNumber', 'Length', 'Width', 'Thickness', 'Quantity', 'passValue','FeedSpeed','Orientation',
                              'OverSizeM1','EdgeMacroLM1','ProgramM1','OverSizeM2','EdgeMacroRM2','ProgramM2']


                path = Folderpath2+'/'+j.split('.')[0]+'.csv'
                # test.save(path)
                df.to_csv(path,sep=';',index=False,header=False,encoding='gb18030')
                with open(Folderpath3 + "/封边后" + fullPath.split('/')[-1], "w", encoding="UTF-8") as fs:
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
        # return IdList
def main(Folderpath,Folderpath2,Folderpath3,Folderpath4,IdList):
    # 创建表格对象
    test = My_sheet()
    # 创建样式对象
    styleCell = style()
    #
    ban=banding()
    SqlUnit = sqlUnit.main()
    file = ban.changeWwidth(Folderpath,Folderpath2,Folderpath3,Folderpath4,test,styleCell,SqlUnit,IdList)
    # sqlUnit.SqlUnit.InserUnit(PaneListNUm)
    return file,SqlUnit


