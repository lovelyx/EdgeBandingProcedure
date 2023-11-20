
# Author: LiYiXiao
# 处理xml,得到相关文件

from xml.dom import minidom
import pymsgbox as mb
import logging
import traceback
import os
import pandas as pd
import datetime

from afterEnd.page.My_sheet import My_sheet
from afterEnd.page.styleT import style
import lib.Sql as sqlUnit


class banding:
    def __init__(self):
        self.changeWwidth

    # 修改csv文件堆垛
    def SawCsv(self,CsvIdFileQ,CsvIdFileQL,CsvIdFileKc,path,file,CsvIdFileKcFive,CsvIdFileKcSix,CsvIdFileA,CsvIdFileT):
        """
        处理满足条件的csv中的堆垛信息
        :param CsvIdFile: 保存满足条件的二维码+
        :param path: 待处理csv存放路径
        :param file: 正在处理的csv的文件名
        :return:
        """
        fullPath = path + "/" + file.split(".")[0] +'.csv'  # 完整路径
        csvData = pd.read_csv(fullPath,encoding='gb18030')
        df = pd.DataFrame(csvData)
        # print(csvData)
        code = df.loc[:,'条形码']
        for index,i in enumerate(code):
            if (i.split('N')[1] in CsvIdFileKc):
                num = df.loc[index,'堆垛']
                if num == 'AC' and i.split('N')[1] not in CsvIdFileQL:
                    # print(f'堆垛为{num,index}')
                    df.loc[index,'堆垛'] = 'A'
                    df.loc[index,'info6'] ='A'
                else:
                    df.loc[index, '堆垛'] = 'E'
                    df.loc[index, 'info6'] = 'E'
            if i.split('N')[1] in CsvIdFileQ:
                df.loc[index, '301'] = 'Q'
                df.loc[index, 'info7'] = 'Q'
            if i.split('N')[1] in CsvIdFileQL:
                df.loc[index, '301'] = 'QL'
                df.loc[index, 'info7'] = 'QL'
            if i.split('N')[1] in CsvIdFileKcFive:
                df.loc[index, '301'] = df.loc[index,'301']+'5'
                df.loc[index, 'info7'] = df.loc[index,'info7']+'5'
            if i.split('N')[1] in CsvIdFileKcSix:
                df.loc[index, '301'] = df.loc[index,'301']+'6'
                df.loc[index, 'info7'] = df.loc[index,'info7']+'6'
            if i.split('N')[1] in CsvIdFileA:
                df.loc[index, '堆垛'] = 'A'
                df.loc[index, 'info6'] = 'A'
            if i.split('N')[1] in CsvIdFileT:
                df.loc[index, '堆垛'] = 'D'
                df.loc[index, 'info6'] = 'D'
        df.to_csv(fullPath, index=False,encoding='gb18030')

        return

    # 二维码，参考
    def QRCode(self,panel,PanelBasics):
        """
        二维码列
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return:QRCode当前行的二维码
        """
        QRCode = panel.getAttribute('ID')
        PanelBasics.append(QRCode)

        return QRCode

    # 备注1
    def notesOne(selfs,panel,PanelBasics):
        """
        备注1 BatchNo属性
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return: RemarksOne 备注1
        """
        RemarksOne = panel.getAttribute('BatchNo')
        PanelBasics.append(RemarksOne)
        return RemarksOne

    # 备注2
    def notesTwo(self,panel,PanelBasics):
        """
        备注2，未赋值
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return:备注2 目前为空
        """
        # RemarksTwo = panel.getAttribute('Info4')
        PanelBasics.append('')
        return

    # 备注3
    def notesThree(self, panel, PanelBasics):
        """
        批次号列
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return: RemarksThree 备注3 批次号
        """
        RemarksThree = panel.getAttribute('Info4')
        PanelBasics.append(RemarksThree)

        return RemarksThree

    # 备注4
    def notesFour(self, panel, PanelBasics):
        """
        合同号列
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return: RemarksFour 备注4 合同号（客户编号）
        """
        RemarksFour = panel.getAttribute('ContractNo')
        PanelBasics.append(RemarksFour)

        return RemarksFour

    # 完工长度
    def FinishedLength(self, panel, PanelBasics):
        """
        成品长度列
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return: FinishedLength 成品长
        """
        FinishedLength = float(panel.getAttribute('Length'))
        return FinishedLength

    # 完工宽度
    def FinishedWidth(self, panel, PanelBasics):
        """
        成品宽度列
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return: FinishedWidth 成品宽
        """
        FinishedWidth = float(panel.getAttribute('Width'))
        return FinishedWidth

     # 厚度
    def Thickness(self, panel, PanelBasics):
        """
        厚度列
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return: FinishedThickness 厚度
        """
        FinishedThickness = float(panel.getAttribute('Thickness'))
        PanelBasics.append(FinishedThickness)

        return FinishedThickness

    # 数量
    def Number(self, panel, PanelBasics):
        """
        数量列
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据
        :return: Quantity 数量
        """
        Quantity = int(panel.getAttribute('Qty'))
        PanelBasics.append(Quantity)

        return Quantity

    # 进给次序
    def FeedSequence(self, i, PanelBasics):
        """
        进给次序列
        :param i: 第几行数据
        :param PanelBasics: 存放当前行数据
        :return:
        """
        if i == 1:
            Order = 1
            PanelBasics.append(int(Order))
        elif i == 2:
            Order = 2
            PanelBasics.append(int(Order))

        return
    # 速度
    def SpeedColumn(self, PanelBasics):
        """
        速度列
        :param PanelBasics: 存放当前行数据
        :return:
        """
        Speed = 30
        PanelBasics.append(float(Speed))

        return

    # 工件旋转
    def WorkpieceRotation(self,i, PanelBasics,WorkpieceRotation,Grain,FinishedLength,FinishedWidth):
        """
        工件旋转，长宽对调
        :param i: 第几行数据
        :param PanelBasics: 存放当前行数据的数组
        :param WorkpieceRotation: 工件旋转变量
        :return:
        """
        # 设置变量宽大于长时判断标准槽时，长宽，坐标也要对应变化
        bianlian = 0;
        if i == 1:
            WorkpieceRotation=1
            if Grain=='W':
                WorkpieceRotation = 0
                # bianlian = 1
            else:
                WorkpieceRotation = 1

            # WorkpieceRotation = 1
            # bianlian = 1
            PanelBasics.append(int(WorkpieceRotation))
        elif i == 2:
            if WorkpieceRotation == 1:
                WorkpieceRotation = 0

                PanelBasics.append(int(WorkpieceRotation))
            elif WorkpieceRotation == 0:
                WorkpieceRotation = 1
                # bianlian = 1
                PanelBasics.append(int(WorkpieceRotation))

        return WorkpieceRotation

    # 左机封边带编码
    def LeftBandingCode(self, i, panel, PanelBasics, BandingCodingDict,Grain):
        """
        左机封边带编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的编码对照字典
        :return:
        """
        # 控制厚封边在右机
        n=0
        if i == 2:
            if Grain =='W':
                LeftBandingCodeOne = panel.getAttribute('EBW1')
                if '▲▲▲▲' in LeftBandingCodeOne:
                    LeftBandingCodeOne = panel.getAttribute('EBW2')
                    n = 1
            else:
                LeftBandingCodeOne = panel.getAttribute('EBL1')
                if '▲▲▲▲' in LeftBandingCodeOne:
                    LeftBandingCodeOne = panel.getAttribute('EBL2')
                    n = 1
            # LeftBandingCodeOne = panel.getAttribute('EBL1')

            if LeftBandingCodeOne == "":
                LeftBandingCodeOne = "无封边"
                LeftBandingCodeOne = BandingCodingDict[LeftBandingCodeOne.strip()]
            else:
                LeftBandingCodeOne = BandingCodingDict[LeftBandingCodeOne.strip()]
            PanelBasics.append(LeftBandingCodeOne)
        elif i == 1:
            if Grain =='W':
                LeftBandingCodeTwo = panel.getAttribute('EBL1')
                if '▲▲▲▲' in LeftBandingCodeTwo:
                    LeftBandingCodeTwo = panel.getAttribute('EBL2')
                    n = 1
            else:
                LeftBandingCodeTwo = panel.getAttribute('EBW1')
                if '▲▲▲▲' in LeftBandingCodeTwo:
                    LeftBandingCodeTwo = panel.getAttribute('EBW2')
                    n = 1
            # LeftBandingCodeTwo = panel.getAttribute('EBW1')

            if LeftBandingCodeTwo == "":
                LeftBandingCodeTwo = "无封边"
                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            else:
                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            # LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            PanelBasics.append(LeftBandingCodeTwo)

        return n


    # 满足开槽
    def Slotting(self,strname,EB2,PanelBasics,BandingCodingDict,panel,m,QRCode,Face,IdFile,CsvIdFileKc,CsvIdFileQL,t):
        """
        满足为标准槽时执行的函数
        :param strname: 保存封边厚薄的变量
        :param EB2: 当前行的封边信息
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的编码对照字典
        :param panel: xml中的panel标签
        :param m: xml中的Machining标签
        :param QRCode: 二维码
        :param Face: 加工面
        :param IdFile: 存放二维码，与mpr文件配对
        :param CsvIdFileKc: 存放二维码，与csv中的数据进行配对
        :param t: 控制存在两边为一薄一厚时，厚边在右机加工变量
        :return: t,strname
        """
        if strname == '1.0mm封边':
            if '▲▲▲▲' in EB2:
                strname = '1.0mm封边'
                PanelBasics[12] = (BandingCodingDict[EB2.strip()])
                t = 1
            elif '△△△△' in EB2:
                strname = '0.5mm封边'
                PanelBasics[12] = (BandingCodingDict[EB2.strip()])
                t = 1
            else:
                strname = ''
                t = 1

        # if f'{Face}面槽' in strname:
        #     strname=strname
        # else:
        strname = strname + f"+{Face}面槽"
        # if '+5面槽' in strname or '+5面槽+6面槽' in strname or '+6面槽+5面槽' in strname:
        # if panel.getAttribute('Info6') == 'AC':
        #     panel.setAttribute('Info6', "A")
        # else:
        #     panel.setAttribute('Info6', "E")
        m.setAttribute('IsGenCode', "0")
        ids = '1' + QRCode + str(Face)
        IdFile.append(ids)
        CsvIdFileKc.append(QRCode)
        return t,strname

    # 左机加工编码
    def LeftProcessCode(self,i, panel, PanelBasics,FinishedLength,FinishedWidth,BandingCodingDict,QRCode,IdFile,CsvIdFileQ,CsvIdFileQL,ProcessingDict,n,fhSum,CsvIdFileT,Grain):
        """
        左机加工编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param FinishedLength: 成品长
        :param FinishedWidth: 成品宽
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :param QRCode: 二维码
        :param IdFile: 存放二维码，与mpr文件配对
        :param CsvIdFileQ: 存放二维码，与csv中的数据进行配对,开槽
        :param CsvIdFileQL:
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: t,当开槽时控制厚边右机封
        """
        # 当开槽时
        # 控制厚边右机封
        t=0
        if i == 2:
            if Grain == 'W':
                if n == 1:
                    EBL1 = panel.getAttribute('EBW2')
                    EBL2 = panel.getAttribute('EBW1')
                elif n == 0:
                    EBL1 = panel.getAttribute('EBW1')
                    EBL2 = panel.getAttribute('EBW2')
            else:
                if n==1:
                    EBL1 = panel.getAttribute('EBL2')
                    EBL2 = panel.getAttribute('EBL1')
                elif n==0:
                    EBL1 = panel.getAttribute('EBL1')
                    EBL2 = panel.getAttribute('EBL2')
            # EBL1 = panel.getAttribute('EBL1')
            # EBL2 = panel.getAttribute('EBL2')
            strname = ''
            if EBL1 == "":
                strname = '无封边'
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
                info6=panel.getAttribute('Info6')
                ebl1 = panel.getAttribute('EBL1')
                ebl2 = panel.getAttribute('EBL2')
                ebw1 = panel.getAttribute('EBW1')
                ebw2 = panel.getAttribute('EBW2')
                Machines = panel.getElementsByTagName("Machines")
                DiameterList = ['6','5','10','12','18']

                for Machining in Machines:
                    macin = Machining.getElementsByTagName("Machining")
                    for m in macin:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        Diameter = m.getAttribute('Diameter')

                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        if '▲▲▲▲' in ebl1 and '▲▲▲▲' in ebl2 and '▲▲▲▲' in ebw1 and '▲▲▲▲' in ebw2:
                            strname = strname.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        if Type == 3:
                            strname = strname.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            strname = strname.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 4:
                            EndX = float(m.getAttribute('EndX'))
                            EndY = float(m.getAttribute('EndY'))
                            widthTool = float(m.getAttribute('Width'))
                            ToolOffset = m.getAttribute('ToolOffset')

                            # 判断槽宽
                            if widthTool == 13:
                                # if bianlian == 0:
                                if 'W' in Grain:
                                # 判断通槽
                                    if (abs(EndX - X) == 0 and abs(EndY - Y) == Width):
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face, IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                    else:
                                        continue
                                else:
                                    if (abs(EndX - X) == Length and abs(EndY - Y) == 0):
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t,strname=self.Slotting(strname,EBL2,PanelBasics,BandingCodingDict,panel,m,QRCode,Face,IdFile,CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t,strname=self.Slotting(strname,EBL2,PanelBasics,BandingCodingDict,panel,m,QRCode,Face,IdFile,CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t,strname=self.Slotting(strname, EBL2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t,strname=self.Slotting(strname, EBL2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t,strname=self.Slotting(strname, EBL2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t,strname=self.Slotting(strname, EBL2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL,t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname,EBL2,PanelBasics,BandingCodingDict,panel,m,QRCode,Face,IdFile,CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)

                                    else:
                                        continue

                            else:
                                continue
                        else:
                            continue
                strname += '+跟踪'
            if '6面槽+5面槽' in strname:
                strname = strname.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in strname:
                strname = strname.replace('5面槽+6面槽', '5&6面槽')
            if '▲▲▲▲' in EBL1:
                strname = strname+'+倒棱'
            # if '极简小射灯' in panel.getAttribute('Name'):


            #     只有六面有标准槽时，不加工
            # if '+6面槽' in strname:
            #     strname = strname.replace('+6面槽', '')
            #     IdFile.remove('1' + QRCode + '6')
            #     CsvIdFileQ.remove(QRCode)
            #     for Machining in Machines:
            #         macin = Machining.getElementsByTagName("Machining")
            #         for q in macin:
            #             if q.getAttribute('Face') =='6':
            #                 q.setAttribute('IsGenCode', "2")

        elif i == 1:
            t = 0
            if Grain =='W':
                if n == 1:
                    EBW1 = panel.getAttribute('EBL2')
                    EBW2 = panel.getAttribute('EBL1')
                elif n == 0:
                    EBW1 = panel.getAttribute('EBL1')
                    EBW2 = panel.getAttribute('EBL2')
            else:
                if n==1:
                    EBW1 = panel.getAttribute('EBW2')
                    EBW2 = panel.getAttribute('EBW1')
                elif n==0:
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
                ID = panel.getAttribute('ID')
                info6 = panel.getAttribute('Info6')
                DiameterList = ['6', '5', '10', '12', '18']
                for Machining in Machines:
                    macin = Machining.getElementsByTagName("Machining")
                    for m in macin:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        Diameter = m.getAttribute('Diameter')
                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        if Type == 3:
                            strname = strname.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            strname = strname.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 4:
                            EndX = float(m.getAttribute('EndX'))
                            EndY = float(m.getAttribute('EndY'))
                            widthTool = float(m.getAttribute('Width'))
                            ToolOffset = m.getAttribute('ToolOffset')
                            # 判断槽宽
                            if widthTool == 13:
                                if 'W' in Grain:
                                # 判断通槽
                                    if (abs(EndX - X) == Length and abs(EndY - Y) == 0):
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif ToolOffset == '左' and ( X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)

                                    else:
                                        continue
                                # if bianlian ==0:
                                else:
                                    if (abs(EndX - X) == 0 and abs(EndY - Y) == Width):
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t,strname=self.Slotting(strname,EBW2,PanelBasics,BandingCodingDict,panel,m,QRCode,Face,IdFile,CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7 or Y - widthTool == 7 or Width - Y == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7 or Y == 7 or Width - Y - widthTool == 7):
                                                    t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL,t)
                                    else:
                                        continue

                            else:
                                continue

                        else:
                            continue

                strname += '+跟踪'
            if '6面槽+5面槽' in strname:
                strname = strname.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in strname:
                strname = strname.replace('5面槽+6面槽', '5&6面槽')
            # if '+6面槽'  in strname:
            #     strname = strname.replace('+6面槽','')
            #     IdFile.remove('1'+QRCode+'6')
            #     CsvIdFileQ.remove(QRCode)
            #     for Machining in Machines:
            #         macin = Machining.getElementsByTagName("Machining")
            #         for q in macin:
            #             if q.getAttribute('Face') == '6':
            #                 q.setAttribute('IsGenCode', "2")

        LeftProcessingCode = ProcessingDict[strname]
        PanelBasics.append(LeftProcessingCode)

        return t,fhSum,strname

    # 侧板 左机加工编码
    def LeftProcessCodeSide(self,i, panel, PanelBasics,FinishedLength,FinishedWidth,BandingCodingDict,QRCode,IdFile,CsvIdFileQ,CsvIdFileQL,ProcessingDict,n,fhSum,CsvIdFileT,Grain):
        """
        左机加工编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param FinishedLength: 成品长
        :param FinishedWidth: 成品宽
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :param QRCode: 二维码
        :param IdFile: 存放二维码，与mpr文件配对
        :param CsvIdFileQ: 存放二维码，与csv中的数据进行配对,开槽
        :param CsvIdFileQL:
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: t,当开槽时控制厚边右机封
        """
        # 当开槽时
        # 控制厚边右机封
        t=0
        if i == 2:
            if Grain == 'W':
                if n == 1:
                    EBL1 = panel.getAttribute('EBW2')
                    EBL2 = panel.getAttribute('EBW1')
                elif n == 0:
                    EBL1 = panel.getAttribute('EBW1')
                    EBL2 = panel.getAttribute('EBW2')
            else:
                if n == 1:
                    EBL1 = panel.getAttribute('EBL2')
                    EBL2 = panel.getAttribute('EBL1')
                elif n == 0:
                    EBL1 = panel.getAttribute('EBL1')
                    EBL2 = panel.getAttribute('EBL2')
            strname = ''
            if EBL1 == "":
                strname = '无封边'
            else:
                if '▲▲▲▲' in EBL1:
                    strname = '1.0mm封边'
                elif '△△△△' in EBL1:
                    strname = '0.5mm封边'
                Length = FinishedLength
                Width = FinishedWidth
                ID = panel.getAttribute('ID')
                info6=panel.getAttribute('Info6')
                Machines = panel.getElementsByTagName("Machines")
                ebl1 = panel.getAttribute('EBL1')
                ebl2 = panel.getAttribute('EBL2')
                ebw1 = panel.getAttribute('EBW1')
                ebw2 = panel.getAttribute('EBW2')
                DiameterList = ['6','5','10','12','18']

                for Machining in Machines:
                    macin = Machining.getElementsByTagName("Machining")
                    for m in macin:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        Diameter = m.getAttribute('Diameter')

                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        if '▲▲▲▲' in ebl1 and '▲▲▲▲' in ebl2 and '▲▲▲▲' in ebw1 and '▲▲▲▲' in ebw2:
                            strname = strname.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        if Type == 3:
                            strname = strname.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            strname = strname.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 4:
                            FiveTypeNum = 0
                            SixTypeNum = 0
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if (q.getAttribute('Face') == '5' and q.getAttribute('Type') == '4'):
                                        FiveTypeNum += 1
                                    elif (q.getAttribute('Face') == '6' and q.getAttribute('Type') == '4'):
                                        SixTypeNum += 1
                            if FiveTypeNum >= 2 or SixTypeNum >= 2:
                                strname = strname.split('+')[0]
                                if ID in CsvIdFileQ:
                                    CsvIdFileQ.remove(ID)
                                panel.setAttribute('Info6', info6)
                                for Machining in Machines:
                                    macin = Machining.getElementsByTagName("Machining")
                                    for q in macin:
                                        if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                            q.setAttribute('IsGenCode', "2")
                                break

                            EndX = float(m.getAttribute('EndX'))
                            EndY = float(m.getAttribute('EndY'))
                            widthTool = float(m.getAttribute('Width'))
                            ToolOffset = m.getAttribute('ToolOffset')

                            # 判断槽宽
                            if widthTool == 13:
                                if (abs(EndY - Y) == 0):
                                    if 'W' in Grain:
                                        if (abs(EndX - X) == 0):
                                            if Face == 5:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)

                                            elif Face == 6:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)


                                    else:
                                        if (abs(EndY - Y) == 0):
                                            if Face == 5:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face,IdFile, CsvIdFileQ,CsvIdFileQL, t)

                                            elif Face == 6:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face, IdFile,CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face, IdFile,CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face, IdFile,CsvIdFileQ, CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face, IdFile,CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face, IdFile,CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBL2, PanelBasics,BandingCodingDict, panel, m,QRCode, Face, IdFile,CsvIdFileQ, CsvIdFileQL, t)

                            else:
                                continue
                        else:
                            continue
                strname += '+跟踪'
            if '6面槽+5面槽' in strname:
                strname = strname.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in strname:
                strname = strname.replace('5面槽+6面槽', '5&6面槽')
            if '▲▲▲▲' in EBL1:
                strname = strname+'+倒棱'
            # if '极简小射灯' in panel.getAttribute('Name'):


            #     只有六面有标准槽时，不加工
            # if '+6面槽' in strname:
            #     strname = strname.replace('+6面槽', '')
            #     IdFile.remove('1' + QRCode + '6')
            #     CsvIdFileQ.remove(QRCode)
            #     for Machining in Machines:
            #         macin = Machining.getElementsByTagName("Machining")
            #         for q in macin:
            #             if q.getAttribute('Face') =='6':
            #                 q.setAttribute('IsGenCode', "2")
        elif i == 1:
            t = 0
            if Grain == 'W':
                if n == 1:
                    EBW1 = panel.getAttribute('EBL2')
                    EBW2 = panel.getAttribute('EBL1')
                elif n == 0:
                    EBW1 = panel.getAttribute('EBL1')
                    EBW2 = panel.getAttribute('EBL2')
            else:
                if n == 1:
                    EBW1 = panel.getAttribute('EBW2')
                    EBW2 = panel.getAttribute('EBW1')
                elif n == 0:
                    EBW1 = panel.getAttribute('EBW1')
                    EBW2 = panel.getAttribute('EBW2')
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
                ID = panel.getAttribute('ID')
                info6=panel.getAttribute('Info6')
                DiameterList = ['6','5','10','12','18']
                for Machining in Machines:
                    macin = Machining.getElementsByTagName("Machining")
                    for m in macin:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        Diameter = m.getAttribute('Diameter')
                        if Type == 3:
                            strname = strname.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            strname = strname.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 4:
                            FiveTypeNum = 0
                            SixTypeNum = 0
                            for Machining in Machines:
                                macin = Machining.getElementsByTagName("Machining")
                                for q in macin:
                                    if (q.getAttribute('Face') == '5' and q.getAttribute('Type') == '4'):
                                        FiveTypeNum +=1
                                    elif (q.getAttribute('Face') == '6' and q.getAttribute('Type') == '4'):
                                        SixTypeNum +=1
                            if FiveTypeNum>=2 or SixTypeNum >=2:
                                strname = strname.split('+')[0]
                                if ID in CsvIdFileQ:
                                    CsvIdFileQ.remove(ID)
                                panel.setAttribute('Info6', info6)
                                for Machining in Machines:
                                    macin = Machining.getElementsByTagName("Machining")
                                    for q in macin:
                                        if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                            q.setAttribute('IsGenCode', "2")
                                break

                            EndX = float(m.getAttribute('EndX'))
                            EndY = float(m.getAttribute('EndY'))
                            widthTool = float(m.getAttribute('Width'))
                            ToolOffset = m.getAttribute('ToolOffset')
                            # 判断槽宽
                            if widthTool == 13:
                                    if 'W' in Grain:
                                        if (abs(EndY - Y) == 0):
                                            if Face == 5:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)

                                            elif Face == 6:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict,panel, m, QRCode, Face, IdFile, CsvIdFileQ,CsvIdFileQL, t)

                                    else:
                                        if (abs(EndX - X) == 0):
                                            if Face == 5:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)

                                            elif Face == 6:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics, BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (
                                                            X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                                    elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                        t, strname = self.Slotting(strname, EBW2, PanelBasics,BandingCodingDict, panel, m, QRCode, Face,IdFile, CsvIdFileQ, CsvIdFileQL, t)
                                            # # 判断通槽
                            else:
                                continue

                        else:
                            continue

                strname += '+跟踪'
            if '6面槽+5面槽' in strname:
                strname = strname.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in strname:
                strname = strname.replace('5面槽+6面槽', '5&6面槽')
            # if '+6面槽'  in strname:
            #     strname = strname.replace('+6面槽','')
            #     IdFile.remove('1'+QRCode+'6')
            #     CsvIdFileQ.remove(QRCode)
            #     for Machining in Machines:
            #         macin = Machining.getElementsByTagName("Machining")
            #         for q in macin:
            #             if q.getAttribute('Face') == '6':
            #                 q.setAttribute('IsGenCode', "2")

        LeftProcessingCode = ProcessingDict[strname]
        PanelBasics.append(LeftProcessingCode)

        return t,fhSum,strname

    # 浮动铣刀
    def FloatingCutter(self,i,PanelBasics,FininshedLength,FinishedWidth):
        """

        :param i: 第几行数据
        :param PanelBasics: 存放当前行数据的数组
        :param FininshedLength: 成品长
        :param FinishedWidth: 成品宽
        :return: 空
        """
        if i == 1 and FininshedLength>=FinishedWidth:
            milling = 0
            PanelBasics.append(milling )
        elif i == 1 and FininshedLength<FinishedWidth:
            milling = 2
            PanelBasics.append(milling)
        elif i == 2 and FininshedLength>=FinishedWidth:
            milling = 2
            PanelBasics.append(milling)
        elif i == 2 and FininshedLength < FinishedWidth:
            milling = 0
            PanelBasics.append(milling)

        return

    # 右机封边带编码
    def RigthBandingCode(self,i, panel, PanelBasics, BandingCodingDict,n,Grain):
        """
        右机分封边带编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :return: 空
        """
        if i == 2:
            if Grain=='W':
                if n == 1:
                    RightBandingCodeOne = panel.getAttribute('EBW1')
                elif n == 0:
                    RightBandingCodeOne = panel.getAttribute('EBW2')
            else:
                if n == 1:
                    RightBandingCodeOne = panel.getAttribute('EBL1')
                elif n == 0:
                    RightBandingCodeOne = panel.getAttribute('EBL2')
            # RightBandingCodeOne = panel.getAttribute('EBL2')
            if RightBandingCodeOne == "":
                RightBandingCodeOne = "无封边"
                RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
            else:
                RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
            # RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
            PanelBasics.append(RightBandingCodeOne)
        elif i == 1:
            if Grain =='W':
                if n == 1:
                    RigthBandingCodeTwo = panel.getAttribute('EBL1')
                elif n == 0:
                    RigthBandingCodeTwo = panel.getAttribute('EBL2')
            else:
                if n == 1:
                    RigthBandingCodeTwo = panel.getAttribute('EBW1')
                elif n==0:
                    RigthBandingCodeTwo = panel.getAttribute('EBW2')
            # RigthBandingCodeTwo = panel.getAttribute('EBW2')
            if RigthBandingCodeTwo == "":
                RigthBandingCodeTwo = "无封边"
                RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
            else:
                RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
            # RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
            PanelBasics.append(RigthBandingCodeTwo)

        return

    # 右击加工编码
    def RigthProcessCode(self,i, panel, PanelBasics, BandingCodingDict,t,ProcessingDict,CheckBoxf,n,Grain):
        """
        右机加工编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :param t:  控制存在两边为一薄一厚时，厚边在右机加工变量
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: RightProcessingCode 加工编码
        """
        if i == 2:
            if t == 0:
                if Grain=='W':
                    if n == 1:
                        EBL2 = panel.getAttribute('EBW1')
                    elif n == 0:
                        EBL2 = panel.getAttribute('EBW2')
                else:
                    if n==1:
                        EBL2 = panel.getAttribute('EBL1')
                    elif n==0:
                        EBL2 = panel.getAttribute('EBL2')
            elif t == 1:
                if Grain=='W':
                    EBL2 = panel.getAttribute('EBW1')
                else:
                    EBL2 = panel.getAttribute('EBL1')
                if CheckBoxf==True:
                    PanelBasics[15] = BandingCodingDict[EBL2.strip()]
                else:
                    PanelBasics[14] = BandingCodingDict[EBL2.strip()]
            strname = ''
            if EBL2 == "":
                strname = '无封边'
            else:
                if '▲▲▲▲' in EBL2:
                    strname = '1.0mm封边'
                elif '△△△△' in EBL2:
                    strname = '0.5mm封边'
                strname += '+跟踪'
            if '▲▲▲▲' in EBL2:
                strname +='+倒棱'
        elif i == 1:
            if t == 0:
                if Grain=='W':
                    if n == 1:
                        EBW2 = panel.getAttribute('EBL1')
                    elif n == 0:
                        EBW2 = panel.getAttribute('EBL2')
                else:
                    if n == 1:
                        EBW2 = panel.getAttribute('EBW1')
                    elif n == 0:
                        EBW2 = panel.getAttribute('EBW2')
                # EBW2 = panel.getAttribute('EBW2')
            elif t == 1:
                if Grain=='W':
                    EBW2 = panel.getAttribute('EBL1')
                else:
                    EBW2 = panel.getAttribute('EBW1')
                if CheckBoxf == True:
                    PanelBasics[15] = BandingCodingDict[EBW2.strip()]
                else:
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
        RightProcessingCode = ProcessingDict[strname]
        PanelBasics.append(RightProcessingCode)

        return RightProcessingCode

    # 左右机预铣
    def PreMilling(self,i, panel, PanelBasics,CheckBoxf):
        """
        左右机预铣，插入指定位置
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :return: 空
        """
        EdgeGroups = panel.getElementsByTagName("EdgeGroup")
        for EdgeGroup in EdgeGroups:
            Edges = EdgeGroup.getElementsByTagName("Edge")
            for Edge in Edges:
                Faces = Edge.getAttribute('Face')
                if i == 1:
                    if int(Faces) == 1:
                        # Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                        Pre_MillingLeft = 0.5
                        if CheckBoxf==True:
                            PanelBasics.insert(12, Pre_MillingLeft)
                        else:
                            PanelBasics.insert(12, Pre_MillingLeft)
                    if int(Faces) == 2:
                        # Pre_MillingRigth = float(Edge.getAttribute('Pre_Milling'))
                        Pre_MillingRigth = 0.5
                        if CheckBoxf == True:
                            PanelBasics.insert(16, Pre_MillingRigth)
                        else:
                            PanelBasics.insert(15, Pre_MillingRigth)
                elif i == 2:
                    if int(Faces) == 3:
                        # Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                        Pre_MillingLeft = 0.5
                        if CheckBoxf == True:
                            PanelBasics.insert(12, Pre_MillingLeft)
                        else:
                            PanelBasics.insert(12, Pre_MillingLeft)
                    if int(Faces) == 4:
                        # Pre_MillingRigth = float(Edge.getAttribute('Pre_Milling'))
                        Pre_MillingRigth = 0.5
                        if CheckBoxf == True:
                            PanelBasics.insert(16, Pre_MillingRigth)
                        else:
                            PanelBasics.insert(15, Pre_MillingRigth)

        return




    def changeWwidth(self,Folderpath,Folderpath2,Folderpath3,Folderpath4,test,styleCell,SqlUnit,IdList,CheckBoxf,username,IdDictNo):
        BandingProcessing, BandingCode = SqlUnit.selectBandingProcessing()
        # BandingProcessingLeft,BandingProcessingRight,BandingCodeLeft,BandingCodeRight = SqlUnit.selectBandingProcessing()
        # print(f'BandingProcessingLeft:{BandingProcessingLeft}')
        # print(f'BandingProcessingRight:{BandingProcessingRight}')
        # print(f'BandingCodeLeft:{BandingCodeLeft}')
        # print(f'BandingCodeRight:{BandingCodeRight}')
        # BandingProcessingLeftDict={}
        # for t in BandingProcessingLeft:
        #     BandingProcessingLeftDict[t[1]]=t[0]
        # BandingProcessingRightDict = {}
        # for t in BandingProcessingRight:
        #     BandingProcessingRightDict[t[1]] = t[0]
        # BandingCodingLeftDict={}
        # BandingCodingRightDict = {}
        # for t in BandingCodeLeft:
        #     BandingCodingLeftDict[t[1]]=t[2]
        # for t in BandingCodeRight:
        #     BandingCodingRightDict[t[1]]=t[2]
        BandingProcessingDict = {}
        for t in BandingProcessing:
            BandingProcessingDict[t[1]] = t[0]
        BandingCodingDict = {}
        for t in BandingCode:
            BandingCodingDict[t[1]]=t[2]
        # 打开文件夹得到文件夹下的所有文件
        listPacth=[]
        Sum=0
        fhSum=0
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
                CsvIdFileQ = []
                CsvIdFileQL =[]
                CsvIdFileKc=[]
                CsvIdFileKcFive = []
                CsvIdFileKcSix = []
                CsvIdFileA=[]
                CsvIdFileT = []

                # 存入数据库
                PanelListSql = []
                for panel in Panelnodes:
                    Sum +=1
                    WorkpieceRotation=0
                    for i in range(1,3):
                        PanelBasics = []
                        PanelBasicsSql=[]
                        Grain = panel.getAttribute('Grain')
                        PanelName = panel.getAttribute('Name')
                        Machines2 = panel.getElementsByTagName("Machines")


                        # 1.二维码，参考
                        QRCode=self.QRCode(panel,PanelBasics)

                        # 2.备注1，装饰
                        RemarksOne =self.notesOne(panel,PanelBasics)

                        # 3.备注2，物料流
                        RemarksTwo =self.notesTwo(panel,PanelBasics)

                        # 4.备注3，批号
                        RemarksThree = self.notesThree(panel, PanelBasics)

                        # 5.备注4，客户编号
                        RemarksFour = self.notesFour(panel, PanelBasics)

                        # 6.完工长度
                        FinishedLength = self.FinishedLength(panel, PanelBasics)
                        if FinishedLength<250:
                            break
                        PanelBasics.append(FinishedLength)


                        # 7.完工宽度
                        FinishedWidth = self.FinishedWidth(panel, PanelBasics)
                        if FinishedWidth<250:
                            break
                        PanelBasics.append(FinishedWidth)

                        if FinishedLength < FinishedWidth and i==1:
                            if Grain == 'L':
                                panel.setAttribute('Info7', "QL")
                                CsvIds = QRCode
                                CsvIdFileQL.append(CsvIds)
                            elif Grain == 'W':
                                panel.setAttribute('Info7', "Q")
                                CsvIds = QRCode
                                CsvIdFileQ.append(CsvIds)
                        elif FinishedLength >= FinishedWidth and i==1:
                            if Grain == 'L':
                                panel.setAttribute('Info7', "Q")
                                CsvIds = QRCode
                                CsvIdFileQ.append(CsvIds)
                            elif Grain == 'W':
                                panel.setAttribute('Info7', "QL")
                                CsvIds = QRCode
                                CsvIdFileQL.append(CsvIds)


                        # 8.完工厚度
                        FinishedThickness = self.Thickness(panel, PanelBasics)

                        # 9.数量
                        Quantity = self.Number(panel, PanelBasics)
                        # 10.进给次序
                        self.FeedSequence(i, PanelBasics)

                        # 11.速度
                        self.SpeedColumn(PanelBasics)

                        # 12.工件旋转
                        WorkpieceRotation=self.WorkpieceRotation(i,PanelBasics,WorkpieceRotation,Grain,FinishedLength,FinishedWidth)

                        # 13.左机封边带编码
                        n=self.LeftBandingCode(i,panel,PanelBasics,BandingCodingDict,Grain)

                        # 14.左机加工编码
                        if ('侧板' in PanelName) or ('左侧' in PanelName) or ('右侧' in PanelName):
                            t,fhSum,strname = self.LeftProcessCodeSide(i,panel,PanelBasics,FinishedLength,FinishedWidth,BandingCodingDict,QRCode,IdFile,CsvIdFileKc,CsvIdFileQL,BandingProcessingDict,n,fhSum,CsvIdFileT,Grain)
                        else:
                            t,fhSum,strname = self.LeftProcessCode(i,panel,PanelBasics,FinishedLength,FinishedWidth,BandingCodingDict,QRCode,IdFile,CsvIdFileKc,CsvIdFileQL,BandingProcessingDict,n,fhSum,CsvIdFileT,Grain)
                        # print(f'左机加工编码:{fhSum}{j}{QRCode}')

                        # 15.浮动铣刀
                        if CheckBoxf == True:
                            self.FloatingCutter(i,PanelBasics,FinishedLength,FinishedWidth,)

                        # 16.右机封边带编码
                        self.RigthBandingCode(i,panel,PanelBasics,BandingCodingDict,n,Grain)


                        # 17.右机加工编码
                        self.RigthProcessCode(i,panel,PanelBasics,BandingCodingDict,t,BandingProcessingDict,CheckBoxf,n,Grain)

                        #18.左右机预铣
                        self.PreMilling(i,panel,PanelBasics,CheckBoxf)

                        PanelList.append(PanelBasics)



                        if '+5' in strname:
                            panel.setAttribute('Info7',panel.getAttribute('Info7')+'5')
                            CsvIdFileKcFive.append(QRCode)
                        # print(strname)
                        if '6' in strname:
                            panel.setAttribute('Info7',panel.getAttribute('Info7')+'6')
                            CsvIdFileKcSix.append(QRCode)
                        if QRCode in CsvIdFileKc:
                            if panel.getAttribute('Info6') == 'AC' and QRCode not in CsvIdFileQL:
                                panel.setAttribute('Info6', "A")
                            elif panel.getAttribute('Info6') == 'A':
                                panel.setAttribute('Info6', "A")
                            else:
                                panel.setAttribute('Info6', "E")


                        if Machines2:
                            # pass
                            if i ==2:
                                FaceOne,FaceSixKc,FaceSixKcAll,FaceFiveKcAll,FaceSixKo,FaceFiveKc,FaceFiveKo,FaceX,FaceDT,FaceEB,FaceThree,FaceFive,FaceSix=0,0,0,0,0,0,0,0,0,0,0,0,0
                                for Machining in Machines2:
                                    macin = Machining.getElementsByTagName("Machining")
                                    for index, q in enumerate(macin):

                                        if q.getAttribute('Face') == '1' or q.getAttribute('Face') == '2':
                                            FaceOne = 1

                                        if q.getAttribute('Face') == '6':
                                            FaceSix = 1

                                        if q.getAttribute('Face') == '6' and q.getAttribute('Type') == '4':
                                            FaceSixKcAll = 1

                                        if q.getAttribute('Face') == '6' and q.getAttribute('Type') == '2':
                                            FaceSixKo = 1

                                        if q.getAttribute('Face') == '5':
                                            FaceFive = 1

                                        if q.getAttribute('Face') == '5' and q.getAttribute('Type') == '4':
                                            FaceFiveKcAll = 1

                                        if q.getAttribute('Face') == '5' and q.getAttribute('Type') == '2':
                                            FaceFiveKo = 1

                                        if q.getAttribute('Type') == '3':
                                            FaceX = 1

                                        if QRCode in CsvIdFileQL:
                                            FaceDT = 1

                                        if '▲' in panel.getAttribute('EBL2') and '△' in panel.getAttribute('EBL1') and '△' in panel.getAttribute('EBW2') and '△' in panel.getAttribute('EBW1'):
                                            FaceEB = 1

                                        if  q.getAttribute('Face') == '3' or q.getAttribute('Face') == '4':
                                            FaceThree=1

                                        if q.getAttribute('Face') == '6' and q.getAttribute(
                                                'Type') == '4' and q.getAttribute('IsGenCode') != '0':
                                            FaceSixKc = 1
                                        if q.getAttribute('Face') == '5' and q.getAttribute(
                                                'Type') == '4' and q.getAttribute('IsGenCode') != '0':
                                            FaceFiveKc = 1

                                        # print(len(macin),index)
                                        if index == len(macin) - 1:
                                            if '层板' in panel.getAttribute('Name') and float(panel.getAttribute('Length')) >= 764:
                                                if FaceOne != 1 and FaceX != 1 and FaceFiveKc !=1 and FaceSixKc !=1 and (FaceSixKc !=1 and FaceSixKo != 1) and (FaceFiveKc !=1 and FaceSixKo != 1 and FaceSixKc != 1) and (FaceSixKc !=1 and FaceSixKo !=1 and FaceFiveKc!=1) and FaceDT != 1 and FaceThree == 1 and FaceEB !=1:
                                                    panel.setAttribute('Info6', "A")
                                                    CsvIdFileA.append(QRCode)
                                                    # print(QRCode)
                                            else:
                                                if FaceOne != 1 and FaceX != 1 and FaceFiveKc !=1 and FaceSixKc !=1 and (FaceSixKc !=1 and FaceSixKo != 1) and (FaceFiveKc !=1 and FaceSixKo != 1 and FaceSixKc != 1) and (FaceSixKc !=1 and FaceSixKo !=1 and FaceFiveKc!=1) and FaceDT != 1  and ('层板' not in panel.getAttribute('Name')) and FaceEB !=1:
                                                    panel.setAttribute('Info6', "A")
                                                    CsvIdFileA.append(QRCode)
                        #                         # print(QRCode)
                        else:
                            if (QRCode not in CsvIdFileQL) or (not ('▲'  in panel.getAttribute('EBL2') and '△' in panel.getAttribute('EBL1') and '△' in panel.getAttribute('EBW2') and '△' in panel.getAttribute('EBW1'))):
                                panel.setAttribute('Info6', "A")
                                # 左侧板 1795.00*400.00*18.00
                                Name = panel.getAttribute('Name')
                                IdDictNo[QRCode] = f'{j.split(".")[0]}*{Name} {FinishedLength}*{FinishedWidth}*{FinishedThickness}'
                                CsvIdFileA.append(QRCode)
                                # print(QRCode)

                        PanelBasicsSql = PanelBasics[:]
                        PanelBasicsSql.append(panel.getAttribute('Info6'))
                        PanelListSql.append(PanelBasicsSql)


                IdList.append(IdFile)
                # IdDictNo
                # print(len(CsvIdFile))
                # 去重
                CsvIdFileQL = list(set(CsvIdFileQL))
                CsvIdFileQ = list(set(CsvIdFileQ))
                CsvIdFileA=list(set(CsvIdFileA))
                CsvIdFileT = list(set(CsvIdFileT))
                if Folderpath4!='':
                    self.SawCsv(CsvIdFileQ,CsvIdFileQL,CsvIdFileKc,Folderpath4,j,CsvIdFileKcFive,CsvIdFileKcSix,CsvIdFileA,CsvIdFileT)


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
                if CheckBoxf == True:
                    df.columns = ['Reference', 'Decor', 'Materialflow', 'BatchNumber', 'CustomerNumber', 'Length', 'Width',
                                  'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation',
                                  'OverSizeM1', 'EdgeMacroLM1', 'ProgramM1','BasicMacroM1', 'OverSizeM2', 'EdgeMacroRM2', 'ProgramM2']
                else:
                    df.columns = ['Reference', 'Decor', 'Materialflow', 'BatchNumber', 'CustomerNumber', 'Length','Width', 'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation','OverSizeM1', 'EdgeMacroLM1', 'ProgramM1', 'OverSizeM2', 'EdgeMacroRM2', 'ProgramM2']


                path = Folderpath2+'/'+j.split('.')[0]+'.csv'
                # test.save(path)
                df.to_csv(path,sep=';',index=False,header=False,encoding='gb18030')
                # with open(Folderpath3 + "/封边后" + fullPath.split('/')[-1], "w", encoding="UTF-8") as fs:
                with open(Folderpath3 +"/"+ fullPath.split('/')[-1], "w", encoding="UTF-8") as fs:
                    fs.write(xmldoc.toxml())
                    fs.close()
                Inserttime = datetime.datetime.now().strftime("%Y%m%d%H%M")
                for i in PanelListSql:
                    i.insert(0,Inserttime)
                    i.append(username)
                #     插入数据库
                print(PanelListSql)
                print(len(PanelListSql))
                SqlUnit.InserUnit(PanelListSql)
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
        print(Sum)
        print(fhSum)
        # return IdList
def main(Folderpath,Folderpath2,Folderpath3,Folderpath4,IdList,CheckBoxf,username,IdDictNo,sqlIp,edt_username,edt_password):
    # 创建表格对象
    test = My_sheet()
    # 创建样式对象
    styleCell = style()
    #
    ban=banding()
    SqlUnit = sqlUnit.main(sqlIp,edt_username,edt_password)
    file = ban.changeWwidth(Folderpath,Folderpath2,Folderpath3,Folderpath4,test,styleCell,SqlUnit,IdList,CheckBoxf,username,IdDictNo)
    # sqlUnit.SqlUnit.InserUnit(PaneListNUm)
    return file,SqlUnit


