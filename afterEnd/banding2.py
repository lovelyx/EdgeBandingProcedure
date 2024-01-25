# Author: LiYiXiao
# 处理xml,得到相关文件
import sys
from xml.dom import minidom
import pymsgbox as mb
import logging
import traceback
import os
import pandas as pd
import datetime
import lib.Sql as sqlUnit


sys.path.append("path")


class FourBanding:
    def __init__(self):
        pass

    # 修改csv文件堆垛
    @staticmethod
    def SawCsv(CsvIdFileQ, CsvIdFileQL, CsvIdFileKc, path, file,
               CsvIdFileKcFive, CsvIdFileKcSix, CsvIdFileA, CsvIdFileT):
        """
        处理满足条件的csv中的堆垛信息
        :param CsvIdFileQ: 保存符合Q条件的条形码
        :param CsvIdFileQL: 保存符合QL条件的条形码，大头板
        :param CsvIdFileKc: 保存封边机开槽的条形码
        :param path: 待处理csv文件路径
        :param file: 正在处理的xml文件名
        :param CsvIdFileKcFive: 五面封边机开槽板件条形码
        :param CsvIdFileKcSix: 六面封边机开槽板件条形码
        :param CsvIdFileA: 满足A标识的板件条形码
        :param CsvIdFileT: 普通先达未装刀无法打孔的板件条形码
        :return:
        """
        fullPath = path + "/" + file.split(".")[0] + '.csv'  # 完整路径
        csvData = pd.read_csv(fullPath, encoding='gb18030')
        df = pd.DataFrame(csvData)
        # print(csvData)
        code = df.loc[:, '条形码']
        for index, i in enumerate(code):
            if i.split('N')[1] in CsvIdFileKc:
                num = df.loc[index, '堆垛']
                if num == 'AC' and i.split('N')[1] not in CsvIdFileQL:
                    # print(f'堆垛为{num,index}')
                    df.loc[index, '堆垛'] = 'A'
                    df.loc[index, 'info6'] = 'A'
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
                df.loc[index, '301'] = df.loc[index, '301'] + '5'
                df.loc[index, 'info7'] = df.loc[index, 'info7'] + '5'
            if i.split('N')[1] in CsvIdFileKcSix:
                df.loc[index, '301'] = df.loc[index, '301'] + '6'
                df.loc[index, 'info7'] = df.loc[index, 'info7'] + '6'
            if i.split('N')[1] in CsvIdFileA:
                df.loc[index, '堆垛'] = 'A'
                df.loc[index, 'info6'] = 'A'
            if i.split('N')[1] in CsvIdFileT:
                df.loc[index, '堆垛'] = 'D'
                df.loc[index, 'info6'] = 'D'
        df.to_csv(fullPath, index=False, encoding='gb18030')
        return

    # 二维码，参考
    @staticmethod
    def QRCode(panel, PanelBasics):
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
    @staticmethod
    def notesOne(panel, PanelBasics):
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
    @staticmethod
    def notesTwo(PanelBasics):
        """
        备注2，未赋值
        :param PanelBasics: 存放当前行数据
        :return:备注2 目前为空
        """
        # RemarksTwo = panel.getAttribute('Info4')
        RemarksTwo = PanelBasics.append('')
        return RemarksTwo

    # 备注3
    @staticmethod
    def notesThree(panel, PanelBasics):
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
    @staticmethod
    def notesFour(panel, PanelBasics):
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
    @staticmethod
    def FinishedLength(panel):
        """
        成品长度列
        :param panel: xml中的panel标签
        :return: FinishedLength 成品长
        """
        FinishedLength = float(panel.getAttribute('Length'))
        return FinishedLength

    # 完工宽度
    @staticmethod
    def FinishedWidth(panel):
        """
        成品宽度列
        :param panel: xml中的panel标签
        :return: FinishedWidth 成品宽
        """
        FinishedWidth = float(panel.getAttribute('Width'))
        return FinishedWidth

    # 厚度
    @staticmethod
    def Thickness(panel, PanelBasics):
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
    @staticmethod
    def Number(panel, PanelBasics):
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
    @staticmethod
    def FeedSequence(i, PanelBasics):
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
    @staticmethod
    def SpeedColumn(PanelBasics):
        """
        速度列
        :param PanelBasics: 存放当前行数据
        :return:
        """
        Speed = 30
        PanelBasics.append(float(Speed))

        return

    # 工件旋转
    @staticmethod
    def WorkpieceRotation(i, PanelBasics, WorkpieceRotation, Grain):
        """
        工件旋转，长宽对调
        :param i: 第几行数据
        :param PanelBasics: 存放当前行数据的数组
        :param WorkpieceRotation: 工件旋转变量
        :param Grain: 纹路
        :return:
        """

        if i == 1:
            if Grain == 'W':
                WorkpieceRotation = 0
            else:
                WorkpieceRotation = 1
            PanelBasics.append(int(WorkpieceRotation))
        elif i == 2:
            if WorkpieceRotation == 1:
                WorkpieceRotation = 0

                PanelBasics.append(int(WorkpieceRotation))
            elif WorkpieceRotation == 0:
                WorkpieceRotation = 1
                PanelBasics.append(int(WorkpieceRotation))

        return WorkpieceRotation

    # 左机封边带编码
    @staticmethod
    def LeftBandingCode(i, panel, PanelBasics, BandingCodingDict, Grain):
        """
        左机封边带编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的编码对照字典
        :param Grain: 纹路
        :return:
        """
        # 控制厚封边在右机
        n = 0
        if i == 2:
            if Grain == 'W':
                LeftBandingCodeOne = panel.getAttribute('EBW1')
                if '▲▲▲▲' in LeftBandingCodeOne:
                    LeftBandingCodeOne = panel.getAttribute('EBW2')
                    n = 1
            else:
                LeftBandingCodeOne = panel.getAttribute('EBL1')
                if '▲▲▲▲' in LeftBandingCodeOne:
                    LeftBandingCodeOne = panel.getAttribute('EBL2')
                    n = 1

            if LeftBandingCodeOne == "":
                LeftBandingCodeOne = "无封边"
                LeftBandingCodeOne = BandingCodingDict[LeftBandingCodeOne.strip()]
            else:
                LeftBandingCodeOne = BandingCodingDict[LeftBandingCodeOne.strip()]
            PanelBasics.append(LeftBandingCodeOne)
        elif i == 1:
            if Grain == 'W':
                LeftBandingCodeTwo = panel.getAttribute('EBL1')
                if '▲▲▲▲' in LeftBandingCodeTwo:
                    LeftBandingCodeTwo = panel.getAttribute('EBL2')
                    n = 1
            else:
                LeftBandingCodeTwo = panel.getAttribute('EBW1')
                if '▲▲▲▲' in LeftBandingCodeTwo:
                    LeftBandingCodeTwo = panel.getAttribute('EBW2')
                    n = 1

            if LeftBandingCodeTwo == "":
                LeftBandingCodeTwo = "无封边"
                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            else:
                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            PanelBasics.append(LeftBandingCodeTwo)

        return n

    # 满足开槽
    @staticmethod
    def Slotting(StrName, EB2, PanelBasics, BandingCodingDict, m, QRCode, Face, IdFile, CsvIdFileKc, t):
        """
        满足为标准槽时执行的函数
        :param StrName: 保存封边厚薄的变量
        :param EB2: 当前行的封边信息
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的编码对照字典
        :param m: xml中的Machining标签
        :param QRCode: 二维码
        :param Face: 加工面
        :param IdFile: 存放二维码，与mpr文件配对
        :param CsvIdFileKc: 存放二维码，与csv中的数据进行配对
        :param t: 控制存在两边为一薄一厚时，厚边在右机加工变量
        :return: t,StrName
        """
        if StrName == '1.0mm封边':
            if '▲▲▲▲' in EB2:
                StrName = '1.0mm封边'
                PanelBasics[12] = (BandingCodingDict[EB2.strip()])
                t = 1
            elif '△△△△' in EB2:
                StrName = '0.5mm封边'
                PanelBasics[12] = (BandingCodingDict[EB2.strip()])
                t = 1
            else:
                StrName = ''
                t = 1
        StrName = StrName + f"+{Face}面槽"
        m.setAttribute('IsGenCode', "0")
        ids = '1' + QRCode + str(Face)
        IdFile.append(ids)
        CsvIdFileKc.append(QRCode)
        return t, StrName

    # 左机加工编码
    def LeftProcessCode(self, i, panel, PanelBasics, FinishedLength, FinishedWidth, BandingCodingDict, QRCode, IdFile,
                        CsvIdFileQ, ProcessingDict, n, CsvIdFileT, Grain, Machines2, CsvIdFileA, CsvIdFileQL):
        """
        左机加工编码
        :param CsvIdFileQL: 大头板
        :param CsvIdFileA: 转为A板件
        :param Machines2: 加工标签，将仅6面槽孔的板件进行A转换
        :param Grain: 纹路
        :param CsvIdFileT: 普通先达未装刀无法打孔的板件条形码
        :param n: 控制厚封边在右机
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param FinishedLength: 成品长
        :param FinishedWidth: 成品宽
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :param QRCode: 二维码
        :param IdFile: 存放二维码，与mpr文件配对
        :param CsvIdFileQ: 存放二维码，与csv中的数据进行配对,开槽
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: t,当开槽时控制厚边右机封
        """
        # 当开槽时
        # 控制厚边右机封
        t = 0
        StrName = ''
        if i == 2:
            EBL1 = ''
            EBL2 = ''
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
            StrName = ''
            if EBL1 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBL1:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBL1:
                    StrName = '0.5mm封边'
                # Length = float(panel.getAttribute('Length'))
                # Width = float(panel.getAttribute('Width'))
                Length = FinishedLength
                Width = FinishedWidth
                ID = panel.getAttribute('ID')
                info6 = panel.getAttribute('Info6')
                ebl1 = panel.getAttribute('EBL1')
                ebl2 = panel.getAttribute('EBL2')
                ebw1 = panel.getAttribute('EBW1')
                ebw2 = panel.getAttribute('EBW2')
                Machines = panel.getElementsByTagName("Machines")
                DiameterList = ['6', '5', '10', '12', '18']

                for Machining in Machines:
                    macing = Machining.getElementsByTagName("Machining")
                    for m in macing:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        Diameter = m.getAttribute('Diameter')

                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        if '▲▲▲▲' in ebl1 and '▲▲▲▲' in ebl2 and '▲▲▲▲' in ebw1 and '▲▲▲▲' in ebw2:
                            StrName = StrName.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining2 in Machines:
                                macing = Machining2.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        if Type == 3:
                            StrName = StrName.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining3 in Machines:
                                macing = Machining3.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            StrName = StrName.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining2 in Machines:
                                macing = Machining2.getElementsByTagName("Machining")
                                for q in macing:
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
                                    if abs(EndX - X) == 0 and abs(EndY - Y) == Width:
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict,
                                                                               m, QRCode, Face, IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                    else:
                                        continue
                                else:
                                    if abs(EndX - X) == Length and abs(EndY - Y) == 0:
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                    else:
                                        continue

                            else:
                                continue
                        else:
                            continue
                StrName += '+跟踪'
            if '6面槽+5面槽' in StrName:
                StrName = StrName.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in StrName:
                StrName = StrName.replace('5面槽+6面槽', '5&6面槽')
            if '▲▲▲▲' in EBL1:
                StrName = StrName + '+倒棱'
        elif i == 1:
            t = 0
            EBW1 = ''
            EBW2 = ''
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
            if EBW1 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBW1:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBW1:
                    StrName = '0.5mm封边'
                Length = float(panel.getAttribute('Length'))
                Width = float(panel.getAttribute('Width'))
                Machines = panel.getElementsByTagName("Machines")
                ID = panel.getAttribute('ID')
                info6 = panel.getAttribute('Info6')
                DiameterList = ['6', '5', '10', '12', '18']
                for Machining in Machines:
                    macing = Machining.getElementsByTagName("Machining")
                    for m in macing:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        Diameter = m.getAttribute('Diameter')
                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        if Type == 3:
                            StrName = StrName.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining2 in Machines:
                                macing = Machining2.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            StrName = StrName.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining3 in Machines:
                                macing = Machining3.getElementsByTagName("Machining")
                                for q in macing:
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
                                    if abs(EndX - X) == Length and abs(EndY - Y) == 0:
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                    else:
                                        continue
                                else:
                                    if abs(EndX - X) == 0 and abs(EndY - Y) == Width:
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7 or
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (
                                                        X == 7 or Length - X - widthTool == 7 or
                                                        Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (
                                                        X - widthTool == 7 or Length - X == 7 or
                                                        Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                    else:
                                        continue

                            else:
                                continue

                        else:
                            continue

                StrName += '+跟踪'
            if '6面槽+5面槽' in StrName:
                StrName = StrName.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in StrName:
                StrName = StrName.replace('5面槽+6面槽', '5&6面槽')
        # if Machines2:
        #     FaceOne, FaceX, FaceFive = 0, 0, 0
        #     for Machining in Machines2:
        #         macing = Machining.getElementsByTagName("Machining")
        #         for index, q in enumerate(macing):
        #             if q.getAttribute('Face') == '1' or q.getAttribute('Face') == '2':
        #                 FaceOne = 1
        #             if q.getAttribute('Face') == '5':
        #                 FaceFive = 1
        #             if q.getAttribute('Type') == '3':
        #                 FaceX = 1
        #             if index == len(macing) - 1:
        #                 if FaceOne != 1 and FaceX != 1 and FaceFive != 1 and ('+6面槽' in StrName) \
        #                         and QRCode not in CsvIdFileQL and Grain == 'L':
        #                     StrName = StrName.replace('6', '5')
        #                     panel.setAttribute('Info6', "A")
        #                     CsvIdFileA.append(QRCode)
        #                     for index2, q2 in enumerate(macing):
        #                         if q2.getAttribute('Type') == '2':
        #                             X = float(q2.getAttribute('X'))
        #                             q2.setAttribute('Face', "5")
        #                             q2.setAttribute('X', f'{FinishedLength - X}')
        #                         elif q2.getAttribute('Type') == '4':
        #                             X = float(q2.getAttribute('X'))
        #                             EndX = float(q2.getAttribute('EndX'))
        #                             q2.setAttribute('Face', "5")
        #                             q2.setAttribute('X', f'{FinishedLength - X}')
        #                             q2.setAttribute('EndX', f'{FinishedLength - EndX}')
        #                         elif q2.getAttribute('Type') == '1':
        #                             Face = q2.getAttribute('Face')
        #                             Thickness = float(panel.getAttribute('Thickness'))
        #                             Z = float(q2.getAttribute('Z'))
        #                             if Face == '3':
        #                                 q2.setAttribute('Face', "4")
        #                                 q2.setAttribute('X', "0")
        #                                 q2.setAttribute('Z', f'{Thickness - Z}')
        #                             elif Face == '4':
        #                                 q2.setAttribute('Face', "3")
        #                                 q2.setAttribute('X', f'{FinishedLength}')
        #                                 q2.setAttribute('Z', f'{Thickness - Z}')
        LeftProcessingCode = ProcessingDict[StrName]
        PanelBasics.append(LeftProcessingCode)

        return t, StrName

    # 侧板 左机加工编码
    def LeftProcessCodeSide(self, i, panel, PanelBasics, FinishedLength, FinishedWidth, BandingCodingDict, QRCode,
                            IdFile, CsvIdFileQ, ProcessingDict, n, CsvIdFileT, Grain, Machines2,
                            CsvIdFileA, CsvIdFileQL):
        """
        左机加工编码
        :param CsvIdFileQL: 大头板
        :param CsvIdFileA: 转为A板件
        :param Machines2: 加工标签，将仅6面槽孔的板件进行A转换
        :param Grain: 纹路
        :param CsvIdFileT: 普通先达未装刀无法打孔的板件条形码
        :param n: 控制厚封边在右机
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param FinishedLength: 成品长
        :param FinishedWidth: 成品宽
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :param QRCode: 二维码
        :param IdFile: 存放二维码，与mpr文件配对
        :param CsvIdFileQ: 存放二维码，与csv中的数据进行配对,开槽
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: t,当开槽时控制厚边右机封
        """
        # 当开槽时
        # 控制厚边右机封
        t = 0
        StrName = ''
        if i == 2:
            EBL1 = ''
            EBL2 = ''
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
            StrName = ''
            if EBL1 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBL1:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBL1:
                    StrName = '0.5mm封边'
                Length = FinishedLength
                Width = FinishedWidth
                ID = panel.getAttribute('ID')
                info6 = panel.getAttribute('Info6')
                Machines = panel.getElementsByTagName("Machines")
                ebl1 = panel.getAttribute('EBL1')
                ebl2 = panel.getAttribute('EBL2')
                ebw1 = panel.getAttribute('EBW1')
                ebw2 = panel.getAttribute('EBW2')
                DiameterList = ['6', '5', '10', '12', '18']

                for Machining in Machines:
                    macing = Machining.getElementsByTagName("Machining")
                    for m in macing:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        Diameter = m.getAttribute('Diameter')

                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        if '▲▲▲▲' in ebl1 and '▲▲▲▲' in ebl2 and '▲▲▲▲' in ebw1 and '▲▲▲▲' in ebw2:
                            StrName = StrName.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining2 in Machines:
                                macing = Machining2.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        if Type == 3:
                            StrName = StrName.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining3 in Machines:
                                macing = Machining3.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            StrName = StrName.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining4 in Machines:
                                macing = Machining4.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 4:
                            FiveTypeNum = 0
                            SixTypeNum = 0
                            for Machining5 in Machines:
                                macing = Machining5.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' and q.getAttribute('Type') == '4':
                                        FiveTypeNum += 1
                                    elif q.getAttribute('Face') == '6' and q.getAttribute('Type') == '4':
                                        SixTypeNum += 1
                            if FiveTypeNum >= 2 or SixTypeNum >= 2:
                                StrName = StrName.split('+')[0]
                                if ID in CsvIdFileQ:
                                    CsvIdFileQ.remove(ID)
                                panel.setAttribute('Info6', info6)
                                for Machining6 in Machines:
                                    macing = Machining6.getElementsByTagName("Machining")
                                    for q in macing:
                                        if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                            q.setAttribute('IsGenCode', "2")
                                break

                            EndX = float(m.getAttribute('EndX'))
                            EndY = float(m.getAttribute('EndY'))
                            widthTool = float(m.getAttribute('Width'))
                            ToolOffset = m.getAttribute('ToolOffset')

                            # 判断槽宽
                            if widthTool == 13:
                                if abs(EndY - Y) == 0:
                                    if 'W' in Grain:
                                        if abs(EndX - X) == 0:
                                            if Face == 5:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (
                                                            X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (
                                                            X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)

                                            elif Face == 6:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (
                                                            X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (
                                                            X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)

                                    else:
                                        if abs(EndY - Y) == 0:
                                            if Face == 5:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (
                                                            Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (
                                                            Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)

                                            elif Face == 6:
                                                # 判断走刀方向
                                                if X < EndX or Y < EndY:
                                                    if ToolOffset == '中' and (
                                                            Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                elif X > EndX or Y > EndY:
                                                    if ToolOffset == '中' and (
                                                            Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)
                                                    elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                        t, StrName = self.Slotting(StrName, EBL2, PanelBasics,
                                                                                   BandingCodingDict, m, QRCode, Face,
                                                                                   IdFile, CsvIdFileQ, t)

                            else:
                                continue
                        else:
                            continue
                StrName += '+跟踪'
            if '6面槽+5面槽' in StrName:
                StrName = StrName.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in StrName:
                StrName = StrName.replace('5面槽+6面槽', '5&6面槽')
            if '▲▲▲▲' in EBL1:
                StrName = StrName + '+倒棱'
        elif i == 1:
            t = 0
            EBW1 = ''
            EBW2 = ''
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
            StrName = ''
            if EBW1 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBW1:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBW1:
                    StrName = '0.5mm封边'
                Length = float(panel.getAttribute('Length'))
                Width = float(panel.getAttribute('Width'))
                Machines = panel.getElementsByTagName("Machines")
                ID = panel.getAttribute('ID')
                info6 = panel.getAttribute('Info6')
                DiameterList = ['6', '5', '10', '12', '18']
                for Machining in Machines:
                    macing = Machining.getElementsByTagName("Machining")
                    for m in macing:
                        Type = int(m.getAttribute('Type'))
                        Face = int(m.getAttribute('Face'))
                        X = float(m.getAttribute('X'))
                        Y = float(m.getAttribute('Y'))
                        Diameter = m.getAttribute('Diameter')
                        if Type == 3:
                            StrName = StrName.split('+')[0]
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', info6)
                            for Machining2 in Machines:
                                macing = Machining2.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 2 and Diameter not in DiameterList:
                            StrName = StrName.split('+')[0]
                            # print(Diameter)
                            if ID in CsvIdFileQ:
                                CsvIdFileQ.remove(ID)
                            panel.setAttribute('Info6', 'D')
                            CsvIdFileT.append(QRCode)
                            for Machining3 in Machines:
                                macing = Machining3.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' or q.getAttribute('Face') == '6':
                                        q.setAttribute('IsGenCode', "2")
                            break
                        elif Type == 4:
                            FiveTypeNum = 0
                            SixTypeNum = 0
                            for Machining4 in Machines:
                                macing = Machining4.getElementsByTagName("Machining")
                                for q in macing:
                                    if q.getAttribute('Face') == '5' and q.getAttribute('Type') == '4':
                                        FiveTypeNum += 1
                                    elif q.getAttribute('Face') == '6' and q.getAttribute('Type') == '4':
                                        SixTypeNum += 1
                            if FiveTypeNum >= 2 or SixTypeNum >= 2:
                                StrName = StrName.split('+')[0]
                                if ID in CsvIdFileQ:
                                    CsvIdFileQ.remove(ID)
                                panel.setAttribute('Info6', info6)
                                for Machining5 in Machines:
                                    macing = Machining5.getElementsByTagName("Machining")
                                    for q in macing:
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
                                    if abs(EndY - Y) == 0:
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        Y - widthTool / 2 == 7 or Width - Y - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (Y - widthTool == 7 or Width - Y == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (Y == 7 or Width - Y - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                else:
                                    if abs(EndX - X) == 0:
                                        if Face == 5:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)

                                        elif Face == 6:
                                            # 判断走刀方向
                                            if X < EndX or Y < EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (X - widthTool == 7 or Length - X == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (X == 7 or Length - X - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                            elif X > EndX or Y > EndY:
                                                if ToolOffset == '中' and (
                                                        X - widthTool / 2 == 7 or Length - X - widthTool / 2 == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '右' and (X == 7 or Length - X - widthTool == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                                elif ToolOffset == '左' and (X - widthTool == 7 or Length - X == 7):
                                                    t, StrName = self.Slotting(StrName, EBW2, PanelBasics,
                                                                               BandingCodingDict, m, QRCode, Face,
                                                                               IdFile, CsvIdFileQ, t)
                                        # # 判断通槽
                            else:
                                continue

                        else:
                            continue

                StrName += '+跟踪'
            if '6面槽+5面槽' in StrName:
                StrName = StrName.replace('6面槽+5面槽', '5&6面槽')
            elif '5面槽+6面槽' in StrName:
                StrName = StrName.replace('5面槽+6面槽', '5&6面槽')

        # if Machines2:
        #     FaceOne, FaceX, FaceFive = 0, 0, 0
        #     for Machining in Machines2:
        #         macing = Machining.getElementsByTagName("Machining")
        #         for index, q in enumerate(macing):
        #             if q.getAttribute('Face') == '1' or q.getAttribute('Face') == '2':
        #                 FaceOne = 1
        #             if q.getAttribute('Face') == '5':
        #                 FaceFive = 1
        #             if q.getAttribute('Type') == '3':
        #                 FaceX = 1
        #             if index == len(macing) - 1:
        #                 if FaceOne != 1 and FaceX != 1 and FaceFive != 1 and ('+6面槽' in StrName) \
        #                         and QRCode not in CsvIdFileQL and Grain == 'L':
        #                     StrName = StrName.replace('6', '5')
        #                     panel.setAttribute('Info6', "A")
        #                     CsvIdFileA.append(QRCode)
        #                     for index2, q2 in enumerate(macing):
        #                         if q2.getAttribute('Type') == '2':
        #                             X = float(q2.getAttribute('X'))
        #                             q2.setAttribute('Face', "5")
        #                             q2.setAttribute('X', f'{FinishedLength - X}')
        #                         elif q2.getAttribute('Type') == '4':
        #                             X = float(q2.getAttribute('X'))
        #                             EndX = float(q2.getAttribute('EndX'))
        #                             q2.setAttribute('Face', "5")
        #                             q2.setAttribute('X', f'{FinishedLength - X}')
        #                             q2.setAttribute('EndX', f'{FinishedLength - EndX}')
        #                         elif q2.getAttribute('Type') == '1':
        #                             Face = q2.getAttribute('Face')
        #                             Thickness = float(panel.getAttribute('Thickness'))
        #                             Z = float(q2.getAttribute('Z'))
        #                             if Face == '3':
        #                                 q2.setAttribute('Face', "4")
        #                                 q2.setAttribute('X', "0")
        #                                 q2.setAttribute('Z', f'{Thickness-Z}')
        #                             elif Face == '4':
        #                                 q2.setAttribute('Face', "3")
        #                                 q2.setAttribute('X', f'{FinishedLength}')
        #                                 q2.setAttribute('Z', f'{Thickness - Z}')
        LeftProcessingCode = ProcessingDict[StrName]
        PanelBasics.append(LeftProcessingCode)

        return t, StrName

    # 浮动铣刀
    @staticmethod
    def FloatingCutter(i, PanelBasics, FinishedLength, FinishedWidth):
        """
        :param i: 第几行数据
        :param PanelBasics: 存放当前行数据的数组
        :param FinishedLength: 成品长
        :param FinishedWidth: 成品宽
        :return: 空
        """
        if i == 1 and FinishedLength >= FinishedWidth:
            milling = 0
            PanelBasics.append(milling)
        elif i == 1 and FinishedLength < FinishedWidth:
            milling = 2
            PanelBasics.append(milling)
        elif i == 2 and FinishedLength >= FinishedWidth:
            milling = 2
            PanelBasics.append(milling)
        elif i == 2 and FinishedLength < FinishedWidth:
            milling = 0
            PanelBasics.append(milling)

        return

    # 右机封边带编码
    @staticmethod
    def RightBandingCode(i, panel, PanelBasics, BandingCodingDict, n, Grain):
        """
        右机分封边带编码
        :param Grain: 纹路
        :param n: 厚封边在右机
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :return: 空
        """
        RightBandingCodeOne = ''
        RightBandingCodeTwo = ''
        if i == 2:
            if Grain == 'W':
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
            if Grain == 'W':
                if n == 1:
                    RightBandingCodeTwo = panel.getAttribute('EBL1')
                elif n == 0:
                    RightBandingCodeTwo = panel.getAttribute('EBL2')
            else:
                if n == 1:
                    RightBandingCodeTwo = panel.getAttribute('EBW1')
                elif n == 0:
                    RightBandingCodeTwo = panel.getAttribute('EBW2')
            # RightBandingCodeTwo = panel.getAttribute('EBW2')
            if RightBandingCodeTwo == "":
                RightBandingCodeTwo = "无封边"
                RightBandingCodeTwo = BandingCodingDict[RightBandingCodeTwo.strip()]
            else:
                RightBandingCodeTwo = BandingCodingDict[RightBandingCodeTwo.strip()]
            # RightBandingCodeTwo = BandingCodingDict[RightBandingCodeTwo.strip()]
            PanelBasics.append(RightBandingCodeTwo)

        return

    # 右击加工编码
    @staticmethod
    def RightProcessCode(i, panel, PanelBasics, BandingCodingDict, t, ProcessingDict, CheckBox, n, Grain):
        """
        右机加工编码
        :param Grain: 纹路
        :param n: 控制厚封边在右机
        :param CheckBox: 是否开浮动铣刀参数，目前统一为true
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :param t:  控制存在两边为一薄一厚时，厚边在右机加工变量
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: RightProcessingCode 加工编码
        """
        StrName = ''
        if i == 2:
            EBL2 = ''
            if t == 0:
                if Grain == 'W':
                    if n == 1:
                        EBL2 = panel.getAttribute('EBW1')
                    elif n == 0:
                        EBL2 = panel.getAttribute('EBW2')
                else:
                    if n == 1:
                        EBL2 = panel.getAttribute('EBL1')
                    elif n == 0:
                        EBL2 = panel.getAttribute('EBL2')
            elif t == 1:
                if Grain == 'W':
                    EBL2 = panel.getAttribute('EBW1')
                else:
                    EBL2 = panel.getAttribute('EBL1')
                if CheckBox:
                    PanelBasics[15] = BandingCodingDict[EBL2.strip()]
                else:
                    PanelBasics[14] = BandingCodingDict[EBL2.strip()]
            StrName = ''
            if EBL2 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBL2:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBL2:
                    StrName = '0.5mm封边'
                StrName += '+跟踪'
            if '▲▲▲▲' in EBL2:
                StrName += '+倒棱'
        elif i == 1:
            EBW2 = ''
            if t == 0:
                if Grain == 'W':
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
                if Grain == 'W':
                    EBW2 = panel.getAttribute('EBL1')
                else:
                    EBW2 = panel.getAttribute('EBW1')
                if CheckBox:
                    PanelBasics[15] = BandingCodingDict[EBW2.strip()]
                else:
                    PanelBasics[14] = BandingCodingDict[EBW2.strip()]

            StrName = ''
            if EBW2 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBW2:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBW2:
                    StrName = '0.5mm封边'

                StrName += '+跟踪'
        RightProcessingCode = ProcessingDict[StrName]
        PanelBasics.append(RightProcessingCode)

        return RightProcessingCode

    # 左右机预铣
    @staticmethod
    def PreMilling(i, panel, PanelBasics, CheckBox):
        """
        左右机预铣，插入指定位置
        :param CheckBox: 浮动铣刀参数
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
                        if CheckBox:
                            PanelBasics.insert(12, Pre_MillingLeft)
                        else:
                            PanelBasics.insert(12, Pre_MillingLeft)
                    if int(Faces) == 2:
                        # Pre_MillingRight = float(Edge.getAttribute('Pre_Milling'))
                        Pre_MillingRight = 0.5
                        if CheckBox:
                            PanelBasics.insert(16, Pre_MillingRight)
                        else:
                            PanelBasics.insert(15, Pre_MillingRight)
                elif i == 2:
                    if int(Faces) == 3:
                        # Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                        Pre_MillingLeft = 0.5
                        if CheckBox:
                            PanelBasics.insert(12, Pre_MillingLeft)
                        else:
                            PanelBasics.insert(12, Pre_MillingLeft)
                    if int(Faces) == 4:
                        # Pre_MillingRight = float(Edge.getAttribute('Pre_Milling'))
                        Pre_MillingRight = 0.5
                        if CheckBox:
                            PanelBasics.insert(16, Pre_MillingRight)
                        else:
                            PanelBasics.insert(15, Pre_MillingRight)

        return

    def changeWidth(self, FolderPath, FolderPath2, FolderPath3, FolderPath4, SqlUnit, IdList, CheckBox, username,
                    IdDictNo):
        BandingProcessing, BandingCode = SqlUnit.selectBandingProcessing()
        # BandingProcessingLeft,BandingProcessingRight,BandingCodeLeft,
        # BandingCodeRight = SqlUnit.selectBandingProcessing()
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
            BandingCodingDict[t[1]] = t[2]
        # 打开文件夹得到文件夹下的所有文件
        listPath = []
        Sum = 0
        xml_list = os.listdir(FolderPath)
        for j in xml_list:  # 遍历所有xml文件
            # noinspection PyBroadException
            try:
                fullPath = FolderPath + "/" + j  # 完整路径
                listPath.append(j)
                # 打开文件
                XmlDoc = minidom.parse(fullPath)
                # 找到Machining节点（标签）
                PanelNodes = XmlDoc.getElementsByTagName("Panel")
                PanelList = []
                IdFile = []
                CsvIdFileQ = []
                CsvIdFileQL = []
                CsvIdFileKc = []
                CsvIdFileKcFive = []
                CsvIdFileKcSix = []
                CsvIdFileA = []
                CsvIdFileT = []

                # 存入数据库
                PanelListSql: list = []
                for panel in PanelNodes:
                    Sum += 1
                    WorkpieceRotation = 0
                    for i in range(1, 3):
                        PanelBasics = []
                        Grain = panel.getAttribute('Grain')
                        PanelName = panel.getAttribute('Name')
                        Machines2 = panel.getElementsByTagName("Machines")

                        # 1.二维码，参考
                        QRCode = self.QRCode(panel, PanelBasics)

                        # 2.备注1，装饰
                        self.notesOne(panel, PanelBasics)

                        # 3.备注2，物料流
                        self.notesTwo(PanelBasics)

                        # 4.备注3，批号
                        self.notesThree(panel, PanelBasics)

                        # 5.备注4，客户编号
                        self.notesFour(panel, PanelBasics)

                        # 6.完工长度
                        FinishedLength = self.FinishedLength(panel)
                        if FinishedLength < 250:
                            break
                        PanelBasics.append(FinishedLength)

                        # 7.完工宽度
                        FinishedWidth = self.FinishedWidth(panel)
                        if FinishedWidth < 250:
                            break
                        PanelBasics.append(FinishedWidth)

                        if FinishedLength < FinishedWidth and i == 1:
                            if Grain == 'L':
                                panel.setAttribute('Info7', "QL")
                                CsvIds = QRCode
                                CsvIdFileQL.append(CsvIds)
                            elif Grain == 'W':
                                panel.setAttribute('Info7', "Q")
                                CsvIds = QRCode
                                CsvIdFileQ.append(CsvIds)
                        elif FinishedLength >= FinishedWidth and i == 1:
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
                        self.Number(panel, PanelBasics)
                        # 10.进给次序
                        self.FeedSequence(i, PanelBasics)

                        # 11.速度
                        self.SpeedColumn(PanelBasics)

                        # 12.工件旋转
                        WorkpieceRotation = self.WorkpieceRotation(i, PanelBasics, WorkpieceRotation, Grain)

                        # 13.左机封边带编码
                        n = self.LeftBandingCode(i, panel, PanelBasics, BandingCodingDict, Grain)

                        # 14.左机加工编码
                        if ('侧板' in PanelName) or ('左侧' in PanelName) or ('右侧' in PanelName):
                            t, StrName = self.LeftProcessCodeSide(i, panel, PanelBasics, FinishedLength, FinishedWidth,
                                                                  BandingCodingDict, QRCode, IdFile, CsvIdFileKc,
                                                                  BandingProcessingDict, n, CsvIdFileT, Grain,
                                                                  Machines2, CsvIdFileA, CsvIdFileQL)
                        else:
                            t, StrName = self.LeftProcessCode(i, panel, PanelBasics, FinishedLength, FinishedWidth,
                                                              BandingCodingDict, QRCode, IdFile, CsvIdFileKc,
                                                              BandingProcessingDict, n, CsvIdFileT, Grain, Machines2,
                                                              CsvIdFileA, CsvIdFileQL)

                        # 15.浮动铣刀
                        if CheckBox:
                            self.FloatingCutter(i, PanelBasics, FinishedLength, FinishedWidth, )

                        # 16.右机封边带编码
                        self.RightBandingCode(i, panel, PanelBasics, BandingCodingDict, n, Grain)

                        # 17.右机加工编码
                        self.RightProcessCode(i, panel, PanelBasics, BandingCodingDict, t, BandingProcessingDict,
                                              CheckBox, n, Grain)

                        # 18.左右机预铣
                        self.PreMilling(i, panel, PanelBasics, CheckBox)

                        PanelList.append(PanelBasics)

                        if '+5' in StrName:
                            panel.setAttribute('Info7', panel.getAttribute('Info7') + '5')
                            CsvIdFileKcFive.append(QRCode)
                        # print(StrName)
                        if '6' in StrName:
                            panel.setAttribute('Info7', panel.getAttribute('Info7') + '6')
                            CsvIdFileKcSix.append(QRCode)
                        if QRCode in CsvIdFileKc:
                            if panel.getAttribute('Info6') == 'AC' and QRCode not in CsvIdFileQL:
                                panel.setAttribute('Info6', "A")
                            elif panel.getAttribute('Info6') == 'A':
                                panel.setAttribute('Info6', "A")
                            else:
                                panel.setAttribute('Info6', "E")

                        if Machines2:
                            if i == 2:
                                FaceOne, FaceSixKc, FaceSixKcAll, FaceFiveKcAll, FaceSixKo, FaceFiveKc, FaceFiveKo,\
                                    FaceX, FaceDT, FaceEB, FaceThree, FaceFive,\
                                    FaceSix = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                                for Machining in Machines2:
                                    macing = Machining.getElementsByTagName("Machining")
                                    for index, q in enumerate(macing):

                                        if q.getAttribute('Face') == '1' or q.getAttribute('Face') == '2':
                                            FaceOne = 1

                                        if q.getAttribute('Face') == '6':
                                            pass
                                            # FaceSix = 1

                                        if q.getAttribute('Face') == '6' and q.getAttribute('Type') == '4':
                                            pass
                                            # FaceSixKcAll = 1

                                        if q.getAttribute('Face') == '6' and q.getAttribute('Type') == '2':
                                            FaceSixKo = 1

                                        if q.getAttribute('Face') == '5':
                                            pass
                                            # FaceFive = 1

                                        if q.getAttribute('Face') == '5' and q.getAttribute('Type') == '4':
                                            pass
                                            # FaceFiveKcAll = 1

                                        if q.getAttribute('Face') == '5' and q.getAttribute('Type') == '2':
                                            pass
                                            # FaceFiveKo = 1

                                        if q.getAttribute('Type') == '3':
                                            FaceX = 1

                                        if QRCode in CsvIdFileQL:
                                            FaceDT = 1

                                        if '▲' in panel.getAttribute('EBL2') and '△' in panel.getAttribute(
                                                'EBL1') and '△' in panel.getAttribute('EBW2') and\
                                                '△' in panel.getAttribute('EBW1'):
                                            FaceEB = 1

                                        if q.getAttribute('Face') == '3' or q.getAttribute('Face') == '4':
                                            FaceThree = 1

                                        if q.getAttribute('Face') == '6' and q.getAttribute(
                                                'Type') == '4' and q.getAttribute('IsGenCode') != '0':
                                            FaceSixKc = 1
                                        if q.getAttribute('Face') == '5' and q.getAttribute(
                                                'Type') == '4' and q.getAttribute('IsGenCode') != '0':
                                            FaceFiveKc = 1

                                        if index == len(macing) - 1:
                                            if '层板' in panel.getAttribute('Name') and float(
                                                    panel.getAttribute('Length')) >= 764:
                                                if FaceOne != 1 and FaceX != 1 and FaceFiveKc != 1 and FaceSixKc != 1\
                                                        and (FaceSixKc != 1 and FaceSixKo != 1) and (
                                                        FaceFiveKc != 1 and FaceSixKo != 1 and FaceSixKc != 1) and (
                                                        FaceSixKc != 1 and FaceSixKo != 1 and FaceFiveKc != 1)\
                                                        and FaceDT != 1 and FaceThree == 1 and FaceEB != 1:
                                                    panel.setAttribute('Info6', "A")
                                                    CsvIdFileA.append(QRCode)
                                                    # print(QRCode)
                                            else:
                                                if FaceOne != 1 and FaceX != 1 and FaceFiveKc != 1 and FaceSixKc != 1\
                                                    and (FaceSixKc != 1 and FaceSixKo != 1) and \
                                                    (FaceFiveKc != 1 and FaceSixKo != 1 and FaceSixKc != 1) and \
                                                    (FaceSixKc != 1 and FaceSixKo != 1 and FaceFiveKc != 1) and \
                                                    FaceDT != 1 and ('层板' not in panel.getAttribute('Name')) and \
                                                        FaceEB != 1:
                                                    panel.setAttribute('Info6', "A")
                                                    CsvIdFileA.append(QRCode)
                        else:
                            if (QRCode not in CsvIdFileQL) or (not (
                                    '▲' in panel.getAttribute('EBL2') and '△' in panel.getAttribute('EBL1') and
                                    '△' in panel.getAttribute('EBW2') and '△' in panel.getAttribute('EBW1'))):
                                panel.setAttribute('Info6', "A")
                                # 左侧板 1795.00*400.00*18.00
                                Name = panel.getAttribute('Name')
                                IdDictNo[
                                    QRCode] = f'{j.split(".")[0]}*{Name} {FinishedLength}*{FinishedWidth}*' \
                                              f'{FinishedThickness}'
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
                CsvIdFileA = list(set(CsvIdFileA))
                CsvIdFileT = list(set(CsvIdFileT))
                if FolderPath4 != '':
                    self.SawCsv(CsvIdFileQ, CsvIdFileQL, CsvIdFileKc, FolderPath4, j, CsvIdFileKcFive, CsvIdFileKcSix,
                                CsvIdFileA, CsvIdFileT)
                    fullCsvPath = FolderPath4 + "/" + j.split(".")[0] + '.csv'
                    CsvDf = pd.read_csv(fullCsvPath, encoding='gb18030')
                    csvDf = pd.DataFrame(CsvDf)
                    InsertTime = datetime.datetime.now().strftime("%Y%m%d%H%M")
                    csvDf.loc[:, 'time'] = InsertTime
                    csvDf.loc[:, 'WorkShop'] = "四楼"
                    csvDf = csvDf.fillna(value='None')
                    # csvDf = csvDf.replace(numpy.nan, None)
                    csvDf = csvDf.values.tolist()
                    SqlUnit.InserUnitProduction(csvDf)

                # Csv
                df = pd.DataFrame(PanelList)
                # df.columns = ['参考', '装饰', '物料流', '批号', '客户编号', '长', '宽', '厚', '数量','进给次序，通过值','速度',
                # '方向(工件旋转)','左机预铣','左机封边','左机加工','浮动铣刀','右机预铣','右机封边','右机加工']
                if CheckBox:
                    df.columns = ['Reference', 'Decor', 'MaterialFlow', 'BatchNumber', 'CustomerNumber', 'Length',
                                  'Width',
                                  'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation',
                                  'OverSizeM1', 'EdgeMacroLM1', 'ProgramM1', 'BasicMacroM1', 'OverSizeM2',
                                  'EdgeMacroRM2', 'ProgramM2']
                else:
                    df.columns = ['Reference', 'Decor', 'MaterialFlow', 'BatchNumber', 'CustomerNumber', 'Length',
                                  'Width', 'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation',
                                  'OverSizeM1', 'EdgeMacroLM1', 'ProgramM1', 'OverSizeM2', 'EdgeMacroRM2', 'ProgramM2']

                path = FolderPath2 + '/' + j.split('.')[0] + '.csv'
                df.to_csv(path, sep=';', index=False, header=False, encoding='gb18030')
                with open(FolderPath3 + "/" + fullPath.split('/')[-1], "w", encoding="UTF-8") as fs:
                    fs.write(XmlDoc.toxml())
                    fs.close()
                InsertTime = datetime.datetime.now().strftime("%Y%m%d%H%M")
                for i in PanelListSql:
                    i.insert(0, InsertTime)
                    i.append(username)
                #     插入数据库
                SqlUnit.InserUnit(PanelListSql)
                yield j
            except:
                logging.basicConfig(filename='log.txt', level=logging.DEBUG,
                                    format='%(ascTime)s - %(levelName)s - %(message)s')
                # 方案一，自己定义一个文件，自己把错误堆栈信息写入文件。
                errorFile = open('log.txt', 'a')
                errorFile.write(traceback.format_exc())
                errorFile.close()
                yield f"{j}文件处理失败，请联系管理员"
                # tkinter.messagebox.showerror(title='错误',message=f"{j.split('.')[0]}文件处理失败，请联系管理员")
                mb.alert(f"{j}文件处理失败，请联系管理员", '报错')
                # win32api.MessageBox(0,f"{j.split('.')[0]}文件处理失败，请联系管理员","错误",win32con.MB_OK)
        print(Sum)
        # return IdList


def main(FolderPath, FolderPath2, FolderPath3, FolderPath4, IdList, CheckBox, username, IdDictNo, sqlIp, edt_username,
         edt_password):
    ban = FourBanding()
    SqlUnit = sqlUnit.main(sqlIp, edt_username, edt_password)
    file = ban.changeWidth(FolderPath, FolderPath2, FolderPath3, FolderPath4, SqlUnit, IdList, CheckBox, username,
                           IdDictNo)
    return file, SqlUnit
