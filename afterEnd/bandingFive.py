
# Author: LiYiXiao
# 处理xml,得到相关文件

from xml.dom import minidom
import pymsgbox as mb
import logging
import traceback
import os
import pandas as pd

from afterEnd.page.My_sheet import My_sheet
from afterEnd.page.styleT import style
import lib.Sql as sqlUnit
import datetime

class banding:
    def __init__(self):
        self.changeWwidth

    # 修改csv文件堆垛
    def SawCsv(self,CsvIdFileQ,CsvIdFileQL,CsvIdFileKc,path,file,CsvIdFileC):
        """
        处理满足条件的csv中的堆垛信息
        :param CsvIdFile: 保存满足条件的二维码+
        :param path: 待处理csv存放路径
        :param file: 正在处理的csv的文件名
        :return:
        """
        # print(CsvIdFileQ)
        # print(CsvIdFileQL)
        # print(CsvIdFileKc)
        fullPath = path + "/" + file.split(".")[0] +'.csv'  # 完整路径
        # print(fullPath)
        csvData = pd.read_csv(fullPath,encoding='gb18030')
        df = pd.DataFrame(csvData)
        # print(csvData)
        code = df.loc[:,'条形码']
        for index,i in enumerate(code):
            if (i.split('N')[1] in CsvIdFileKc):
                num = df.loc[index,'堆垛']
                if num == 'AC':
                    # print(f'堆垛为{num,index}')
                    df.loc[index,'堆垛'] = 'A'
                    df.loc[index,'info6'] ='A'
            if i.split('N')[1] in CsvIdFileQ:
                df.loc[index, '301'] = 'Q'
                df.loc[index, 'info7'] = 'Q'
                if i.split('N')[1] in CsvIdFileC:
                    length = df.loc[index,'开料长']
                    width = df.loc[index, '开料宽']
                    if float(length)>float(width):
                        df.loc[index, '开料宽'] = float(df.loc[index, '开料宽']) + 1
                    elif float(length)<float(width):
                        df.loc[index, '开料长'] = float(df.loc[index, '开料长']) + 1
            elif i.split('N')[1] in CsvIdFileQL:
                df.loc[index, '301'] = 'QL'
                df.loc[index, 'info7'] = 'QL'
                if i.split('N')[1] in CsvIdFileC:
                    length = df.loc[index, '开料长']
                    width = df.loc[index, '开料宽']
                    if float(length) > float(width):
                        df.loc[index, '开料宽'] = float(df.loc[index, '开料宽']) + 1
                    elif float(length) < float(width):
                        df.loc[index, '开料长'] = float(df.loc[index, '开料长']) + 1
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
        Speed = 25
        PanelBasics.append(float(Speed))

        return

    # 工件旋转
    def WorkpieceRotation(self,i, PanelBasics,WorkpieceRotation):
        """
        工件旋转，长宽对调
        :param i: 第几行数据
        :param PanelBasics: 存放当前行数据的数组
        :param WorkpieceRotation: 工件旋转变量
        :return:
        """
        # 设置变量宽大于长时判断标准槽时，长宽，坐标也要对应变化
        if i == 1:
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
    def LeftBandingCode(self, i, panel, PanelBasics, BandingCodingDict,Thick):
        """
        左机封边带编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的编码对照字典
        :param BandingCodingDict: 板件厚度
        :return:
        """
        if i == 2:
            LeftBandingCodeOne = panel.getAttribute('EBL1')
            if LeftBandingCodeOne == "":
                LeftBandingCodeOne = "无封边"
                LeftBandingCodeOne = BandingCodingDict[LeftBandingCodeOne.strip()]
            else:
                LeftBandingCodeOne = LeftBandingCodeOne+'+'+str(int(Thick))
                LeftBandingCodeOne = BandingCodingDict[LeftBandingCodeOne.strip()]
            PanelBasics.append(LeftBandingCodeOne)
        elif i == 1:
            LeftBandingCodeTwo = panel.getAttribute('EBW1')
            if LeftBandingCodeTwo == "":
                LeftBandingCodeTwo = "无封边"
                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            else:
                LeftBandingCodeTwo = LeftBandingCodeTwo +'+'+ str(int(Thick))
                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            PanelBasics.append(LeftBandingCodeTwo)

        return


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

        strname = strname + f"+{Face}面槽"
        if panel.getAttribute('Info6') == 'AC':
            panel.setAttribute('Info6', "A")
        m.setAttribute('IsGenCode', "0")
        ids = '1' + QRCode + str(Face)
        IdFile.append(ids)
        CsvIdFileKc.append(QRCode)
        return t,strname

    # 左机加工编码
    def LeftProcessCode(self,i, panel, PanelBasics,ProcessingDict,floatingCutter):
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
            EBL1 = panel.getAttribute('EBL1')
            strname = ''
            if EBL1 == "":
                strname = '无封边'
            else:
                if '▲▲▲▲' in EBL1:
                    strname = '1.0mm封边'
                elif '△△△△' in EBL1:
                    strname = '0.5mm封边'
                strname += '+跟踪'
            if '▲▲▲▲' in EBL1:
                strname = strname+'+倒棱'
        elif i == 1:
            EBW1 = panel.getAttribute('EBW1')
            strname = ''
            if EBW1 == "":
                strname = '无封边'
            else:
                if '▲▲▲▲' in EBW1:
                    strname = '1.0mm封边'
                elif '△△△△' in EBW1:
                    strname = '0.5mm封边'
                strname += '+跟踪'
        if floatingCutter !=2:
            strname=strname+'+横向'
        LeftProcessingCode = ProcessingDict[strname]
        PanelBasics.insert(13, LeftProcessingCode)

        return t

    # 浮动铣刀
    def FloatingCutter(self,i,PanelBasics,FininshedLength,FinishedWidth,floatingCutter):
        """

        :param i: 第几行数据
        :param PanelBasics: 存放当前行数据的数组
        :param FininshedLength: 成品长
        :param FinishedWidth: 成品宽
        :return: 空
        """
        milling=floatingCutter
        if i == 1 and FininshedLength>=FinishedWidth:
            milling = 0
            PanelBasics.append(milling )
        elif i == 1 and FininshedLength<FinishedWidth:
            if FinishedWidth>=1800:
                milling = 2
            else:
                milling = 0
            PanelBasics.append(milling)
        elif i == 2 and FininshedLength>=FinishedWidth:
            if FininshedLength>=1800:
                milling = 2
            else:
                milling=0
            PanelBasics.append(milling)
        elif i == 2 and FininshedLength < FinishedWidth:
            milling = 0
            PanelBasics.append(milling)

        return milling

    # 右机封边带编码
    def RigthBandingCode(self,i, panel, PanelBasics, BandingCodingDict,Thick):
        """
        右机分封边带编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param BandingCodingDict: 从数据库中获取的封边编码对照字典
        :return: 空
        """
        if i == 2:
            RightBandingCodeOne = panel.getAttribute('EBL2')
            if RightBandingCodeOne == "":
                RightBandingCodeOne = "无封边"
                RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
            else:
                RightBandingCodeOne= RightBandingCodeOne+'+'+str(int(Thick))
                RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
            PanelBasics.append(RightBandingCodeOne)
        elif i == 1:
            RigthBandingCodeTwo = panel.getAttribute('EBW2')
            if RigthBandingCodeTwo == "":
                RigthBandingCodeTwo = "无封边"
                RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
            else:
                RigthBandingCodeTwo = RigthBandingCodeTwo +'+'+ str(int(Thick))
                RigthBandingCodeTwo = BandingCodingDict[RigthBandingCodeTwo.strip()]
            PanelBasics.append(RigthBandingCodeTwo)

        return

    # 右击加工编码
    def RigthProcessCode(self,i, panel, PanelBasics,ProcessingDict):
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
            EBL2 = panel.getAttribute('EBL2')
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
            EBW2 = panel.getAttribute('EBW2')
            strname = ''
            if EBW2 == "":
                strname = '无封边'
            else:
                if '▲▲▲▲' in EBW2:
                    strname = '1.0mm封边'
                elif '△△△△' in EBW2:
                    strname = '0.5mm封边'
                strname += '+跟踪'
        strname = strname + '+横向'
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
                        Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                        if CheckBoxf==True:
                            PanelBasics.insert(12, Pre_MillingLeft)
                        else:
                            PanelBasics.insert(12, Pre_MillingLeft)
                    if int(Faces) == 2:
                        Pre_MillingRigth = float(Edge.getAttribute('Pre_Milling'))
                        if CheckBoxf == True:
                            PanelBasics.insert(16, Pre_MillingRigth)
                        else:
                            PanelBasics.insert(15, Pre_MillingRigth)
                elif i == 2:
                    if int(Faces) == 3:
                        Pre_MillingLeft = float(Edge.getAttribute('Pre_Milling'))
                        if CheckBoxf == True:
                            PanelBasics.insert(12, Pre_MillingLeft)
                        else:
                            PanelBasics.insert(12, Pre_MillingLeft)
                    if int(Faces) == 4:
                        Pre_MillingRigth = float(Edge.getAttribute('Pre_Milling'))
                        if CheckBoxf == True:
                            PanelBasics.insert(16, Pre_MillingRigth)
                        else:
                            PanelBasics.insert(15, Pre_MillingRigth)

        return


    def changeWwidth(self,Folderpath,Folderpath2,Folderpath3,Folderpath4,test,styleCell,SqlUnit,IdList,CheckBoxf,username):
        BandingProcessing,BandingCode = SqlUnit.selectBandingProcessingFive()
        ProcessingDict={}
        for t in BandingProcessing:
            ProcessingDict[t[1]]=t[0]
        BandingCodingDict={}
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
                CsvIdFileQ = []
                CsvIdFileQL =[]
                CsvIdFileKc=[]
                CsvIdFileC=[]
                for panel in Panelnodes:
                    WorkpieceRotation=0
                    for i in range(1,3):
                        PanelBasics = []
                        Grain = panel.getAttribute('Grain')

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

                        if FinishedLength>=1800 or FinishedWidth>=1800:
                            CsvIdFileC.append(QRCode)


                        # if i == 2:
                        #     CsvIds = QRCode
                        #     if FinishedLength < FinishedWidth:
                        #         if Grain == 'L':
                        #             panel.setAttribute('Info7', "QL")
                        #             CsvIdFileQL.append(CsvIds)
                        #         elif Grain == 'W':
                        #             panel.setAttribute('Info7', "Q")
                        #             CsvIdFileQ.append(CsvIds)
                        #     elif FinishedLength >= FinishedWidth:
                        #         if Grain == 'L':
                        #             panel.setAttribute('Info7', "Q")
                        #             CsvIdFileQ.append(CsvIds)
                        #         elif Grain == 'W':
                        #             panel.setAttribute('Info7', "QL")
                        #             CsvIdFileQL.append(CsvIds)

                        # 8.完工厚度
                        FinishedThickness = self.Thickness(panel, PanelBasics)

                        # 9.数量
                        Quantity = self.Number(panel, PanelBasics)
                        # 10.进给次序
                        self.FeedSequence(i, PanelBasics)

                        # 11.速度
                        self.SpeedColumn(PanelBasics)

                        # 12.工件旋转
                        WorkpieceRotation=self.WorkpieceRotation(i,PanelBasics,WorkpieceRotation)

                        # 13.左机封边带编码
                        self.LeftBandingCode(i,panel,PanelBasics,BandingCodingDict,FinishedThickness)

                        # 15.浮动铣刀
                        floatingCutter = 0
                        if CheckBoxf == True:
                            floatingCutter =self.FloatingCutter(i, PanelBasics, FinishedLength, FinishedWidth,floatingCutter)

                        # 14.左机加工编码
                        t = self.LeftProcessCode(i,panel,PanelBasics,ProcessingDict,floatingCutter)


                        # 16.右机封边带编码
                        self.RigthBandingCode(i,panel,PanelBasics,BandingCodingDict,FinishedThickness)


                        # 17.右机加工编码
                        self.RigthProcessCode(i,panel,PanelBasics,ProcessingDict)

                        #18.左右机预铣
                        self.PreMilling(i,panel,PanelBasics,CheckBoxf)

                        PanelList.append(PanelBasics)

                        if i == 2:
                            if FinishedLength<FinishedWidth:
                                panel.setAttribute('Info7', "QL")
                                if FinishedWidth >=1800:
                                # 整数还是小数
                                    if panel.getAttribute('CutLength').isalnum():
                                        panel.setAttribute('CutLength', str(int(panel.getAttribute('CutLength')) + 1))
                                    else:
                                        panel.setAttribute('CutLength', str(float(panel.getAttribute('CutLength')) + 1))
                                CsvIds=QRCode
                                CsvIdFileQL.append(CsvIds)
                            elif FinishedLength>=FinishedWidth:
                                panel.setAttribute('Info7', "Q")
                                if FinishedLength >= 1800:
                                    if panel.getAttribute('CutWidth').isalnum():
                                        panel.setAttribute('CutWidth', str(int(panel.getAttribute('CutWidth')) + 1))
                                    else:
                                        panel.setAttribute('CutWidth', str(float(panel.getAttribute('CutWidth')) + 1))
                                CsvIds =QRCode
                                CsvIdFileQ.append(CsvIds)
                IdList.append(IdFile)
                # print(len(CsvIdFile))
                # 去重
                CsvIdFileQL = list(set(CsvIdFileQL))
                CsvIdFileQ = list(set(CsvIdFileQ))
                CsvIdFileC= list(set(CsvIdFileC))
                if Folderpath4!='':
                    self.SawCsv(CsvIdFileQ,CsvIdFileQL,CsvIdFileKc,Folderpath4,j,CsvIdFileC)


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
                # print(PanelList)
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
                # print(Folderpath3)
                with open(Folderpath3 + "/" + fullPath.split('/')[-1], "w", encoding="UTF-8") as fs:
                    fs.write(xmldoc.toxml())
                    fs.close()
                Inserttime = datetime.datetime.now().strftime("%Y%m%d%H%M")
                for i in PanelList:
                    i.insert(0, Inserttime)
                    i.append(username)
                #     插入数据库
                # SqlUnit.InserUnitFive(PanelList)
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
def main(Folderpath,Folderpath2,Folderpath3,Folderpath4,IdList,CheckBoxf,username,sqlIp,edt_username,edt_password):
    # 创建表格对象
    test = My_sheet()
    # 创建样式对象
    styleCell = style()
    #
    ban=banding()
    SqlUnit = sqlUnit.main(sqlIp,edt_username,edt_password)
    file = ban.changeWwidth(Folderpath,Folderpath2,Folderpath3,Folderpath4,test,styleCell,SqlUnit,IdList,CheckBoxf,username)
    # sqlUnit.SqlUnit.InserUnit(PaneListNUm)
    return file,SqlUnit


