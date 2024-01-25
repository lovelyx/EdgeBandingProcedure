
# Author: LiYiXiao
# 处理xml,得到相关文件

from xml.dom import minidom
import pymsgbox as mb
import logging
import traceback
import os
import pandas as pd
import lib.Sql as sqlUnit
import datetime


class banding:
    def __init__(self):
        pass

    # 修改csv文件堆垛
    @staticmethod
    def SawCsv(CsvIdFileQ, CsvIdFileQL, path, file):
        """
        处理满足条件的csv中的堆垛信息
        :param CsvIdFileQL: 上柔性封边机且宽大于长的条形码
        :param CsvIdFileQ: 上柔性封边机条形码
        :param path: 待处理csv存放路径
        :param file: 正在处理的csv的文件名
        :return:
        """
        fullPath = path + "/" + file.split(".")[0] + '.csv'  # 完整路径
        csvData = pd.read_csv(fullPath, encoding='gb18030')
        df = pd.DataFrame(csvData)
        code = df.loc[:, '条形码']
        for index, i in enumerate(code):
            if i.split('N')[1] in CsvIdFileQ:
                df.loc[index, '301'] = 'Q'
                df.loc[index, 'info7'] = 'Q'
                # if i.split('N')[1] in CsvIdFileC:
                #     length = df.loc[index,'开料长']
                #     width = df.loc[index, '开料宽']
                #     if float(length)>float(width):
                #         df.loc[index, '开料宽'] = float(df.loc[index, '开料宽']) + 1
                #     elif float(length)<float(width):
                #         df.loc[index, '开料长'] = float(df.loc[index, '开料长']) + 1
            elif i.split('N')[1] in CsvIdFileQL:
                df.loc[index, '301'] = 'QL'
                df.loc[index, 'info7'] = 'QL'
                # if i.split('N')[1] in CsvIdFileC:
                #     length = df.loc[index, '开料长']
                #     width = df.loc[index, '开料宽']
                #     if float(length) > float(width):
                #         df.loc[index, '开料宽'] = float(df.loc[index, '开料宽']) + 1
                #     elif float(length) < float(width):
                #         df.loc[index, '开料长'] = float(df.loc[index, '开料长']) + 1
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
        PanelBasics.append('')
        return

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
        Speed = 25
        PanelBasics.append(float(Speed))

        return

    # 工件旋转
    @staticmethod
    def WorkpieceRotation(i, PanelBasics, WorkpieceRotation):
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
    @staticmethod
    def LeftBandingCode(i, panel, PanelBasics, BandingCodingDict, Thick):
        """
        左机封边带编码
        :param Thick: 封边条厚度
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
                LeftBandingCodeTwo = LeftBandingCodeTwo + '+' + str(int(Thick))
                LeftBandingCodeTwo = BandingCodingDict[LeftBandingCodeTwo.strip()]
            PanelBasics.append(LeftBandingCodeTwo)

        return

    # 左机加工编码
    @staticmethod
    def LeftProcessCode(i, panel, PanelBasics, ProcessingDict, floatingCutter):
        """
        左机加工编码
        :param floatingCutter: 是否开了浮动铣刀
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: t,当开槽时控制厚边右机封
        """
        # 当开槽时
        # 控制厚边右机封
        t = 0
        StrName = ''
        if i == 2:
            EBL1 = panel.getAttribute('EBL1')
            StrName = ''
            if EBL1 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBL1:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBL1:
                    StrName = '0.5mm封边'
                StrName += '+跟踪'
            if '▲▲▲▲' in EBL1:
                StrName = StrName+'+倒棱'
        elif i == 1:
            EBW1 = panel.getAttribute('EBW1')
            StrName = ''
            if EBW1 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBW1:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBW1:
                    StrName = '0.5mm封边'
                StrName += '+跟踪'
        if floatingCutter != 2:
            StrName = StrName + '+横向'
        LeftProcessingCode = ProcessingDict[StrName]
        PanelBasics.insert(13, LeftProcessingCode)

        return t

    # 浮动铣刀
    @staticmethod
    def FloatingCutter(PanelBasics):
        """
        :param PanelBasics: 存放当前行数据的数组
        :return: 空
        """
        # milling=floatingCutter
        # if i == 1 and FinishedLength>=FinishedWidth:
        #     milling = 0
        #     PanelBasics.append(milling )
        # elif i == 1 and FinishedLength<FinishedWidth:
        #     if FinishedWidth>=1800:
        #         milling = 2
        #     else:
        #         milling = 0
        #     PanelBasics.append(milling)
        # elif i == 2 and FinishedLength>=FinishedWidth:
        #     if FinishedLength>=1800:
        #         milling = 2
        #     else:
        #         milling=0
        #     PanelBasics.append(milling)
        # elif i == 2 and FinishedLength < FinishedWidth:
        #     milling = 0
        #     PanelBasics.append(milling)
        milling = 0
        PanelBasics.append(milling)

        return milling

    # 右机封边带编码
    @staticmethod
    def RightBandingCode(i, panel, PanelBasics, BandingCodingDict, Thick):
        """
        右机分封边带编码
        :param Thick: 封边条厚度
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
                RightBandingCodeOne = RightBandingCodeOne+'+'+str(int(Thick))
                RightBandingCodeOne = BandingCodingDict[RightBandingCodeOne.strip()]
            PanelBasics.append(RightBandingCodeOne)
        elif i == 1:
            RightBandingCodeTwo = panel.getAttribute('EBW2')
            if RightBandingCodeTwo == "":
                RightBandingCodeTwo = "无封边"
                RightBandingCodeTwo = BandingCodingDict[RightBandingCodeTwo.strip()]
            else:
                RightBandingCodeTwo = RightBandingCodeTwo + '+' + str(int(Thick))
                RightBandingCodeTwo = BandingCodingDict[RightBandingCodeTwo.strip()]
            PanelBasics.append(RightBandingCodeTwo)

        return

    # 右击加工编码
    @staticmethod
    def RightProcessCode(i, panel, PanelBasics, ProcessingDict):
        """
        右机加工编码
        :param i: 第几行数据
        :param panel: xml中的panel标签
        :param PanelBasics: 存放当前行数据的数组
        :param ProcessingDict: 从数据库中获取的加工编码对照字典
        :return: RightProcessingCode 加工编码
        """
        StrName = ''
        if i == 2:
            EBL2 = panel.getAttribute('EBL2')
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
            EBW2 = panel.getAttribute('EBW2')
            StrName = ''
            if EBW2 == "":
                StrName = '无封边'
            else:
                if '▲▲▲▲' in EBW2:
                    StrName = '1.0mm封边'
                elif '△△△△' in EBW2:
                    StrName = '0.5mm封边'
                StrName += '+跟踪'
        StrName = StrName + '+横向'
        RightProcessingCode = ProcessingDict[StrName]
        PanelBasics.append(RightProcessingCode)

        return RightProcessingCode

    # 左右机预铣
    @staticmethod
    def PreMilling(i, panel, PanelBasics, CheckBox):
        """
        左右机预铣，插入指定位置
        :param CheckBox: 浮动铣刀是否开启
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

    def changeWidth(self, FolderPath, FolderPath2, FolderPath3, FolderPath4, SqlUnit, IdList, CheckBox, username):
        BandingProcessing, BandingCode = SqlUnit.selectBandingProcessingFive()
        ProcessingDict = {}
        for t in BandingProcessing:
            ProcessingDict[t[1]] = t[0]
        BandingCodingDict = {}
        for t in BandingCode:
            BandingCodingDict[t[1]] = t[2]
        # 打开文件夹得到文件夹下的所有文件
        listPath = []
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
                PanelList: list = []
                IdFile = []
                CsvIdFileQ = []
                CsvIdFileQL = []
                # CsvIdFileC = []
                for panel in PanelNodes:
                    WorkpieceRotation = 0
                    for i in range(1, 3):
                        PanelBasics = []

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

                        #  7.完工宽度
                        FinishedWidth = self.FinishedWidth(panel)
                        if FinishedWidth < 250:
                            break
                        PanelBasics.append(FinishedWidth)
                        # if FinishedLength>=1800 or FinishedWidth>=1800:
                        #     CsvIdFileC.append(QRCode)

                        # 8.完工厚度
                        FinishedThickness = self.Thickness(panel, PanelBasics)

                        # 9.数量
                        self.Number(panel, PanelBasics)
                        # 10.进给次序
                        self.FeedSequence(i, PanelBasics)

                        # 11.速度
                        self.SpeedColumn(PanelBasics)

                        # 12.工件旋转
                        WorkpieceRotation = self.WorkpieceRotation(i, PanelBasics, WorkpieceRotation)

                        # 13.左机封边带编码
                        self.LeftBandingCode(i, panel, PanelBasics, BandingCodingDict, FinishedThickness)

                        # 15.浮动铣刀
                        floatingCutter = 0
                        if CheckBox:
                            floatingCutter = self.FloatingCutter(PanelBasics)

                        # 14.左机加工编码
                        self.LeftProcessCode(i, panel, PanelBasics, ProcessingDict, floatingCutter)

                        # 16.右机封边带编码
                        self.RightBandingCode(i, panel, PanelBasics, BandingCodingDict, FinishedThickness)

                        # 17.右机加工编码
                        self.RightProcessCode(i, panel, PanelBasics, ProcessingDict)

                        # 18.左右机预铣
                        self.PreMilling(i, panel, PanelBasics, CheckBox)

                        PanelList.append(PanelBasics)

                        if i == 2:
                            if FinishedLength < FinishedWidth:
                                panel.setAttribute('Info7', "QL")
                                # if FinishedWidth >=1800:
                                # # 整数还是小数
                                #     if panel.getAttribute('CutLength').isalnum():
                                #         panel.setAttribute('CutLength', str(int(panel.getAttribute('CutLength')) + 1))
                                #     else:
                                #      panel.setAttribute('CutLength', str(float(panel.getAttribute('CutLength')) + 1))
                                CsvIds = QRCode
                                CsvIdFileQL.append(CsvIds)
                            elif FinishedLength >= FinishedWidth:
                                panel.setAttribute('Info7', "Q")
                                # if FinishedLength >= 1800:
                                #     if panel.getAttribute('CutWidth').isalnum():
                                #         panel.setAttribute('CutWidth', str(int(panel.getAttribute('CutWidth')) + 1))
                                #     else:
                                #         panel.setAttribute('CutWidth', str(float(panel.getAttribute('CutWidth')) + 1))
                                CsvIds = QRCode
                                CsvIdFileQ.append(CsvIds)
                IdList.append(IdFile)
                # 去重
                CsvIdFileQL = list(set(CsvIdFileQL))
                CsvIdFileQ = list(set(CsvIdFileQ))
                # CsvIdFileC = list(set(CsvIdFileC))
                if FolderPath4 != '':
                    self.SawCsv(CsvIdFileQ, CsvIdFileQL, FolderPath4, j)
                    fullCsvPath = FolderPath4 + "/" + j.split(".")[0] + '.csv'
                    CsvDf = pd.read_csv(fullCsvPath, encoding='gb18030')
                    CsvDf = pd.DataFrame(CsvDf)
                    InsertTime = datetime.datetime.now().strftime("%Y%m%d%H%M")
                    CsvDf.loc[:, 'time'] = InsertTime
                    CsvDf.loc[:, 'WorkShop'] = "五楼"
                    CsvDf = CsvDf.fillna(value='None')
                    CsvDf = CsvDf.values.tolist()
                    SqlUnit.InserUnitProduction(CsvDf)

                # Csv
                df = pd.DataFrame(PanelList)
                # df.columns = ['参考', '装饰', '物料流', '批号', '客户编号', '长', '宽', '厚', '数量','进给次序，通过值','速度',
                # '方向(工件旋转)','左机预铣','左机封边','左机加工','浮动铣刀','右机预铣','右机封边','右机加工']
                if CheckBox:
                    df.columns = ['Reference', 'Decor', 'MaterialFlow', 'BatchNumber', 'CustomerNumber', 'Length',
                                  'Width', 'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation',
                                  'OverSizeM1', 'EdgeMacroLM1', 'ProgramM1', 'BasicMacroM1', 'OverSizeM2',
                                  'EdgeMacroRM2',  'ProgramM2']
                else:
                    df.columns = ['Reference', 'Decor', 'MaterialFlow', 'BatchNumber', 'CustomerNumber', 'Length',
                                  'Width', 'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation',
                                  'OverSizeM1', 'EdgeMacroLM1', 'ProgramM1', 'OverSizeM2', 'EdgeMacroRM2', 'ProgramM2']

                path = FolderPath2+'/'+j.split('.')[0]+'.csv'
                df.to_csv(path, sep=';', index=False, header=False, encoding='gb18030')
                with open(FolderPath3 + "/" + fullPath.split('/')[-1], "w", encoding="UTF-8") as fs:
                    fs.write(XmlDoc.toxml())
                    fs.close()
                InsertTime = datetime.datetime.now().strftime("%Y%m%d%H%M")
                for i in PanelList:
                    i.insert(0, InsertTime)
                    i.append(username)
                #     插入数据库
                SqlUnit.InserUnitFive(PanelList)
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
        # return IdList


def main(FolderPath, FolderPath2, FolderPath3, FolderPath4, IdList, CheckBox, username, sqlIp,
         edt_username, edt_password):
    ban = banding()
    SqlUnit = sqlUnit.main(sqlIp, edt_username, edt_password)
    file = ban.changeWidth(FolderPath, FolderPath2, FolderPath3, FolderPath4, SqlUnit, IdList, CheckBox, username)
    return file, SqlUnit
