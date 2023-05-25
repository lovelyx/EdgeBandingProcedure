import xlrd
import xlwt
from xlwt import Style
import datetime as dt
from afterEnd.page import styleT
import datetime
import lib.Sql as sqlUnit

class My_sheet(object):
    def __init__(self, sheet_name='数据', re_write=True):
        '''
        自定义类说明：
        :param sheet_name:默认sheet表对象名称，默认值为 'sheet_1'
        :param re_write: 单元格重写写功能默认开启
        '''
        # self.work_book = xlwt.Workbook()
        # self.sheet = self.work_book.add_sheet(sheet_name,cell_overwrite_ok=re_write)
        # self.col_data = {}
        # self.book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        self.book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        self.sheet = self.book.add_sheet(sheet_name, cell_overwrite_ok=re_write)
        self.sheet.col(0).width = 256 * 13
        self.sheet.col(1).width = 256 * 30
        self.sheet.col(2).width = 256 * 47
        self.sheet.col(3).width = 256 * 20
        self.sheet.col(4).width = 256 * 25
        self.sheet.col(5).width = 256 * 25
        self.sheet.col(6).width = 256 * 30
        self.sheet.col(7).width = 256 * 30
        self.sheet.col(8).width = 256 * 15
        self.sheet.col(9).width = 256 * 12
        self.listSum=[0,0,0,0]
        self.listSum2=[0,0]
        self.MaterialresultSum=0
        self.fileName=""
        # col_data字典用来收集所有写入sheet表的列信息，为后续设置自动调整列宽用。
        self.col_data = {}
    def save(self, file_name):
        self.book.save(file_name)

    def set_panes_frozen(self, i, j, z):
        # 设置冻结窗口
        # 设置冻结为真
        self.sheet.set_panes_frozen(i)
        # 水平冻结
        self.sheet.set_horz_split_pos(j)
        # 垂直冻结
        self.sheet.set_vert_split_pos(z)

    def diy_style(self,font_name,font_height,bold = True,horz = 2):
        '''
        创建单元格格式：（默认垂直居中）
        :param font_name: 字体名称
        :param font_height: 字体高度
        :param bold: 默认加粗
        :param horz: 水平对齐方式，默认水平居中：2，左对齐：1，右对齐：3
        :return: 返回设置好的格式
        '''
        style = xlwt.XFStyle()
        # 字体设置
        font = style.font
        font.name = font_name
        font.height = font_height*20
        font.bold = bold
        # 对齐方式
        alignment = style.alignment
        # 水平居中
        alignment.horz = horz
        # 垂直居中
        alignment.vert = 1

        return style

    def sum(self,data2):
        '''
        累加余料相关数据
        :param data2: 导入的一行数据
        '''
        for i in range(0,4):
            self.listSum[i] = self.listSum[i]+data2[i+2]
        # self.listSum[0] = self.listSum[0] + data2[2]
        # self.listSum[1] = self.listSum[1 ] + data2[3]
        # self.listSum[2] = self.listSum[2] + data2[4]
        # self.listSum[3] = self.listSum[3] + data2[5]

    def sum2(self, data2):
        '''
        累加余料相关数据
        :param data2: 导入的一行数据
        '''
        # for i in range(0, 4):
        #     self.listSum[i] = self.listSum[i] + data2[i + 4]
        self.listSum2[0] = self.listSum2[0] + data2[4]
        self.listSum2[1] = self.listSum2[1 ] + data2[6]

    def Day(self,numList,numListDay,filname):
        self.fileName=filname
        j = 0
        numDay2 = 1
        while True:
        # for i in range(0,len(numListDay)):
            # if float(numListDay[j + 1]) == float(numListDay[j]):
            if j == len(numListDay) - 1:
                numDay = 0
                for i in range(0, j + 1):
                    numDay = numList[i] + numDay
                    # print(numDay)
                self.totalDay(numDay2, numDay)
                self.write_merge(numDay2, numDay, 9, 9, self.MaterialresultSum, styleT.style.style2(self=""))
                self.write_merge(numDay2, numDay, 0, 0, numListDay[j], styleT.style.style0(self=""))
                break
            if numListDay[j + 1] == numListDay[j]:
                j = j + 1
            else:
                numDay = 0
                for i in range(0, j + 1):
                    numDay = numList[i] + numDay
                    # print(numDay)
                self.write_merge(numDay2, numDay, 0, 0, numListDay[j], styleT.style.style0(self=""))
                # self.write(timeIndex, 0, numListDay[timeIndex], style.style.style0(self=""))
                self.totalDay(numDay2,numDay)
                self.write_merge(numDay2, numDay, 9, 9, self.MaterialresultSum, styleT.style.style2(self=""))
                numDay2 = numDay + 1
                j = j + 1
                continue


    def totalFile(self, num, start_row, style=styleT.style.style3(self='')):
        '''
        向表中添加合计数据
        :param num: 表中一共又多少行
        :param start_row:起始行号
        '''
        time = dt.datetime.now().strftime("%m")
        time=time+"月"
        self.write(start_row + num, 1, time, style)
        self.write(start_row + num, 2, "合计", style)
        for i in range(3, 7):
            self.write(start_row + num, i, self.listSum[i - 3], style)
        UtilizationSum = (float(self.listSum[1]) / float(self.listSum[3]))
        self.MaterialresultSum = '{:.2%}'.format(UtilizationSum)
        self.write(start_row + num,7,self.MaterialresultSum, style)

    def read_excel(self,row):
        """
        从xls文件中读取数据
        """
        time = dt.datetime.now().strftime("%m-%d")
        filename = self.fileName
        workbook = xlrd.open_workbook(filename)
        # print(workbook)
        # 可以使用workbook对象的sheet_names()方法获取到excel文件中哪些表有数据
        # print(workbook.sheet_names())
        # 可以通过sheet_by_index()方法或sheet_by_name()方法获取到一张表，返回一个对象
        # table = workbook.sheet_by_index(0)
        # print(table)
        table = workbook.sheet_by_name('优化数据')
        # print(table)
        # 通过nrows和ncols获取到表格中数据的行数和列数
        # rows = table.nrows
        # cols = table.ncols
        # 可以通过row.values()按行获取数据，返回一个列表，也可以按列
        # for row in range(rows):
        row_data = table.row_values(row)
        return row_data
            # print(''.join(row_data))
        # 也可以根据单元格获取每一个单元格的数据
        # for row in range(rows):
        #     for col in range(cols):
        #         data = table.cell(row, col).value
        #         print(data, end=' ')

    def totalDay(self, num, start_row, style=styleT.style.style3(self='')):
        '''
        向表中添加合计数据
        :param num: 起始行
        :param start_row:终点行
        '''
        time = dt.datetime.now().strftime("%m-%d")
        self.listSum2=[0,0]
        for i in range(num,start_row+1):
            self.sum2(self.read_excel(i))
        # print(self.listSum2)
        UtilizationSum = (float(self.listSum2[0]) / float(self.listSum2[1]))
        # print(UtilizationSum)
        self.MaterialresultSum = '{:.2%}'.format(UtilizationSum)
        # self.write(start_row + num,7,self.MaterialresultSum, style)

    def write_merge(self, rowStart, rowEnd, colStart, colEnd, label, style=styleT.style.style0(self="")):
        '''
        合并单元格
        :param rowStart: 开始行
        :param colStart: 开始列
        :param rowEnd: 结束行
        :param colEnd: 结束列
        :param label: 写入数据
        '''
        self.sheet.write_merge(rowStart,rowEnd,colStart,colEnd,label,style)


    def write(self, row, col, label, style=Style.default_style):
        '''
        在默认sheet表对象一个单元格内写入数据
        :param row: 写入行
        :param col: 写入列
        :param label: 写入数据
        '''
        self.sheet.write(row, col, label, style)

        # 将列数据加入到col_data字典中
        if col not in self.col_data.keys():
            self.col_data[col] = []
            self.col_data[col].append(label)
        else:
            self.col_data[col].append(label)

    def write_row(self, start_row, start_col, date_list,
                  style=Style.default_style):
        '''
        按行写入一行数据
        :param start_row:写入行序号
        :param start_col: 写入列序号
        :param date_list: 写入数据：列表
        :return: 返回行对象
        '''
        # print(date_list)
        for col, label in enumerate(date_list):
            self.write(start_row, start_col + col, label, style)
        #  冻结为真
        self.sheet.set_panes_frozen('1')
        # 首行冻结
        self.sheet.set_horz_split_pos(1)

        return self.sheet.row(start_row)

    def write_rows(self, start_row, start_col, data_lists,
                   style=Style.default_style):
        '''
        按行写入多组数据
        :param start_row: 开始写入行序号
        :param start_col: 写入列序号
        :param data_lists: 列表嵌套列表数据
        :return: 返回写入行对象列表
        '''
        row_obj = []
        num=1
        # print(data_lists)
        for row_, data in enumerate(data_lists):
            if isinstance(data, list):
                # self.sum(data)
                # print(data)
                self.write_row(start_row + row_, start_col, data, style)
                row_obj.append(self.sheet.row(start_row + row_))
                num=row_+1
                now = datetime.datetime.now()
                CreatTime = now.strftime("%Y-%m-%d-%H:%M:%S")
                data.insert(0,CreatTime)
                # sqlUnit.mainInsert(data)
            else:
                msg = '数据列表不是嵌套列表数据，而是%s' % type(data)
                raise Exception(msg)
        # self.totalFile(num,start_row)``````````````````
        return row_obj

    def write_col(self, start_row, start_col, date_list,
                  style=Style.default_style):
        '''
        按列写入一列数据
        :param start_row:写入行序号
        :param start_col:写入列
        :param start_col: 写入列序号
        :param date_list: 写入数据：列表
        :return: 返回写入的列对象
        '''
        for row, label in enumerate(date_list):
            self.write(row + start_row, start_col, label, style)

        return self.sheet.col(start_col)

    def write_cols(self, start_row, start_col, data_lists,
                   style=Style.default_style):
        '''
        按列写入多列数据
        :param start_row:开始写入行序号
        :param start_col: 开始写入列序号
        :param data_lists: 列表嵌套列表数据
        :return: 返回列对象列表
        '''
        col_obj = []
        for col_, data in enumerate(data_lists):
            if isinstance(data, list):
                self.write_col(start_row, start_col + col_, data, style)
                col_obj.append(self.sheet.col(start_col + col_))
            else:
                msg = '数据列表不是嵌套列表数据，而是%s' % type(data)
                raise Exception(msg)

        return col_obj






