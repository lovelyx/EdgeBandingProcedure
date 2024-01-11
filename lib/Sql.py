import pymssql
from pymssql import _mssql
from pymssql import _pymssql
import uuid
import decimal

class SqlUnit():
    def __init__(self,sqlIp,edt_username,edt_password):
        print(sqlIp,edt_username,edt_password)
        self.conn = pymssql.connect(sqlIp, edt_username,edt_password ,'HT')  # 单纯连接数据库
        if self.conn:
            print("连接成功")

    def CreateDataUnit(self):
        # 1:创建一个数据库
        # conn = pymssql.connect('192.168.10.16','sa','admin@325')#单纯连接数据库
        # if conn:
        #     print("连接成功")
        cursor = self.conn.cursor() #创建执行语句
        self.conn.autocommit(True) #创建库的核心!!!
        sql_DATA ="""
        CREATE DATABASE PY_DATA
        ON   PRIMARY
         (NAME = 'PY_DATA',
        FILENAME = 'D:\DATA\PY_DATA.MDF' ,
        SIZE = 5MB,
        MAXSIZE = 20MB,
        FILEGROWTH = 20%)
        LOG ON
        (NAME ='PY_DATA_LOG',
        FILENAME = 'D:\DATA\PY_DATA_LOG. LDF',
        SIZE = 5MB,
        MAXSIZE = 10MB,
        FILEGROWTH = 2MB)
        """
        cursor.execute(sql_DATA)
        cursor.close()
        self.conn.autocommit(False)
        self.conn.close()


    def creatTableUnit(self):
        #2:创建一个表
        # import pymssql
        # ##连接
        # conn = pymssql.connect('.','sa','123456','School')
        # if conn:
        #     print("连接成功")

        ##操作
        cursor = self.conn.cursor()
        self.conn.autocommit(True)#是修改表的结构都要有吗?
        sql_TABLE = """
        Create Table HT_User(
        ID int primary key,
        username varchar(50),
        password varchar(50),
        authority int)
        """
        cursor.execute(sql_TABLE)
        print("创建表成功")
        self.conn.autocommit(False)
        self.conn.close()

    def InserUnit(self,data):
        # CreateTime, code, MemoOne, MemoTwo, MemoThree, LeftMachine, RightMachine, Lenght, Width, Thick, Qty, FeedSequence, WorkpieceRotation, FloatingCutter, velocity, LeftEdgeCode, LeftMachiningCode, RightEdgeCode, RightMachiningCode

        # 3:为这个表插入数据
        # import  pymssql
        # conn = pymssql.connect('.','sa','123456','PY_DATA')
        # if conn:
        #     print("True")

        cursor = self.conn.cursor()
        # a = "松仁、秉峰、泳纪海奉、威剑、颂和、祥益、腾恩、柏铄、孟深、忠庄、轩哲、铠鑫、仕伦、儒亿、积进信钦、贤元、程基、安泉、树昌、祝斌、一科、游湖、普济、中坚"
        # 'Reference', 'Decor', 'Materialflow', 'BatchNumber', 'CustomerNumber', 'Length','Width', 'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation','OverSizeM1', 'EdgeMacroLM1', 'ProgramM1', 'OverSizeM2', 'EdgeMacroRM2', 'ProgramM2'
        # a = a.split("、")
        # for i in range(len(a)):
        #     sql_insert = f"insert into Student Values({i},'{a[i]}',18)"
        #     print(sql_insert)
        #     cursor.execute(sql_insert)
        #     self.conn.commit()
        # VALUES( % s);' % ', '.join(params)
        # ",".join(map(str, data))
        sql_insert = 'INSERT INTO HT_EdgeData (CreateTime,Reference,MemoOne,MemoTwo,MemoThree,CustomerNumber,Length,Width,Thickness,Quantity,passValue,FeedSpeed,Orientation,OverSizeM1,EdgeMacroLM1,ProgramM1,OverSizeM2,EdgeMacroRM2,ProgramM2,identifying,UserName) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
        # print(data)
        DataTuple=list(map(tuple, data))
        # cursor.execute(sql_insert)
        cursor.executemany(sql_insert,DataTuple)
        self.conn.commit()
        cursor.close()
        return
        # self.conn.close()

    def InserUnitFive(self,data):
        # CreateTime, code, MemoOne, MemoTwo, MemoThree, LeftMachine, RightMachine, Lenght, Width, Thick, Qty, FeedSequence, WorkpieceRotation, FloatingCutter, velocity, LeftEdgeCode, LeftMachiningCode, RightEdgeCode, RightMachiningCode

        # 3:为这个表插入数据
        # import  pymssql
        # conn = pymssql.connect('.','sa','123456','PY_DATA')
        # if conn:
        #     print("True")

        cursor = self.conn.cursor()
        # a = "松仁、秉峰、泳纪海奉、威剑、颂和、祥益、腾恩、柏铄、孟深、忠庄、轩哲、铠鑫、仕伦、儒亿、积进信钦、贤元、程基、安泉、树昌、祝斌、一科、游湖、普济、中坚"
        # 'Reference', 'Decor', 'Materialflow', 'BatchNumber', 'CustomerNumber', 'Length','Width', 'Thickness', 'Quantity', 'passValue', 'FeedSpeed', 'Orientation','OverSizeM1', 'EdgeMacroLM1', 'ProgramM1', 'OverSizeM2', 'EdgeMacroRM2', 'ProgramM2'
        # a = a.split("、")
        # for i in range(len(a)):
        #     sql_insert = f"insert into Student Values({i},'{a[i]}',18)"
        #     print(sql_insert)
        #     cursor.execute(sql_insert)
        #     self.conn.commit()
        # VALUES( % s);' % ', '.join(params)
        # ",".join(map(str, data))
        sql_insert = 'INSERT INTO HT_EdgeDataFive (CreateTime,Reference,MemoOne,MemoTwo,MemoThree,CustomerNumber,Length,Width,Thickness,Quantity,passValue,FeedSpeed,Orientation,OverSizeM1,EdgeMacroLM1,ProgramM1,BasicMacroM1,OverSizeM2,EdgeMacroRM2,ProgramM2,UserName) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
        # print(data)
        DataTuple=list(map(tuple, data))
        # cursor.execute(sql_insert)
        cursor.executemany(sql_insert,DataTuple)
        self.conn.commit()
        cursor.close()
        return
        # self.conn.close()
    def UpdataUnit(self):
        # ----------------------------------------------------------------------------
        #
        # 4:更新插入的数据
        # import pymssql
        conn = pymssql.connect('.','sa','123456','PY_DATA')
        if conn:
            print("True")

        cursor = conn.cursor()
        sql_updata = "update student set Name='消息' where ID = 12"
        cursor.execute(sql_updata)
        conn.commit()
        cursor.close()
        conn.close()

    def selectUnit(self,username,password):
        # 5:查询前三行数据
        # conn = pymssql.connect('.','sa','123456','PY_DATA',charset='cp936')
        # conn = pymssql.connect('192.168.10.16','sa','admin@325','HT')
        # if conn:
        #     print("True")
        cursor = self.conn.cursor()
        sql_select = f"Select * From HT_User where username = '{username}' and password ='{password}'"
        print(sql_select)
        cursor.execute(sql_select)
        re = cursor.fetchone()
        print(re)

        # if re == None:
        #     re = False
        # elif re != '':
        #     re=True


        cursor.close()
        self.conn.close()
        print("已关闭连接")

        return re

    def selectBandingProcessing(self):
        cursor=self.conn.cursor()
        sql_select = "select * from HT_BandingProcessing"
        cursor.execute(sql_select)
        row = cursor.fetchall()
        list = row

        # sql_select = "select * from HT_BandingProcessingRight"
        # cursor.execute(sql_select)
        # row4 = cursor.fetchall()
        # list4 = row4

        sql_select = "select * from HT_BandingCode"
        cursor.execute(sql_select)
        row2 = cursor.fetchall()
        list2=row2

        # sql_select = "select * from HT_BandingCodeRight"
        # cursor.execute(sql_select)
        # row3 = cursor.fetchall()
        # list3 = row3

        cursor.close()
        # self.conn.close()

        # return list,list4,list2,list3

        return list, list2

    def selectBandingProcessingFive(self):
        cursor=self.conn.cursor()
        sql_select = "select * from HT_BandingProcessingFive"
        cursor.execute(sql_select)
        row = cursor.fetchall()
        list = row
        sql_select = "select * from HT_BandingCodeFive"
        cursor.execute(sql_select)
        row2 = cursor.fetchall()
        list2=row2
        cursor.close()
        # self.conn.close()

        return list,list2



def main(sqlIp,edt_username,edt_password):
    sqlUnit = SqlUnit(sqlIp,edt_username,edt_password)
    # sqlUnit.creatTableUnit()
    # sqlUnit.InserUnit()

    return sqlUnit

# def mainInsert(data):
    # sqlUnit = SqlUnit()
    # sqlUnit.InserUnit(data)
