import pymssql
from pymssql import _mssql
from pymssql import _pymssql
import uuid
import decimal

class SqlUnit():
    def __init__(self):
        self.conn = pymssql.connect('192.168.10.16', 'sa', 'admin@325','HT')  # 单纯连接数据库
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
        # a = a.split("、")
        # for i in range(len(a)):
        #     sql_insert = f"insert into Student Values({i},'{a[i]}',18)"
        #     print(sql_insert)
        #     cursor.execute(sql_insert)
        #     self.conn.commit()
        # VALUES( % s);' % ', '.join(params)
        # ",".join(map(str, data))
        sql_insert = 'INSERT INTO HT_EdgeData  VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
        DataTuple=list(map(tuple, data))
        # cursor.execute(sql_insert)
        cursor.executemany(sql_insert,DataTuple)
        self.conn.commit()
        cursor.close()
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
        sql_select = f"Select  * From HT_User where username = '{username}' and password ='{password}'"
        print(sql_select)
        cursor.execute(sql_select)
        re = cursor.fetchone()

        if re == None:
            re = False
        elif re != '':
            re=True


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
        sql_select = "select * from HT_BandingCode"
        cursor.execute(sql_select)
        row2 = cursor.fetchall()
        list2=row2
        cursor.close()
        # self.conn.close()

        return list,list2



def main():
    sqlUnit = SqlUnit()
    # sqlUnit.creatTableUnit()
    # sqlUnit.InserUnit()

    return sqlUnit

# def mainInsert(data):
    # sqlUnit = SqlUnit()
    # sqlUnit.InserUnit(data)
