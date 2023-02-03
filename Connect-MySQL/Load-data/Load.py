import xlrd
import mysql.connector


def load_data(filepath, name):
    """
    Excel文件数据入库
    :param filepath: 文件路径
    :param name: sheet名
    """
    try:
        # 打开excel数据表
        data = xlrd.open_workbook(filepath)  # 修改表名称
        sheet = data.sheet_by_name(name)  # 选择打开的sheet

        # 数据过渡存储
        row_num = sheet.nrows

        data_list = []  # 定义列表用来存放数据
        for i in range(1, row_num):
            row_data = sheet.row_values(i)
            value = (row_data[0], row_data[1], row_data[2], ...)
            data_list.append(value)

        # 连接数据库
        mydb = mysql.connector.connect(
            # host="", # 主机地址
            user="",  # 用户名
            passwd="",  # MySQL密码
            database=""  # 选择要导入数据的数据表所在的数据库
        )

        cursor = mydb.cursor()

        # 数据导入数据库
        sql = """INSERT INTO 数据表(column1,column2,column3,....) VALUES (%s,%s,%s,....) """  # 导入语句
        cursor.executemany(sql, data_list)

        mydb.commit()
        cursor.close()
        mydb.close()

        print('数据导入成功！')

    except Exception as e:
        print("数据导入发生错误！", e)
