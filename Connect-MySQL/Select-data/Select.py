import pymysql
import pandas as pd


def select_data():
    # 连接数据库
    mydb = pymysql.connect(
        # host="",
        user="",
        passwd="",
        database="",
        charset='utf8'
    )

    sql = """SELECT * FROM test ; """  # 查询语句

    # 查询
    cursor = mydb.cursor()
    cursor.execute(sql)
    result = cursor.fetchall()  # fetchall() 获取所有记录
    # 移动指针到某一行.如果mode='relative',则表示从当前所在行移动value条,如果mode='absolute',则表示从结果集的第一行移动value条.
    cursor.scroll(0, mode='absolute')
    # cursor.description获取表格的字段信息
    fields = cursor.description
    # 关闭游标和数据库的连接
    cursor.close()
    mydb.close()

    header = []
    for field in range(len(fields)):
        header.append(fields[field][0])

    result_df = pd.DataFrame(result, columns=header)

    return result_df


