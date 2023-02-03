import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.dialects.mysql import VARCHAR


# 指定列类型
# temporary_ces = {'USER_PHONE': VARCHAR(100), 'CITY': VARCHAR(100), 'SCORE': VARCHAR(100)}

def load_data(file):
    # 创建sql连接
    engine = create_engine("mysql+pymysql://用户名:密码@ip:端口/库名", max_overflow=5)

    # 读取数据 量级较大的txt文件
    data = pd.read_csv(file, sep='',
                       names=[],  # 给定表头，与数据库表的表头相同最好
                       dtype=str,
                       on_bad_lines='skip',
                       chunksize=50000,  # 分块导入
                       header=None,
                       )

    for i in data:
        i.to_sql('表名', con=engine, if_exists='append', index=False)
