import pymssql
import pandas as pd
from dbutils.pooled_db import PooledDB

class MSSQLDB:
    def __init__(self, config):
        self.db_config = config
        self.pool = PooledDB(**self.db_config)

    def read_table(self, table_name):
        with self.pool.connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM {table_name}'.format(table_name=table_name))
            df = pd.DataFrame(cursor.fetchall())
            df.columns = [desc[0] for desc in cursor.description]
        return df