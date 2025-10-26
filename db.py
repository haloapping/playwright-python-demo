from psycopg.rows import dict_row
from psycopg_pool import ConnectionPool
import config


DB_DSN = f"host={config.HOST} dbname={config.DBNAME} user={config.USER} password={config.PASSWORD}"

pool = ConnectionPool(
    open=True,
    conninfo=DB_DSN,
    min_size=1,
    max_size=10,
    kwargs={
        "row_factory": dict_row,
    },
)
