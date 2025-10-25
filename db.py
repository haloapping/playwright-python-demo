import psycopg
import config
from psycopg.rows import dict_row


def init_db():
    """
    Initialize and return a PostgreSQL connection using psycopg3.
    """
    try:
        conn = psycopg.connect(
            host=config.HOST,
            dbname=config.DBNAME,
            user=config.USER,
            password=config.PASSWORD,
            row_factory=dict_row,  # Optional: rows as dict instead of tuple
        )
        print("Database connection established!")
        return conn
    except Exception as e:
        print(f"Database connection failed: {e}")
        raise
