# /home/lusca/py_excel/tkinter campssa/database_connection.py

import sqlite3

class DatabaseConnection:
    """Context manager para conexões com banco de dados"""

    def __init__(self, db_name: str):
        self.db_name = db_name

    def __enter__(self) -> sqlite3.Connection:
        self.conn = sqlite3.connect(self.db_name)
        return self.conn

    def __exit__(self, exc_type, exc_value, traceback):
        if self.conn:
            self.conn.close()