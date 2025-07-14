import os
import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv
import urllib.parse

# Load environment variables from conn.env
load_dotenv(dotenv_path="conn.env")

class DatabaseHandler:
    def __init__(self):
        self.engine = None

    def connect(self):
        raise NotImplementedError("Subclasses must implement the 'connect' method.")

    def dispose(self):
        if self.engine:
            self.engine.dispose()
            print("✅ Connection disposed.")

class SQLServerHandler(DatabaseHandler):
    def connect(self):
        # Build ODBC connection string and URL-encode it
        odbc_str = (
            f"DRIVER={os.getenv('DB_DRIVER')};"
            f"SERVER={os.getenv('DB_SERVER')};"
            f"DATABASE={os.getenv('DB_DATABASE')};"
            f"Trusted_Connection={os.getenv('DB_TRUSTED_CONNECTION')};"
        )
        odbc_str_encoded = urllib.parse.quote_plus(odbc_str)
        self.engine = create_engine(f"mssql+pyodbc:///?odbc_connect={odbc_str_encoded}")
        print("✅ Connected to SQL Server.")

class PostgreSQLHandler(DatabaseHandler):
    def connect(self):
        self.engine = create_engine(
            f"postgresql://{os.getenv('DB_USER')}:{os.getenv('DB_PASS')}"
            f"@{os.getenv('DB_HOST')}/{os.getenv('DB_NAME')}"
        )
        print("✅ Connected to PostgreSQL.")

class MySQLHandler(DatabaseHandler):
    def connect(self):
        self.engine = create_engine(
            f"mysql+pymysql://{os.getenv('DB_USER')}:{os.getenv('DB_PASS')}"
            f"@{os.getenv('DB_HOST')}/{os.getenv('DB_NAME')}"
        )
        print("✅ Connected to MySQL.")

class SQLiteHandler(DatabaseHandler):
    def connect(self):
        self.engine = create_engine(f"sqlite:///{os.getenv('DB_PATH')}")
        print("✅ Connected to SQLite.")

def get_database_handler(db_type):
    handlers = {
        "sqlserver": SQLServerHandler,
        "postgresql": PostgreSQLHandler,
        "mysql": MySQLHandler,
        "sqlite": SQLiteHandler
    }
    if db_type in handlers:
        return handlers[db_type]()
    else:
        raise ValueError(f"❌ Unsupported database type: {db_type}")

def export_data_to_excel(db_handler, sql_file: str, excel_file: str):
    try:
        # Read the SQL query from the file
        with open(sql_file, 'r', encoding='utf-8') as file:
            sql_query = file.read()

        # Execute the SQL query and fetch the data into a DataFrame
        with db_handler.engine.connect() as connection:
            df = pd.read_sql(sql_query, connection)

        # Export the DataFrame to an Excel file
        df.to_excel(excel_file, index=False)
        print(f"✅ Data exported successfully to {excel_file}")

    except Exception as e:
        print(f"❌ Error: {e}")

    finally:
        db_handler.dispose()

if __name__ == "__main__":
    print("=== SQL to Excel Exporter ===")
    # Prompt user for database type
    db_type = input("Enter database type (sqlserver/postgresql/mysql/sqlite): ").strip().lower()
    db_handler = get_database_handler(db_type)
    db_handler.connect()

    # Prompt user for SQL file path
    sql_file_path = os.getenv("SQL_FILE_PATH")
    if not sql_file_path:
        sql_file_path = input("Enter path to SQL file: ").strip()
    if not os.path.isfile(sql_file_path):
        print(f"❌ SQL file not found: {sql_file_path}")
        exit(1)

    # Prompt user for Excel output path
    excel_file_path = os.getenv("EXCEL_FILE_PATH")
    if not excel_file_path:
        excel_file_path = input("Enter path for output Excel file: ").strip()
    excel_dir = os.path.dirname(excel_file_path)
    if excel_dir and not os.path.exists(excel_dir):
        try:
            os.makedirs(excel_dir)
        except Exception as e:
            print(f"❌ Could not create directory {excel_dir}: {e}")
            exit(1)

    # Export data
    export_data_to_excel(db_handler, sql_file_path, excel_file_path)