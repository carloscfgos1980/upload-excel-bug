# pip install SQLAlchemy. SQLAlchemy to create the connection bridge to connect python to SQL server
from sqlalchemy import create_engine
from sqlalchemy.engine import URL  # URL class to construct the connection string
import pypyodbc  # pip install pypyodbc. pypyodbc to connect to SQL server
import pandas as pd  # pip install pandas. pandas to load the data from excels files

SERVER_NAME = 'AIR-VAN-CARLOS'
DATABASE_NAME = 'JJ'
TABLE_NAME = 'demo_table'

connection_string = f"""
DRIVER={{SQL Server}};
SERVER={SERVER_NAME};
DATABASE={DATABASE_NAME};
Trusted_Connection=yes;
"""
excel_file = 'amecillo.xlsx'

connection_url = URL.create(
    'mssql+pyodbc', query={'odbc_connect': connection_string})
enigne = create_engine(connection_url, module=pypyodbc)

excel_file = pd.read_excel(excel_file, sheet_name=None)
for sheet_name, df_data in excel_file.items():
    print(f'Loading worksheet {sheet_name}...')
    # {'fail', 'replace', 'append'}
    df_data.to_sql(TABLE_NAME, enigne, if_exists='append', index=False)
