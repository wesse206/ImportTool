import pyodbc
def connectDB(server, database, username, password):
    SERVER = server
    DATABASE = database
    USERNAME = username
    PASSWORD = password
    connectionString = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};Encrypt=no;'
    
    conn = pyodbc.connect(connectionString) 
    return conn
