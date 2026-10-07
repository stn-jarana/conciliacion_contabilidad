import pyodbc

def get_connection():
    try:
        return pyodbc.connect(
            "DRIVER={ODBC Driver 18 for SQL Server};"
            "SERVER=192.168.30.55;"
            "DATABASE=SIGE_STN;"
            "Trusted_Connection=yes;"
            "TrustServerCertificate=yes;"
        )
    except pyodbc.Error as e:
        print("Error al conectar a la base de datos:", e)
        return None
