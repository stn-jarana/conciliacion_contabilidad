import pyodbc

conn = pyodbc.connect(
    "DRIVER={ODBC Driver 18 for SQL Server};"
    "SERVER=192.168.30.55;"
    "DATABASE=SIGE_STN;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)
cursor = conn.cursor()
cursor.execute("SELECT DISTINCT Ano FROM CN_PlanContable ORDER BY Ano")
for r in cursor.fetchall():
    print(r[0])
conn.close()
