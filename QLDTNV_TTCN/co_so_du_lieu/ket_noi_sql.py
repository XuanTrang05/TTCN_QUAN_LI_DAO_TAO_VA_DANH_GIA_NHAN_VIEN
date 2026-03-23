import pyodbc

def ket_noi_sql():
    try:
        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=LAPTOP-H6D1IIMQ\\XUANTRANG;"
            "DATABASE=QLDTNV_TTCN;"
            "Trusted_Connection=yes;"
            "Encrypt=yes;"
            "TrustServerCertificate=yes;"
        )
        return conn
    except Exception as e:
        print("❌ Lỗi kết nối SQL Server:", e)
        return None
