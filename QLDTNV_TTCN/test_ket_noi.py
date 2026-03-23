from co_so_du_lieu.ket_noi_sql import ket_noi_sql

conn = ket_noi_sql()

if conn:
    print("✅ Kết nối SQL Server thành công!")
else:
    print("❌ Kết nối thất bại!")
