from flask import Flask, render_template, request, redirect, url_for, session
from co_so_du_lieu.ket_noi_sql import ket_noi_sql

from chuc_nang.admin import admin_bp
from chuc_nang.nguoi_danh_gia import nguoi_danh_gia_bp
from chuc_nang.nhan_vien import nhan_vien_bp

app = Flask(__name__)
app.secret_key = "qldtnv_secret_key"

# ===== KẾT NỐI DATABASE (DÙNG CHUNG) =====
try:
    app.config['DB_CONN'] = ket_noi_sql()
    print("✅ Kết nối SQL Server thành công")
except Exception as e:
    print("❌ Lỗi kết nối DB:", e)

# ===== ĐĂNG KÝ MODULE (BLUEPRINT) =====
app.register_blueprint(admin_bp, url_prefix="/admin")
app.register_blueprint(nguoi_danh_gia_bp, url_prefix="/nguoi_danh_gia")
app.register_blueprint(nhan_vien_bp, url_prefix="/nhan_vien")

# ===== TRANG ĐĂNG NHẬP =====
@app.route('/', methods=['GET'])
def trang_dang_nhap():
    return render_template('dang_nhap.html')

# ===== XỬ LÝ ĐĂNG NHẬP (ĐÚNG DB – KHÔNG 404) =====
@app.route('/dang-nhap', methods=['POST'])
def dang_nhap():
    username = request.form.get('username')
    password = request.form.get('password')

    conn = app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT vai_tro 
        FROM Tai_Khoan
        WHERE ten_dang_nhap = ? AND mat_khau = ?
    """, (username, password))

    user = cursor.fetchone()

    if user:
        vai_tro = user[0]
        session['username'] = username
        session['vai_tro'] = vai_tro

        # ===== ĐIỀU HƯỚNG ĐÚNG THEO VAI TRÒ =====
        if vai_tro == 'admin':
            return redirect(url_for('admin.trang_chu'))

        elif vai_tro == 'nguoi_danh_gia':
            return redirect(url_for('nguoi_danh_gia.trang_chu'))

        elif vai_tro == 'nhan_vien':
            return redirect(url_for('nhan_vien.trang_chu'))

    # ❌ Sai tài khoản
    return render_template(
        'dang_nhap.html',
        error="Sai tên đăng nhập hoặc mật khẩu"
    )

# ===== ĐĂNG XUẤT =====
@app.route('/dang-xuat')
def dang_xuat():
    session.clear()
    return redirect(url_for('trang_dang_nhap'))

# ===== CHẠY APP =====
if __name__ == '__main__':
    app.run(debug=True)  