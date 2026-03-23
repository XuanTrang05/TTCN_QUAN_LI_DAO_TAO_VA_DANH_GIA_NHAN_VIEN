from flask import Blueprint, render_template, session, redirect, url_for, current_app
from co_so_du_lieu.ket_noi_sql import ket_noi_sql  # giữ nguyên theo yêu cầu
from openpyxl import Workbook
from flask import send_file
import io
from flask import request


nhan_vien_bp = Blueprint(
    "nhan_vien",
    __name__,
    url_prefix="/nhan-vien"
)

# ==========================
# Trang chủ nhân viên
# ==========================
@nhan_vien_bp.route("/trang-chu")
def trang_chu():
    # Kiểm tra đăng nhập + vai trò
    if 'username' not in session or session.get('vai_tro') != 'nhan_vien':
        return redirect(url_for('trang_dang_nhap'))

    ten_dang_nhap = session.get("username")
    ten_nhan_vien = "Nhân viên"

    conn = ket_noi_sql()
    cursor = conn.cursor()

    try:
        # 🔥 LẤY TÊN NHÂN VIÊN QUA TÊN ĐĂNG NHẬP
        cursor.execute("""
            SELECT nv.ho_ten
            FROM Tai_Khoan tk
            JOIN Nhan_Vien nv ON tk.ma_nhan_vien = nv.ma_nhan_vien
            WHERE tk.ten_dang_nhap = ?
        """, (ten_dang_nhap,))

        row = cursor.fetchone()
        if row:
            ten_nhan_vien = row[0]

    finally:
        cursor.close()
        conn.close()

    return render_template(
        "nhan_vien/trang_chu.html",
        ten_nhan_vien=ten_nhan_vien
    )

# ==========================
# KHÓA ĐÀO TẠO CỦA TÔI
# ==========================
@nhan_vien_bp.route("/khoa-dao-tao")
def khoa_dao_tao():
    if 'username' not in session or session.get('vai_tro') != 'nhan_vien':
        return redirect(url_for('trang_dang_nhap'))

    username = session['username']
    conn = ket_noi_sql()
    cursor = conn.cursor()

    # Lấy mã nhân viên
    cursor.execute("""
        SELECT ma_nhan_vien 
        FROM Tai_Khoan 
        WHERE ten_dang_nhap = ?
    """, (username,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return redirect(url_for('trang_dang_nhap'))

    ma_nv = row[0]

    # 1. Khóa được phân công / đang học
    cursor.execute("""
        SELECT kdt.ten_khoa_dao_tao, kdt.ngay_bat_dau, kdt.ngay_ket_thuc, pc.trang_thai
        FROM Phan_Cong_Dao_Tao pc
        JOIN Khoa_Dao_Tao kdt ON pc.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
        WHERE pc.ma_nhan_vien = ? AND pc.trang_thai = N'Đang học'
    """, (ma_nv,))
    dang_hoc = cursor.fetchall()

    # 2. Khóa đã hoàn thành
    cursor.execute("""
        SELECT 
        kdt.ten_khoa_dao_tao,
        dg.diem_so,
        dg.nhan_xet
    FROM Danh_Gia dg
    JOIN Khoa_Dao_Tao kdt 
        ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
    WHERE dg.ma_nhan_vien = ?
""", (ma_nv,))
    da_hoan_thanh = cursor.fetchall()

    # 3. Khóa sắp diễn ra
    cursor.execute("""
        SELECT ma_khoa_dao_tao, ten_khoa_dao_tao, ngay_bat_dau
        FROM Khoa_Dao_Tao
        WHERE ngay_bat_dau > GETDATE()
    """)
    sap_dien_ra = cursor.fetchall()

    conn.close()

    return render_template(
        "nhan_vien/khoa_dao_tao_cua_toi.html",
        dang_hoc=dang_hoc,
        da_hoan_thanh=da_hoan_thanh,
        sap_dien_ra=sap_dien_ra
    )

# ==========================
# ĐĂNG KÝ KHÓA HỌC  ✅ (BỔ SUNG – HẾT BuildError)
# ==========================
@nhan_vien_bp.route("/dang-ky-khoa-hoc/<int:khoa_id>", methods=["POST"])
def dang_ky_khoa_hoc(khoa_id):
    if 'username' not in session or session.get('vai_tro') != 'nhan_vien':
        return redirect(url_for('trang_dang_nhap'))

    username = session['username']
    conn = ket_noi_sql()
    cursor = conn.cursor()

    # Lấy mã nhân viên
    cursor.execute("""
        SELECT ma_nhan_vien FROM Tai_Khoan WHERE ten_dang_nhap = ?
    """, (username,))
    ma_nv = cursor.fetchone()[0]

    # Tránh đăng ký trùng
    cursor.execute("""
        IF NOT EXISTS (
            SELECT 1 FROM Phan_Cong_Dao_Tao
            WHERE ma_nhan_vien = ? AND ma_khoa_dao_tao = ?
        )
        INSERT INTO Phan_Cong_Dao_Tao(ma_nhan_vien, ma_khoa_dao_tao, trang_thai)
        VALUES (?, ?, N'Đang học')
    """, (ma_nv, khoa_id, ma_nv, khoa_id))

    conn.commit()
    conn.close()

    return redirect(url_for('nhan_vien.khoa_dao_tao'))

# ==========================
# KẾT QUẢ ĐÁNH GIÁ
# ==========================
# ==========================
# KẾT QUẢ ĐÁNH GIÁ
# ==========================
@nhan_vien_bp.route("/ket-qua")
def ket_qua():
    if 'username' not in session or session.get('vai_tro') != 'nhan_vien':
        return redirect(url_for('trang_dang_nhap'))

    username = session.get('username')

    conn = ket_noi_sql()
    cursor = conn.cursor()

    try:
        # 🔥 LẤY MÃ NHÂN VIÊN
        cursor.execute("""
            SELECT ma_nhan_vien
            FROM Tai_Khoan
            WHERE ten_dang_nhap = ?
        """, (username,))
        row = cursor.fetchone()

        if not row:
            return redirect(url_for('trang_dang_nhap'))

        ma_nv = row[0]

        # 🔥 LẤY KẾT QUẢ ĐÁNH GIÁ
        cursor.execute("""
            SELECT 
                kdt.ten_khoa_dao_tao,
                dg.diem_so,
                dg.nhan_xet,
                dg.ngay_danh_gia
            FROM Danh_Gia dg
            JOIN Khoa_Dao_Tao kdt 
                ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
            WHERE dg.ma_nhan_vien = ?
              AND dg.diem_so IS NOT NULL
        """, (ma_nv,))

        rows = cursor.fetchall()

        # 🔥 FIX JSON ERROR — convert Row → Dict
        khoa_hoan_thanh = []
        for r in rows:
            khoa_hoan_thanh.append({
                "ten_khoa": r[0],
                "diem_so": float(r[1]) if r[1] is not None else 0,
                "nhan_xet": r[2] if r[2] is not None else "",
                "ngay_danh_gia": r[3].strftime("%d/%m/%Y") if r[3] else""
            })

    finally:
        cursor.close()
        conn.close()

    return render_template(
        "nhan_vien/ket_qua_danh_gia.html",
        khoa_hoan_thanh=khoa_hoan_thanh
    )

# ==========================
# Xuất Excel bảng điểm
# ==========================
@nhan_vien_bp.route("/xuat-excel-diem")
def xuat_excel_diem():
    if 'username' not in session or session.get('vai_tro') != 'nhan_vien':
        return redirect(url_for('trang_dang_nhap'))

    username = session.get("username")

    conn = ket_noi_sql()
    cursor = conn.cursor()

    try:
        # 1️⃣ Lấy mã nhân viên từ tài khoản
        cursor.execute("""
            SELECT ma_nhan_vien
            FROM Tai_Khoan
            WHERE ten_dang_nhap = ?
        """, (username,))
        row = cursor.fetchone()
        if not row:
            return redirect(url_for('trang_dang_nhap'))

        ma_nv = row[0]

        # 2️⃣ Lấy bảng điểm từ bảng Danh_Gia
        cursor.execute("""
            SELECT 
                kdt.ten_khoa_dao_tao,
                dg.diem_so,
                dg.nhan_xet,
                dg.ngay_danh_gia
            FROM Danh_Gia dg
            JOIN Khoa_Dao_Tao kdt 
                ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
            WHERE dg.ma_nhan_vien = ?
        """, (ma_nv,))

        data = cursor.fetchall()

    finally:
        cursor.close()
        conn.close()

    # 3️⃣ Tạo file Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng điểm đào tạo"

    # Header
    ws.append([
        "Tên khóa đào tạo",
        "Điểm số",
        "Nhận xét",
        "Ngày đánh giá"
    ])

    # Dữ liệu
    for row in data:
        ws.append(list(row))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name="bang_diem_dao_tao.xlsx",
        as_attachment=True
    )

# ==========================
# Xuất Excel tất cả khóa đào tạo
# ==========================
@nhan_vien_bp.route("/xuat-excel-khoa-dao-tao")
def xuat_excel_khoa_dao_tao():
    if 'username' not in session or session.get('vai_tro') != 'nhan_vien':
        return redirect(url_for('trang_dang_nhap'))

    return "Chức năng xuất Excel khóa đào tạo (đang phát triển)"
#=============== thong tin NHÂN VIÊN============
@nhan_vien_bp.route("/thong-tin", methods=["GET", "POST"])
def thong_tin_ca_nhan():
    if "username" not in session:
        return redirect(url_for("trang_dang_nhap"))

    username = session["username"]

    conn = ket_noi_sql()
    cursor = conn.cursor()

    try:
        # 1️⃣ Lấy mã nhân viên từ tài khoản
        cursor.execute("""
            SELECT ma_nhan_vien
            FROM Tai_Khoan
            WHERE ten_dang_nhap = ?
        """, (username,))
        row = cursor.fetchone()

        if not row:
            return redirect(url_for("trang_dang_nhap"))

        ma_nv = row[0]

        # ================== POST: CẬP NHẬT ==================
        if request.method == "POST":
            email = request.form.get("email")
            so_dien_thoai = request.form.get("so_dien_thoai")

            cursor.execute("""
                UPDATE Nhan_Vien
                SET email = ?, so_dien_thoai = ?
                WHERE ma_nhan_vien = ?
            """, (email, so_dien_thoai, ma_nv))

            conn.commit()

        # ================== GET: LẤY THÔNG TIN ==================
        cursor.execute("""
            SELECT 
                nv.ma_nhan_vien,
                nv.ho_ten,
                nv.ngay_sinh,
                nv.email,
                nv.so_dien_thoai,
                nv.chuc_vu,
                pb.ten_phong
            FROM Nhan_Vien nv
            LEFT JOIN Phong_Ban pb 
                ON nv.ma_phong = pb.ma_phong
            WHERE nv.ma_nhan_vien = ?
        """, (ma_nv,))

        row = cursor.fetchone()

        if not row:
            nhan_vien = None
        else:
            nhan_vien = {
                "MaNhanVien": row[0],
                "HoTen": row[1],
                "NgaySinh": row[2],
                "Email": row[3],
                "SoDienThoai": row[4],
                "TenChucVu": row[5],
                "TenPhongBan": row[6]
            }

    finally:
        cursor.close()
        conn.close()

    return render_template(
        "nhan_vien/thong_tin_ca_nhan.html",
        nv=nhan_vien
    )
