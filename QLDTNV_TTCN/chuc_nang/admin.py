from flask import Blueprint, render_template, request, redirect, url_for, current_app, jsonify, send_file
from datetime import date
from openpyxl import Workbook
from io import BytesIO
from co_so_du_lieu.ket_noi_sql import ket_noi_sql
from flask import send_file, session
import pandas as pd
import io
from openpyxl import Workbook
from datetime import datetime
import os

admin_bp = Blueprint(
    "admin",
    __name__,
    url_prefix="/admin"
)

# ================= TRANG CHỦ ADMIN =================
@admin_bp.route("/trang-chu")
def trang_chu():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM Nhan_Vien")
    tong_nhan_vien = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM Khoa_Dao_Tao")
    tong_khoa = cursor.fetchone()[0]

    return render_template(
        "admin/trang_chu.html",
        tong_nhan_vien=tong_nhan_vien,
        tong_khoa=tong_khoa
    )


# ================= QUẢN LÝ NHÂN VIÊN =================
@admin_bp.route("/nhan-vien")
def quan_ly_nhan_vien():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            nv.ngay_sinh,
            nv.gioi_tinh,
            nv.chuc_vu,
            nv.email,
            nv.so_dien_thoai,
            pb.ten_phong
        FROM Nhan_Vien nv
        LEFT JOIN Phong_Ban pb ON nv.ma_phong = pb.ma_phong
    """)
    nhan_vien = cursor.fetchall()

    cursor.execute("SELECT * FROM Phong_Ban")
    phong_ban = cursor.fetchall()

    return render_template(
        "admin/quan_ly_nhan_vien.html",
        nhan_vien=nhan_vien,
        phong_ban=phong_ban
    )


@admin_bp.route("/them-nhan-vien", methods=["POST"])
def them_nhan_vien():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO Nhan_Vien 
        (ho_ten, ngay_sinh, gioi_tinh, chuc_vu, email, so_dien_thoai, ma_phong)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        request.form["ho_ten"],
        request.form["ngay_sinh"],
        request.form["gioi_tinh"],
        request.form["chuc_vu"],
        request.form["email"],
        request.form["so_dien_thoai"],
        request.form["ma_phong"]
    ))

    conn.commit()
    return redirect(url_for("admin.quan_ly_nhan_vien"))


@admin_bp.route("/xoa-nhan-vien/<int:ma_nv>")
def xoa_nhan_vien(ma_nv):
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("DELETE FROM Nhan_Vien WHERE ma_nhan_vien = ?", (ma_nv,))
    conn.commit()

    return redirect(url_for("admin.quan_ly_nhan_vien"))


# ================= QUẢN LÝ KHÓA ĐÀO TẠO =================
@admin_bp.route("/khoa-dao-tao")
def khoa_dao_tao_card():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()
    today = date.today()

    # ⭐⭐⭐ FIX QUAN TRỌNG — THÊM JOIN NHAN_VIEN ĐỂ LẤY GIẢNG VIÊN
    base_sql = """
        SELECT 
            kdt.ma_khoa_dao_tao,
            kdt.ten_khoa_dao_tao,
            kdt.mo_ta,
            kdt.si_so_toi_da,
            kdt.ngay_bat_dau,
            kdt.ngay_ket_thuc,
            nv.ho_ten AS ten_giang_vien,
            COUNT(pc.ma_nhan_vien) AS so_hoc_vien
        FROM Khoa_Dao_Tao kdt
        LEFT JOIN Nhan_Vien nv
            ON kdt.ma_giang_vien = nv.ma_nhan_vien
        LEFT JOIN Phan_Cong_Dao_Tao pc 
            ON kdt.ma_khoa_dao_tao = pc.ma_khoa_dao_tao
    """

    # ⭐⭐⭐ FIX GROUP BY — THÊM nv.ho_ten
    group_by = """
        GROUP BY 
            kdt.ma_khoa_dao_tao,
            kdt.ten_khoa_dao_tao,
            kdt.mo_ta,
            kdt.si_so_toi_da,
            kdt.ngay_bat_dau,
            kdt.ngay_ket_thuc,
            nv.ho_ten
    """

    cursor.execute(base_sql + """
        WHERE kdt.ngay_bat_dau <= ? AND kdt.ngay_ket_thuc >= ?
    """ + group_by, (today, today))
    dang_dien_ra = cursor.fetchall()

    cursor.execute(base_sql + """
        WHERE kdt.ngay_bat_dau > ?
    """ + group_by, (today,))
    sap_dien_ra = cursor.fetchall()

    cursor.execute(base_sql + """
        WHERE kdt.ngay_ket_thuc < ?
    """ + group_by, (today,))
    da_ket_thuc = cursor.fetchall()

    cursor.execute(base_sql + group_by)
    tat_ca_khoa = cursor.fetchall()

    cursor.execute("""
        SELECT ma_nhan_vien, ho_ten
        FROM Nhan_Vien
    """)
    danh_sach_nhan_vien = cursor.fetchall()

    return render_template(
        "admin/quan_ly_khoa_dao_tao.html",
        dang_dien_ra=dang_dien_ra,
        sap_dien_ra=sap_dien_ra,
        da_ket_thuc=da_ket_thuc,
        tat_ca_khoa=tat_ca_khoa,
        danh_sach_nhan_vien=danh_sach_nhan_vien
    )



@admin_bp.route("/quan-ly-khoa-dao-tao")
def quan_ly_khoa_cu():
    return redirect(url_for("admin.khoa_dao_tao_card"))


@admin_bp.route("/them-khoa-dao-tao", methods=["POST"])
def them_khoa():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    ten_khoa = request.form.get("ten_khoa")
    mo_ta = request.form.get("mo_ta")
    si_so = request.form.get("si_so")
    ngay_bat_dau = request.form.get("ngay_bat_dau")
    ngay_ket_thuc = request.form.get("ngay_ket_thuc")
    ma_giang_vien = request.form.get("ma_giang_vien")

    # DEBUG — xem terminal Flask
    print("MA GIANG VIEN GUI LEN =", ma_giang_vien)

    cursor.execute("""
        INSERT INTO Khoa_Dao_Tao
        (ten_khoa_dao_tao, mo_ta, si_so_toi_da, ngay_bat_dau, ngay_ket_thuc, ma_giang_vien)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        ten_khoa,
        mo_ta,
        si_so,
        ngay_bat_dau,
        ngay_ket_thuc,
        ma_giang_vien
    ))

    conn.commit()
    return redirect(url_for("admin.khoa_dao_tao_card"))



@admin_bp.route("/xoa-khoa-dao-tao/<int:ma_khoa>")
def xoa_khoa_dao_tao(ma_khoa):
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("DELETE FROM Phan_Cong_Dao_Tao WHERE ma_khoa_dao_tao = ?", (ma_khoa,))
    cursor.execute("DELETE FROM Khoa_Dao_Tao WHERE ma_khoa_dao_tao = ?", (ma_khoa,))

    conn.commit()
    return redirect(url_for("admin.khoa_dao_tao_card"))


# ================= PHÂN CÔNG NHÂN VIÊN =================
@admin_bp.route("/phan-cong", methods=["POST"])
def phan_cong():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    ma_nv = request.form.get("ma_nhan_vien")
    ma_khoa = request.form.get("ma_khoa_dao_tao")

    # ===== CHECK DỮ LIỆU ĐẦU VÀO =====
    if not ma_nv or not ma_khoa:
        return redirect(url_for("admin.khoa_dao_tao_card"))

    # ===== CHỐNG PHÂN CÔNG TRÙNG =====
    cursor.execute("""
        SELECT COUNT(*) FROM Phan_Cong_Dao_Tao
        WHERE ma_nhan_vien = ? AND ma_khoa_dao_tao = ?
    """, (ma_nv, ma_khoa))
    if cursor.fetchone()[0] > 0:
        return redirect(url_for("admin.khoa_dao_tao_card"))

    # ===== ĐẾM SĨ SỐ HIỆN TẠI =====
    cursor.execute("""
        SELECT COUNT(*) FROM Phan_Cong_Dao_Tao
        WHERE ma_khoa_dao_tao = ?
    """, (ma_khoa,))
    si_so_hien_tai = cursor.fetchone()[0] or 0

    # ===== LẤY SĨ SỐ TỐI ĐA =====
    cursor.execute("""
        SELECT si_so_toi_da FROM Khoa_Dao_Tao
        WHERE ma_khoa_dao_tao = ?
    """, (ma_khoa,))
    row = cursor.fetchone()

    # ❗ FIX LỖI NoneType
    if not row or row[0] is None:
        return redirect(url_for("admin.khoa_dao_tao_card"))

    si_so_toi_da = int(row[0])

    # ===== CHECK ĐỦ SĨ SỐ =====
    if si_so_hien_tai >= si_so_toi_da:
        return redirect(url_for("admin.khoa_dao_tao_card"))

    # ===== PHÂN CÔNG =====
    cursor.execute("""
        INSERT INTO Phan_Cong_Dao_Tao
        (ma_nhan_vien, ma_khoa_dao_tao, trang_thai)
        VALUES (?, ?, N'Chưa học')
    """, (ma_nv, ma_khoa))

    conn.commit()
    return redirect(url_for("admin.khoa_dao_tao_card"))


# ================= DANH SÁCH HỌC VIÊN (JSON) =================
@admin_bp.route("/khoa-dao-tao/<int:ma_khoa>/hoc-vien")
def danh_sach_hoc_vien(ma_khoa):
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            pb.ten_phong,
            dg.diem_so,
            dg.nhan_xet
        FROM Phan_Cong_Dao_Tao pc
        JOIN Nhan_Vien nv ON pc.ma_nhan_vien = nv.ma_nhan_vien
        LEFT JOIN Phong_Ban pb ON nv.ma_phong = pb.ma_phong
        LEFT JOIN Danh_Gia dg 
            ON dg.ma_nhan_vien = nv.ma_nhan_vien
            AND dg.ma_khoa_dao_tao = pc.ma_khoa_dao_tao
        WHERE pc.ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    rows = cursor.fetchall()

    ds = []
    for r in rows:
        diem = r[3]

        # ✅ TỰ TÍNH KẾT QUẢ
        if diem is None:
            ket_qua = ""
        elif diem >= 5:
            ket_qua = "Đạt"
        else:
            ket_qua = "Không đạt"

        ds.append({
            "ma_nhan_vien": r[0],
            "ho_ten": r[1],
            "phong_ban": r[2],
            "diem": diem,
            "ket_qua": ket_qua,
            "danh_gia": r[4]
        })

    return jsonify(ds)


# ================= XUẤT EXCEL HỌC VIÊN =================
@admin_bp.route("/khoa-dao-tao/<int:ma_khoa>/xuat-excel")
def xuat_excel_hoc_vien(ma_khoa):
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            pb.ten_phong,
            dg.diem_so,
            dg.nhan_xet
        FROM Phan_Cong_Dao_Tao pc
        JOIN Nhan_Vien nv ON pc.ma_nhan_vien = nv.ma_nhan_vien
        LEFT JOIN Phong_Ban pb ON nv.ma_phong = pb.ma_phong
        LEFT JOIN Danh_Gia dg 
            ON dg.ma_nhan_vien = nv.ma_nhan_vien
            AND dg.ma_khoa_dao_tao = pc.ma_khoa_dao_tao
        WHERE pc.ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    rows = cursor.fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Danh sách học viên"

    # ✅ HEADER MỚI
    ws.append(["Mã NV", "Họ tên", "Phòng ban", "Điểm", "Kết quả", "Đánh giá"])

    for r in rows:
        diem = r[3]

        if diem is None:
            ket_qua = ""
        elif diem >= 5:
            ket_qua = "Đạt"
        else:
            ket_qua = "Không đạt"

        ws.append([
            r[0],  # mã NV
            r[1],  # tên
            r[2],  # phòng ban
            diem,
            ket_qua,
            r[4]   # đánh giá
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"danh_sach_hoc_vien_khoa_{ma_khoa}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ================= QUẢN LÝ TÀI KHOẢN =================
@admin_bp.route("/quan_ly_tai_khoan")
def quan_ly_tai_khoan():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    tu_khoa = request.args.get("tu_khoa")
    sua = request.args.get("sua")

    # 🔹 JOIN với bảng nhân viên để lấy tên
    sql = """
        SELECT 
            tk.ten_dang_nhap,
            tk.vai_tro,
            tk.ma_nhan_vien,
            nv.ho_ten
        FROM Tai_Khoan tk
        LEFT JOIN Nhan_Vien nv
            ON tk.ma_nhan_vien = nv.ma_nhan_vien
    """
    params = []

    if tu_khoa:
        sql += """
            WHERE tk.ten_dang_nhap LIKE ?
               OR tk.vai_tro LIKE ?
               OR nv.ho_ten LIKE ?
        """
        params.extend([f"%{tu_khoa}%", f"%{tu_khoa}%", f"%{tu_khoa}%"])

    cursor.execute(sql, params)
    rows = cursor.fetchall()

    danh_sach_tai_khoan = [
        {
            "ten_dang_nhap": r[0],
            "vai_tro": r[1],
            "ma_nhan_vien": r[2],
            "ten_nhan_vien": r[3]
        }
        for r in rows
    ]

    # ===== TÀI KHOẢN ĐANG SỬA =====
    tai_khoan_sua = None
    if sua:
        cursor.execute("""
            SELECT 
                tk.ten_dang_nhap,
                tk.vai_tro,
                tk.ma_nhan_vien,
                nv.ho_ten
            FROM Tai_Khoan tk
            LEFT JOIN Nhan_Vien nv
                ON tk.ma_nhan_vien = nv.ma_nhan_vien
            WHERE tk.ten_dang_nhap = ?
        """, (sua,))
        tk = cursor.fetchone()
        if tk:
            tai_khoan_sua = {
                "ten_dang_nhap": tk[0],
                "vai_tro": tk[1],
                "ma_nhan_vien": tk[2],
                "ten_nhan_vien": tk[3]
            }

    return render_template(
        "admin/quan_ly_tai_khoan.html",
        danh_sach_tai_khoan=danh_sach_tai_khoan,
        tai_khoan_sua=tai_khoan_sua
    )


# ===== ALIAS CHỐNG 404 =====
@admin_bp.route("/quan-ly-tai-khoan")
def quan_ly_tai_khoan_alias():
    return redirect(url_for("admin.quan_ly_tai_khoan", **request.args))


# ================= THÊM TÀI KHOẢN =================
@admin_bp.route("/them_tai_khoan", methods=["POST"])
def them_tai_khoan():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO Tai_Khoan (ten_dang_nhap, mat_khau, vai_tro, ma_nhan_vien)
        VALUES (?, ?, ?, ?)
    """, (
        request.form["ten_dang_nhap"],
        request.form["mat_khau"],
        request.form["vai_tro"],
        request.form.get("ma_nhan_vien")
    ))

    conn.commit()
    return redirect(url_for("admin.quan_ly_tai_khoan"))


# ================= CẬP NHẬT TÀI KHOẢN =================
@admin_bp.route("/cap_nhat_tai_khoan", methods=["POST"])
def cap_nhat_tai_khoan():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    if request.form.get("mat_khau"):
        cursor.execute("""
            UPDATE Tai_Khoan
            SET mat_khau = ?, vai_tro = ?, ma_nhan_vien = ?
            WHERE ten_dang_nhap = ?
        """, (
            request.form["mat_khau"],
            request.form["vai_tro"],
            request.form.get("ma_nhan_vien"),
            request.form["ten_dang_nhap"]
        ))
    else:
        cursor.execute("""
            UPDATE Tai_Khoan
            SET vai_tro = ?, ma_nhan_vien = ?
            WHERE ten_dang_nhap = ?
        """, (
            request.form["vai_tro"],
            request.form.get("ma_nhan_vien"),
            request.form["ten_dang_nhap"]
        ))

    conn.commit()
    return redirect(url_for("admin.quan_ly_tai_khoan"))

# ================= XÓA TÀI KHOẢN =================
@admin_bp.route("/xoa_tai_khoan/<ten_dang_nhap>")
def xoa_tai_khoan(ten_dang_nhap):
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM Tai_Khoan WHERE ten_dang_nhap = ?",
        (ten_dang_nhap,)
    )

    conn.commit()
    return redirect(url_for("admin.quan_ly_tai_khoan"))
# ========================= BÁO CÁO – THỐNG KÊ (ĐÃ FIX THEO DB) =========================
@admin_bp.route("/bao-cao-thong-ke")
def bao_cao_thong_ke():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ho_ten,
            kdt.ten_khoa_dao_tao,
            dg.diem_so,
            dg.nhan_xet,
            dg.ngay_danh_gia
        FROM Danh_Gia dg
        JOIN Nhan_Vien nv ON dg.ma_nhan_vien = nv.ma_nhan_vien
        JOIN Khoa_Dao_Tao kdt ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
    """)

    bao_cao = cursor.fetchall()

    return render_template(
        "admin/bao_cao_thong_ke.html",
        bao_cao=bao_cao
    )


# ===== FIX 404: ALIAS CHO /admin/bao-cao =====
@admin_bp.route("/bao-cao")
def bao_cao_alias():
    return redirect(url_for("admin.bao_cao_thong_ke"))

# ================= API BÁO CÁO (FIX: THÊM MÃ NV) =================
@admin_bp.route("/api/bao-cao-danh-gia")
def api_bao_cao_danh_gia():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            kdt.ten_khoa_dao_tao,
            dg.diem_so,
            dg.nhan_xet,
            dg.ngay_danh_gia
        FROM Danh_Gia dg
        JOIN Nhan_Vien nv ON dg.ma_nhan_vien = nv.ma_nhan_vien
        JOIN Khoa_Dao_Tao kdt ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
    """)

    data = []
    for r in cursor.fetchall():
        data.append({
            "maNhanVien": r[0],
            "nhanVien": r[1],
            "khoa": r[2],
            "diem": r[3],
            "nhanXet": r[4],
            "ngayDanhGia": str(r[5])
        })

    return jsonify(data)


# ================= API BÁO CÁO THEO NHÂN VIÊN (FIX INDEX) =================
@admin_bp.route("/api/bao-cao-danh-gia/nhan-vien/<int:ma_nv>")
def api_bao_cao_theo_nhan_vien(ma_nv):
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            kdt.ten_khoa_dao_tao,
            dg.diem_so,
            dg.nhan_xet,
            dg.ngay_danh_gia
        FROM Danh_Gia dg
        JOIN Nhan_Vien nv ON dg.ma_nhan_vien = nv.ma_nhan_vien
        JOIN Khoa_Dao_Tao kdt ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
        WHERE nv.ma_nhan_vien = ?
    """, (ma_nv,))

    data = []
    for r in cursor.fetchall():
        data.append({
            "maNhanVien": r[0],
            "nhanVien": r[1],
            "khoa": r[2],
            "diem": r[3],
            "nhanXet": r[4],
            "ngayDanhGia": str(r[5])
        })

    return jsonify(data)


# ================= API BÁO CÁO THEO KHÓA (FIX LOGIC) =================
@admin_bp.route("/api/bao-cao-danh-gia/theo-khoa")
def api_bao_cao_theo_khoa():
    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            kdt.ten_khoa_dao_tao,
            nv.ma_nhan_vien,
            nv.ho_ten,
            dg.diem_so,
            dg.nhan_xet
        FROM Danh_Gia dg
        JOIN Nhan_Vien nv ON dg.ma_nhan_vien = nv.ma_nhan_vien
        JOIN Khoa_Dao_Tao kdt ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
        ORDER BY kdt.ten_khoa_dao_tao
    """)

    ket_qua = {}

    for r in cursor.fetchall():
        ten_khoa = r[0]
        if ten_khoa not in ket_qua:
            ket_qua[ten_khoa] = []

        ket_qua[ten_khoa].append({
            "maNhanVien": r[1],
            "nhanVien": r[2],
            "diem": r[3],
            "nhanXet": r[4]
        })

    return jsonify(ket_qua)
# ================= API THỐNG KÊ =================
@admin_bp.route("/api/thong-ke/diem-trung-binh-theo-khoa")
def api_thong_ke_diem_tb():
    conn = ket_noi_sql()
    if conn is None:
        return {"error": "Không kết nối được CSDL"}, 500

    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            kdt.ten_khoa_dao_tao,
            AVG(dg.diem_so) AS diem_trung_binh
        FROM Danh_Gia dg
        JOIN Khoa_Dao_Tao kdt 
            ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
        GROUP BY kdt.ten_khoa_dao_tao
    """)

    data = [
        {"khoa": r[0], "diemTB": float(r[1])}
        for r in cursor.fetchall()
    ]

    cursor.close()
    conn.close()

    return jsonify(data)
# ================= XUẤT EXCEL =================
@admin_bp.route("/api/xuat-excel")
def xuat_excel():
    conn = ket_noi_sql()
    cursor = conn.cursor()

    # 👉 Nếu là nhân viên thì chỉ xem dữ liệu của mình
    ma_nv = session.get("ma_nhan_vien")  # ví dụ: lưu khi login

    sql = """
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            kdt.ten_khoa_dao_tao,
            dg.diem_so,
            dg.nhan_xet,
            dg.ngay_danh_gia
        FROM Danh_Gia dg
        JOIN Nhan_Vien nv ON dg.ma_nhan_vien = nv.ma_nhan_vien
        JOIN Khoa_Dao_Tao kdt ON dg.ma_khoa_dao_tao = kdt.ma_khoa_dao_tao
    """

    if ma_nv:
        sql += " WHERE nv.ma_nhan_vien = ?"
        cursor.execute(sql, ma_nv)
    else:
        cursor.execute(sql)

    rows = cursor.fetchall()

    df = pd.DataFrame(rows, columns=[
        "Mã NV", "Nhân viên", "Khóa đào tạo",
        "Điểm", "Nhận xét", "Ngày đánh giá"
    ])

    output = io.BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    return send_file(
        output,
        download_name="bao_cao_dao_tao.xlsx",
        as_attachment=True
    )
