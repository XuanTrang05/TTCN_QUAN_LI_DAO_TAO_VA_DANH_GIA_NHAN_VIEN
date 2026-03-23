from flask import Blueprint, render_template, request, redirect, url_for, current_app, jsonify, send_file, session
from datetime import date, datetime
from openpyxl import Workbook
from io import BytesIO
from co_so_du_lieu.ket_noi_sql import ket_noi_sql
import pandas as pd
import io
import os  


nguoi_danh_gia_bp = Blueprint(
    "nguoi_danh_gia",
    __name__,
    url_prefix="/nguoi_danh_gia"
)

# ================= TRANG CHỦ =================
@nguoi_danh_gia_bp.route("/", endpoint="trang_chu")
def trang_chu():

    username = session.get("username")
    if not username:
        return redirect("/login")

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT nv.ho_ten
        FROM Tai_Khoan tk
        JOIN Nhan_Vien nv ON tk.ma_nhan_vien = nv.ma_nhan_vien
        WHERE tk.ten_dang_nhap = ?
    """, (username,))

    user = cursor.fetchone()
    ten_user = user[0] if user else username

    return render_template(
        "nguoi_danh_gia/trang_chu.html",
        ten_user=ten_user
    )


# ================= DANH SÁCH KHÓA ĐƯỢC PHÂN CÔNG =================
@nguoi_danh_gia_bp.route("/khoa-dao-tao")
def danh_sach_khoa():

    username = session.get("username")
    if not username:
        return redirect("/login")

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ma_nhan_vien 
        FROM Tai_Khoan
        WHERE ten_dang_nhap = ?
    """, (username,))
    
    row = cursor.fetchone()
    if not row:
        return "Không tìm thấy người dùng"

    ma_nv = row[0]

    cursor.execute("""
        SELECT 
            k.ma_khoa_dao_tao,
            k.ten_khoa_dao_tao,
            k.ngay_bat_dau,
            k.ngay_ket_thuc,
            k.si_so_toi_da,
            nv.ho_ten as ten_nguoi_danh_gia
        FROM Khoa_Dao_Tao k
        LEFT JOIN Nhan_Vien nv 
            ON k.ma_giang_vien = nv.ma_nhan_vien
        WHERE k.ma_giang_vien = ?
        ORDER BY k.ngay_bat_dau DESC
    """, (ma_nv,))

    ds = cursor.fetchall()

    today = date.today()

    dang_dien_ra = []
    sap_dien_ra = []
    da_hoan_thanh = []

    for k in ds:
        if k.ngay_bat_dau and k.ngay_ket_thuc:
            if k.ngay_bat_dau <= today <= k.ngay_ket_thuc:
                dang_dien_ra.append(k)
            elif today < k.ngay_bat_dau:
                sap_dien_ra.append(k)
            else:
                da_hoan_thanh.append(k)

    return render_template(
        "nguoi_danh_gia/danh_sach_khoa.html",
        dang_dien_ra=dang_dien_ra,
        sap_dien_ra=sap_dien_ra,
        da_hoan_thanh=da_hoan_thanh
    )


# ================= DS NHÂN VIÊN TRONG KHÓA =================
@nguoi_danh_gia_bp.route("/khoa/<ma_khoa>")
def ds_nhan_vien_khoa(ma_khoa):

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()
# ===== LẤY TÊN KHÓA =====
    cursor.execute("""
        SELECT ten_khoa_dao_tao
        FROM Khoa_Dao_Tao
        WHERE ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    khoa_row = cursor.fetchone()
    ten_khoa = khoa_row[0] if khoa_row else ""


# ===== LẤY DS HỌC VIÊN + ĐIỂM =====
    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            pc.trang_thai,
            dg.diem_so
        FROM Phan_Cong_Dao_Tao pc
        JOIN Nhan_Vien nv 
            ON pc.ma_nhan_vien = nv.ma_nhan_vien
        LEFT JOIN Danh_Gia dg
            ON dg.ma_nhan_vien = nv.ma_nhan_vien
            AND dg.ma_khoa_dao_tao = pc.ma_khoa_dao_tao
        WHERE pc.ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    rows = cursor.fetchall()

    ds_nv = []
    for r in rows:

        diem = r[3]
        pc_trang_thai = r[2]

    # ===== Có điểm → Đạt / Chưa đạt =====
        if diem is not None:

            if float(diem) >= 5:
                trang_thai = "Đạt"
            else:
                trang_thai = "Chưa đạt"

    # ===== Chưa có điểm =====
        else:
            if pc_trang_thai:

                if pc_trang_thai.lower() == "chưa học":
                    trang_thai = "Đang học"
                else:
                    trang_thai = pc_trang_thai

            else:
                trang_thai = "Đang học"

        ds_nv.append((
            r[0],
            r[1],
            trang_thai
    ))
    # ⭐ FIX TRANG TRẮNG: luôn render cham_diem.html để dùng chung
    return render_template(
        "nguoi_danh_gia/cham_diem.html",
        ds_nv=ds_nv,
        ma_khoa=ma_khoa,
        ten_khoa=ten_khoa,
        nv=None
    )
# ================= FORM CHẤM ĐIỂM =================
@nguoi_danh_gia_bp.route("/cham-diem/<ma_khoa>/<ma_nv>", methods=["GET","POST"])
def cham_diem(ma_khoa, ma_nv):

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    # ===== POST = LƯU ĐIỂM =====
    if request.method == "POST":

        diem = request.form.get("diem")
        nhan_xet = request.form.get("nhan_xet")

        # CHECK EXIST
        cursor.execute("""
            SELECT 1 FROM Danh_Gia
            WHERE ma_nhan_vien = ?
            AND ma_khoa_dao_tao = ?
        """, (ma_nv, ma_khoa))

        exist = cursor.fetchone()

        if exist:
            cursor.execute("""
                UPDATE Danh_Gia
                SET diem_so=?, nhan_xet=?, ngay_danh_gia=?
                WHERE ma_nhan_vien=? AND ma_khoa_dao_tao=?
            """, (
                diem,
                nhan_xet,
                date.today(),
                ma_nv,
                ma_khoa
            ))
        else:
            cursor.execute("""
                INSERT INTO Danh_Gia
                (ma_nhan_vien, ma_khoa_dao_tao, diem_so, nhan_xet, ngay_danh_gia)
                VALUES (?, ?, ?, ?, ?)
            """, (
                ma_nv,
                ma_khoa,
                diem,
                nhan_xet,
                date.today()
            ))

        conn.commit()

        return redirect(url_for(
            "nguoi_danh_gia.ds_nhan_vien_khoa",
            ma_khoa=ma_khoa
        ))

    # ===== GET = HIỆN FORM =====
    cursor.execute("""
        SELECT ho_ten 
        FROM Nhan_Vien
        WHERE ma_nhan_vien = ?
    """, (ma_nv,))

    nv = cursor.fetchone()

    return render_template(
        "nguoi_danh_gia/cham_diem.html",
        nv=nv,
        ma_nv=ma_nv,
        ma_khoa=ma_khoa,
        ds_nv=None
    )    
# =====================================================
# API LẤY DANH SÁCH HỌC VIÊN
# =====================================================
@nguoi_danh_gia_bp.route("/api/hoc_vien/<ma_khoa>")
def api_hoc_vien(ma_khoa):

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            nv.email
        FROM Phan_Cong_Dao_Tao pc
        JOIN Nhan_Vien nv 
            ON pc.ma_nhan_vien = nv.ma_nhan_vien
        WHERE pc.ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    ds = cursor.fetchall()

    data = []
    for row in ds:
        data.append({
            "ma": row[0],
            "ten": row[1],
            "email": row[2]
        })

    return jsonify(data)
# =====================================================
# XUẤT EXCEL DANH SÁCH HỌC VIÊN
# =====================================================
@nguoi_danh_gia_bp.route("/xuat_excel/<ma_khoa>")
def xuat_excel(ma_khoa):

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ma_nhan_vien,
            nv.ho_ten,
            nv.email,
            pc.trang_thai
        FROM Phan_Cong_Dao_Tao pc
        JOIN Nhan_Vien nv 
            ON pc.ma_nhan_vien = nv.ma_nhan_vien
        WHERE pc.ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    rows = cursor.fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Danh sách học viên"

    ws.append([
        "Mã nhân viên",
        "Họ tên",
        "Email",
        "Trạng thái"
    ])

    for r in rows:
        ws.append(list(r))

    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=f"danh_sach_hoc_vien_khoa_{ma_khoa}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ================= MỞ TRANG THỰC HIỆN ĐÁNH GIÁ =================
@nguoi_danh_gia_bp.route("/thuc_hien_danh_gia")
def thuc_hien_danh_gia():
    return render_template("nguoi_danh_gia/cham_diem.html")


# ================= API KHÓA HOÀN THÀNH =================
@nguoi_danh_gia_bp.route("/api/khoa_hoan_thanh")
def api_khoa_hoan_thanh():

    username = session.get("username")

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ma_nhan_vien 
        FROM Tai_Khoan 
        WHERE ten_dang_nhap = ?
    """, (username,))

    row = cursor.fetchone()
    if not row:
        return jsonify({
            "chua_cham_xong": [],
            "da_cham_xong": []
        })

    ma_nv = row[0]

    cursor.execute("""
        SELECT 
            k.ma_khoa_dao_tao,
            k.ten_khoa_dao_tao,
            k.ngay_bat_dau,
            k.ngay_ket_thuc,
            nv.ho_ten
        FROM Khoa_Dao_Tao k
        LEFT JOIN Nhan_Vien nv 
            ON k.ma_giang_vien = nv.ma_nhan_vien
        WHERE k.ma_giang_vien = ?
        AND k.ngay_ket_thuc < GETDATE()
    """, (ma_nv,))

    khoa_list = cursor.fetchall()

    chua_cham = []
    da_cham = []

    for k in khoa_list:

        ma_khoa = k[0]

        cursor.execute("""
            SELECT COUNT(*) 
            FROM Phan_Cong_Dao_Tao
            WHERE ma_khoa_dao_tao = ?
        """, (ma_khoa,))
        tong = cursor.fetchone()[0]

        cursor.execute("""
            SELECT COUNT(DISTINCT ma_nhan_vien) 
            FROM Danh_Gia
            WHERE ma_khoa_dao_tao = ?
        """, (ma_khoa,))
        cham = cursor.fetchone()[0]

        data = {
            "ma": ma_khoa,
            "ten": k[1],
            "bd": str(k[2]),
            "kt": str(k[3]),
            "gv": k[4] if k[4] else "",
            "tong": tong,
            "cham": cham
        }

        if tong == cham and tong != 0:
            da_cham.append(data)
        else:
            chua_cham.append(data)

    return jsonify({
        "chua_cham_xong": chua_cham,
        "da_cham_xong": da_cham
    })


# ================= API BIỂU ĐỒ =================
@nguoi_danh_gia_bp.route("/api/chart_diem")
def api_chart_diem():

    username = session.get("username")

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ma_nhan_vien 
        FROM Tai_Khoan 
        WHERE ten_dang_nhap = ?
    """, (username,))

    row = cursor.fetchone()
    if not row:
        return jsonify([])

    ma_nv = row[0]

    cursor.execute("""
        SELECT 
            k.ten_khoa_dao_tao,
            AVG(d.diem_so)
        FROM Khoa_Dao_Tao k
        JOIN Danh_Gia d 
            ON k.ma_khoa_dao_tao = d.ma_khoa_dao_tao
        WHERE k.ma_giang_vien = ?
        GROUP BY k.ten_khoa_dao_tao
    """, (ma_nv,))

    data = []
    for r in cursor.fetchall():
        data.append({
            "ten": r[0],
            "diem": float(r[1]) if r[1] else 0
        })

    return jsonify(data)
# ================= XEM LẠI ĐÁNH GIÁ =================
@nguoi_danh_gia_bp.route("/xem_lai/<ma_khoa>/<ma_nv>")
def xem_lai_danh_gia(ma_khoa, ma_nv):

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            nv.ho_ten,
            d.diem_so,
            d.nhan_xet,
            d.ngay_danh_gia
        FROM Danh_Gia d
        JOIN Nhan_Vien nv 
            ON d.ma_nhan_vien = nv.ma_nhan_vien
        WHERE d.ma_khoa_dao_tao = ?
        AND d.ma_nhan_vien = ?
    """, (ma_khoa, ma_nv))

    row = cursor.fetchone()

    if not row:
        return "Chưa có dữ liệu đánh giá"

    data = {
        "ten_nv": row[0],
        "diem": row[1],
        "nhan_xet": row[2],
        "ngay": str(row[3]) if row[3] else ""
    }
# ===== LẤY TÊN KHÓA =====
    cursor.execute("""
        SELECT ten_khoa_dao_tao
        FROM Khoa_Dao_Tao
        WHERE ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    khoa_row = cursor.fetchone()
    ten_khoa = khoa_row[0] if khoa_row else ""

    return render_template(
        "nguoi_danh_gia/cham_diem.html",
        xem_lai=data,
        ma_khoa=ma_khoa,
        ma_nv=ma_nv,
        ten_khoa=ten_khoa,
        ds_nv=None,
        nv=None
    )

# ================= XUẤT EXCEL FULL =================
@nguoi_danh_gia_bp.route("/xuat_excel_full/<ma_khoa>")
def xuat_excel_full(ma_khoa):

    conn = current_app.config['DB_CONN']
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            k.ten_khoa_dao_tao,
            k.ngay_bat_dau,
            k.ngay_ket_thuc,
            gv.ho_ten,
            nv.ho_ten,
            nv.ngay_sinh,
            d.diem_so,
            d.nhan_xet
        FROM Khoa_Dao_Tao k
        LEFT JOIN Nhan_Vien gv 
            ON k.ma_giang_vien = gv.ma_nhan_vien
        LEFT JOIN Danh_Gia d 
            ON k.ma_khoa_dao_tao = d.ma_khoa_dao_tao
        LEFT JOIN Nhan_Vien nv 
            ON d.ma_nhan_vien = nv.ma_nhan_vien
        WHERE k.ma_khoa_dao_tao = ?
    """, (ma_khoa,))

    rows = cursor.fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Bao cao khoa"

    ws.append([
        "Tên khóa",
        "Ngày bắt đầu",
        "Ngày kết thúc",
        "Giảng viên",
        "Tên học viên",
        "Ngày sinh",
        "Điểm",
        "Nhận xét"
    ])

    for r in rows:
        ws.append(list(r))

    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=f"bao_cao_khoa_{ma_khoa}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ================= API THỐNG KÊ =================
@nguoi_danh_gia_bp.route("/api/thong-ke-danh-gia")
def api_thong_ke():

    khoa = request.args.get("khoa")
    from_date = request.args.get("from")
    to_date = request.args.get("to")

    conn = ket_noi_sql()
    cursor = conn.cursor()

    query = """
        SELECT 
            k.ten_khoa_dao_tao,
            nv.ho_ten,
            dg.diem_so,
            dg.ngay_danh_gia
        FROM Danh_Gia dg
        JOIN Nhan_Vien nv 
            ON dg.ma_nhan_vien = nv.ma_nhan_vien
        JOIN Khoa_Dao_Tao k 
            ON dg.ma_khoa_dao_tao = k.ma_khoa_dao_tao
        WHERE 1=1
    """

    params = []

    # ===== FILTER KHÓA =====
    if khoa and khoa != "all":
        query += " AND k.ma_khoa_dao_tao = ?"
        params.append(khoa)

    # ===== FILTER FROM DATE =====
    if from_date:
        query += " AND dg.ngay_danh_gia >= ?"
        params.append(from_date)

    # ===== FILTER TO DATE =====
    if to_date:
        query += " AND dg.ngay_danh_gia <= ?"
        params.append(to_date)

    cursor.execute(query, params)
    rows = cursor.fetchall()

    data = []

    for r in rows:
        data.append({
            "ten_khoa": r[0] if r[0] else "",
            "ten_nhan_vien": r[1] if r[1] else "",
            "diem": float(r[2]) if r[2] else 0,
            "ngay_danh_gia": r[3].strftime("%Y-%m-%d") if r[3] else ""
        })

    cursor.close()
    conn.close()

    return jsonify(data)
# ================= PAGE THỐNG KÊ =================
@nguoi_danh_gia_bp.route("/thong_ke_danh_gia")
def page_thong_ke_danh_gia():
    return render_template("nguoi_danh_gia/danh_gia_thong_ke.html")
