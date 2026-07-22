import os
from dotenv import load_dotenv
from flask import Flask, request, jsonify
from flask_cors import CORS
from config import get_db_connection
from datetime import date, datetime, time, timedelta, timezone
from math import ceil, radians, sin, cos, sqrt, atan2
import jwt
import locale
from openpyxl import Workbook
from io import BytesIO
from flask import send_file
from functools import wraps

load_dotenv()

app = Flask(__name__)

app.secret_key = os.getenv("SECRET_KEY")

SECRET_KEY = os.getenv("SECRET_KEY")

CORS(app, supports_credentials=True, origins=["http://localhost:5173", "https://web-presensi-seminar-hasil-nvvz.vercel.app"])

try:
    locale.setlocale(locale.LC_TIME, "id_ID.UTF-8")   # Linux/Mac
except:
    try:
        locale.setlocale(locale.LC_TIME, "Indonesian_Indonesia.1252")   # Windows
    except:
        pass

TIMEZONE_INDO = timezone(timedelta(hours=8))

#Fungsi helper untuk format waktu
def parse_date(val):
    if not val:
        return date.today()
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        try:
            return datetime.strptime(val.split(" ")[0], "%Y-%m-%d").date()
        except ValueError:
            return date.today()
    return date.today()


def parse_time(val):
    if not val:
        return time(0, 0)
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    if isinstance(val, timedelta):  # Tipe data TIME MySQL
        total_sec = int(val.total_seconds())
        h = (total_sec // 3600) % 24
        m = (total_sec % 3600) // 60
        return time(h, m)
    if isinstance(val, str):
        try:
            clean = val.replace(".", ":").strip()
            p = clean.split(":")
            if len(p) >= 2:
                return time(int(p[0]), int(p[1]))
        except Exception:
            return time(0, 0)
    return time(0, 0)

def format_waktu(val):
    """
    Menerima input jam dari database (None, timedelta, string, time)
    dan mengembalikan string format 'HH:MM' tanpa melempar error.
    """
    if val is None:
        return "-"
    t = parse_time(val)
    return t.strftime("%H:%M")

#Fungsi helper untuk menghitung jarak antara lokasi dan perangkat mahasiswa ketika melakukan presensi menggunakan Haversine
def hitung_jarak(lat1, lon1, lat2, lon2):
    R = 6371000 #meter

    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)

    a = (sin(dlat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon / 2) ** 2)

    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    return R * c

# Fungsi helper untuk memeriksa apakah user sudah login 
def login_required(f):
    @wraps(f)
    def decorated(*args, ** kwargs):
        token = request.headers.get("Authorization")

        print("Authorization Header:", token)

        if not token:
            return jsonify({
                "success": False,
                "message": "Token tidak ditemukan"
            }), 401
        
        try:
            token = token.replace("Bearer ", "")

            print("SECRET_KEY:", SECRET_KEY)
            print("TOKEN:", token)

            payload = jwt.decode(
                token,
                SECRET_KEY,
                algorithms=["HS256"]
            )
            print("PAYLOAD:", payload)

            request.user = payload

        except jwt.ExpiredSignatureError:
            return jsonify({
                "success": False,
                "where": "login_required",
                "message": "Token sudah kedaluwarsa"
            }), 401
        
        except jwt.InvalidTokenError as e:
            print(e)

            return jsonify({
                "success": False,
                "where": "login_required",
                "message": "Token tidak valid"
            }), 401
        
        return f(*args, ** kwargs)
    
    return decorated

def role_required(*roles):
    def wrapper(f):
        @wraps(f)
        def decorated(*args, **kwargs):

            if request.user["role"] not in roles:
                return jsonify({
                    "success": False,
                    "message": "Forbidden"
                }), 403

            return f(*args, **kwargs)

        return decorated
    return wrapper

@app.route("/debug-time")
def debug_time():
    from datetime import datetime, timezone
    import time

    return jsonify({
        "python_now": str(datetime.now()),
        "python_utc": str(datetime.now(timezone.utc)),
        "python_utcnow": str(datetime.utcnow()),
        "time_time": time.time()
    })

#Menampilkan detail riwayat verifikasi, fitur search, fitur sort, fitur filter, card dan progress bar halaman Riwayat Verifikasi Lihat Detail - Verifikator
@app.route("/riwayat-verifikasi/<int:id_seminar>")
@login_required
@role_required("verifikator")
def detail_riwayat_verifikasi(id_seminar):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Pagination
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 5))
    offset = (page - 1) * limit

    #Parameter sorting
    sort_by = request.args.get("sort_by", "waktu_scan")
    sort_order = request.args.get("sort_order", "desc").lower()

    # Search
    search = request.args.get("search", "").strip()

    #Filter
    status_verifikasi = request.args.get("status_verifikasi", "")

    #Validasi agar aman dari SQL Injection
    allowed_columns = {"nama":"m.nama", "nim":"m.nim", "waktu_scan":"p.waktu_scan", "waktu_verifikasi":"p.waktu_verifikasi"}
    allowed_orders = ["asc", "desc"]

    if sort_by not in allowed_columns:
        sort_by = "waktu_scan"

    if sort_order not in allowed_orders:
        sort_order = "desc"

    conditions = ["p.id_seminar = %s"]
    params = [id_seminar]

    #Search
    if search:
        keyword = f"%{search}%"
        conditions.append("(m.nama LIKE %s OR m.nim LIKE %s)")
        params.extend([keyword, keyword])

    # Filter status
    if status_verifikasi:
        conditions.append("p.status_verifikasi = %s")
        params.append(status_verifikasi)

    where_clause = "WHERE " + " AND ".join(conditions)
    order_clause = f"{allowed_columns[sort_by]} {sort_order.upper()}"

    #Hitung data
    count_query = f"""
        SELECT COUNT(*) AS total
        FROM presensi p
        JOIN mahasiswa m
            ON m.id_user = p.id_mahasiswa
        {where_clause}
    """

    cursor.execute(count_query, params)
    total_data = cursor.fetchone()["total"]
    total_pages = ceil(total_data / limit)

    #Hitung data untuk card total peserta, dan total status (pending, valid, invalid)
    card_query = """
        SELECT
            COUNT(*) AS total_peserta,
            COALESCE(SUM(CASE WHEN status_verifikasi = 'pending' THEN 1 ELSE 0 END),0) AS total_pending,
            COALESCE(SUM(CASE WHEN status_verifikasi = 'valid' THEN 1 ELSE 0 END),0) AS total_valid,
            COALESCE(SUM(CASE WHEN status_verifikasi = 'invalid' THEN 1 ELSE 0 END),0) AS total_invalid
        FROM presensi
        WHERE id_seminar = %s
    """

    cursor.execute(card_query, (id_seminar,))
    card = cursor.fetchone()

    #Ambil data sesuai halaman
    query = f"""
        SELECT
            p.id_presensi,
            m.nama,
            m.nim,
            p.waktu_scan,
            p.status_verifikasi,
            p.waktu_verifikasi
        FROM presensi p
        JOIN mahasiswa m
            ON m.id_user = p.id_mahasiswa
        {where_clause}
        ORDER BY {order_clause}
        LIMIT %s OFFSET %s
    """

    data_params = params.copy()
    data_params.extend([limit, offset])

    cursor.execute(query, data_params)
    data = cursor.fetchall()

    for item in data:
        # Format waktu scan
        if item["waktu_scan"]:
            # Anggap naive datetime sebagai UTC, lalu konversi ke TIMEZONE_INDO
            dt_scan = item["waktu_scan"].replace(tzinfo=timezone.utc).astimezone(TIMEZONE_INDO)
            item["waktu_scan"] = dt_scan.isoformat()

        # Format waktu verifikasi
        if item["waktu_verifikasi"]:
            dt_verif = item["waktu_verifikasi"].replace(tzinfo=timezone.utc).astimezone(TIMEZONE_INDO)
            item["waktu_verifikasi"] = dt_verif.isoformat()
        else:
            item["waktu_verifikasi"] = None
    
    cursor.close()
    conn.close()

    return jsonify({
        "data": data,
        "card": card,
        "pagination": {"page": page, "limit": limit, "total_data": total_data, "total_pages": total_pages}
    })

#Menampilkan data progress riwayat verifikasi, fitur search dan fitur filter halaman Riwayat Verifikasi - Verifikator
@app.route("/riwayat-verifikasi")
@login_required
@role_required("verifikator")
def riwayat_verifikasi():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Search
    search = request.args.get("search", "").strip()

    #Filter
    status = request.args.get("status", "")
    tanggal_filter = request.args.get("tanggal", "Semua")
    tanggal_awal = request.args.get("tanggal_awal")
    tanggal_akhir = request.args.get("tanggal_akhir")

    conditions = []
    params = []

    having_conditions = []

    #Search
    if search:
        conditions.append("m.nama LIKE %s")
        params.append(f"%{search}%")

    #Filter
    #Status
    if status == "belum":
        having_conditions.append("""
            (COUNT(p.id_presensi) = 0 OR COALESCE(SUM(CASE WHEN p.status_verifikasi IN ('valid', 'invalid') THEN 1 ELSE 0 END), 0) = 0)
        """)
    elif status == "sedang":
        having_conditions.append("""
            (COALESCE(SUM(CASE WHEN p.status_verifikasi IN ('valid', 'invalid') THEN 1 ELSE 0 END), 0) > 0) AND COALESCE(SUM(CASE WHEN p.status_verifikasi IN ('valid', 'invalid') THEN 1 ELSE 0 END), 0) < COUNT(p.id_presensi)
        """)
    elif status == "selesai":
        having_conditions.append("""
            COUNT(p.id_presensi) > 0 AND COALESCE(SUM(CASE WHEN p.status_verifikasi IN ('valid', 'invalid') THEN 1 ELSE 0 END), 0) = COUNT(p.id_presensi)
        """)

    #Tanggal
    if tanggal_filter == "Hari Ini":
        conditions.append("DATE(s.tanggal) = CURDATE()")
    elif tanggal_filter == "Minggu Ini":
        conditions.append("YEARWEEK(s.tanggal,1)=YEARWEEK(CURDATE(),1)")
    elif tanggal_filter == "Bulan Ini":
        conditions.append("""
            MONTH(s.tanggal)=MONTH(CURDATE())
            AND YEAR(s.tanggal)=YEAR(CURDATE())
        """)
    #Rentang tanggal
    elif tanggal_awal and tanggal_akhir:
        conditions.append("DATE(s.tanggal) BETWEEN %s AND %s")
        params.extend([tanggal_awal, tanggal_akhir])
    
    where_clause = ""

    having_clause = ""

    if conditions:
        where_clause = "WHERE " + " AND ".join(conditions)

    if having_conditions:
        having_clause = "HAVING " + " AND ".join(having_conditions)

    #Ambil data
    query = f"""
        SELECT
            s.id_seminar,
            m.nama,
            s.tanggal,
            s.waktu_mulai,
            s.waktu_selesai,
        COUNT(p.id_presensi) AS total_presensi,
        SUM(CASE WHEN p.status_verifikasi IN ('valid', 'invalid') THEN 1 ELSE 0 END) AS selesai_diproses
        FROM seminar s
        JOIN mahasiswa m
            ON m.id_user = s.id_mahasiswa
        LEFT JOIN presensi p
            ON p.id_seminar = s.id_seminar
        {where_clause}
        GROUP BY
            s.id_seminar,
            m.nama,
            s.tanggal,
            s.waktu_mulai,
            s.waktu_selesai
        {having_clause}
        ORDER BY s.tanggal DESC
    """

    cursor.execute(query, params)
    data = cursor.fetchall()

    for item in data:
        # Format tanggal
        item["tanggal"] = item["tanggal"].isoformat()

        # Format jam
        item["waktu_mulai"] = format_waktu(item["waktu_mulai"])
        item["waktu_selesai"] = format_waktu(item["waktu_selesai"])

        #Status
        item["selesai_diproses"] = item["selesai_diproses"] or 0

    cursor.close()
    conn.close()

    return jsonify({
        "data": data
    })

#Menampilkan data riwayat presensi, fitur search, fitur filter dan card statistik kehadiran di halaman Riwayat Presensi - Mahasiswa
@app.route("/riwayat-presensi-mahasiswa")
@login_required
@role_required("mahasiswa")
def riwayat_presensi_mahasiswa():
    conn = get_db_connection ()
    cursor = conn.cursor(dictionary=True)

    id_mahasiswa = request.user.get("id_user")

    if not id_mahasiswa:
        return jsonify({"message": "Unauthorized"}), 401

    #Search
    search = request.args.get("search", "").strip()

    #Filter
    status = request.args.get("status", "")
    tanggal_filter = request.args.get("tanggal", "Semua")
    tanggal_awal = request.args.get("tanggal_awal")
    tanggal_akhir = request.args.get("tanggal_akhir")

    conditions = ["p.id_mahasiswa = %s"]
    params = [id_mahasiswa]

    #Search
    if search:
        keyword = f"%{search}%"
        conditions.append("(m.nama LIKE %s OR s.judul_penelitian LIKE %s OR s.dosen_pembimbing LIKE %s OR s.dosen_penguji_1 LIKE %s OR s.dosen_penguji_2 LIKE %s)")
        params.extend([keyword, keyword, keyword, keyword, keyword])

    #Filter
    #Status
    if status:
        conditions.append("p.status_verifikasi = %s")
        params.append(status)

    #Tanggal
    if tanggal_filter == "Hari Ini":
        conditions.append("DATE(s.tanggal) = CURDATE()")
    elif tanggal_filter == "Minggu Ini":
        conditions.append("YEARWEEK(s.tanggal,1)=YEARWEEK(CURDATE(),1)")
    elif tanggal_filter == "Bulan Ini":
        conditions.append("""
            MONTH(s.tanggal)=MONTH(CURDATE())
            AND YEAR(s.tanggal)=YEAR(CURDATE())
        """)
    #Rentang tanggal
    elif tanggal_awal and tanggal_akhir:
        conditions.append("DATE(s.tanggal) BETWEEN %s AND %s")
        params.extend([tanggal_awal, tanggal_akhir])

    where_clause = "WHERE " + " AND ".join(conditions)

    #Data statistik kehadiran
    statistik_query = f"""
        SELECT
            COUNT(*) AS total_kehadiran,
            COALESCE(SUM(CASE WHEN p.status_verifikasi = 'valid' THEN 1 ELSE 0 END),0) AS kehadiran_valid,
            COALESCE(SUM(CASE WHEN p.status_verifikasi = 'pending' THEN 1 ELSE 0 END),0) AS kehadiran_pending
        FROM presensi p
        WHERE id_mahasiswa = %s
    """
    cursor.execute(statistik_query, (id_mahasiswa,))
    statistik = cursor.fetchone()

    #Ambil data
    data_query = f"""
        SELECT
            p.id_presensi,
            p.status_verifikasi,
            
            s.judul_penelitian,
            s.tanggal,
            s.waktu_mulai,
            s.waktu_selesai,

            m.nama AS nama_mahasiswa,
            s.dosen_pembimbing,
            s.dosen_penguji_1,
            s.dosen_penguji_2
        FROM presensi p
        JOIN seminar s
            ON s.id_seminar = p.id_seminar
        JOIN mahasiswa m
            ON m.id_user = s.id_mahasiswa
        {where_clause}
        ORDER BY s.tanggal DESC
    """

    cursor.execute(data_query, params)
    data = cursor.fetchall()

    for item in data:
        # Format tanggal
        item["tanggal"] = item["tanggal"].isoformat()

        # Format jam
        item["waktu_mulai"] = format_waktu(item["waktu_mulai"])
        item["waktu_selesai"] = format_waktu(item["waktu_selesai"])

    cursor.close()
    conn.close()

    return jsonify({
        "statistik": statistik,
        "data": data
    })

#Export laporan presensi ke excel
@app.route("/laporan-presensi/export")
@login_required
@role_required("admin")
def export_laporan_presensi():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Sort
    sort_by = request.args.get("sort_by", "nama")
    sort_order = request.args.get("sort_order", "asc").lower()

    # Search
    search = request.args.get("search", "").strip()

    #Filter
    angkatan = request.args.get("angkatan", "")
    status = request.args.get("status", "")

    #Validasi agar aman dari SQL Injection
    allowed_columns = {"nama":"m.nama", "nim":"m.nim", "angkatan":"m.angkatan"}
    allowed_orders = ["asc", "desc"]

    if sort_by not in allowed_columns:
        sort_by = "nama"

    if sort_order not in allowed_orders:
        sort_order = "asc"

    conditions = []
    params = []

    #Search
    if search:
        keyword = f"%{search}%"
        conditions.append("(m.nama LIKE %s OR m.nim LIKE %s)")
        params.extend([keyword, keyword])

    #Filter angkatan
    if angkatan:
        conditions.append("m.angkatan = %s")
        params.append(angkatan)

    where_clause = ""

    if conditions:
        where_clause = "WHERE " + " AND ".join(conditions)

    having_clause = ""

    if status == "Memenuhi":
        having_clause = """
            HAVING COUNT(
                CASE
                    WHEN p.status_verifikasi = 'valid'
                    THEN p.id_presensi
                END
            ) >= 3
        """
    elif status == "Belum Memenuhi":
        having_clause = """
            HAVING COUNT(
                CASE
                    WHEN p.status_verifikasi = 'valid'
                    THEN p.id_presensi
                END
            ) < 3
        """

    #Hitung total data
    count_query = f"""
        SELECT COUNT(*) AS total_data
        FROM (
            SELECT
                m.id_user

            FROM mahasiswa m
            LEFT JOIN presensi p
                ON p.id_mahasiswa = m.id_user

            {where_clause}

            GROUP BY
                m.id_user,
                m.nama,
                m.nim,
                m.angkatan

            {having_clause}
        ) AS hasil
    """

    #Ambil data sesuai halaman
    data_query = f"""
        SELECT
            m.id_user,
            m.nama,
            m.nim,
            m.angkatan,

            COUNT(CASE WHEN p.status_verifikasi = 'valid' THEN p.id_presensi END) AS kehadiran

        FROM mahasiswa m
        LEFT JOIN presensi p
            ON m.id_user = p.id_mahasiswa
        {where_clause}
        GROUP BY
            m.id_user, m.nama, m.nim, m.angkatan
        {having_clause}
        ORDER BY {allowed_columns[sort_by]} {sort_order.upper()}
    """ 

    cursor.execute(data_query, params)
    data = cursor.fetchall()

    #Filter status
    for item in data:
        item["status"] = (
            "Memenuhi"
            if item["kehadiran"] >= 3
            else "Belum Memenuhi"
        )

    #File excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan Presensi"

    #Header
    ws.append(["Nama", "NIM", "Angkatan", "Kehadiran", "Status"])

    #Isi data
    for item in data:
        ws.append([item["nama"], item["nim"], item["angkatan"], item["kehadiran"], item["status"]])

    #Simpan file
    file = BytesIO()
    wb.save(file)
    file.seek(0)

    cursor.close()
    conn.close()

    return send_file(
        file,
        as_attachment=True,
        download_name="laporan_presensi.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

#Menampilkan data angkatan untuk masuk ke filter
@app.route("/data-angkatan-laporan")
@login_required
@role_required("admin")
def get_data_angkatan_laporan():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT DISTINCT angkatan
        FROM mahasiswa
        ORDER BY angkatan DESC
    """)

    data = cursor.fetchall()

    cursor.close()
    conn.close()

    return jsonify(data)

#Menampilkan data laporan presensi, pagination, fitur search dan fitur filter dan download di halaman Laporan Presensi - Admin
@app.route("/laporan-presensi")
@login_required
@role_required("admin")
def laporan_presensi():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Pagination
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 10))
    offset = (page - 1) * limit

    #Sort
    sort_by = request.args.get("sort_by", "nama")
    sort_order = request.args.get("sort_order", "asc").lower()

    # Search
    search = request.args.get("search", "").strip()

    #Filter
    angkatan = request.args.get("angkatan", "")
    status = request.args.get("status", "")

    #Validasi agar aman dari SQL Injection
    allowed_columns = {"nama":"m.nama", "nim":"m.nim", "angkatan":"m.angkatan"}
    allowed_orders = ["asc", "desc"]

    if sort_by not in allowed_columns:
        sort_by = "nama"

    if sort_order not in allowed_orders:
        sort_order = "asc"

    conditions = []
    params = []

    #Search
    if search:
        keyword = f"%{search}%"
        conditions.append("(m.nama LIKE %s OR m.nim LIKE %s)")
        params.extend([keyword, keyword])

    #Filter angkatan
    if angkatan:
        conditions.append("m.angkatan = %s")
        params.append(angkatan)

    where_clause = ""

    if conditions:
        where_clause = "WHERE " + " AND ".join(conditions)

    having_clause = ""
    if status == "Memenuhi":
        having_clause = """
            HAVING COUNT(
                CASE
                    WHEN p.status_verifikasi = 'valid'
                    THEN p.id_presensi
                END
            ) >= 3
    """
    elif status == "Belum Memenuhi":
        having_clause = """
            HAVING COUNT(
                CASE
                    WHEN p.status_verifikasi = 'valid'
                    THEN p.id_presensi
                END
            ) < 3
    """

    #Hitung total data
    count_query = f"""
        SELECT COUNT(*) AS total_data
        FROM (
            SELECT
                m.id_user

            FROM mahasiswa m
            LEFT JOIN presensi p
                ON p.id_mahasiswa = m.id_user

            {where_clause}

            GROUP BY
                m.id_user,
                m.nama,
                m.nim,
                m.angkatan

            {having_clause}
        ) AS hasil
    """

    cursor.execute(count_query, params)
    total_data = cursor.fetchone()["total_data"]
    total_pages = ceil(total_data / limit)

    #Ambil data sesuai halaman
    data_query = f"""
        SELECT
            m.id_user,
            m.nama,
            m.nim,
            m.angkatan,

            COUNT(CASE WHEN p.status_verifikasi = 'valid' THEN p.id_presensi END) AS kehadiran

        FROM mahasiswa m
        LEFT JOIN presensi p
            ON m.id_user = p.id_mahasiswa
        {where_clause}
        GROUP BY
            m.id_user, m.nama, m.nim, m.angkatan
        {having_clause}
        ORDER BY {allowed_columns[sort_by]} {sort_order.upper()}
        LIMIT %s OFFSET %s
    """
    
    data_params = params.copy()
    data_params.extend([limit, offset])

    cursor.execute(data_query, data_params)
    data = cursor.fetchall()

    #Filter status
    for item in data:
            item["status"] = (
                "Memenuhi"
                if item["kehadiran"] >= 3
                else "Belum Memenuhi"
            )

    cursor.close()
    conn.close()

    return jsonify({
        "data": data,
        "pagination": {"page": page, "limit": limit, "total_data": total_data, "total_pages": total_pages}
    })

#Update status presensi
@app.route("/verifikator-update-status-presensi/<int:id_presensi>", methods=["PUT"])
@login_required
@role_required("verifikator")
def update_status_presensi(id_presensi):
    conn = get_db_connection()
    cursor = conn.cursor()

    if request.user["role"] != "verifikator":
        return jsonify({
            "success": False,
            "message": "Akses ditolak"
        }), 403

    data = request.get_json()
    status = data.get("status")

    id_verifikator = request.user["id_user"]

    if status not in ["pending", "valid", "invalid"]:
        return jsonify({
            "success": False,
            "message": "Status tidak valid"
        }), 400
    
    cursor.execute("""
        UPDATE presensi
        SET
            status_verifikasi = %s,
            id_user_verifikator = %s,
            waktu_verifikasi = NOW()
        WHERE id_presensi = %s
    """, (status, id_verifikator, id_presensi))

    conn.commit()

    cursor.close()
    conn.close()

    return jsonify({
        "success": True,
        "message": "Status berhasil diperbarui"
    })

#Menampilkan data daftar hadir, pagination, fitur search dan fitur filter Verifikasi Presensi - Verfikator
@app.route("/verifikator-lihat-daftar-hadir/<int:id_seminar>", methods=["GET"])
@login_required
@role_required("verifikator")
def lihat_daftar_hadir(id_seminar):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Pagination
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 5))
    offset = (page - 1) * limit

    #Parameter sorting
    sort_by = request.args.get("sort_by", "waktu_scan")
    sort_order = request.args.get("sort_order", "desc").lower()

    # Search
    search = request.args.get("search", "").strip()

    #Filter
    status_verifikasi = request.args.get("status_verifikasi", "")

    #Validasi agar aman dari SQL Injection
    allowed_columns = {"nama":"m.nama", "nim":"m.nim", "waktu_scan":"p.waktu_scan"}
    allowed_orders = ["asc", "desc"]

    if sort_by not in allowed_columns:
        sort_by = "waktu_scan"

    if sort_order not in allowed_orders:
        sort_order = "desc"

    #Untuk card total peserta
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM presensi
        WHERE id_seminar = %s
    """, (id_seminar,))

    total_peserta = cursor.fetchone()["total"]

    #Detail Seminar
    cursor.execute("""
        SELECT
            s.id_seminar,
            s.judul_penelitian,
            s.tanggal,
            s.waktu_mulai,
            s.waktu_selesai,
            m.nama
        FROM seminar s
        JOIN mahasiswa m
            ON s.id_mahasiswa = m.id_user
        WHERE s.id_seminar = %s
    """, (id_seminar,))

    seminar = cursor.fetchone()

    if seminar:
        try:
            # 1. Ambil waktu sekarang sesuai Timezone
            now_dt = datetime.now(TIMEZONE_INDO)

            # 2. Parse tanggal & waktu langsung dari objek 'seminar'
            d_obj = parse_date(seminar.get("tanggal"))
            tm_mulai = parse_time(seminar.get("waktu_mulai"))
            tm_selesai = parse_time(seminar.get("waktu_selesai"))

            # 3. Gabungkan tanggal & waktu
            dt_start = datetime.combine(d_obj, tm_mulai).replace(tzinfo=TIMEZONE_INDO)
            dt_end = datetime.combine(d_obj, tm_selesai).replace(tzinfo=TIMEZONE_INDO)

            # 4. Hitung status seminar
            if now_dt < dt_start:
                status_val = "Belum Dimulai"
            elif dt_start <= now_dt <= dt_end:
                status_val = "Sedang Berlangsung"
            else:
                status_val = "Selesai"

            # Simpan langsung ke dictionary 'seminar'
            seminar["status"] = status_val
            seminar["status_seminar"] = status_val

            # Format string untuk dikirim ke React
            seminar["tanggal"] = d_obj.strftime("%Y-%m-%d")
            seminar["waktu_mulai"] = tm_mulai.strftime("%H:%M")
            seminar["waktu_selesai"] = tm_selesai.strftime("%H:%M")

        except Exception as e:
            print(f"[WARNING] Gagal memproses status detail seminar ID {id_seminar}: {e}")
            seminar["status_seminar"] = "-"

    #Hitung total data
    count_query = """
        SELECT COUNT(*) AS total
        FROM presensi p
        JOIN mahasiswa m
            ON p.id_mahasiswa = m.id_user
        WHERE id_seminar = %s
    """

    count_params = [id_seminar]

    if search:
        count_query += """
        AND (
            m.nama LIKE %s
            OR m.nim LIKE %s
        )
        """
        keyword = f"%{search}%"
        count_params.extend([keyword, keyword])

    if status_verifikasi:
        count_query += " AND p.status_verifikasi = %s"
        count_params.append(status_verifikasi)

    cursor.execute(count_query, tuple(count_params))
    total_data = cursor.fetchone()["total"]

    total_pages = ceil(total_data / limit)

    #Daftar Hadir
    data_query = f"""
        SELECT
            p.id_presensi,
            p.waktu_scan,
            p.latitude,
            p.longitude,
            p.status_verifikasi,
            m.nama,
            m.nim,
            l.latitude AS lokasi_latitude,
            l.longitude AS lokasi_longitude
        FROM presensi p
        JOIN mahasiswa m
            ON p.id_mahasiswa = m.id_user
        JOIN seminar s
            ON p.id_seminar = s.id_seminar
        JOIN lokasi_seminar l
            ON s.id_lokasi = l.id_lokasi
        WHERE p.id_seminar = %s
    """

    data_params = [id_seminar]

    if search:
        data_query += """
        AND (
            m.nama LIKE %s
            OR m.nim LIKE %s
        )
        """
        keyword = f"%{search}%"
        data_params.extend([keyword, keyword])

    if status_verifikasi:
        data_query += " AND p.status_verifikasi = %s"
        data_params.append(status_verifikasi)

    data_query += f"""
        ORDER BY {allowed_columns[sort_by]} {sort_order.upper()}
        LIMIT %s OFFSET %s
    """
    
    data_params.extend([limit, offset])

    cursor.execute(data_query, tuple(data_params))
    presensi = cursor.fetchall()

    for item in presensi:
        #Format waktu scan di tabel
        item["waktu_scan"] = item["waktu_scan"].isoformat()

        #Menghitung jarak lokasi peserta saat scan
        jarak = hitung_jarak(
            float(item["latitude"]),
            float(item["longitude"]),
            float(item["lokasi_latitude"]),
            float(item["lokasi_longitude"])
        )

        item["jarak"] = round(jarak)

        if jarak <= 15:
            item["status_lokasi"] = "dekat"
        else:
            item["status_lokasi"] = "sedang"

    cursor.close()
    conn.close()

    return jsonify({
        "seminar": seminar,
        "data": presensi,
        "total_peserta": total_peserta,
        "pagination": {"page": page, "limit": limit, "total_data": total_data, "total_pages": total_pages}
    })

#Menampilkan data daftar seminar, fitur search dan fitur filter Verifikasi Presensi - Verfikator
@app.route("/verifikasi-presensi", methods=["GET"])
@login_required
@role_required("verifikator")
def verifikasi_presensi():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Search
    search = request.args.get("search", "").strip()

    # Filter
    tanggal_filter = request.args.get("tanggal", "Semua")
    tanggal_awal = request.args.get("tanggal_awal")
    tanggal_akhir = request.args.get("tanggal_akhir")

    #Ambil data
    query = """
        SELECT
            s.id_seminar,
            s.judul_penelitian,
            s.tanggal,
            s.waktu_mulai,
            s.waktu_selesai,
            s.dosen_pembimbing,
            s.dosen_penguji_1,
            s.dosen_penguji_2,
    
            m.nama,
            m.nim
        FROM seminar s
        JOIN mahasiswa m
            ON s.id_mahasiswa = m.id_user
        WHERE 1=1
    """

    params = []

    if search:
        query += """
        AND m.nama LIKE %s
        """
        keyword = f"%{search}%"
        params.append(keyword)

    if tanggal_filter == "Hari Ini":
        query += """
        AND DATE(s.tanggal) = CURDATE()
        """

    elif tanggal_filter == "Minggu Ini":
        query += """
        AND YEARWEEK(s.tanggal,1)=YEARWEEK(CURDATE(),1)
        """

    elif tanggal_filter == "Bulan Ini":
        query += """
        AND MONTH(s.tanggal)=MONTH(CURDATE())
        AND YEAR(s.tanggal)=YEAR(CURDATE())
        """

    elif tanggal_awal and tanggal_akhir:
        query += """
        AND DATE(s.tanggal) BETWEEN %s AND %s
        """
        params.extend([tanggal_awal, tanggal_akhir])

    query += """
        ORDER BY s.tanggal DESC, s.waktu_mulai ASC
    """

    cursor.execute(query, tuple(params))
    
    data = cursor.fetchall()

    for item in data:
        # Format tanggal
        item["tanggal"] = item["tanggal"].isoformat()

        # Format jam
        item["waktu_mulai"] = format_waktu(item["waktu_mulai"])
        item["waktu_selesai"] = format_waktu(item["waktu_selesai"])

    cursor.close()
    conn.close()

    return jsonify({
        "data": data
    })
    
#Menampilkan data daftar hadir, fitur search dan fitur sort Lihat Daftar Hadir - Mahasiswa Penyelenggara Seminar
@app.route("/daftar-hadir/<int:id_seminar>", methods=["GET"])
@login_required
@role_required("mahasiswa")
def daftar_hadir(id_seminar):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Pagination
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 10))
    offset = (page - 1) * limit

    #Parameter sorting
    sort_by = request.args.get("sort_by", "waktu_scan")
    sort_order = request.args.get("sort_order", "desc").lower()

    #Search
    search = request.args.get("search", "").strip()

    #Validasi agar aman dari SQL Injection
    allowed_columns = {"nama":"m.nama", "nim":"m.nim", "waktu_scan":"p.waktu_scan"}
    allowed_orders = ["asc", "desc"]

    if sort_by not in allowed_columns:
        sort_by = "waktu_scan"
    
    if sort_order not in allowed_orders:
        sort_order = "desc"

    #Hitung total data
    count_query = """
        SELECT COUNT(*) AS total
        FROM presensi p
        JOIN mahasiswa m
            ON p.id_mahasiswa = m.id_user
        WHERE p.id_seminar = %s
    """

    count_params = [id_seminar]

    if search:
        count_query += """
        AND (
            m.nama LIKE %s
            OR m.nim LIKE %s
        )
        """
        keyword = f"%{search}%"
        count_params.extend([keyword, keyword])

    cursor.execute(count_query, tuple(count_params))
    total_data = cursor.fetchone()["total"]

    #Ambil data sesuai halaman
    data_query = f"""
        SELECT
            p.id_presensi,
            p.waktu_scan,
            p.latitude,
            p.longitude,
            p.status_verifikasi,
            
            m.nama,
            m.nim,
                   
            l.latitude AS lokasi_latitude,
            l.longitude AS lokasi_longitude
        FROM presensi p
        JOIN mahasiswa m
            ON p.id_mahasiswa = m.id_user
        JOIN seminar s
            ON p.id_seminar = s.id_seminar
        JOIN lokasi_seminar l
            ON s.id_lokasi = l.id_lokasi
        WHERE p.id_seminar = %s
    """

    data_params = [id_seminar]

    if search:
        data_query += """
        AND (
            m.nama LIKE %s
            OR m.nim LIKE %s
        )
        """
        keyword = f"%{search}%"
        data_params.extend([keyword, keyword])

    data_query += f"""
        ORDER BY {allowed_columns[sort_by]} {sort_order.upper()}
        LIMIT %s OFFSET %s
    """
    
    data_params.extend([limit, offset])

    cursor.execute(data_query, tuple(data_params))
    data = cursor.fetchall()

    cursor.close()
    conn.close()

    #Menghitung jarak lokasi peserta saat scan
    for item in data:
        jarak = hitung_jarak(
            float(item["latitude"]),
            float(item["longitude"]),
            float(item["lokasi_latitude"]),
            float(item["lokasi_longitude"])
        )

        item["jarak"] = round(jarak)

        if jarak <= 15:
            item["status_lokasi"] = "dekat"
        else:
            item["status_lokasi"] = "sedang"

        #Format waktu saat peserta scan
        item["waktu_scan"] = item["waktu_scan"] = item["waktu_scan"].isoformat()

    return jsonify({
        "data": data,
        "pagination": {
            "page": page,
            "total": total_data,
            "total_pages": ceil(total_data / limit)
        }
    })
    
#Untuk menghapus data lokasi dari tabel
@app.route("/lokasi-seminar/<int:id_lokasi>", methods=["DELETE"])
@login_required
@role_required("admin")
def delete_lokasi(id_lokasi):

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Melakukan pengecekan apakah lokasi telah dipakai untuk seminar atau tidak
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM seminar
        WHERE id_lokasi = %s
    """, (id_lokasi,))

    jumlah = cursor.fetchone()["total"]

    # Jika masih digunakan, batalkan penghapusan
    if jumlah > 0:
        cursor.close()
        conn.close()

        return jsonify({
            "message": "Lokasi masih digunakan oleh seminar.",
            "used": jumlah
        }), 400

    # Jika tidak digunakan, hapus lokasi
    cursor.execute("""
        DELETE FROM lokasi_seminar
        WHERE id_lokasi = %s
    """, (id_lokasi,))

    conn.commit()

    if cursor.rowcount == 0:
        cursor.close()
        conn.close()

        return jsonify({
            "message": "Lokasi tidak ditemukan"
        }), 404

    cursor.close()
    conn.close()

    return jsonify({
        "message": "Lokasi berhasil dihapus"
    })

#Untuk form edit data lokasi
@app.route("/lokasi-seminar/<int:id_lokasi>", methods=["PUT"])
@login_required
@role_required("admin")
def update_lokasi(id_lokasi):

    data = request.get_json()

    nama_lokasi = data.get("nama_lokasi")
    latitude = data.get("latitude")
    longitude = data.get("longitude")
    radius = data.get("radius")

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE lokasi_seminar
        SET
            nama_lokasi = %s,
            latitude = %s,
            longitude = %s,
            radius = %s
        WHERE id_lokasi = %s
    """, (
        nama_lokasi,
        latitude,
        longitude,
        radius,
        id_lokasi
    ))

    conn.commit()

    if cursor.rowcount == 0:
        cursor.close()
        conn.close()

        return jsonify({
            "message": "Lokasi tidak ditemukan"
        }), 404

    cursor.close()
    conn.close()

    return jsonify({
        "message": "Lokasi berhasil diperbarui"
    })

#Menampilkan data lokasi seminar, fitur search dan fitur sort di halaman Kelola Data Lokasi - Admin
@app.route("/lokasi-seminar", methods=["GET"])
@login_required
@role_required("admin")
def get_lokasi_seminar():

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Pagination
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 10))
    offset = (page - 1) * limit

    #Search
    search = request.args.get("search", "").strip()
    keyword = f"%{search}%"

    #Sort
    sort = request.args.get("sort", "asc").lower()

    if sort not in ["asc", "desc"]:
        sort = "asc"

    #Hitung total data
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM lokasi_seminar
        WHERE nama_lokasi LIKE %s
    """, (keyword,))

    total = cursor.fetchone()["total"]

    #Ambil data sesuai halaman
    cursor.execute(f"""
        SELECT
            id_lokasi,
            nama_lokasi,
            latitude,
            longitude,
            radius
        FROM lokasi_seminar
        WHERE nama_lokasi LIKE %s
        ORDER BY nama_lokasi {sort.upper()}
        LIMIT %s OFFSET %s
    """, (keyword, limit, offset))

    lokasi = cursor.fetchall()

    cursor.close()
    conn.close()

    return jsonify({
        "data": lokasi,
        "page": page,
        "limit": limit,
        "total": total,
        "total_page": ceil(total / limit)
    })

#Untuk form tambah lokasi
@app.route("/lokasi-seminar", methods=["POST"])
@login_required
@role_required("admin")
def add_lokasi():

    data = request.get_json()

    nama_lokasi = data.get("nama_lokasi")
    latitude = data.get("latitude")
    longitude = data.get("longitude")
    radius = data.get("radius")

    if not nama_lokasi or latitude is None or longitude is None:
        return jsonify({"message": "Data belum lengkap"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO lokasi_seminar
        (nama_lokasi, latitude, longitude, radius)
        VALUES (%s, %s, %s, %s)
    """, (
        nama_lokasi,
        latitude,
        longitude,
        radius
    ))

    conn.commit()

    cursor.close()
    conn.close()

    return jsonify({
        "message": "Lokasi berhasil ditambahkan"
    }), 201

#Untuk menghapus data seminar dari tabel
@app.route("/delete-seminar/<int:id_seminar>", methods=["DELETE"])
@login_required
@role_required("admin")
def delete_seminar(id_seminar):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        DELETE FROM seminar
        WHERE id_seminar = %s
    """, (id_seminar,))

    conn.commit()

    cursor.close()
    conn.close()

    return jsonify({
        "success": True,
        "message": "Data seminar berhasil dihapus"
    })

#Untuk fitur search mahasiswa di form add
@app.route("/search/mahasiswa", methods=["GET"])
@login_required
@role_required("admin")
def search_mahasiswa():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    search = request.args.get("search", "").strip()

    keyword = f"%{search}%"

    cursor.execute("""
        SELECT 
            m.id_user, 
            m.nama, 
            m.nim,
            CASE WHEN s.id_seminar IS NOT NULL THEN TRUE ELSE FALSE END AS memiliki_seminar
        FROM mahasiswa m
        LEFT JOIN seminar s ON m.id_user = s.id_mahasiswa
        WHERE m.nama LIKE %s OR m.nim LIKE %s
    """, (keyword, keyword))

    data = cursor.fetchall()

    cursor.close()
    conn.close()

    return jsonify(data)

#Menampilkan data lokasi untuk masuk ke filter
@app.route("/filter/lokasi", methods=["GET"])
@login_required
@role_required("admin")
def get_filter_lokasi():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT
            id_lokasi,
            nama_lokasi,
            latitude,
            longitude
        FROM lokasi_seminar
        ORDER BY nama_lokasi
    """)

    data = cursor.fetchall()

    cursor.close()
    conn.close()

    return jsonify(data)

#Untuk form edit data seminar
@app.route("/edit-seminar/<int:id_seminar>", methods=["PUT"])
@login_required
@role_required("admin")
def edit_seminar(id_seminar):
    data = request.json

    conn = get_db_connection();
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE seminar
        SET
            id_mahasiswa = %s,
            id_lokasi = %s,
            judul_penelitian = %s,
            tanggal = %s,
            waktu_mulai = %s,
            waktu_selesai = %s,
            dosen_pembimbing = %s,
            dosen_penguji_1 = %s,
            dosen_penguji_2 = %s
        WHERE id_seminar = %s
    """, (
        data["id_mahasiswa"],
        data["id_lokasi"],
        data["judul_penelitian"],
        data["tanggal"],
        data["waktu_mulai"],
        data["waktu_selesai"],
        data["dosen_pembimbing"],
        data["dosen_penguji_1"],
        data["dosen_penguji_2"],
        id_seminar
    ))

    conn.commit()

    cursor.close()
    conn.close()

    return jsonify({
        "success": True,
        "message": "Data seminar berhasil diperbarui"
    })

#Menampilkan data seminar, fitur search, fitur sort dan fitur filter di halaman Kelola Data Seminar - Admin
@app.route("/data-seminar", methods=["GET"])
@login_required
@role_required("admin")
def get_data_seminar():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Pagination
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 5))
    offset = (page - 1) * limit

    #Search
    search = request.args.get("search", "").strip()
    keyword = f"%{search}%"

    #Filter
    lokasi = request.args.get("lokasi", "Semua")
    tanggal_filter = request.args.get("tanggal", "Semua")
    tanggal_awal = request.args.get("tanggal_awal")
    tanggal_akhir = request.args.get("tanggal_akhir")

    #Sort
    sort_by = request.args.get("sort_by", "tanggal")
    sort_order = request.args.get("sort_order", "desc").upper()

    conditions =[
        "(m.nama LIKE %s OR m.nim LIKE %s)"
    ]

    params = [keyword, keyword]

    if lokasi != "Semua":
        conditions.append("l.nama_lokasi = %s")
        params.append(lokasi)

    sort_columns = {
        "nama": "m.nama",
        "judul": "s.judul_penelitian",
        "tanggal": "s.tanggal"
    }

    sort_column = sort_columns.get(sort_by, "s.tanggal")

    if sort_order not in ["ASC", "DESC"]:
        sort_order = "DESC"

    #Kondisi untuk filter tanggal
    if tanggal_filter == "Hari Ini":
        conditions.append("DATE(s.tanggal) = CURDATE()")
    elif tanggal_filter == "Minggu Ini":
        conditions.append("YEARWEEK(s.tanggal,1)=YEARWEEK(CURDATE(),1)")
    elif tanggal_filter == "Bulan Ini":
        conditions.append("""
            MONTH(s.tanggal)=MONTH(CURDATE())
            AND YEAR(s.tanggal)=YEAR(CURDATE())
        """)
    #Kondisi filter tanggal (rentang tanggal)
    elif tanggal_awal and tanggal_akhir:
        conditions.append("DATE(s.tanggal) BETWEEN %s AND %s")
        params.extend([tanggal_awal, tanggal_akhir])

    where = "WHERE " + " AND ".join(conditions) 

    #Hitung total data
    count_query = f"""
        SELECT COUNT(*) AS total
        FROM seminar s
        JOIN mahasiswa m
            ON s.id_mahasiswa = m.id_user
        LEFT JOIN lokasi_seminar l
            ON s.id_lokasi = l.id_lokasi
        {where}
    """

    cursor.execute(count_query, tuple(params))
    total = cursor.fetchone()["total"]

    #Ambil data sesuai halaman
    data_query = f"""
        SELECT
            s.id_seminar,
            s.id_lokasi,
            UPPER(s.judul_penelitian) AS judul_penelitian,
            s.tanggal,
            s.waktu_mulai,
            s.waktu_selesai,
            s.dosen_pembimbing,
            s.dosen_penguji_1,
            s.dosen_penguji_2,

            l.nama_lokasi,
            l.latitude,
            l.longitude,
            l.radius,
                   
            m.id_user,
            m.nama,
            m.nim,
            m.angkatan
        FROM seminar s
        JOIN mahasiswa m
            ON s.id_mahasiswa = m.id_user
        LEFT JOIN lokasi_seminar l
            ON s.id_lokasi = l.id_lokasi
        {where}
        
        ORDER BY {sort_column} {sort_order}, s.waktu_mulai ASC
        LIMIT %s OFFSET %s
    """

    data_params = params.copy()
    data_params.extend([limit, offset])

    cursor.execute(data_query, tuple(data_params))
    data = cursor.fetchall()

    for item in data:
        #Format tanggal
        item["tanggal"] = item["tanggal"].isoformat()

        #Format jam
        def format_time_iso(td):
            total_seconds = int(td.total_seconds())
            jam = total_seconds // 3600
            menit = (total_seconds % 3600) // 60
            detik = total_seconds % 60

            return f"{jam:02}:{menit:02}:{detik:02}"

        item["waktu_mulai_asli"] = format_time_iso(item["waktu_mulai"])
        item["waktu_selesai_asli"] = format_time_iso(item["waktu_selesai"])

        item["waktu_mulai"] = format_waktu(item["waktu_mulai"])
        item["waktu_selesai"] = format_waktu(item["waktu_selesai"])

    cursor.close()
    conn.close()

    return jsonify({
        "data": data,
        "page": page,
        "limit": limit,
        "total": total,
        "total_page": max(1, ceil(total / limit))
    })

#Untuk form tambah seminar
@app.route("/data-seminar", methods=["POST"])
@login_required
@role_required("admin")
def tambah_seminar():
    data = request.json
    
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO seminar(
            id_mahasiswa,
            id_user_admin,
            id_lokasi,
            judul_penelitian,
            tanggal,
            waktu_mulai,
            waktu_selesai,
            dosen_pembimbing,
            dosen_penguji_1,
            dosen_penguji_2
        )
        VALUES(
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s
        )
    """,(
        data["id_mahasiswa"],
        data["id_user_admin"],
        data["id_lokasi"],
        data["judul_penelitian"],
        data["tanggal"],
        data["waktu_mulai"],
        data["waktu_selesai"],
        data["dosen_pembimbing"],
        data["dosen_penguji_1"],
        data["dosen_penguji_2"]
    ))

    conn.commit()

    cursor.close()
    conn.close()

    return jsonify({"message": "Berhasil"})

@app.route("/server-time")
def server_time():
    return jsonify({
        "utc": datetime.now(timezone.utc).isoformat(),
        "local": datetime.now().isoformat()
    })

#Mengecek status QR Code apakah aktif atau tidak
@app.route("/qr-status/<int:id_seminar>", methods=["GET"])
@login_required
@role_required("mahasiswa")
def qr_status(id_seminar):

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT
            qr_code,
            status_qr, 
            expired_at
        FROM 
            qr_codes
        WHERE id_seminar = %s
    """, (id_seminar,))

    qr = cursor.fetchone()

    if not qr:
        cursor.close()
        conn.close()

        return jsonify({
            "success": False,
            "message": "QR belum dibuat"
        }), 404

    now = datetime.now(timezone.utc)

    expired_at_aware = qr["expired_at"].replace(tzinfo=timezone.utc) if qr["expired_at"] else None

    if qr["status_qr"] == "active":
        if expired_at_aware is None:
            qr["status_qr"] = "inactive"
        elif expired_at_aware <= now:
            cursor.execute("""
                UPDATE qr_codes
                SET
                    status_qr='inactive',
                    activated_at=NULL,
                    expired_at=NULL
                WHERE id_seminar=%s
            """, (id_seminar,))

            conn.commit()

            qr["status_qr"] = "inactive"
            expired_at_aware = None

        cursor.close()
        conn.close()

    print("DATABASE expired_at :", qr["expired_at"])
    print("DATABASE isoformat  :", qr["expired_at"].isoformat() if qr["expired_at"] else None)

    return jsonify({
        "success": True,
        "qr_code": qr["qr_code"],
        "status_qr": qr["status_qr"],
        "expired_at": expired_at_aware.isoformat() if expired_at_aware else None,
        "server_time": now.isoformat()
    })

#Menghubungkan data QR Code dengan data seminar
@app.route("/generate-qr", methods=["POST"])
@login_required
@role_required("mahasiswa")
def generate_qr():
    token = request.headers.get("Authorization")

    if not token:
        return jsonify({
            "success": False,
            "message": "Token tidak ditemukan"
        }), 401
    
    token = token.replace("Bearer ", "")

    try:
        payload = jwt.decode(
            token,
            SECRET_KEY,
            algorithms=["HS256"]
        )

        #Mengambil data seminar dari database
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)

        cursor.execute("""
            SELECT *
            FROM seminar
            WHERE id_mahasiswa = %s
        """, (payload["id_user"],))

        seminar = cursor.fetchone()

        if not seminar:
            cursor.close()
            conn.close()

            return jsonify({
                "success": False,
                "message": "Data seminar tidak ditemukan"
            }), 404

        #Membuat JWT QR
        qr_payload = {
            "id_user": payload["id_user"],
            "id_seminar": seminar["id_seminar"],
            "role": payload["role"],
            "exp": datetime.now(timezone.utc) + timedelta(minutes=10)
        }

        qr_token = jwt.encode(
            qr_payload,
            SECRET_KEY,
            algorithm="HS256"
        )

        # Cek apakah QR sudah ada
        cursor.execute("""
            SELECT *
            FROM qr_codes
            WHERE id_seminar = %s
        """, (seminar["id_seminar"],))

        existing_qr = cursor.fetchone()

        now = datetime.now(timezone.utc)

        # Kalau QR belum ada, masukkan data ke tabel qr_codes
        if existing_qr:
            # Kalau QR masih aktif
            ex_expired_aware = existing_qr["expired_at"].replace(tzinfo=timezone.utc) if existing_qr["expired_at"] else None
            
            if existing_qr["status_qr"] == "active" and ex_expired_aware and ex_expired_aware > now:
                cursor.close()
                conn.close()

                return jsonify({
                    "success": True,
                    "qr_code": existing_qr["qr_code"],
                    "status_qr": "active",
                    "expired_at": existing_qr["expired_at"].isoformat(),
                    "server_time": now.isoformat()
                })
            
            # QR sudah ada tapi inactive/expired
            cursor.execute("""
                UPDATE qr_codes
                SET
                    qr_code=%s,
                    status_qr='inactive',
                    generated_at=NOW(),
                    activated_at=NULL,
                    expired_at=NULL
                WHERE id_seminar=%s
            """, (
                qr_token,
                seminar["id_seminar"]
            ))
        else:
            # QR belum ada
            cursor.execute("""
                INSERT INTO qr_codes(
                    id_seminar,
                    qr_code,
                    status_qr,
                    generated_at
                )
                VALUES(%s,%s,'inactive', NOW())
            """, (
                seminar["id_seminar"],
                qr_token
            ))

        conn.commit()
        cursor.close()
        conn.close()

        return jsonify({
            "success": True,
            "qr_code": qr_token,
            "status_qr": "inactive",
            "expired_at": None,
            "server_time": now.isoformat()
        }), 200
    
    except jwt.ExpiredSignatureError:
        return jsonify({
            "success": False,
            "message": "Token expired"
        }), 401
    
    except jwt.InvalidTokenError:
        return jsonify({
            "success": False,
            "message": "Token tidak valid"
        }), 401

#Mengaktifkan QR Code
@app.route("/activate-qr", methods=["POST"])
@login_required
@role_required("mahasiswa")
def activate_qr():
    data = request.get_json()
    id_seminar = data.get("id_seminar")

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    #Memastikan QR ada
    cursor.execute("""
        SELECT *
        FROM qr_codes
        WHERE id_seminar = %s
    """, (id_seminar,))

    if not cursor.fetchone():
        cursor.close()
        conn.close()
        return jsonify({"success": False, "message": "QR Code belum dibuat"}), 404
    
    now = datetime.now(timezone.utc)
    expired = now + timedelta(minutes=10)

    cursor.execute("""
        UPDATE qr_codes
        SET
            status_qr = 'active',
            activated_at = %s,
            expired_at = %s
        WHERE id_seminar = %s           
    """, (now, expired, id_seminar))

    conn.commit()
    cursor.close()
    conn.close()

    return jsonify ({
        "success": True,
        "message": "QR Code berhasil diaktifkan",
        "status_qr": "active",
        "expired_at": expired.isoformat(),
        "server_time": now.isoformat()
    })

#Menonaktifkan QR Code ketika waktu 10 menit selesai
@app.route("/deactivate-qr", methods=["POST"])
@login_required
@role_required("mahasiswa")
def deactivate_qr():
    data = request.get_json()

    id_seminar = data.get("id_seminar")

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        UPDATE qr_codes
        SET
            status_qr = 'inactive'
        WHERE id_seminar = %s               
    """, (id_seminar,))

    conn.commit()

    cursor.close()
    conn.close()

    return jsonify({
        "success": True,
        "message": "QR Code telah dinonaktifkan"
    })

#Menghubungkan scanner ke backend
@app.route("/scan-qr", methods=["POST"])
@login_required
@role_required("mahasiswa")
def scan_qr():
    #Mengambil token login peserta seminar
    token = request.headers.get("Authorization")

    if not token:
        return jsonify({
            "success": False,
            "message": "Token tidak ditemukan"
        }), 401
    
    token = token.replace("Bearer ", "")

    #Mengambil data QR Code dari frontend
    data = request.get_json()

    qr_token = data.get("qr_code")

    print("QR TOKEN DARI FRONTEND:")
    print(qr_token)

    latitude = data.get("latitude")
    longitude = data.get("longitude")

    if not qr_token:
        print("QR CODE TIDAK DITEMUKAN")

        return jsonify({
            "success": False,
            "message": "QR Code tidak ditemukan"
        }), 400
    
    if latitude is None or longitude is None:
        print("LOKASI GPS TIDAK TERSEDIA DARI FRONTEND")
        return jsonify({
            "success": False,
            "code": "LOCATION_MISSING",
            "message": "Pastikan izin GPS/Lokasi di HP Anda sudah aktif."
        }), 400
    
    try:
        #Decode token login peserta seminar
        peserta = request.user

        #Mengecek agar hanya mahasiswa yang bisa melakukan presensi
        if peserta["role"] != "mahasiswa":
            return jsonify({
                "success": False,
                "code": "INVALID_ROLE",
                "message": "Hanya mahasiswa yang dapat melakukan presensi"
            }), 403

        print("QR YANG DITERIMA:", qr_token)

        #Decode QR Code
        qr_payload = jwt.decode(
            qr_token,
            SECRET_KEY,
            algorithms=["HS256"]
        )

        #Mengecek QR di database
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True, buffered=True)

        cursor.execute("""
            SELECT *
            FROM qr_codes
            WHERE id_seminar = %s
        """, (qr_payload["id_seminar"],))

        qr = cursor.fetchone()

        if not qr:
            cursor.close()
            conn.close()

            return jsonify({
                "success": False,
                "code": "QR_NOT_FOUND",
                "message": "QR Code tidak ditemukan"
            }), 404
        
        #Memastikan QR tersebut aktif
        if qr["status_qr"] != "active":
            cursor.close()
            conn.close()

            print("QR CODE BELUM DIAKTIFKAN")

            return jsonify({
                "success": False,
                "code": "QR_NOT_ACTIVE",
                "message": "QR Code belum diaktifkan"
            }), 400
        
        cursor.execute("""
            SELECT
                l.latitude,
                l.longitude,
                l.radius
            FROM seminar s
            JOIN lokasi_seminar l
                ON s.id_lokasi = l.id_lokasi
            WHERE s.id_seminar = %s
        """, (qr_payload["id_seminar"],))

        lokasi = cursor.fetchone()
        
        #Mengecek apakah penyelenggara mencoba scan qr seminar miliknya sendiri
        if peserta["id_user"] == qr_payload["id_user"]:
            cursor.close()
            conn.close()

            print("PENYELENGGARA TIDAK DAPAT MELAKUKAN PRESENSI")

            return jsonify({
                "success": False,
                "code": "PENYELENGGARA",
                "message": "Penyelenggara seminar tidak dapat melakukan presensi"
            }), 400
        
        now = datetime.now()

        #Mengecek apakah qr yang akan di scan sudah kedaluwarsa atau belum
        if qr["expired_at"] is not None and now > qr["expired_at"]:
            cursor.close()
            conn.close()

            print("QR CODE SUDAH KEDALUWARSA")

            return jsonify({
            "success": False,
            "code": "QR_EXPIRED",
            "message": "QR Code sudah kedaluwarsa"
        }), 400

        #Mengecek apakah peserta sudah pernah melakukan presensi sebelumnya
        cursor.execute("""
            SELECT *
            FROM presensi
            WHERE id_mahasiswa = %s
            AND id_seminar = %s       
        """, (
            peserta["id_user"],
            qr_payload["id_seminar"]
        ))

        existing = cursor.fetchone()
        if existing:
            cursor.close()
            conn.close()

            print("ANDA SUDAH MELAKUKAN PRESENSI")

            return jsonify({
                "success": False,
                "code": "ALREADY_ATTENDED",
                "message": "Anda sudah melakukan presensi"
            }), 400
        
        if lokasi is None or lokasi["latitude"] is None or lokasi["longitude"] is None:
            cursor.close()
            conn.close()
            
            return jsonify({
                "success": False,
                "code": "INVALID_SEMINAR_LOCATION",
                "message": "Data koordinat lokasi seminar di database belum diatur."
            }), 400
        
        jarak = hitung_jarak(
            float(latitude),
            float(longitude),
            float(lokasi["latitude"]),
            float(lokasi["longitude"])
        )

        if jarak > lokasi["radius"]:
            cursor.close()
            conn.close()

            print("ANDA BERADA DI LUAR AREA SEMINAR")

            return jsonify({
                "success": False,
                "code": "OUT_OF_RADIUS",
                "message": "Anda berada di luar area seminar",
                "distance": round(jarak, 2),
                "radius": lokasi["radius"]
            }), 400
        
        #Menyimpan data presensi
        cursor.execute("""
            INSERT INTO presensi(
                id_mahasiswa,
                id_seminar,
                waktu_scan,
                latitude,
                longitude
            )
            VALUES(%s,%s,%s,%s,%s)
        """,(
            peserta["id_user"],
            qr_payload["id_seminar"],
            now,
            latitude,
            longitude
        ))

        conn.commit()

        cursor.close()
        conn.close()

        return jsonify({
            "success": True,
            "code": "SCAN_SUCCESS",
            "message": "QR berhasil dibaca",
            "peserta_id": peserta["id_user"],
            "seminar_id": qr_payload["id_seminar"],
            "role": peserta["role"]
        })
    
    except jwt.ExpiredSignatureError:
        return jsonify({
            "success": False,
            "code": "QR_EXPIRED",
            "where": "scan_qr",
            "message": "QR Code sudah kedaluwarsa"
        }), 401
    
    except jwt.InvalidTokenError as e:
        print("QR ERROR:", e)

        return jsonify({
            "success": False,
            "code": "QR_INVALID",
            "where": "scan_qr",
            "message": "QR Code tidak valid"
        }), 401
    
    except Exception as e:
        print("SYSTEM ERROR ON SCAN_QR:", str(e)) # Cek terminal Flask Anda untuk detail error ini!
        return jsonify({
            "success": False,
            "code": "SERVER_ERROR",
            "message": f"Terjadi kesalahan internal server: {str(e)}"
        }), 500
    
#Menghubungkan data di halaman seminar saya (Penyelenggara)
@app.route("/detail-seminar/<int:id_user>")
@login_required
@role_required("mahasiswa")
def detail_seminar(id_user):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
        SELECT
            s.*,
            m.nama,
            m.nim,
            m.angkatan,
            l.nama_lokasi,
            l.latitude,
            l.longitude,
            l.radius,

            (
                SELECT COUNT(*)
                FROM presensi p
                WHERE p.id_seminar = s.id_seminar
            ) AS total_peserta,
            (
                SELECT COUNT(*)
                FROM presensi p
                WHERE p.id_seminar = s.id_seminar
                AND p.status_verifikasi <> 'pending'
            ) AS telah_diverifikasi,
            (
                SELECT COUNT(*)
                FROM presensi p
                WHERE p.id_seminar = s.id_seminar
                AND p.status_verifikasi = 'pending'
            ) AS pending
        
        FROM seminar s
        JOIN mahasiswa m
            ON s.id_mahasiswa = m.id_user
        LEFT JOIN lokasi_seminar l
            ON s.id_lokasi = l.id_lokasi
        WHERE s.id_mahasiswa = %s
    """

    cursor.execute(query, (id_user,))
    seminar = cursor.fetchone()

    cursor.close()
    conn.close()

    #Statistik presensi seminar
    if seminar:
        seminar["total_peserta"] = seminar["total_peserta"] or 0
        seminar["telah_diverifikasi"] = seminar["telah_diverifikasi"] or 0
        seminar["pending"] = seminar["pending"] or 0

        for key, value in seminar.items():
            if isinstance(value, timedelta):
                seminar[key] = str(value)

    return jsonify(seminar)

#Menghubungkan halaman login dengan BE dan mengecek apakah mahasiswa yang login memiliki jadwal seminar atau tidak, untuk menyesuaikan tampilan halaman seminar saya
@app.route("/login", methods=["POST"])
def login():
    data = request.get_json()

    username = data.get("username")
    password = data.get("password")

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT
            u.*,
            m.nim,
            m.nama,
            m.angkatan
        FROM users u
        LEFT JOIN mahasiswa m
            ON u.id_user = m.id_user
        WHERE u.username = %s
    """, (username,))

    user = cursor.fetchone()

    # Username salah
    if not user:
        cursor.close()
        conn.close()

        return jsonify({
            "success": False,
            "field": "username",
            "message": "Username salah"
        }), 401

    # Password salah
    if password != user["password"]:
        cursor.close()
        conn.close()

        return jsonify({
            "success": False,
            "field": "password",
            "message": "Password salah"
        }), 401

    # Default
    memiliki_seminar = False

    # Jika mahasiswa, cek apakah memiliki seminar
    if user["role"] == "mahasiswa":
        cursor.execute("""
            SELECT 1
            FROM seminar
            WHERE id_mahasiswa = %s
            LIMIT 1
        """, (user["id_user"],))

        memiliki_seminar = cursor.fetchone() is not None

    payload = {
        "id_user": user["id_user"],
        "username": user["username"],
        "role": user["role"],
        "memiliki_seminar": memiliki_seminar,
        "exp": datetime.utcnow() + timedelta(hours=3)
    }

    token = jwt.encode(
        payload,
        SECRET_KEY,
        algorithm="HS256"
    )

    cursor.close()
    conn.close()

    return jsonify({
        "success": True,
        "token": token,
        "user": {
            "id_user": user["id_user"],
            "id_mahasiswa": user["id_user"],
            "username": user["username"],
            "role": user["role"],
            "nim": user["nim"],
            "nama": user["nama"],
            "memiliki_seminar": memiliki_seminar
        }
    })

#Testing BE
@app.route("/")
def home():
    return "Backend Flask Berjalan!"

if __name__ == "__main__":
    app.run(debug=True)