import os
import json
from datetime import datetime
from typing import Optional
import gspread
from google.oauth2.service_account import Credentials
from fastapi import FastAPI, Request
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from pathlib import Path
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles

BASE_DIR = Path(__file__).resolve().parent.parent
app = FastAPI(title="Baraka Express Executive Controller")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

SPREADSHEET_ID = "1RopoqGAtSgZl5UGURcBZyhorEfV1ZzhoZ2SGc6-2Sqo"

DAFTAR_CABANG = [
    "PURWOKERTO", "CILACAP", "PURBALINGGA",
    "BANJARNEGARA", "KEBUMEN", "BANYUMAS SELATAN"
]

BRANCH_MAP = {
    "PURWOKERTO":       {"tunai": 1,  "tujuan": 2,  "va": 3,  "total": 4},
    "KEBUMEN":          {"tunai": 5,  "tujuan": 6,  "va": 7,  "total": 8},
    "CILACAP":          {"tunai": 9,  "tujuan": 10, "va": 11, "total": 12},
    "PURBALINGGA":      {"tunai": 13, "tujuan": 14, "va": 15, "total": 16},
    "BANJARNEGARA":     {"tunai": 17, "tujuan": 18, "va": 19, "total": 20},
    "BANYUMAS SELATAN": {"tunai": 21, "tujuan": 22, "va": 23, "total": 24},
}

GUDANG_BRANCH_MAP = {
    "PURWOKERTO":       {"msk": 1,  "bbm": 2,  "lb": 3,  "tot_setor": 4,  "lokalan": 6,  "tag_msk": 7,  "tag_dist": 8,  "s_cash": 9,  "s_tf": 10},
    "KEBUMEN":          {"msk": 11, "bbm": 12, "lb": 13, "tot_setor": 14, "lokalan": 16, "tag_msk": 17, "tag_dist": 18, "s_cash": 19, "s_tf": 20},
    "CILACAP":          {"msk": 21, "bbm": 22, "lb": 23, "tot_setor": 24, "lokalan": 26, "tag_msk": 27, "tag_dist": 28, "s_cash": 29, "s_tf": 30},
    "PURBALINGGA":      {"msk": 31, "bbm": 32, "lb": 33, "tot_setor": 34, "lokalan": 36, "tag_msk": 37, "tag_dist": 38, "s_cash": 39, "s_tf": 40},
    "BANJARNEGARA":     {"msk": 41, "bbm": 42, "lb": 43, "tot_setor": 44, "lokalan": 46, "tag_msk": 47, "tag_dist": 48, "s_cash": 49, "s_tf": 50},
    "BANYUMAS SELATAN": {"msk": 51, "bbm": 52, "lb": 53, "tot_setor": 54, "lokalan": 56, "tag_msk": 57, "tag_dist": 58, "s_cash": 59, "s_tf": 60},
}

PENGELUARAN_GUDANG_MAP = {
    "PURWOKERTO":       {"tgl": 0,  "rincian": 1,  "nom": 2,  "ket": 3},
    "KEBUMEN":          {"tgl": 5,  "rincian": 6,  "nom": 7,  "ket": 8},
    "CILACAP":          {"tgl": 10, "rincian": 11, "nom": 12, "ket": 13},
    "PURBALINGGA":      {"tgl": 15, "rincian": 16, "nom": 17, "ket": 18},
    "BANJARNEGARA":     {"tgl": 20, "rincian": 21, "nom": 22, "ket": 23},
    "BANYUMAS SELATAN": {"tgl": 25, "rincian": 26, "nom": 27, "ket": 28},
}

FEE_OMSET_MAP = {
    "PURWOKERTO":       {"t_tunai": 1,  "t_tujuan": 2,  "t_va": 3,  "t_tot": 4,  "k_tunai": 5,  "k_tujuan": 6,  "k_va": 7,  "k_tot": 8,  "fee_terima": 9,  "fee_kirim": 10, "kewajiban": 11, "estimasi": 12, "setor_pusat": 13},
    "KEBUMEN":          {"t_tunai": 15, "t_tujuan": 16, "t_va": 17, "t_tot": 18, "k_tunai": 19, "k_tujuan": 20, "k_va": 21, "k_tot": 22, "fee_terima": 23, "fee_kirim": 24, "kewajiban": 25, "estimasi": 26, "setor_pusat": 27},
    "CILACAP":          {"t_tunai": 29, "t_tujuan": 30, "t_va": 31, "t_tot": 32, "k_tunai": 33, "k_tujuan": 34, "k_va": 35, "k_tot": 36, "fee_terima": 37, "fee_kirim": 38, "kewajiban": 39, "estimasi": 40, "setor_pusat": 41},
    "PURBALINGGA":      {"t_tunai": 43, "t_tujuan": 44, "t_va": 45, "t_tot": 46, "k_tunai": 47, "k_tujuan": 48, "k_va": 49, "k_tot": 50, "fee_terima": 51, "fee_kirim": 52, "kewajiban": 53, "estimasi": 54, "setor_pusat": 55},
    "BANJARNEGARA":     {"t_tunai": 57, "t_tujuan": 58, "t_va": 59, "t_tot": 60, "k_tunai": 61, "k_tujuan": 62, "k_va": 63, "k_tot": 64, "fee_terima": 65, "fee_kirim": 66, "kewajiban": 67, "estimasi": 68, "setor_pusat": 69},
    "BANYUMAS SELATAN": {"t_tunai": 71, "t_tujuan": 72, "t_va": 73, "t_tot": 74, "k_tunai": 75, "k_tujuan": 76, "k_va": 77, "k_tot": 78, "fee_terima": 79, "fee_kirim": 80, "kewajiban": 81, "estimasi": 82, "setor_pusat": 83},
}

# --- KONEKSI GOOGLE SHEETS ---
def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise ValueError("Environment variable GOOGLE_CREDENTIALS belum diatur di Vercel!")
    
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def fetch_all_sheets_batch(ss) -> dict:
    """Membaca seluruh sheet dalam 1x API call (values_batch_get)."""
    sheet_names = [
        "SETORAN CABANG",
        "SETORAN_CABANG",
        "SETOR_TUNAI_BANK",
        "PENGELUARAN",
        "OMSET_TERIMA",
        "OMSET_KIRIM",
        "PENGELUARAN_GUDANG",
        "KAS_Gudang",
        "CONTROL_GUDANG",
        "FEE_OMSET",
    ]
    existing_titles = {ws.title for ws in ss.worksheets()}
    valid_names = [name for name in sheet_names if name in existing_titles]
    if not valid_names:
        return {}

    ranges = [f"'{name}'!A:DZ" for name in valid_names]
    res = ss.values_batch_get(ranges)
    value_ranges = res.get("valueRanges", [])

    data_map = {}
    for name, vr in zip(valid_names, value_ranges):
        data_map[name] = vr.get("values", [])
    return data_map

# --- HELPER PARSE ANGKA & TANGGAL ---
def parse_num(val) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    clean = str(val).replace("Rp", "").replace(".", "").replace(",", "").strip()
    try:
        return float(clean)
    except ValueError:
        return 0.0

def parse_date_slash(raw_tgl) -> Optional[datetime]:
    if not raw_tgl:
        return None
    s = str(raw_tgl).strip().split(" ")[0]
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None

def parse_date_matrix(raw_tgl) -> Optional[datetime]:
    if not raw_tgl:
        return None
    s = str(raw_tgl).strip().split(" ")[0]
    dt = parse_date_slash(s)
    if dt:
        return dt

    months = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5,
        "jun": 6, "jul": 7, "agu": 8, "aug": 8, "sep": 9,
        "okt": 10, "oct": 10, "nov": 11, "nop": 11, "des": 12, "dec": 12
    }

    for sep in ("-", "/", " "):
        parts = [p.strip() for p in s.split(sep) if p.strip()]
        if len(parts) == 3:
            m_str = parts[1].lower()[:3]
            m = months.get(m_str)
            if m:
                try:
                    day = int(parts[0])
                    yr_raw = int(parts[2])
                    year = 2000 + yr_raw if yr_raw < 100 else yr_raw
                    return datetime(year, m, day)
                except ValueError:
                    pass
        elif len(parts) == 2:
            try:
                day = int(parts[0])
                m = int(parts[1]) if parts[1].isdigit() else months.get(parts[1].lower()[:3])
                if m and 1 <= m <= 12 and 1 <= day <= 31:
                    return datetime(datetime.now().year, m, day)
            except ValueError:
                pass
    return None

def parse_periode_bulan_tahun(raw_val) -> Optional[tuple]:
    if not raw_val:
        return None
    s = str(raw_val).strip()
    for sep in ("/", "-"):
        parts = s.split(sep)
        if len(parts) == 2:
            try:
                p1, p2 = int(parts[0]), int(parts[1])
                if 1 <= p1 <= 12 and p2 >= 2000:
                    return (p1, p2)
                if p1 >= 2000 and 1 <= p2 <= 12:
                    return (p2, p1)
            except ValueError:
                pass
    dt = parse_date_matrix(s)
    if dt:
        return (dt.month, dt.year)
    return None

# --- AUTO SYNC BIAYA ADMIN BANK DARI SHEET FEE_OMSET ---
def sync_admin_bank_from_fee(ss, rows_fee=None, rows_keluar=None):
    try:
        if rows_fee is None:
            rows_fee = ss.worksheet("FEE_OMSET").get_all_values()
        if rows_keluar is None:
            rows_keluar = ss.worksheet("PENGELUARAN").get_all_values()
    except gspread.WorksheetNotFound:
        return

    cabang_per_tgl = {}
    if len(rows_fee) >= 3:
        for r in rows_fee[2:]:
            if not r or not r[0]:
                continue
            d = parse_date_matrix(r[0])
            if not d:
                continue
            tgl_str = d.strftime("%d/%m/%Y")
            for cab, m in FEE_OMSET_MAP.items():
                if len(r) > m["setor_pusat"]:
                    realisasi = parse_num(r[m["setor_pusat"]])
                    if realisasi > 0:
                        cabang_per_tgl.setdefault(tgl_str, set()).add(cab)

    baris_admin = {}
    if len(rows_keluar) > 1:
        for idx, r in enumerate(rows_keluar[1:], start=2):
            if len(r) >= 5:
                tgl_str = str(r[1]).strip()
                kat = str(r[4])
                if "Biaya Admin Bank" in kat:
                    baris_admin[tgl_str] = {"row": idx, "nominal": parse_num(r[3])}

    sheet_keluar = None
    for tgl, set_cb in cabang_per_tgl.items():
        jml_cb = len(set_cb)
        total_admin = jml_cb * 2500
        ket_baru = f"Biaya Admin Transfer Setoran Pusat ({jml_cb} Cabang x Rp 2.500)"
        if tgl in baris_admin:
            if baris_admin[tgl]["nominal"] != total_admin:
                if sheet_keluar is None:
                    sheet_keluar = ss.worksheet("PENGELUARAN")
                sheet_keluar.update_cell(baris_admin[tgl]["row"], 4, total_admin)
                sheet_keluar.update_cell(baris_admin[tgl]["row"], 6, ket_baru)
        else:
            if sheet_keluar is None:
                sheet_keluar = ss.worksheet("PENGELUARAN")
            no = len(rows_keluar) if len(rows_keluar) > 1 else 1
            now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            sheet_keluar.append_row([no, tgl, "REKENING BANK", total_admin, "Biaya Admin Bank", ket_baru, now_str])

# --- PYDANTIC SCHEMAS UNTUK INPUT ---
class SetorBankInput(BaseModel):
    tanggal: str
    nominal: float
    bank: str
    keterangan: Optional[str] = ""

class SetorPusatInput(BaseModel):
    tanggal: str
    cabang: str
    estimasi: float
    setoran: float
    keterangan: Optional[str] = ""

class PengeluaranInput(BaseModel):
    tanggal: str
    sumberDana: str
    nominal: float
    kategori: str
    keterangan: str

# --- ROUTES ---
@app.get("/", response_class=HTMLResponse)
@app.get("/api/main", response_class=HTMLResponse)
async def serve_dashboard():
    return FileResponse(str(BASE_DIR / "templates" / "index.html"))

@app.get("/api/dashboard")
async def get_dashboard(bulan: int, tahun: int):
    try:
        ss = get_spreadsheet()
        sheets_data = fetch_all_sheets_batch(ss)

        rows_fee = sheets_data.get("FEE_OMSET", [])
        raw_keluar = sheets_data.get("PENGELUARAN", [])
        sync_admin_bank_from_fee(ss, rows_fee, raw_keluar)

        start_date = datetime(tahun, bulan, 1)

        # Variabel Kas Penampung
        saldo_awal_cash = 0.0
        saldo_awal_bank = 0.0
        cash_masuk = 0.0
        tf_masuk = 0.0
        setor_bank = 0.0
        keluar_cash = 0.0
        keluar_bank = 0.0
        setor_pusat_tf = 0.0
        total_setoran_cabang = 0.0

        riwayat_setoran = []
        riwayat_bank = []
        riwayat_keluar = []
        riwayat_pusat = []

        # --- 0. BACA SETORAN_CABANG (PENAMBAHAN SALDO PEMILIK / SALDO AWAL / NON-GUDANG) ---
        existing_setoran_keys = set()
        rows_setoran_manual = (
            sheets_data.get("SETORAN CABANG")
            or sheets_data.get("SETORAN_CABANG")
            or []
        )[1:]
        for r in rows_setoran_manual:
            if not r:
                continue
            # Deteksi apakah kolom A berisi Tanggal atau Nomor Urut
            d_col0 = parse_date_matrix(r[0]) if len(r) > 0 else None
            d_col1 = parse_date_matrix(r[1]) if len(r) > 1 else None

            if d_col0 and len(r) >= 4:
                d = d_col0
                cb = str(r[1]).replace(" .", "").strip().upper() or "PEMILIK"
                nom = parse_num(r[2])
                jenis = str(r[3]).strip().upper()
                ket = str(r[4]).strip() if len(r) > 4 else ""
            elif d_col1 and len(r) >= 5:
                d = d_col1
                cb = str(r[2]).replace(" .", "").strip().upper() or "PEMILIK"
                nom = parse_num(r[3])
                jenis = str(r[4]).strip().upper()
                ket = str(r[5]).strip() if len(r) > 5 else ""
            else:
                continue

            if d and nom > 0:
                is_cash = "CASH" in jenis or "TUNAI" in jenis
                jenis_clean = "CASH" if is_cash else "TRANSFER"
                tgl_fmt = d.strftime("%d/%m/%Y")
                existing_setoran_keys.add((tgl_fmt, cb, jenis_clean))

                if d < start_date:
                    if is_cash:
                        saldo_awal_cash += nom
                    else:
                        saldo_awal_bank += nom
                elif d.month == bulan and d.year == tahun:
                    if is_cash:
                        cash_masuk += nom
                    else:
                        tf_masuk += nom
                    total_setoran_cabang += nom
                    riwayat_setoran.insert(0, {
                        "tanggal": tgl_fmt,
                        "cabang": cb,
                        "nominal": nom,
                        "jenis": jenis_clean,
                        "keterangan": ket,
                        "is_manual": True
                    })

        # --- 1. MODUL KAS GUDANG (SEKALIGUS SUMBER SETORAN CABANG OTOMATIS KE PENAMPUNG) ---
        rekap_gudang_cab = {
            cab: {
                "msk": 0.0, "bbm": 0.0, "lb": 0.0, "tot_setor": 0.0,
                "lokalan": 0.0, "tag_msk": 0.0, "tag_dist": 0.0,
                "s_cash": 0.0, "s_tf": 0.0, "keluar_lain": 0.0,
                "sisa_kas": 0.0
            } for cab in DAFTAR_CABANG
        }
        riwayat_gudang_harian = []
        riwayat_keluar_gudang = []

        rows_kg = sheets_data.get("PENGELUARAN_GUDANG", [])
        if len(rows_kg) >= 3:
            for r in rows_kg[2:]:
                for cab, m in PENGELUARAN_GUDANG_MAP.items():
                    if len(r) > m["nom"]:
                        d = parse_date_matrix(r[m["tgl"]])
                        nom = parse_num(r[m["nom"]])
                        if d and d.month == bulan and d.year == tahun and nom > 0:
                            rekap_gudang_cab[cab]["keluar_lain"] += nom
                            riwayat_keluar_gudang.insert(0, {
                                "tanggal": d.strftime("%d/%m/%Y"),
                                "cabang": cab,
                                "nominal": nom,
                                "kategori": r[m["rincian"]] if len(r) > m["rincian"] and r[m["rincian"]] else "Operasional",
                                "keterangan": r[m["ket"]] if len(r) > m["ket"] and r[m["ket"]] else "-"
                            })

        rows_gdg = sheets_data.get("KAS_Gudang", [])
        if len(rows_gdg) >= 3:
            for r in rows_gdg[2:]:
                if not r or not r[0]:
                    continue
                if "TOTAL" in str(r[0]).upper() or "TAGIHAN" in str(r[0]).upper():
                    continue
                d = parse_date_matrix(r[0])
                if not d:
                    continue
                tgl_fmt = d.strftime("%d/%m/%Y")

                for cab, m in GUDANG_BRANCH_MAP.items():
                    if len(r) > m["s_tf"]:
                        v_msk  = parse_num(r[m["msk"]])
                        v_bbm  = parse_num(r[m["bbm"]])
                        v_lb   = parse_num(r[m["lb"]])
                        v_tot  = parse_num(r[m["tot_setor"]])
                        if v_tot == 0 and (v_msk > 0 or v_bbm > 0 or v_lb > 0):
                            v_tot = v_msk - v_bbm + v_lb
                        v_lok  = parse_num(r[m["lokalan"]])
                        v_tmsk = parse_num(r[m["tag_msk"]])
                        v_tdst = parse_num(r[m["tag_dist"]])
                        v_cash = parse_num(r[m["s_cash"]])
                        v_tf   = parse_num(r[m["s_tf"]])

                        if v_cash > 0 and (tgl_fmt, cab, "CASH") not in existing_setoran_keys:
                            if d < start_date:
                                saldo_awal_cash += v_cash
                            elif d.month == bulan and d.year == tahun:
                                cash_masuk += v_cash
                                total_setoran_cabang += v_cash
                                riwayat_setoran.insert(0, {
                                    "tanggal": tgl_fmt,
                                    "cabang": cab,
                                    "nominal": v_cash,
                                    "jenis": "CASH",
                                    "keterangan": "",
                                    "is_manual": False
                                })

                        tf_ke_penampung = 0.0 if cab == "PURWOKERTO" else v_tf
                        if tf_ke_penampung > 0 and (tgl_fmt, cab, "TRANSFER") not in existing_setoran_keys:
                            if d < start_date:
                                saldo_awal_bank += tf_ke_penampung
                            elif d.month == bulan and d.year == tahun:
                                tf_masuk += tf_ke_penampung
                                total_setoran_cabang += tf_ke_penampung
                                riwayat_setoran.insert(0, {
                                    "tanggal": tgl_fmt,
                                    "cabang": cab,
                                    "nominal": tf_ke_penampung,
                                    "jenis": "TRANSFER",
                                    "keterangan": "",
                                    "is_manual": False
                                })

                        if d.month == bulan and d.year == tahun:
                            if any([v_msk, v_bbm, v_lb, v_tot, v_lok, v_tmsk, v_tdst, v_cash, v_tf]):
                                rec = rekap_gudang_cab[cab]
                                rec["msk"] += v_msk
                                rec["bbm"] += v_bbm
                                rec["lb"] += v_lb
                                rec["tot_setor"] += v_tot
                                rec["lokalan"] += v_lok
                                rec["tag_msk"] += v_tmsk
                                rec["tag_dist"] += v_tdst
                                rec["s_cash"] += v_cash
                                rec["s_tf"] += v_tf

                                riwayat_gudang_harian.append({
                                    "tanggal": tgl_fmt,
                                    "cabang": cab,
                                    "msk": v_msk, "bbm": v_bbm, "lb": v_lb, "tot_setor": v_tot,
                                    "lokalan": v_lok, "tag_msk": v_tmsk, "tag_dist": v_tdst,
                                    "s_cash": v_cash, "s_tf": v_tf
                                })

        ctrl_gudang = {
            cab: {
                "saldo_awal": 0.0, "aktual_cash": 0.0,
                "kasbon_lama": 0.0, "kasbon_baru": 0.0, "cicilan_masuk": 0.0,
                "tf_cust_cabang": 0.0, "setor_pemilik": 0.0
            }
            for cab in DAFTAR_CABANG
        }
        rows_ctrl = sheets_data.get("CONTROL_GUDANG", [])[1:]
        for r in rows_ctrl:
            if len(r) < 3:
                continue
            per = parse_periode_bulan_tahun(r[0])
            if per and per == (bulan, tahun):
                cb = str(r[1]).replace(" .", "").strip().upper()
                if cb in ctrl_gudang:
                    ctrl_gudang[cb]["saldo_awal"] = parse_num(r[2]) if len(r) > 2 else 0.0
                    ctrl_gudang[cb]["aktual_cash"] = parse_num(r[3]) if len(r) > 3 else 0.0
                    ctrl_gudang[cb]["kasbon_lama"] = parse_num(r[4]) if len(r) > 4 else 0.0
                    ctrl_gudang[cb]["kasbon_baru"] = parse_num(r[5]) if len(r) > 5 else 0.0
                    ctrl_gudang[cb]["cicilan_masuk"] = parse_num(r[6]) if len(r) > 6 else 0.0
                    ctrl_gudang[cb]["tf_cust_cabang"] = parse_num(r[7]) if len(r) > 7 else 0.0
                    ctrl_gudang[cb]["setor_pemilik"] = parse_num(r[8]) if len(r) > 8 else 0.0

        for cab, rec in rekap_gudang_cab.items():
            c_data = ctrl_gudang[cab]
            rec["saldo_awal"] = c_data["saldo_awal"]
            rec["aktual_cash"] = c_data["aktual_cash"]
            rec["kasbon_lama"] = c_data["kasbon_lama"]
            rec["kasbon_baru"] = c_data["kasbon_baru"]
            rec["cicilan_masuk"] = c_data["cicilan_masuk"]
            rec["tf_cust_cabang"] = c_data["tf_cust_cabang"]
            rec["setor_pemilik"] = c_data["setor_pemilik"]
            rec["sisa_kasbon"] = c_data["kasbon_lama"] + c_data["kasbon_baru"] - c_data["cicilan_masuk"]

            tot_pemasukan = rec["msk"] + rec["lb"] + rec["lokalan"] + rec["tag_msk"] + rec["cicilan_masuk"]
            rec["total_pemasukan"] = tot_pemasukan

            if cab == "PURWOKERTO":
                tot_pengeluaran = rec["bbm"] + rec["keluar_lain"] + rec["s_cash"] + rec["kasbon_baru"]
            else:
                tot_pengeluaran = rec["bbm"] + rec["keluar_lain"] + rec["s_cash"] + rec["s_tf"] + rec["kasbon_baru"]

            rec["total_pengeluaran"] = tot_pengeluaran
            rec["sisa_kas"] = rec["saldo_awal"] + tot_pemasukan - tot_pengeluaran
            rec["selisih_aktual"] = rec["aktual_cash"] - rec["sisa_kas"] if rec["aktual_cash"] > 0 else 0.0

        # --- 2. SETOR_TUNAI_BANK & PENGELUARAN KAS PENAMPUNG ---
        rows_bank = sheets_data.get("SETOR_TUNAI_BANK", [])[1:]
        for r in rows_bank:
            if len(r) < 3:
                continue
            d = parse_date_slash(r[1])
            nom = parse_num(r[2])
            if d and nom > 0:
                if d < start_date:
                    saldo_awal_cash -= nom
                    saldo_awal_bank += nom
                elif d.month == bulan and d.year == tahun:
                    setor_bank += nom
                    riwayat_bank.insert(0, {
                        "tanggal": d.strftime("%d/%m/%Y"),
                        "nominal": nom,
                        "bank": r[3] if len(r) > 3 else "MANDIRI",
                        "keterangan": r[4] if len(r) > 4 else "-"
                    })

        for r in raw_keluar[1:]:
            if len(r) < 4:
                continue
            d = parse_date_slash(r[1])
            sumber = str(r[2]).upper()
            nom = parse_num(r[3])
            if d and nom > 0:
                is_cash = "CASH" in sumber
                if d < start_date:
                    if is_cash:
                        saldo_awal_cash -= nom
                    else:
                        saldo_awal_bank -= nom
                elif d.month == bulan and d.year == tahun:
                    if is_cash:
                        keluar_cash += nom
                    else:
                        keluar_bank += nom
                    riwayat_keluar.insert(0, {
                        "tanggal": d.strftime("%d/%m/%Y"),
                        "sumber": sumber,
                        "nominal": nom,
                        "keterangan": r[5] if len(r) > 5 else "-"
                    })

        # --- 3. MODUL OMSET MATRIKS (4 SUB-KOLOM, TERMASUK BARIS SISA BULAN SEBELUMNYA) ---
        def baca_matriks_omset(nama_sheet):
            rekap_cb = {
                cab: {
                    "tunai": 0.0, "tujuan": 0.0, "va": 0.0, "total": 0.0,
                    "sisa_lalu": 0.0, "bulan_ini": 0.0
                } for cab in DAFTAR_CABANG
            }
            total_all = 0.0
            riwayat = []
            rows = sheets_data.get(nama_sheet, [])
            if len(rows) >= 4:
                data_rows = rows[3:]
                parsed_dates = [
                    parse_date_matrix(r[0]) if (r and r[0]) else None
                    for r in data_rows
                ]

                for idx, r in enumerate(data_rows):
                    if not r or not r[0]:
                        continue
                    label_col0 = str(r[0]).strip()
                    upper_col0 = label_col0.upper()
                    if "TOTAL" in upper_col0:
                        continue

                    d = parsed_dates[idx]
                    is_sisa_row = False
                    target_bulan, target_tahun = None, None
                    tgl_fmt = ""

                    if d:
                        target_bulan, target_tahun = d.month, d.year
                        tgl_fmt = d.strftime("%d/%m/%Y")
                    elif "SISA" in upper_col0:
                        is_sisa_row = True
                        context_dt = None
                        for next_d in parsed_dates[idx + 1:]:
                            if next_d:
                                context_dt = next_d
                                break
                        if not context_dt:
                            for prev_d in reversed(parsed_dates[:idx]):
                                if prev_d:
                                    context_dt = prev_d
                                    break
                        if context_dt:
                            target_bulan, target_tahun = context_dt.month, context_dt.year
                        else:
                            target_bulan, target_tahun = bulan, tahun
                        tgl_fmt = f"00/00 ({upper_col0})"
                    else:
                        continue

                    if target_bulan == bulan and target_tahun == tahun:
                        for cab, m in BRANCH_MAP.items():
                            if len(r) > m["total"]:
                                t_val = parse_num(r[m["tunai"]])
                                tuj_val = parse_num(r[m["tujuan"]])
                                va_val = parse_num(r[m["va"]])
                                tot_val = parse_num(r[m["total"]])
                                if tot_val == 0:
                                    tot_val = t_val + tuj_val + va_val
                                if tot_val > 0:
                                    rekap_cb[cab]["tunai"] += t_val
                                    rekap_cb[cab]["tujuan"] += tuj_val
                                    rekap_cb[cab]["va"] += va_val
                                    rekap_cb[cab]["total"] += tot_val
                                    if is_sisa_row:
                                        rekap_cb[cab]["sisa_lalu"] += tot_val
                                    else:
                                        rekap_cb[cab]["bulan_ini"] += tot_val
                                    total_all += tot_val
                                    riwayat.append({
                                        "tanggal": tgl_fmt,
                                        "is_sisa": is_sisa_row,
                                        "label_sisa": upper_col0 if is_sisa_row else "",
                                        "cabang": cab,
                                        "tunai": t_val,
                                        "tujuan": tuj_val,
                                        "va": va_val,
                                        "jumlah": tot_val
                                    })
            return total_all, rekap_cb, riwayat

        tot_terima, rekap_terima, riw_terima = baca_matriks_omset("OMSET_TERIMA")
        tot_kirim, rekap_kirim, riw_kirim = baca_matriks_omset("OMSET_KIRIM")

        # --- 4. MODUL FEE TERIMA, FEE KIRIM, ESTIMASI & SETORAN PUSAT (DARI SHEET FEE_OMSET) ---
        rekap_fee_cab = {
            cab: {
                "omset_terkirim": 0.0,
                "omset_kirim_fee": 0.0,
                "fee_terima": 0.0,
                "fee_kirim": 0.0,
                "total_fee": 0.0,
                "kewajiban": 0.0,
                "estimasi": 0.0,
                "setor_pusat": 0.0,
                "selisih_setor_pusat": 0.0
            } for cab in DAFTAR_CABANG
        }
        riwayat_fee_harian = []
        estimasi_kewajiban = {cab: 0.0 for cab in DAFTAR_CABANG}
        rincian_estimasi_tanggal = []

        if len(rows_fee) >= 3:
            # Auto-deteksi indeks kolom ESTIMASI & SETOR PUSAT dari baris header (baris 1-2) tiap blok cabang
            dynamic_fee_map = {cab: dict(m) for cab, m in FEE_OMSET_MAP.items()}
            header_rows = rows_fee[:2]
            for cab, m in dynamic_fee_map.items():
                start_col = m["t_tunai"]
                end_col = min(start_col + 15, max((len(hr) for hr in header_rows), default=0))
                for c_idx in range(start_col, end_col):
                    h_text = " ".join(
                        str(hr[c_idx]).strip().upper()
                        for hr in header_rows if len(hr) > c_idx
                    )
                    if "ESTIMASI" in h_text or "EST" in h_text:
                        m["estimasi"] = c_idx
                    elif "KEWAJIBAN" in h_text:
                        m["kewajiban"] = c_idx
                    elif "SETOR" in h_text and "PUSAT" in h_text:
                        m["setor_pusat"] = c_idx

            for r in rows_fee[2:]:
                if not r or not r[0]:
                    continue
                if "TOTAL" in str(r[0]).upper():
                    continue
                d = parse_date_matrix(r[0])
                if not d:
                    continue

                if d < start_date:
                    for cab, m in dynamic_fee_map.items():
                        if len(r) > m["setor_pusat"]:
                            v_pusat_prev = parse_num(r[m["setor_pusat"]])
                            if v_pusat_prev > 0:
                                saldo_awal_bank -= v_pusat_prev
                elif d.month == bulan and d.year == tahun:
                    tgl_fmt = d.strftime("%d/%m/%Y")
                    for cab, m in dynamic_fee_map.items():
                        if len(r) > m["fee_kirim"]:
                            v_ttot = parse_num(r[m["t_tot"]])
                            if v_ttot == 0:
                                v_ttot = parse_num(r[m["t_tunai"]]) + parse_num(r[m["t_tujuan"]]) + parse_num(r[m["t_va"]])

                            v_ktot = parse_num(r[m["k_tot"]])
                            if v_ktot == 0:
                                v_ktot = parse_num(r[m["k_tunai"]]) + parse_num(r[m["k_tujuan"]]) + parse_num(r[m["k_va"]])

                            v_terima = parse_num(r[m["fee_terima"]])
                            v_kirim  = parse_num(r[m["fee_kirim"]])
                            v_tot_fee = v_terima + v_kirim
                            v_kew    = parse_num(r[m["kewajiban"]]) if len(r) > m["kewajiban"] else 0.0
                            v_est    = parse_num(r[m["estimasi"]]) if len(r) > m["estimasi"] else 0.0
                            v_pusat  = parse_num(r[m["setor_pusat"]]) if len(r) > m["setor_pusat"] else 0.0

                            if any([v_ttot, v_ktot, v_terima, v_kirim, v_kew, v_est, v_pusat]):
                                rec = rekap_fee_cab[cab]
                                rec["omset_terkirim"] += v_ttot
                                rec["omset_kirim_fee"] += v_ktot
                                rec["fee_terima"] += v_terima
                                rec["fee_kirim"] += v_kirim
                                rec["total_fee"] += v_tot_fee
                                rec["kewajiban"] += v_kew
                                rec["estimasi"] += v_est
                                rec["setor_pusat"] += v_pusat

                                if v_pusat > 0:
                                    setor_pusat_tf += v_pusat
                                    riwayat_pusat.insert(0, {
                                        "tanggal": tgl_fmt,
                                        "cabang": cab,
                                        "estimasi": v_est,
                                        "setoran": v_pusat
                                    })

                                # Hanya tampilkan rincian per tanggal jika estimasi > 0 dan setor pusat masih kosong
                                if v_est > 0 and v_pusat == 0:
                                    rincian_estimasi_tanggal.append({
                                        "tanggal": tgl_fmt,
                                        "cabang": cab,
                                        "nominal": v_est
                                    })

                                riwayat_fee_harian.append({
                                    "tanggal": tgl_fmt,
                                    "cabang": cab,
                                    "omset_terkirim": v_ttot,
                                    "omset_kirim": v_ktot,
                                    "fee_terima": v_terima,
                                    "fee_kirim": v_kirim,
                                    "total_fee": v_tot_fee,
                                    "kewajiban": v_kew,
                                    "estimasi": v_est,
                                    "setor_pusat": v_pusat
                                })

        tgl_estimasi_terakhir = ""
        for cab in DAFTAR_CABANG:
            f_rec = rekap_fee_cab[cab]
            target_setor = f_rec["estimasi"]
            f_rec["selisih_setor_pusat"] = f_rec["setor_pusat"] - target_setor
            estimasi_kewajiban[cab] = max(0.0, target_setor - f_rec["setor_pusat"])

        if riwayat_fee_harian:
            tgl_estimasi_terakhir = sorted([x["tanggal"] for x in riwayat_fee_harian])[-1]

        saldo_cash_akhir = saldo_awal_cash + cash_masuk - setor_bank - keluar_cash
        saldo_bank_akhir = saldo_awal_bank + tf_masuk + setor_bank - keluar_bank - setor_pusat_tf

        return JSONResponse({
            "kas": {
                "saldoAwal": {
                    "cash": saldo_awal_cash,
                    "rekening": saldo_awal_bank,
                    "total": saldo_awal_cash + saldo_awal_bank
                },
                "saldoAkhir": {
                    "cashGudang": saldo_cash_akhir,
                    "rekeningBank": saldo_bank_akhir,
                    "totalSaldo": saldo_cash_akhir + saldo_bank_akhir
                },
                "rekapMutasi": {
                    "totalSetoranCabang": total_setoran_cabang,
                    "totalSetoranPusat": setor_pusat_tf,
                    "totalPengeluaran": keluar_cash + keluar_bank
                },
                "estimasiKewajiban": estimasi_kewajiban,
                "tanggalEstimasiTerakhir": tgl_estimasi_terakhir,
                "rincianEstimasiTanggal": rincian_estimasi_tanggal,
                "riwayatSetoran": riwayat_setoran,
                "riwayatSetorBank": riwayat_bank,
                "riwayatPusat": riwayat_pusat,
                "riwayatPengeluaran": riwayat_keluar
            },
            "omset": {
                "totalTerima": tot_terima,
                "totalKirim": tot_kirim,
                "totalGabungan": tot_terima + tot_kirim,
                "terimaPerCabang": rekap_terima,
                "kirimPerCabang": rekap_kirim,
                "riwayatTerima": riw_terima,
                "riwayatKirim": riw_kirim
            },
            "gudang": {
                "rekapPerCabang": rekap_gudang_cab,
                "riwayatHarian": riwayat_gudang_harian,
                "riwayatPengeluaran": riwayat_keluar_gudang
            },
            "fee": {
                "rekapPerCabang": rekap_fee_cab,
                "riwayatHarian": riwayat_fee_harian
            },
        })
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/api/setor-bank")
async def post_setor_bank(data: SetorBankInput):
    try:
        ss = get_spreadsheet()
        ws = ss.worksheet("SETOR_TUNAI_BANK")
        tgl_fmt = datetime.strptime(data.tanggal, "%Y-%m-%d").strftime("%d/%m/%Y")
        no = len(ws.get_all_values())
        ws.append_row([no, tgl_fmt, data.nominal, data.bank, data.keterangan or "Setor Tunai ke Rekening", datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
        return {"success": True, "message": "Setor tunai bank berhasil dicatat!"}
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.post("/api/setor-pusat")
async def post_setor_pusat(data: SetorPusatInput):
    try:
        ss = get_spreadsheet()
        ws = ss.worksheet("FEE_OMSET")
        cabang = data.cabang.strip().upper()
        if cabang not in FEE_OMSET_MAP:
            return {"success": False, "message": "Cabang tidak valid!"}

        target_dt = datetime.strptime(data.tanggal, "%Y-%m-%d")
        all_rows = ws.get_all_values()
        target_row_idx = None

        for idx, r in enumerate(all_rows[2:], start=3):
            if r and r[0]:
                d = parse_date_matrix(r[0])
                if d and d.year == target_dt.year and d.month == target_dt.month and d.day == target_dt.day:
                    target_row_idx = idx
                    break

        col_pusat = FEE_OMSET_MAP[cabang]["setor_pusat"] + 1  # 1-indexed gspread
        if target_row_idx:
            existing_val = parse_num(all_rows[target_row_idx - 1][col_pusat - 1]) if len(all_rows[target_row_idx - 1]) >= col_pusat else 0.0
            ws.update_cell(target_row_idx, col_pusat, existing_val + data.setoran)
        else:
            next_row = len(all_rows) + 1
            tgl_str = target_dt.strftime("%d-%b-%y")
            ws.update_cell(next_row, 1, tgl_str)
            ws.update_cell(next_row, col_pusat, data.setoran)

        return {"success": True, "message": f"Setoran pusat cabang {cabang} berhasil dicatat ke FEE_OMSET!"}
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.post("/api/pengeluaran")
async def post_pengeluaran(data: PengeluaranInput):
    try:
        ss = get_spreadsheet()
        ws = ss.worksheet("PENGELUARAN")
        tgl_fmt = datetime.strptime(data.tanggal, "%Y-%m-%d").strftime("%d/%m/%Y")
        no = len(ws.get_all_values())
        ws.append_row([no, tgl_fmt, data.sumberDana, data.nominal, data.kategori, data.keterangan, datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
        return {"success": True, "message": "Pengeluaran berhasil dicatat!"}
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.post("/api/pengeluaran-gudang")
async def tambah_pengeluaran_gudang(request: Request):
    try:
        body = await request.json()
        ss = get_spreadsheet()
        ws = ss.worksheet("PENGELUARAN_GUDANG")
        cabang = str(body.get("cabang", "")).upper()
        if cabang not in PENGELUARAN_GUDANG_MAP:
            return JSONResponse({"success": False, "message": "Cabang tidak valid"}, status_code=400)

        col_idx = PENGELUARAN_GUDANG_MAP[cabang]["tgl"] + 1
        col_values = ws.col_values(col_idx)
        next_row = max(3, len(col_values) + 1)

        tgl_str = datetime.strptime(body.get("tanggal"), "%Y-%m-%d").strftime("%d/%m/%Y")
        rincian = str(body.get("kategori", "Operasional Gudang"))
        nominal = float(body.get("nominal", 0))
        ket = str(body.get("keterangan", ""))

        start_col_letter = gspread.utils.rowcol_to_a1(next_row, col_idx)
        end_col_letter = gspread.utils.rowcol_to_a1(next_row, col_idx + 3)
        ws.update(f"{start_col_letter}:{end_col_letter}", [[tgl_str, rincian, nominal, ket]], value_input_option="USER_ENTERED")

        return JSONResponse({"success": True, "message": f"Pengeluaran gudang {cabang} berhasil dicatat!"})
    except Exception as e:
        return JSONResponse({"success": False, "message": f"Gagal menyimpan: {str(e)}"}, status_code=500)