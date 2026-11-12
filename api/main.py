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

BASE_DIR = Path(__file__).resolve().parent.parent
app = FastAPI(title="Baraka Express Executive Controller")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

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

# --- KONEKSI GOOGLE SHEETS ---
def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    # Mengambil JSON credential dari Environment Variable Vercel
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("Environment variable GOOGLE_CREDENTIALS_JSON belum diatur di Vercel!")
    
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

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
    s = str(raw_tgl).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None

def parse_date_matrix(raw_tgl) -> Optional[datetime]:
    if not raw_tgl:
        return None
    s = str(raw_tgl).strip()
    # Format dd/mm/yyyy
    dt = parse_date_slash(s)
    if dt:
        return dt
    # Format 1-Okt-26
    parts = s.split("-")
    if len(parts) == 3:
        months = {
            "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5,
            "jun": 6, "jul": 7, "agu": 8, "aug": 8, "sep": 9,
            "okt": 10, "oct": 10, "nov": 11, "des": 12, "dec": 12
        }
        m_str = parts[1].lower()[:3]
        m = months.get(m_str)
        if m:
            try:
                day = int(parts[0])
                yr_raw = int(parts[2])
                year = 2000 + yr_raw if yr_raw < 100 else yr_raw
                return datetime(year, m, day)
            except ValueError:
                return None
    return None

# --- AUTO SYNC BIAYA ADMIN BANK ---
def sync_admin_bank(ss):
    try:
        sheet_pusat = ss.worksheet("SETORAN_PUSAT")
        sheet_keluar = ss.worksheet("PENGELUARAN")
    except gspread.WorksheetNotFound:
        return

    rows_pusat = sheet_pusat.get_all_values()
    cabang_per_tgl = {}
    if len(rows_pusat) > 1:
        for r in rows_pusat[1:]:
            if len(r) >= 5:
                tgl_str = str(r[1]).strip()
                cb = str(r[2]).strip().upper()
                realisasi = parse_num(r[4])
                if realisasi > 0 and tgl_str:
                    cabang_per_tgl.setdefault(tgl_str, set()).add(cb)

    rows_keluar = sheet_keluar.get_all_values()
    baris_admin = {}
    if len(rows_keluar) > 1:
        for idx, r in enumerate(rows_keluar[1:], start=2):
            if len(r) >= 5:
                tgl_str = str(r[1]).strip()
                kat = str(r[4])
                if "Biaya Admin Bank" in kat:
                    baris_admin[tgl_str] = {"row": idx, "nominal": parse_num(r[3])}

    for tgl, set_cb in cabang_per_tgl.items():
        jml_cb = len(set_cb)
        total_admin = jml_cb * 2500
        ket_baru = f"Biaya Admin Transfer Setoran Pusat ({jml_cb} Cabang x Rp 2.500)"
        if tgl in baris_admin:
            if baris_admin[tgl]["nominal"] != total_admin:
                sheet_keluar.update_cell(baris_admin[tgl]["row"], 4, total_admin)
                sheet_keluar.update_cell(baris_admin[tgl]["row"], 6, ket_baru)
        else:
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
    ss = get_spreadsheet()
    sync_admin_bank(ss)

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

    estimasi_kewajiban = {cab: 0.0 for cab in DAFTAR_CABANG}
    riwayat_setoran = []
    riwayat_bank = []
    riwayat_keluar = []
    riwayat_pusat = []
    tgl_estimasi_terakhir = ""

    # 1. SETORAN CABANG
    try:
        rows_setoran = ss.worksheet("SETORAN CABANG").get_all_values()[1:]
        for r in rows_setoran:
            if len(r) < 4: continue
            d = parse_date_slash(r[1])
            nom = parse_num(r[3])
            ket = str(r[4]).upper() if len(r) > 4 else ""
            is_tf = "TRANSFER" in ket
            if d and nom > 0:
                if d < start_date:
                    if is_tf: saldo_awal_bank += nom
                    else: saldo_awal_cash += nom
                elif d.month == bulan and d.year == tahun:
                    if is_tf: tf_masuk += nom
                    else: cash_masuk += nom
                    total_setoran_cabang += nom
                    cb = str(r[2]).replace(" .", "").strip().upper()
                    riwayat_setoran.insert(0, {
                        "tanggal": d.strftime("%d/%m/%Y"),
                        "cabang": cb or "TANPA CABANG",
                        "nominal": nom,
                        "jenis": "TRANSFER" if is_tf else "CASH"
                    })
    except Exception:
        pass

    # 2. SETOR_TUNAI_BANK
    try:
        rows_bank = ss.worksheet("SETOR_TUNAI_BANK").get_all_values()[1:]
        for r in rows_bank:
            if len(r) < 3: continue
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
    except Exception:
        pass

    # 3. PENGELUARAN
    try:
        rows_keluar = ss.worksheet("PENGELUARAN").get_all_values()[1:]
        for r in rows_keluar:
            if len(r) < 4: continue
            d = parse_date_slash(r[1])
            sumber = str(r[2]).upper()
            nom = parse_num(r[3])
            if d and nom > 0:
                is_cash = "CASH" in sumber
                if d < start_date:
                    if is_cash: saldo_awal_cash -= nom
                    else: saldo_awal_bank -= nom
                elif d.month == bulan and d.year == tahun:
                    if is_cash: keluar_cash += nom
                    else: keluar_bank += nom
                    riwayat_keluar.insert(0, {
                        "tanggal": d.strftime("%d/%m/%Y"),
                        "sumber": sumber,
                        "nominal": nom,
                        "keterangan": r[5] if len(r) > 5 else "-"
                    })
    except Exception:
        pass

    # 4. SETORAN_PUSAT
    try:
        rows_pusat = ss.worksheet("SETORAN_PUSAT").get_all_values()[1:]
        for r in rows_pusat:
            if len(r) < 5: continue
            d = parse_date_slash(r[1])
            cb = str(r[2]).replace(" .", "").strip().upper()
            est = parse_num(r[3])
            real = parse_num(r[4])
            if d:
                if d < start_date:
                    if real > 0: saldo_awal_bank -= real
                elif d.month == bulan and d.year == tahun:
                    tgl_fmt = d.strftime("%d/%m/%Y")
                    if est > 0: tgl_estimasi_terakhir = tgl_fmt
                    if real > 0:
                        setor_pusat_tf += real
                        riwayat_pusat.insert(0, {
                            "tanggal": tgl_fmt,
                            "cabang": cb,
                            "estimasi": est,
                            "setoran": real
                        })
                    elif est > 0 and cb in estimasi_kewajiban:
                        estimasi_kewajiban[cb] += est
    except Exception:
        pass

    saldo_cash_akhir = saldo_awal_cash + cash_masuk - setor_bank - keluar_cash
    saldo_bank_akhir = saldo_awal_bank + tf_masuk + setor_bank - keluar_bank - setor_pusat_tf

    # --- MODUL OMSET MATRIKS (4 SUB-KOLOM) ---
    def baca_matriks_omset(nama_sheet):
        rekap_cb = {cab: {"tunai": 0.0, "tujuan": 0.0, "va": 0.0, "total": 0.0} for cab in DAFTAR_CABANG}
        total_all = 0.0
        riwayat = []
        try:
            ws = ss.worksheet(nama_sheet)
            rows = ws.get_all_values()
            if len(rows) >= 4:
                for r in rows[3:]:
                    if not r: continue
                    d = parse_date_matrix(r[0])
                    if d and d.month == bulan and d.year == tahun:
                        tgl_fmt = d.strftime("%d/%m/%Y")
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
                                    total_all += tot_val
                                    riwayat.append({
                                        "tanggal": tgl_fmt,
                                        "cabang": cab,
                                        "tunai": t_val,
                                        "tujuan": tuj_val,
                                        "va": va_val,
                                        "jumlah": tot_val
                                    })
        except Exception:
            pass
        return total_all, rekap_cb, riwayat

    tot_terima, rekap_terima, riw_terima = baca_matriks_omset("OMSET_TERIMA")
    tot_kirim, rekap_kirim, riw_kirim = baca_matriks_omset("OMSET_KIRIM")

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
        }
    })

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
        ws = ss.worksheet("SETORAN_PUSAT")
        tgl_fmt = datetime.strptime(data.tanggal, "%Y-%m-%d").strftime("%d/%m/%Y")
        no = len(ws.get_all_values())
        ws.append_row([no, tgl_fmt, data.cabang, data.estimasi, data.setoran, data.keterangan or "Setoran Transfer Pusat", datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
        return {"success": True, "message": f"Setoran pusat cabang {data.cabang} berhasil dicatat!"}
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