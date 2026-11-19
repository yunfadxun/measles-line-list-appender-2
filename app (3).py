"""Updater Database Linelist Provinsi (Campak/Rubella) - W41 sebagai database dasar.

Apa yang berubah dibanding versi sebelumnya
  * File yang template-nya "ditarik" sampai baris 1.048.576 (mis. NTB, Banten) sekarang
    dibaca per baris dan BERHENTI setelah 30 baris berturut-turut tanpa data inti
    (EPID, Kabupaten, Nama Kasus, Tgl Terima, Tgl Investigasi, Sumber Laporan semuanya kosong).
    Lebar kolom dibatasi (A:BH) sehingga cell nyasar di kolom jauh (AY, XFD) tidak ikut terbaca.
  * Kolom hasil mengikuti file "Data Line" terbaru (A:AV, 46 kolom bernama, nama & urutan sama). Kolom lain
    di file provinsi (mis. "Tgl Meninggal", "Investigasi Adekuat", "Shipment") tidak diambil. Nama kolom provinsi yang berbeda-beda
    ("Gejala Batuk" -> "Batuk", "Tanggal Mulai Rash (Tidak Boleh Kosong)" -> "Tanggal Mulai Rash",
    dst.) dicocokkan otomatis.
  * Provinsi dibaca dari blok judul sheet ("Provinsi : NUSA_TENGGARA_BARAT"), lalu kolom
    Provinsi, lalu nama file, lalu dipilih manual. Kolom "Provinsi (Otomatis)" yang kosong
    diisi dengan provinsi file tersebut.
  * Update mingguan: unggah data gabungan minggu lalu sebagai data dasar, lalu file baru.
    Semua baris lama milik provinsi yang mengirim file terbaru DIGANTI utuh; provinsi lain tetap.
  * Kolom "No" dinomori ulang 1..N saat unduh.

Langkah 3 - NAR (baru)
  * Unggah file tarikan aplikasi NAR (.xls bergaya HTML / .xlsx / .csv). Dari NAR yang punya ~70 kolom,
    HANYA kolom yang setara dengan 46 kolom Data Line yang diambil (nama NAR dipetakan ke nama Data Line).
  * Semua baris NAR tetap diolah. Tgl Mulai Rash kosong/tidak valid/tahun < RASH_MIN_YEAR hanya menjadi warning.
  * Baris NAR di-APPEND di sheet yang sama (di bawah data provinsi). Satu kolom tambahan "Sumber Data"
    (tepat setelah 46 kolom) berisi "NAR" untuk baris dari NAR dan "Excel Provinsi" untuk baris lainnya.
  * Data NAR tidak ikut "mengganti data per provinsi": NAR diganti utuh hanya bila file NAR baru diunggah.

Flag Tanggal Rash (baru)
  * Kolom "Flag Tgl Rash" (tepat setelah "Sumber Data") dihitung ulang setiap kali hasil disusun, untuk baris
    data dasar, file provinsi, maupun NAR. Tidak ada baris yang dibuang; baris hanya ditandai:
      - "Rash < 2026"       : tahun Tanggal Mulai Rash lebih kecil dari RASH_MIN_YEAR
      - "Rash kosong"       : Tanggal Mulai Rash kosong / "-" / placeholder
      - "Rash tidak valid"  : isinya teks yang tidak terbaca sebagai tanggal (mis. "YA", "#VALUE!", "23/042026")
    Baris yang aman dibiarkan kosong. Flag lama dari data dasar diabaikan dan dihitung ulang.
  * Ringkasan jumlah flag per provinsi ditampilkan di layar, dan ada unduhan khusus "baris ter-flag".
"""
import codecs
import gc
import importlib.util
import os
import re
import shutil
import tempfile
import uuid
from collections import Counter, deque
from datetime import date, datetime, timedelta
from functools import lru_cache
from html.parser import HTMLParser

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ----------------------------------------------------------------------------
# Pengaturan
# ----------------------------------------------------------------------------
SHEET_NAME = "List Kasus"
HEADER_TOP_ROW = 9        # baris judul kelompok (Laporan, Identitas Kasus, ...)
HEADER_SUB_ROW = 10       # baris nama kolom
FIRST_DATA_ROW = 12       # baris 11 = nomor kolom (1,2,3,...), data mulai baris 12
MAX_COL = 60              # baca maksimal sampai kolom BH (cell nyasar di luar ini diabaikan)
STOP_AFTER_BLANK = 30     # berhenti setelah sekian baris berturut-turut tanpa data inti
EXCEL_MAX_DATA_ROWS = 1_048_575
EXCEL_EPOCH = datetime(1899, 12, 30)
DATE_DISPLAY_FORMAT = "dd/mm/yyyy"
AMBIGUOUS_MAX_DISTANCE_DAYS = 366

INDONESIAN_MONTHS = {
    "januari": 1, "jan": 1,
    "februari": 2, "feb": 2,
    "maret": 3, "mar": 3,
    "april": 4, "apr": 4,
    "mei": 5,
    "juni": 6, "jun": 6,
    "juli": 7, "jul": 7,
    "agustus": 8, "agu": 8, "ags": 8, "agt": 8,
    "september": 9, "sep": 9, "sept": 9,
    "oktober": 10, "okt": 10,
    "november": 11, "nov": 11,
    "desember": 12, "des": 12,
}

RASH_MIN_YEAR = 2026      # batas flag Tanggal Mulai Rash (tahun < ini di-flag); NAR: juga untuk warning
SOURCE_COL = "Sumber Data"        # kolom tambahan (tepat setelah 46 kolom Data Line)
SOURCE_PROV = "Excel Provinsi"    # isi kolom untuk baris dari file Excel provinsi
SOURCE_NAR = "NAR"                # isi kolom untuk baris dari NAR

RASH_COL = "Tanggal Mulai Rash"
FLAG_COL = "Flag Tgl Rash"        # kolom tambahan (tepat setelah 'Sumber Data'); dihitung ulang tiap unduh
FLAG_LT = f"Rash < {RASH_MIN_YEAR}"
FLAG_BLANK = "Rash kosong"
FLAG_INVALID = "Rash tidak valid"
FLAG_TYPES = [FLAG_LT, FLAG_BLANK, FLAG_INVALID]

HAS_CALAMINE = importlib.util.find_spec("python_calamine") is not None
BASE_ENGINE = "calamine" if HAS_CALAMINE else "openpyxl"   # untuk membaca data dasar / NAR Excel

PLACEHOLDERS = ("nama kabupaten/kota tidak berdasarkan list / belum diisi", "data belum terisi")

# ----------------------------------------------------------------------------
# Kolom hasil = kolom pada file "Data Line" (nama & urutan sama)
# ----------------------------------------------------------------------------
# (nama di file Data Line). Tanda kurung otomatis memakai teks asli; baris baru -> spasi.
SRC_COLUMNS = [
    "No", "Nomor EPID", "Status KLB", "Nomor KLB",
    "Sumber Laporan", "Tanggal Terima", "Tanggal Investigasi",
    "Kabupaten (Mulai_mencari_dengan_mengetik_nama_Kabupaten/Kota)", "Provinsi (Otomatis)",
    "Nama Kasus", "Jenis Kelamin", "Tahun", "Bulan", "Alamat", "Desa", "Kecamatan",
    "Nama Orangtua / Wali",
    "Tanggal Mulai Demam", "Tanggal Mulai Rash", "Batuk", "Pilek", "Mata Merah",
    "Komplikasi (Sebutkan)", "Apakah kasus dirawat di RS?",
    "MCV1 (9 bln)", "MCV2 (18 bln)", "Campak BIAS kelas 1 SD",
    "Pernah terima imuniasasi MR saat kampanye", "Tanggal Vaksin MR terakhir",
    "Jumlah vaksin imunisasi campak diterima (Otomatis)",
    "Tanggal Ambil Spesimen Serum Darah", "Tanggal Kirim Spesimen Serum Darah ke Lab",
    "Tanggal Terima Spesimen Serum Darah di Lab", "Kondisi spesimen darah saat diterima lab",
    "Tanggal Terima Hasil Spesimen Serum Darah", "Hasil IgM Campak", "Hasil IgM Rubella",
    "Tanggal Ambil Spesimen Urin", "Tanggal Kirim Spesimen Urin", "Hasil Spesimen Urin",
    "Genotype",
    "VitA", "Bepergian dalam 1 bulan terakhir", "Kondisi Kasus",
    "Klasifikasi KLB (Isi jika suspek merupakan kasus KLB)", "Klasifikasi Akhir (Otomatis)",
]
FULL_LAYOUT = list(SRC_COLUMNS)  # susunan kolom hasil = file Data Line (A:AV)

PROV_COL = "Provinsi (Otomatis)"
NO_COL = "No"


def canon(value):
    """Kunci pencocokan nama kolom: baris pertama, huruf kecil, tanpa 'Gejala ' dan '_'."""
    if value is None:
        return ""
    text = str(value).split("\n")[0].replace("_", " ")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return re.sub(r"^gejala ", "", text)


def _first_line_key(out_header):
    """Kunci kolom dari nama keluaran (buang keterangan dalam kurung 'Otomatis' dst.)."""
    base = re.sub(r"\s*\((Otomatis|Isi jika.*|Mulai.*)\)\s*$", "", out_header)
    return canon(base)


KEY_TO_OUT = {_first_line_key(c): c for c in SRC_COLUMNS}
SRC_KEYS = list(KEY_TO_OUT)
# variasi penulisan yang mungkin dipakai provinsi lain
KEY_ALIASES = {
    "pernah terima imunisasi mr saat kampanye": "pernah terima imuniasasi mr saat kampanye",
    "tanggal mulai rash (tidak boleh kosong)": "tanggal mulai rash",
    "komplikasi": "komplikasi (sebutkan)",
    "dirawat di rs": "apakah kasus dirawat di rs?",
}
DATE_KEYS = {k for k in SRC_KEYS if k.startswith("tanggal")}
CORE_KEYS = ["nomor epid", "sumber laporan", "tanggal terima", "tanggal investigasi",
             "kabupaten", "nama kasus"]

# Pencarian nama kolom data dasar (keluaran aplikasi / versi singkat)
_BASE_LOOKUP = {}
for _out in FULL_LAYOUT:
    _BASE_LOOKUP[re.sub(r"\s+", " ", _out).strip().lower()] = _out
for _k, _out in KEY_TO_OUT.items():
    _BASE_LOOKUP.setdefault(_k, _out)

# ----------------------------------------------------------------------------
# Daftar provinsi (format seperti kolom provinsi di file sumber)
# ----------------------------------------------------------------------------
PROVINCE_ALIASES = {
    "ACEH": ["nanggroe aceh darussalam", "nad", "di aceh"],
    "SUMATERA_UTARA": ["sumut", "sumatra utara"],
    "SUMATERA_BARAT": ["sumbar", "sumatra barat"],
    "RIAU": [],
    "JAMBI": [],
    "SUMATERA_SELATAN": ["sumsel", "sumatra selatan"],
    "BENGKULU": [],
    "LAMPUNG": [],
    "JAKARTA": ["dki jakarta", "dki", "prov dki jakarta"],
    "JAWA_BARAT": ["jabar"],
    "JAWA_TENGAH": ["jateng"],
    "YOGYAKARTA": ["di yogyakarta", "diy", "jogja", "daerah istimewa yogyakarta"],
    "JAWA_TIMUR": ["jatim"],
    "KALIMANTAN_BARAT": ["kalbar"],
    "KALIMANTAN_TENGAH": ["kalteng"],
    "KALIMANTAN_SELATAN": ["kalsel"],
    "KALIMANTAN_TIMUR": ["kaltim"],
    "SULAWESI_UTARA": ["sulut"],
    "SULAWESI_TENGAH": ["sulteng"],
    "SULAWESI_SELATAN": ["sulsel"],
    "SULAWESI_TENGGARA": ["sultra"],
    "BALI": [],
    "NUSA_TENGGARA_BARAT": ["ntb"],
    "NUSA_TENGGARA_TIMUR": ["ntt"],
    "MALUKU": [],
    "PAPUA": [],
    "BANTEN": [],
    "MALUKU_UTARA": ["malut"],
    "GORONTALO": [],
    "BANGKA_BELITUNG": ["kepulauan bangka belitung", "kep bangka belitung", "babel"],
    "PAPUA_BARAT": ["pabar"],
    "KEPULAUAN_RIAU": ["kepri", "kep riau"],
    "SULAWESI_BARAT": ["sulbar"],
    "KALIMANTAN_UTARA": ["kaltara"],
    "PAPUA_TENGAH": [],
    "PAPUA_SELATAN": [],
    "PAPUA_PEGUNUNGAN": [],
    "PAPUA_BARAT_DAYA": ["pbd"],
}
PROVINCES = list(PROVINCE_ALIASES)


# ----------------------------------------------------------------------------
# Logika inti (tidak bergantung pada Streamlit, mudah diuji)
# ----------------------------------------------------------------------------
class SheetNotFound(Exception):
    pass


def _text(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def is_blank(v):
    t = _text(v).lower()
    return t == "" or t in PLACEHOLDERS


def _slug(text):
    return re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()


def _build_alias_list():
    pairs = set()
    for canon_name, aliases in PROVINCE_ALIASES.items():
        for name in [canon_name, *aliases]:
            s = _slug(name)
            pairs.add((s, canon_name))
            if "sumatera" in s:
                pairs.add((s.replace("sumatera", "sumatra"), canon_name))
    return sorted(pairs, key=lambda p: -len(p[0]))  # terpanjang dulu


ALIASES_SORTED = _build_alias_list()


@lru_cache(maxsize=4096)
def match_province(text, compact=False):
    """Cocokkan teks ke nama provinsi; None bila tidak ada. compact=True untuk nama file."""
    s = _slug(text)
    if not s:
        return None
    padded = f" {s} "
    for alias, name in ALIASES_SORTED:
        if f" {alias} " in padded:
            return name
    if compact:
        joined = s.replace(" ", "")
        for alias, name in ALIASES_SORTED:
            a = alias.replace(" ", "")
            if len(a) >= 5 and a in joined:
                return name
    return None


def find_sheet(names):
    for name in names:
        if re.sub(r"\s+", " ", str(name)).strip().lower() == SHEET_NAME.lower():
            return name
    raise SheetNotFound(f"Sheet '{SHEET_NAME}' tidak ditemukan (sheet yang ada: {', '.join(map(str, names))})")


def build_colmap(top_row, sub_row):
    """Peta kunci-kolom -> indeks kolom, dari dua baris judul (baris 10 diutamakan, lalu baris 9)."""
    colmap = {}
    for i in range(max(len(top_row), len(sub_row))):
        sub = sub_row[i] if i < len(sub_row) else None
        top = top_row[i] if i < len(top_row) else None
        key = canon(sub) or canon(top)
        key = KEY_ALIASES.get(key, key)
        if key in KEY_TO_OUT and key not in colmap:
            colmap[key] = i
    return colmap


def clean_value(v, is_date=False):
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return v or None
    if isinstance(v, float) and v != v:
        return None
    if is_date and isinstance(v, (int, float)) and not isinstance(v, bool) and 36526 <= v <= 73050:
        return EXCEL_EPOCH + timedelta(days=float(v))  # serial Excel -> tanggal
    return v


def _safe_date(year, month, day):
    try:
        return datetime(int(year), int(month), int(day))
    except (TypeError, ValueError, OverflowError):
        return None


def _unambiguous_reference_date(value):
    """Parse reference date without guessing an ambiguous numeric date."""
    parsed, method, ambiguous = parse_flexible_date(value, reference=None, default_dmy=False)
    return None if ambiguous else parsed if isinstance(parsed, datetime) else None


def parse_flexible_date(value, reference=None, default_dmy=True, prefer_before_reference=False):
    """Return ``(value, method, ambiguous)``.

    Ambiguous dates such as 01/02/2026 produce DMY and MDY candidates. If a
    reference date is available, choose the chronologically closest candidate.
    For onset-type columns, a candidate on/before investigation is preferred.
    If no usable reference exists, DMY is used and marked as ambiguous.
    """
    if value is None or value is pd.NaT or (isinstance(value, float) and value != value):
        return None, "blank", False
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime(), "excel-date", False
    if isinstance(value, datetime):
        return value.replace(tzinfo=None), "excel-date", False
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day), "excel-date", False
    if isinstance(value, bool):
        return value, "invalid", False
    if isinstance(value, (int, float)):
        number = float(value)
        if number.is_integer() and 1900 <= int(number) <= 2100:
            return value, "invalid-year-only", False
        if 1 <= number <= 2_958_465:
            try:
                return EXCEL_EPOCH + timedelta(days=number), "excel-serial", False
            except (OverflowError, ValueError):
                return value, "invalid", False
        return value, "invalid", False

    text = re.sub(r"\s+", " ", str(value).strip())
    if not text or text in {"-", "/", ","} or is_blank(text):
        return None, "blank", False
    if text.upper() in {"#VALUE!", "#N/A", "#REF!", "#NAME?", "#NUM!", "#NULL!", "N/A", "NA", "NAN", "NAT", "NONE", "NULL"}:
        return value, "invalid", False
    if re.fullmatch(r"\d{4}", text):
        return value, "invalid-year-only", False

    # Indonesian month names: 13 Agustus 2026 / 13 Agu 2026.
    m = re.fullmatch(r"(\d{1,2})[\s\-/]+([A-Za-z]+)[\s\-/]+(\d{2}|\d{4})", text, flags=re.I)
    if m:
        month = INDONESIAN_MONTHS.get(m.group(2).lower())
        year = int(m.group(3))
        if year < 100:
            year += 2000
        parsed = _safe_date(year, month, int(m.group(1))) if month else None
        return (parsed, "bulan-indonesia", False) if parsed else (value, "invalid", False)

    # ISO yyyy-mm-dd, optionally with time.
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$", text)
    if m:
        parsed = _safe_date(m.group(1), m.group(2), m.group(3))
        return (parsed, "iso", False) if parsed else (value, "invalid", False)

    # Numeric d/m/y or m/d/y, including two-digit years.
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})", text)
    if not m:
        return value, "invalid", False
    first, second, year = map(int, m.groups())
    if year < 100:
        year += 2000

    dmy = _safe_date(year, second, first)
    mdy = _safe_date(year, first, second)
    if dmy and not mdy:
        return dmy, "dmy-unambiguous", False
    if mdy and not dmy:
        return mdy, "mdy-unambiguous", False
    if not dmy and not mdy:
        return value, "invalid", False
    if dmy == mdy:
        return dmy, "same-day-month", False

    # Both candidates are valid: choose using reference where possible.
    if isinstance(reference, datetime):
        candidates = [(dmy, "ambiguous-reference-dmy"), (mdy, "ambiguous-reference-mdy")]
        if prefer_before_reference:
            before = [item for item in candidates if item[0] <= reference]
            if before:
                candidates = before
        candidates.sort(key=lambda item: abs((item[0] - reference).days))
        best, method = candidates[0]
        if abs((best - reference).days) <= AMBIGUOUS_MAX_DISTANCE_DAYS:
            return best, method, True

    if default_dmy:
        return dmy, "ambiguous-default-dmy", True
    return value, "ambiguous-unresolved", True


def normalize_date_columns(df):
    """Normalize date columns in-place, using investigation as row context."""
    date_columns = [c for c in df.columns if canon(c).startswith("tanggal")]
    investigation = "Tanggal Investigasi" if "Tanggal Investigasi" in df.columns else None

    # First normalize investigation. Ambiguous values default to DMY.
    if investigation:
        df[investigation] = [parse_flexible_date(v, default_dmy=True)[0] for v in df[investigation]]

    onset_columns = {
        "Tanggal Mulai Demam", "Tanggal Mulai Rash", "Tanggal Terima"
    }
    for column in date_columns:
        if column == investigation:
            continue
        normalized = []
        for row_index, value in enumerate(df[column]):
            reference = None
            if investigation:
                ref_value = df.iloc[row_index][investigation]
                if isinstance(ref_value, datetime):
                    reference = ref_value
            parsed, _, _ = parse_flexible_date(
                value,
                reference=reference,
                default_dmy=True,
                prefer_before_reference=column in onset_columns,
            )
            normalized.append(parsed)
        df[column] = normalized
    return df


def header_province(head_rows):
    """Cari 'Provinsi : NAMA' di blok judul sheet (baris 1-8)."""
    for row in head_rows:
        for i, v in enumerate(row):
            if canon(v) == "provinsi":
                for w in row[i + 1:i + 6]:
                    t = _text(w)
                    if t and t != ":":
                        p = match_province(t)
                        if p:
                            return p
    return None


def read_linelist(file):
    """Baca sheet 'List Kasus' satu file provinsi.

    Return (DataFrame kolom Data Line, provinsi_dari_judul, kolom_tidak_ditemukan, baris_dibuang).
    Membaca streaming dan berhenti saat template kosong dimulai, jadi aman untuk file yang
    barisnya ditarik sampai 1.048.576.
    """
    if hasattr(file, "seek"):
        file.seek(0)
    wb = load_workbook(file, read_only=True, data_only=True, keep_vba=False)
    try:
        ws = wb[find_sheet(wb.sheetnames)]
        head_rows, top, colmap, core_idx, records = [], None, {}, [], []
        blank_run, dropped = 0, 0
        for r, row in enumerate(ws.iter_rows(min_row=1, max_col=MAX_COL, values_only=True), start=1):
            if r < HEADER_TOP_ROW:
                head_rows.append(row)
            elif r == HEADER_TOP_ROW:
                top = row
            elif r == HEADER_SUB_ROW:
                colmap = build_colmap(top or (), row)
                if "nomor epid" not in colmap:
                    raise ValueError("Header tidak dikenali: kolom 'Nomor EPID' tidak ditemukan "
                                     f"di baris {HEADER_TOP_ROW}-{HEADER_SUB_ROW}")
                core_idx = [colmap[k] for k in CORE_KEYS if k in colmap]
            elif r >= FIRST_DATA_ROW:
                width = len(row)
                if all(is_blank(row[i] if i < width else None) for i in core_idx):
                    blank_run += 1
                    dropped += 1
                    if blank_run >= STOP_AFTER_BLANK:
                        break
                    continue
                blank_run = 0
                records.append([
                    clean_value(row[colmap[k]] if (k in colmap and colmap[k] < width) else None, k in DATE_KEYS)
                    for k in SRC_KEYS
                ])
        if not colmap:
            raise ValueError("Sheet terlalu pendek: baris header tidak ditemukan")
    finally:
        wb.close()

    df = pd.DataFrame(records, columns=[KEY_TO_OUT[k] for k in SRC_KEYS], dtype=object)
    normalize_date_columns(df)
    missing = [KEY_TO_OUT[k] for k in SRC_KEYS if k not in colmap and k != NO_COL.lower()]
    return df, header_province(head_rows), missing, dropped


def detect_province(df, filename, header_prov):
    """Tentukan provinsi satu file. Return (provinsi | None, sumber, catatan)."""
    file_prov = match_province(filename, compact=True) if filename else None
    if header_prov:
        note = "dari judul sheet"
        if file_prov and file_prov != header_prov:
            note += f" (PERIKSA: nama file menunjuk {file_prov})"
        return header_prov, "judul sheet", note
    votes = Counter()
    if PROV_COL in df.columns:
        for value, n in df[PROV_COL].dropna().astype(str).str.strip().value_counts().items():
            p = match_province(value)
            if p:
                votes[p] += n
    if votes:
        return votes.most_common(1)[0][0], "kolom", f"dari kolom '{PROV_COL}'"
    if file_prov:
        return file_prov, "nama file", "dari nama file"
    return None, None, "provinsi tidak terdeteksi - pilih manual"


def fill_province(df, province):
    """Isi kolom Provinsi (Otomatis) yang kosong / placeholder / error Excel (#ERROR!, #N/A) dengan provinsi file."""
    if PROV_COL in df.columns and province:
        df[PROV_COL] = [
            province if (is_blank(v) or _text(v).startswith("#") or not match_province(_text(v))) else v
            for v in df[PROV_COL]
        ]
    return df


# ---- flag Tanggal Mulai Rash -------------------------------------------------
def rash_flag(v):
    """Flag satu nilai Tanggal Mulai Rash. Return '' bila aman, selain itu salah satu FLAG_TYPES.

    - kosong / '-' / placeholder           -> FLAG_BLANK
    - teks yang bukan tanggal              -> FLAG_INVALID
    - tanggal (datetime/date/teks/serial)  -> FLAG_LT bila tahun < RASH_MIN_YEAR
    """
    if v is None or v is pd.NaT or (isinstance(v, float) and v != v):
        return FLAG_BLANK
    if isinstance(v, str):
        t = v.strip()
        if is_blank(t) or t == "-":
            return FLAG_BLANK
        d = _parse_date_text(t)
        if not isinstance(d, datetime):
            return FLAG_INVALID
        year = d.year
    elif isinstance(v, date):  # datetime juga subclass date (termasuk pd.Timestamp)
        year = v.year
    elif isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            year = (EXCEL_EPOCH + timedelta(days=float(v))).year   # serial Excel
        except (OverflowError, ValueError):
            return FLAG_INVALID
    else:
        return FLAG_INVALID
    return FLAG_LT if year < RASH_MIN_YEAR else ""


def flag_series(df):
    """Daftar flag per baris DataFrame (urutan sama dengan df)."""
    if RASH_COL not in df.columns:
        return [FLAG_BLANK] * len(df)
    return [rash_flag(v) for v in df[RASH_COL]]


def count_flags(df):
    """Counter jenis flag -> jumlah baris (baris aman tidak dihitung)."""
    return Counter(f for f in flag_series(df) if f)


def _save_record(df, workdir, name, size, kind, province, via, note, dropped=0, missing=None):
    rec_id = uuid.uuid4().hex[:8]
    path = os.path.join(workdir, f"{rec_id}.pkl")
    df.to_pickle(path)
    return {
        "id": rec_id, "name": name, "size": size, "kind": kind, "path": path,
        "rows": len(df), "columns": list(df.columns), "province": province, "via": via,
        "note": note, "dropped": dropped, "missing": missing or [],
        "flags": count_flags(df),
        "status": "ok" if province else "pending", "detail": "",
    }


def process_new_file(file, workdir):
    """Proses 1 file provinsi -> 1 record."""
    df, head_prov, missing, dropped = read_linelist(file)
    prov, via, note = detect_province(df, file.name, head_prov)
    df = fill_province(df, prov)
    rec = _save_record(df, workdir, file.name, file.size, "Baru", prov, via, note, dropped, missing)
    del df
    gc.collect()
    return rec


# ---- data dasar (hasil gabungan minggu lalu) --------------------------------
def coerce_dates(df):
    """Normalize all date columns in the base database before splitting."""
    return normalize_date_columns(df)


def read_base(file):
    """Baca database dasar W41 (CSV / Excel) dan samakan nama kolomnya.

    Untuk Excel, hanya sheet bernama ``Combined_Linelist`` atau
    ``Combined_Linelist_2``, dst. yang dibaca. Ini mencegah sheet dashboard,
    referensi, atau sheet lain ikut ter-concat secara tidak sengaja.
    """
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file, dtype=str, encoding="utf-8-sig")
    else:
        sheets = pd.read_excel(file, sheet_name=None, dtype=object, engine=BASE_ENGINE)
        combined = [
            frame for name, frame in sheets.items()
            if re.fullmatch(r"Combined_Linelist(?:_\d+)?", str(name).strip(), flags=re.I)
        ]
        if not combined:
            # Fallback for a simple one-sheet database exported manually.
            if len(sheets) == 1:
                combined = list(sheets.values())
            else:
                raise ValueError(
                    "Sheet database tidak ditemukan. Gunakan sheet 'Combined_Linelist' "
                    "(atau Combined_Linelist_2, dst.)."
                )
        df = pd.concat(combined, ignore_index=True)
    rename, seen = {}, set()
    for c in df.columns:
        target = _BASE_LOOKUP.get(re.sub(r"\s+", " ", str(c)).strip().lower())
        if target and target not in seen:
            rename[c], _ = target, seen.add(target)
    df = df.rename(columns=rename)
    if "Nomor EPID" not in df.columns or PROV_COL not in df.columns:
        raise ValueError(
            "Database dasar tidak dikenali: kolom 'Nomor EPID' dan/atau "
            "'Provinsi (Otomatis)' tidak ada. Gunakan W41 Untouched Data "
            "atau file hasil unduhan aplikasi ini."
        )
    df = df.drop(columns=[FLAG_COL], errors="ignore")   # flag lama dibuang; dihitung ulang
    df = df.astype(object).where(df.notna(), None)
    return coerce_dates(df)

def split_base(df):
    """Pecah data dasar per provinsi (berdasarkan kolom Provinsi (Otomatis)).

    Return list (provinsi | None, catatan, label, sub_df).
    """
    key = df[PROV_COL].map(lambda v: "" if is_blank(v) else _text(v))
    groups = []
    for value, sub in df.groupby(key, sort=False):
        prov = match_province(value) if value else None
        if prov:
            groups.append((prov, "dari data dasar", prov, sub))
        else:
            groups.append((None, "provinsi kosong/tidak dikenali - pilih manual", value or "(kosong)", sub))
    # gabungkan grup yang menunjuk provinsi sama (mis. 'Aceh' dan 'ACEH')
    merged = {}
    for prov, note, label, sub in groups:
        k = prov or f"?{label}"
        merged[k] = (prov, note, label, pd.concat([merged[k][3], sub]) if k in merged else sub)
    return list(merged.values())


def load_base(file, workdir):
    """Muat data dasar. Return (records_provinsi, DataFrame_NAR | None).

    Baris dengan kolom 'Sumber Data' = NAR (dari hasil gabungan minggu lalu) dipisahkan: tidak ikut
    mekanisme 'ganti per provinsi', tetap berstatus NAR sampai file NAR baru diunggah.
    """
    df = read_base(file)
    nar_df = None
    if SOURCE_COL in df.columns:
        is_nar = df[SOURCE_COL].map(lambda v: _text(v).upper() == SOURCE_NAR)
        if is_nar.any():
            nar_df = df[is_nar].drop(columns=[SOURCE_COL]).reindex(columns=FULL_LAYOUT).reset_index(drop=True)
        df = df[~is_nar].drop(columns=[SOURCE_COL])
    recs = []
    for prov, note, label, sub in split_base(df):
        sub = sub.reset_index(drop=True)
        if prov:
            sub[PROV_COL] = prov  # samakan penulisan
        recs.append(_save_record(sub, workdir, f"{file.name} :: {label}", file.size,
                                 "Data dasar", prov, "data dasar", note))
    del df
    gc.collect()
    return recs, nar_df


# ---- aktivasi & penggantian per provinsi ------------------------------------
def activate(rec, records):
    """Aktifkan record; buang semua data lama (aktif) milik provinsi yang sama."""
    rec["status"] = "ok"
    replaced = 0
    for other in records:
        if other is not rec and other["status"] == "ok" and other["province"] == rec["province"]:
            other["status"] = "replaced"
            other["detail"] = f"Digantikan oleh {rec['name']}"
            replaced += other["rows"]
    tail = f"menggantikan {replaced:,} baris lama" if replaced else "provinsi baru, tidak ada yang diganti"
    rec["detail"] = (f"Berhasil - {rec['province']} ({rec['note']}); {tail}; "
                     f"data terbaru: {rec['rows']:,} baris")
    n_flag = sum(rec["flags"].values())
    if n_flag:
        rec["detail"] += f"; {n_flag:,} baris ter-flag Tgl Rash"


def active_records(records):
    return sorted((r for r in records if r["status"] == "ok"), key=lambda r: PROVINCES.index(r["province"]))


# ============================================================================
# LANGKAH 3 - NAR
# ============================================================================
# Kolom Data Line -> nama kolom NAR (huruf kecil). Kolom NAR lain tidak diambil.
KAB_COL = "Kabupaten (Mulai_mencari_dengan_mengetik_nama_Kabupaten/Kota)"
NAR_DIRECT = {
    "Nomor EPID": "nomor epid",
    "Status KLB": "status klb",
    "Nomor KLB": "nomor klb",
    "Sumber Laporan": "sumber laporan kasus",
    "Tanggal Terima": "tanggal terima",
    "Tanggal Investigasi": "tanggal pelacakan",
    KAB_COL: "kabupaten/kota",
    PROV_COL: "provinsi",
    "Nama Kasus": "nama",
    "Jenis Kelamin": "jenis kelamin",
    "Tahun": "tahun",
    "Bulan": "bulan",
    "Alamat": "alamat",
    "Desa": "desa",
    "Kecamatan": "kecamatan",
    "Tanggal Mulai Demam": "tgl mulai demam",
    "Tanggal Mulai Rash": "tgl mulai rash",
    "Batuk": "batuk",
    "Pilek": "pilek",
    "Mata Merah": "mata merah",
    "Komplikasi (Sebutkan)": "komplikasi",
    "Apakah kasus dirawat di RS?": "dirawat di rs",
    "MCV1 (9 bln)": "imun. mr1",
    "MCV2 (18 bln)": "imun. mr2",
    "Campak BIAS kelas 1 SD": "imun. mr bias",
    "Pernah terima imuniasasi MR saat kampanye": "imun. imunisasi tambahan (ori, bian, crash program)",
    "VitA": "pemberian vitamin a",
    "Bepergian dalam 1 bulan terakhir": "riwayat perjalanan",
    "Kondisi Kasus": "kondisi saat ini",
}
# "Tanggal Vaksin MR terakhir" = tanggal paling baru dari keempat kolom vaksin NAR ini
NAR_VAX_DATES = ["tgl terakhir vaksin mr1", "tgl terakhir vaksin mr2", "tgl terakhir vaksin mr bias",
                 "tgl terakhir vaksin imunisasi tambahan (ori, bian, crash program)"]
VAX_COL = "Tanggal Vaksin MR terakhir"
# Spesimen 1 / 2 NAR -> blok serum / urin Data Line (berdasarkan kolom 'Jenis' spesimen)
NAR_SERUM = {
    "Tanggal Ambil Spesimen Serum Darah": "tgl ambil",
    "Tanggal Kirim Spesimen Serum Darah ke Lab": "tgl kirim",
    "Tanggal Terima Spesimen Serum Darah di Lab": "tgl terima (lab)",
    "Kondisi spesimen darah saat diterima lab": "kondisi",
    "Tanggal Terima Hasil Spesimen Serum Darah": "tgl hasil",
    "Hasil IgM Campak": "hasil campak",
    "Hasil IgM Rubella": "hasil rubela",
}
NAR_URINE = {"Tanggal Ambil Spesimen Urin": "tgl ambil", "Tanggal Kirim Spesimen Urin": "tgl kirim"}
NAR_REQUIRED = ("nomor epid", "tgl mulai rash")
NAR_DATE_OUT = [c for c in SRC_COLUMNS if c.startswith("Tanggal")]


class _RowParser(HTMLParser):
    """Cadangan bila lxml tidak terpasang: kumpulkan <tr> sebagai (punya_th, [teks sel])."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows, self._cur, self._cell, self._th = deque(), None, None, False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._cur, self._th = [], False
        elif tag in ("td", "th") and self._cur is not None:
            self._cell = []
            self._th = self._th or tag == "th"

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None and self._cur is not None:
            self._cur.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._cur is not None:
            self.rows.append((self._th, self._cur))
            self._cur = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _iter_html_rows(file):
    """Hasilkan (punya_th, [teks sel]) per <tr> dari file HTML (.xls bergaya HTML). Streaming."""
    file.seek(0)
    try:
        from lxml import etree
    except ImportError:
        etree = None
    if etree is not None:
        for _, el in etree.iterparse(file, events=("end",), tag="tr", html=True,
                                     encoding="utf-8", huge_tree=True):
            yield el.find("th") is not None, ["".join(c.itertext()).strip() for c in el if c.tag in ("td", "th")]
            el.clear()
            while el.getprevious() is not None:
                del el.getparent()[0]
        return
    decoder = codecs.getincrementaldecoder("utf-8")(errors="replace")
    parser = _RowParser()
    while True:
        chunk = file.read(1 << 20)
        if not chunk:
            break
        parser.feed(decoder.decode(chunk))
        while parser.rows:
            yield parser.rows.popleft()
    parser.feed(decoder.decode(b"", final=True))
    parser.close()
    while parser.rows:
        yield parser.rows.popleft()


def _sniff_kind(file):
    file.seek(0)
    head = file.read(512)
    file.seek(0)
    h = head.lstrip(b"\xef\xbb\xbf \t\r\n")
    if h[:2] == b"PK":
        return "xlsx"
    if h[:4] == b"\xd0\xcf\x11\xe0":
        return "xls"
    if h[:1] == b"<":
        return "html"
    return "csv"


def _iter_nar_cells(file):
    """Hasilkan baris NAR sebagai list teks/nilai: pertama baris header, lalu baris data."""
    kind = _sniff_kind(file)
    if kind == "html":
        header_seen = False
        for has_th, cells in _iter_html_rows(file):
            if not header_seen:
                if has_th or any(canon(c) == "nomor epid" for c in cells):
                    header_seen = True
                    yield cells
            else:
                yield cells
        return
    if kind == "xls":
        raise ValueError("Format .xls biner lama tidak didukung. Simpan file sebagai .xlsx / .csv, "
                         "atau unggah file tarikan NAR apa adanya (.xls bergaya HTML).")
    file.seek(0)
    if kind == "csv":
        raw = pd.read_csv(file, header=None, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    else:
        raw = pd.read_excel(file, header=None, dtype=object, engine=BASE_ENGINE)
    hdr_i = next((i for i in range(min(len(raw), 30))
                  if any(canon(c) == "nomor epid" for c in raw.iloc[i].tolist())), None)
    if hdr_i is None:
        raise ValueError("Header NAR tidak ditemukan (kolom 'Nomor EPID' tidak ada).")
    for row in raw.iloc[hdr_i:].itertuples(index=False, name=None):
        yield list(row)


def _nar_val(v):
    """Nilai sel NAR: kosong / '-' -> None; teks di-strip."""
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, str):
        v = v.strip()
        return None if v in ("", "-") else v
    return v


@lru_cache(maxsize=65536)
def _parse_date_text(t):
    """'dd-mm-yyyy' / 'dd/mm/yyyy' / 'yyyy-mm-dd' -> datetime; selain itu kembalikan teks aslinya."""
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", t)
    parts = (int(m[3]), int(m[2]), int(m[1])) if m else None
    if parts is None:
        m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$", t)
        parts = (int(m[1]), int(m[2]), int(m[3])) if m else None
    if parts is None or parts[0] < 1900:
        return t
    try:
        return datetime(*parts)
    except ValueError:
        return t


def _to_datetime(v):
    """Nilai sel -> datetime bila bisa dibaca sebagai tanggal; selain itu kembalikan apa adanya."""
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        return _parse_date_text(v)
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 36526 <= v <= 73050:
        return EXCEL_EPOCH + timedelta(days=float(v))
    return v


def nar_rash_ok(v):
    """True bila Tgl Mulai Rash ADA (tidak kosong), terbaca sebagai tanggal, dan tahunnya >= RASH_MIN_YEAR."""
    v = _nar_val(v)
    if v is None:
        return False
    d = _to_datetime(v)
    return isinstance(d, datetime) and d.year >= RASH_MIN_YEAR


def _map_nar_row(cells, idx):
    """Petakan satu baris NAR ke daftar nilai sesuai urutan FULL_LAYOUT."""
    def get(name):
        i = idx.get(name)
        return _nar_val(cells[i]) if i is not None and i < len(cells) else None

    out = {c: get(n) for c, n in NAR_DIRECT.items()}

    # Spesimen 1 / 2: ambil spesimen serum pertama dan spesimen urin pertama
    serum_done = urin_done = False
    for k in (1, 2):
        jenis = (get(f"spesimen {k} jenis") or "").lower()
        if "serum" in jenis and not serum_done:
            serum_done = True
            for col, suffix in NAR_SERUM.items():
                out[col] = get(f"spesimen {k} {suffix}")
        elif "urin" in jenis and not urin_done:
            urin_done = True
            for col, suffix in NAR_URINE.items():
                out[col] = get(f"spesimen {k} {suffix}")
            hasil = [f"{lab}: {v}" for lab, v in (("Campak", get(f"spesimen {k} hasil campak")),
                                                  ("Rubella", get(f"spesimen {k} hasil rubela"))) if v]
            out["Hasil Spesimen Urin"] = "; ".join(hasil) or None

    # tanggal -> datetime (teks yang bukan tanggal valid dibiarkan apa adanya)
    for col in NAR_DATE_OUT:
        if out.get(col) is not None:
            out[col] = _to_datetime(out[col])

    # Tanggal Vaksin MR terakhir = paling baru dari 4 kolom vaksin NAR
    vax = [d for d in (_to_datetime(get(n)) for n in NAR_VAX_DATES) if isinstance(d, datetime)]
    out[VAX_COL] = max(vax) if vax else None

    for col in ("Tahun", "Bulan"):
        v = out.get(col)
        if isinstance(v, str) and v.isdigit():
            out[col] = int(v)

    prov = out.get(PROV_COL)
    if prov is not None:
        out[PROV_COL] = match_province(_text(prov)) or prov
    return [out.get(c) for c in FULL_LAYOUT]


def read_nar(file):
    """Baca file tarikan NAR (.xls bergaya HTML / .xlsx / .csv).

    Return (DataFrame 46 kolom Data Line, total_baris_di_file, warning_rash, kolom_NAR_tidak_ditemukan).
    Semua baris data NAR tetap diolah. Tgl Mulai Rash kosong, tidak valid, atau tahun
    < RASH_MIN_YEAR hanya dihitung sebagai warning. Kolom 'No' dinomori ulang 1..N.
    """
    it = _iter_nar_cells(file)
    header = next(it, None)
    if header is None:
        raise ValueError("Header NAR tidak ditemukan (kolom 'Nomor EPID' tidak ada).")
    idx = {}
    for i, name in enumerate(header):
        idx.setdefault(re.sub(r"\s+", " ", str(name)).strip().lower(), i)
    absent = [n for n in NAR_REQUIRED if n not in idx]
    if absent:
        raise ValueError("Kolom NAR tidak ditemukan: " + ", ".join(absent))
    rash_i = idx["tgl mulai rash"]

    records, total, rash_warnings = [], 0, 0
    min_cells = max(5, len(header) // 2)
    for cells in it:
        if len(cells) < min_cells or all(_nar_val(c) is None for c in cells):
            continue    # baris judul / catatan / kosong; bukan record data
        total += 1

        # Warning saja: record tetap diproses walaupun rash kosong/tidak valid/< batas tahun.
        rash_value = cells[rash_i] if rash_i < len(cells) else None
        if not nar_rash_ok(rash_value):
            rash_warnings += 1

        records.append(_map_nar_row(cells, idx))

    wanted = set(NAR_DIRECT.values()) | set(NAR_VAX_DATES)
    missing = sorted(n for n in wanted if n not in idx)
    df = pd.DataFrame(records, columns=FULL_LAYOUT, dtype=object)
    normalize_date_columns(df)
    df[NO_COL] = range(1, len(df) + 1)
    return df, total, rash_warnings, missing


def _save_nar(df, workdir, name, size, total=None, dropped=0, missing=None):
    path = os.path.join(workdir, f"nar_{uuid.uuid4().hex[:8]}.pkl")
    df.to_pickle(path)
    return {"name": name, "size": size, "path": path, "rows": len(df),
            "total": total if total is not None else len(df), "dropped": dropped, "missing": missing or [],
            "flags": count_flags(df)}


def process_nar(file, workdir):
    """Proses file NAR -> info (data tersimpan sebagai pickle)."""
    df, total, rash_warnings, missing = read_nar(file)
    info = _save_nar(df, workdir, file.name, file.size, total, rash_warnings, missing)
    del df
    gc.collect()
    return info


# ---- keluaran ---------------------------------------------------------------
def output_columns(records):
    """Kolom hasil = 46 kolom Data Line, 'Sumber Data', 'Flag Tgl Rash', lalu kolom ekstra dari data dasar."""
    extras = []
    for r in records:
        for c in r["columns"]:
            if c not in FULL_LAYOUT and c not in (SOURCE_COL, FLAG_COL) and c not in extras:
                extras.append(c)
    return FULL_LAYOUT + [SOURCE_COL, FLAG_COL] + extras, extras


def iter_frames(records, columns, for_csv=False, nar=None, only_flagged=False):
    """DataFrame per record (satu per satu), kolom seragam, 'No' berlanjut 1..N.

    Baris provinsi lebih dulu, lalu (bila ada) baris NAR di-append di bawahnya.
    'Flag Tgl Rash' selalu dihitung ulang dari kolom Tanggal Mulai Rash.
    only_flagged=True: hanya baris ter-flag (kolom 'No' tetap nomor urut di data gabungan lengkap).
    """
    counter = 1
    date_cols = [c for c in columns if canon(c).startswith("tanggal") or _first_line_key(c).startswith("tanggal")]
    sources = [(rec["path"], SOURCE_PROV) for rec in records]
    if nar:
        sources.append((nar["path"], SOURCE_NAR))
    for path, label in sources:
        df = pd.read_pickle(path).reindex(columns=columns)
        df[SOURCE_COL] = label
        df[FLAG_COL] = flag_series(df)
        df[NO_COL] = range(counter, counter + len(df))
        counter += len(df)
        if only_flagged:
            df = df[df[FLAG_COL] != ""]
            if df.empty:
                continue
        if for_csv:
            for c in date_cols:
                df[c] = [v.strftime("%Y-%m-%d") if hasattr(v, "strftime") and not pd.isna(v) else v for v in df[c]]
        yield df


def build_csv(records, out_path, nar=None):
    columns, _ = output_columns(records)
    total = 0
    with open(out_path, "w", encoding="utf-8-sig", newline="") as fh:
        for i, df in enumerate(iter_frames(records, columns, for_csv=True, nar=nar)):
            df.to_csv(fh, index=False, header=(i == 0))
            total += len(df)
    return total


def build_xlsx(records, out_path, nar=None, only_flagged=False):
    import xlsxwriter

    columns, _ = output_columns(records)
    wb = xlsxwriter.Workbook(
        out_path,
        {"constant_memory": True, "nan_inf_to_errors": True, "default_date_format": DATE_DISPLAY_FORMAT},
    )
    base_title = "Baris_Flag_Rash" if only_flagged else "Combined_Linelist"
    state = {"ws": None, "row": 0, "n": 0}

    def new_sheet():
        state["n"] += 1
        title = base_title if state["n"] == 1 else f"{base_title}_{state['n']}"
        state["ws"] = wb.add_worksheet(title)
        state["ws"].write_row(0, 0, columns)
        state["row"] = 1

    new_sheet()
    total = 0
    for df in iter_frames(records, columns, nar=nar, only_flagged=only_flagged):
        values = df.astype(object).where(df.notna(), None)
        for row in values.itertuples(index=False, name=None):
            if state["row"] > EXCEL_MAX_DATA_ROWS:
                new_sheet()
            state["ws"].write_row(state["row"], 0, row)
            state["row"] += 1
        total += len(df)
        del values
        gc.collect()
    wb.close()
    return total, state["n"]


def flag_summary_rows(records, nar=None):
    """Baris ringkasan flag per provinsi (dan NAR): hanya yang punya flag."""
    rows = []
    for rec in records:
        f = rec.get("flags") or {}
        if sum(f.values()):
            rows.append({"Provinsi": rec["province"], "Jumlah Baris": rec["rows"],
                         **{t: f.get(t, 0) for t in FLAG_TYPES}, "Total Flag": sum(f.values())})
    if nar and sum((nar.get("flags") or {}).values()):
        f = nar["flags"]
        rows.append({"Provinsi": "NAR (semua provinsi)", "Jumlah Baris": nar["rows"],
                     **{t: f.get(t, 0) for t in FLAG_TYPES}, "Total Flag": sum(f.values())})
    return rows


# ----------------------------------------------------------------------------
# Antarmuka Streamlit
# ----------------------------------------------------------------------------
def _init_state():
    if "workdir" not in st.session_state:
        st.session_state.workdir = tempfile.mkdtemp(prefix="linelist_")
    defaults = {"records": [], "uploader_key": 0, "base_key": 0, "nar_key": 0, "nar": None,
                "csv": None, "xlsx": None, "flagged": None}
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def _reset_all():
    shutil.rmtree(st.session_state.workdir, ignore_errors=True)
    for key in ("workdir", "records", "csv", "xlsx", "flagged", "nar"):
        st.session_state.pop(key, None)
    st.session_state.uploader_key += 1
    st.session_state.base_key += 1
    st.session_state.nar_key += 1


def _invalidate_outputs():
    st.session_state.csv = None
    st.session_state.xlsx = None
    st.session_state.flagged = None


def _fail(file, exc):
    return {"id": uuid.uuid4().hex[:8], "name": file.name, "size": file.size, "kind": "Baru", "rows": 0,
            "dropped": 0, "missing": [], "flags": Counter(), "province": None, "status": "error",
            "detail": f"Gagal: {exc}"}


def _drop_nar():
    old = st.session_state.nar
    if old:
        try:
            os.remove(old["path"])
        except OSError:
            pass
    st.session_state.nar = None


def _process_base(file):
    try:
        with st.spinner("Membaca data dasar..."):
            recs, nar_df = load_base(file, st.session_state.workdir)
        for rec in recs:
            if rec["status"] == "ok":
                activate(rec, [])
                rec["detail"] = f"Data dasar - {rec['province']}"
                n_flag = sum(rec["flags"].values())
                if n_flag:
                    rec["detail"] += f"; {n_flag:,} baris ter-flag Tgl Rash"
            else:
                rec["detail"] = f"Perlu dipilih manual: {rec['note']}"
        st.session_state.records.extend(recs)
        if nar_df is not None and len(nar_df) and not st.session_state.nar:
            st.session_state.nar = _save_nar(nar_df, st.session_state.workdir,
                                             f"{file.name} (baris NAR dari data dasar)", file.size)
    except Exception as exc:  # noqa: BLE001
        st.session_state.records.append(_fail(file, exc))
    _invalidate_outputs()


def _process_batch(files):
    records = st.session_state.records
    done = {(r["name"], r["size"]) for r in records if r["status"] in ("ok", "replaced", "pending")}
    progress = st.progress(0.0)
    for i, file in enumerate(files, start=1):
        progress.progress((i - 1) / len(files), text=f"Memproses {file.name} ({i}/{len(files)})")
        if (file.name, file.size) in done:
            rec = _fail(file, "")
            rec.update(status="skip", detail="Dilewati: file ini sudah diproses sebelumnya")
        else:
            try:
                rec = process_new_file(file, st.session_state.workdir)
                if rec["status"] == "ok":
                    activate(rec, records)
                    if rec["missing"]:
                        rec["detail"] += f"; kolom tidak ditemukan: {', '.join(rec['missing'])}"
                else:
                    rec["detail"] = f"Perlu dipilih manual: {rec['note']}"
                done.add((file.name, file.size))
            except Exception as exc:  # noqa: BLE001  (sheet tidak ada, file rusak, dll.)
                rec = _fail(file, exc)
        records.append(rec)
    progress.empty()
    _invalidate_outputs()


def _process_nar_upload(file):
    try:
        with st.spinner("Membaca file NAR... file besar bisa memakan waktu sekitar setengah menit"):
            info = process_nar(file, st.session_state.workdir)
        _drop_nar()
        st.session_state.nar = info
        _invalidate_outputs()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Gagal membaca file NAR '{file.name}': {exc}")


def _pending_form(records):
    pending = [r for r in records if r["status"] == "pending"]
    if not pending:
        return
    st.warning(f"{len(pending)} file/kelompok belum diketahui provinsinya. Pilih provinsi lalu klik Terapkan; "
               "datanya belum ikut digabung.")
    with st.form("pending_form"):
        choices = {
            r["id"]: st.selectbox(f"{r['name']} ({r['rows']:,} baris)", ["-- pilih --"] + PROVINCES,
                                  key=f"assign_{r['id']}")
            for r in pending
        }
        if st.form_submit_button("Terapkan pilihan provinsi"):
            for rec in pending:
                choice = choices[rec["id"]]
                if choice in PROVINCES:
                    df = pd.read_pickle(rec["path"])
                    fill_province(df, choice)
                    df.to_pickle(rec["path"])
                    rec.update(province=choice, via="manual", note="dipilih manual")
                    activate(rec, records)
            _invalidate_outputs()
            st.rerun()


def main():
    st.set_page_config(page_title="Updater Database W41 per Provinsi", layout="wide")
    _init_state()

    st.title("Updater Database W41 per Provinsi")
    st.write(
        f"Gunakan **W41 Untouched Data** sebagai database dasar. Setelah itu unggah file Excel provinsi "
        f"terbaru (.xlsx / .xlsm). Hanya sheet '{SHEET_NAME}' dan kolom format Data Line yang dibaca. "
        "Jika file YOGYAKARTA diunggah, seluruh baris lama YOGYAKARTA akan **diganti utuh** oleh data "
        "terbaru, sedangkan semua provinsi lain tetap sama seperti di W41. "
        "Tanggal dinormalisasi sebelum proses replace/append. Format US yang tidak ambigu seperti "
        "1/15/2026 menjadi 15/01/2026; tanggal ambigu dipilih berdasarkan kedekatan dengan "
        "Tanggal Investigasi, lalu default ke dd/mm/yyyy bila referensi tidak tersedia. "
        "Nama bulan Indonesia seperti 13 Agustus 2026 juga dikenali."
    )

    records = st.session_state.records
    nar = st.session_state.nar

    if not records:
        st.subheader("Langkah 1: unggah database dasar W41")
        base = st.file_uploader("Unggah W41 Untouched Data (CSV / XLSX / XLSM)",
                                type=["csv", "xlsx", "xlsm"], key=f"base_{st.session_state.base_key}")
        if base and st.button("Muat database W41", type="primary"):
            _process_base(base)
            st.session_state.base_key += 1
            st.rerun()

    st.subheader("Langkah 2: unggah file provinsi terbaru" if records else "Langkah 2: unggah file provinsi terbaru")
    files = st.file_uploader("Unggah file Excel provinsi terbaru", type=["xlsx", "xlsm"],
                             accept_multiple_files=True, key=f"uploader_{st.session_state.uploader_key}")
    if files and st.button("Proses file", type="primary"):
        _process_batch(files)
        st.session_state.uploader_key += 1
        st.rerun()

    st.subheader("Langkah 3 (opsional): data NAR")
    st.caption(
        f"Hanya kolom NAR yang setara kolom Data Line yang diambil. Semua baris data NAR tetap diolah. "
        f"Tgl Mulai Rash kosong, tidak valid, atau tahun < {RASH_MIN_YEAR} hanya ditampilkan sebagai warning. "
        f"Baris NAR di-append di sheet yang sama, dengan kolom "
        f"'{SOURCE_COL}' = '{SOURCE_NAR}'. Mengunggah NAR baru menggantikan seluruh data NAR sebelumnya."
    )
    nar_file = st.file_uploader("Unggah file tarikan NAR (.xls / .xlsx / .csv)", type=["xls", "xlsx", "csv"],
                                key=f"nar_{st.session_state.nar_key}")
    if nar_file and st.button("Proses file NAR"):
        _process_nar_upload(nar_file)
        st.session_state.nar_key += 1
        st.rerun()
    if nar:
        n1, n2, n3 = st.columns(3)
        n1.metric("Baris NAR di file", f"{nar['total']:,}")
        n2.metric("Dipakai / diolah", f"{nar['rows']:,}")
        n3.metric("Warning Tgl Mulai Rash", f"{nar['dropped']:,}")
        st.caption(f"Sumber: {nar['name']}")
        if nar["rows"] == 0:
            st.warning("Tidak ada record data NAR yang dapat dibaca.")
        elif nar["dropped"] > 0:
            st.warning(
                f"{nar['dropped']:,} dari {nar['total']:,} baris NAR memiliki Tgl Mulai Rash "
                f"kosong, tidak valid, atau tahun < {RASH_MIN_YEAR}. Semua baris tersebut tetap diolah."
            )
        if st.button("Hapus data NAR"):
            _drop_nar()
            _invalidate_outputs()
            st.rerun()

    _pending_form(records)

    if records or nar:
        st.subheader("Ringkasan File")
        rows = [
            {
                "Nama File": r["name"],
                "Jenis": r.get("kind", ""),
                "Provinsi": r.get("province") or "-",
                "Jumlah Kasus": r["rows"],
                "Baris template dibuang": r.get("dropped", 0),
                "Baris ter-flag Tgl Rash": sum((r.get("flags") or {}).values()),
                "Status": {"ok": "Dipakai", "replaced": "Digantikan"}.get(r["status"], r["status"]),
                "Keterangan": r["detail"],
            }
            for r in records
        ]
        if nar:
            rows.append({
                "Nama File": nar["name"], "Jenis": SOURCE_NAR, "Provinsi": "(sesuai kolom Provinsi NAR)",
                "Jumlah Kasus": nar["rows"], "Baris template dibuang": 0,
                "Baris ter-flag Tgl Rash": sum((nar.get("flags") or {}).values()), "Status": "Dipakai",
                "Keterangan": (f"{nar['dropped']:,} warning Tgl Mulai Rash kosong/tidak valid/"
                               f"tahun < {RASH_MIN_YEAR}; semua baris tetap diolah"),
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True)

    ok = active_records(records)
    if ok or nar:
        _, extras = output_columns(ok)
        if extras:
            st.info("Data dasar punya kolom tambahan (mis. hasil cleaning) yang dipertahankan: "
                    f"{', '.join(extras)}. Untuk provinsi yang diganti dan baris NAR, kolom ini akan kosong.")

        n_prov = sum(r["rows"] for r in ok)
        n_nar = nar["rows"] if nar else 0
        flag_rows = flag_summary_rows(ok, nar)
        n_flag = sum(r["Total Flag"] for r in flag_rows)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Kasus Tergabung", f"{n_prov + n_nar:,}")
        c2.metric("Provinsi dalam data", f"{len({r['province'] for r in ok})} provinsi")
        c3.metric("Dari NAR", f"{n_nar:,}")
        c4.metric("Baris ter-flag Tgl Rash", f"{n_flag:,}")

        st.subheader("Flag Tanggal Mulai Rash")
        st.caption(
            f"Baris tidak dibuang, hanya ditandai pada kolom '{FLAG_COL}' di hasil unduhan: "
            f"'{FLAG_LT}' (tahun < {RASH_MIN_YEAR}), '{FLAG_BLANK}', atau '{FLAG_INVALID}' (teks yang bukan tanggal)."
        )
        if flag_rows:
            st.dataframe(pd.DataFrame(flag_rows).sort_values("Total Flag", ascending=False), hide_index=True)
        else:
            st.success(f"Tidak ada baris dengan Tgl Mulai Rash < {RASH_MIN_YEAR}, kosong, atau tidak valid.")

        st.subheader("Unduh Hasil Penggabungan")
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("1. Siapkan file CSV (cepat, disarankan)"):
                with st.spinner("Menyusun CSV..."):
                    path = os.path.join(st.session_state.workdir, "Linelist_Combined.csv")
                    total = build_csv(ok, path, nar)
                st.session_state.csv = {"path": path, "rows": total}
            if st.session_state.csv:
                info = st.session_state.csv
                size_mb = os.path.getsize(info["path"]) / 1e6
                with open(info["path"], "rb") as fh:
                    st.download_button(f"Unduh CSV ({info['rows']:,} baris, {size_mb:,.1f} MB)", data=fh,
                                       file_name="Linelist_Combined.csv", mime="text/csv", type="primary")

        with col2:
            if st.button("2. Siapkan file Excel (.xlsx) - lebih lama"):
                with st.spinner("Menyusun Excel... data besar bisa memakan beberapa menit"):
                    path = os.path.join(st.session_state.workdir, "Linelist_Combined.xlsx")
                    total, n_sheets = build_xlsx(ok, path, nar)
                st.session_state.xlsx = {"path": path, "rows": total, "sheets": n_sheets}
            if st.session_state.xlsx:
                info = st.session_state.xlsx
                size_mb = os.path.getsize(info["path"]) / 1e6
                if info["sheets"] > 1:
                    st.info(f"Data melebihi batas baris Excel, jadi dipecah ke {info['sheets']} sheet.")
                with open(info["path"], "rb") as fh:
                    st.download_button(
                        f"Unduh Excel ({info['rows']:,} baris, {size_mb:,.1f} MB)", data=fh,
                        file_name="Linelist_Combined.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with col3:
            if n_flag and st.button("3. Siapkan daftar baris ter-flag (.xlsx)"):
                with st.spinner("Menyusun daftar baris ter-flag..."):
                    path = os.path.join(st.session_state.workdir, "Linelist_Flag_Tgl_Rash.xlsx")
                    total, n_sheets = build_xlsx(ok, path, nar, only_flagged=True)
                st.session_state.flagged = {"path": path, "rows": total}
            if not n_flag:
                st.caption("Tidak ada baris ter-flag.")
            if st.session_state.flagged:
                info = st.session_state.flagged
                size_mb = os.path.getsize(info["path"]) / 1e6
                with open(info["path"], "rb") as fh:
                    st.download_button(
                        f"Unduh baris ter-flag ({info['rows']:,} baris, {size_mb:,.1f} MB)", data=fh,
                        file_name="Linelist_Flag_Tgl_Rash.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if st.session_state.csv:
            st.subheader("Pratinjau Data Tergabung")
            st.dataframe(pd.read_csv(st.session_state.csv["path"], nrows=10, encoding="utf-8-sig"))

    if (records or nar) and st.button("Hapus semua hasil & mulai dari awal"):
        _reset_all()
        st.rerun()


if __name__ == "__main__":
    main()
