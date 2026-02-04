
import streamlit as st
import face_recognition
import pandas as pd
import os
import pickle
import numpy as np
from datetime import datetime, date
import shutil
import socket
import qrcode
import uuid
from PIL import Image, ImageOps
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import zipfile
import time
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from copy import copy
import json
import glob
from copy import copy
# ==========================================================
# 1) CẤU HÌNH APP & SESSION STATE
# ==========================================================
st.set_page_config(page_title="IUH Smart Attendance", layout="wide", page_icon="🏫")

if "menu_selection" not in st.session_state:
    st.session_state.menu_selection = "Trang Chủ"
if "menu_selection_label" not in st.session_state:
    st.session_state.menu_selection_label = "🏠 Trang Chủ"
if "attendance_key" not in st.session_state:
    st.session_state.attendance_key = 0

# Lưu kết quả quét để review
if "scan_result" not in st.session_state:
    st.session_state.scan_result = None
if "save_result" not in st.session_state:
    st.session_state.save_result = None

PAGE_LABELS = {
    "Trang Chủ": "🏠 Trang Chủ",
    "Quản Lý Lớp & SV": "📂 Quản Lý Lớp & SV",
    "Điểm Danh": "📸 Điểm Danh",
    "Báo Cáo": "📊 Báo Cáo",
}
LABEL_TO_PAGE = {v: k for k, v in PAGE_LABELS.items()}

def navigate_to(page_name: str):
    st.session_state.menu_selection = page_name
    st.session_state.menu_selection_label = PAGE_LABELS.get(page_name, "🏠 Trang Chủ")

# ==========================================================
# 2) THEME CSS (ĐẸP + ANIMATION NHẸ) + STATUS PILL
# ==========================================================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&display=swap');

:root{
  --bg0:#f7f9fc;
  --bg1:#eef2ff;
  --card:rgba(255,255,255,.78);
  --card2:rgba(255,255,255,.92);
  --text:#0f172a;
  --muted:#64748b;
  --border:rgba(148,163,184,.35);
  --shadow: 0 10px 30px rgba(2, 6, 23, .08);
  --shadow2: 0 18px 60px rgba(2, 6, 23, .12);
  --radius: 18px;
  --radius2: 22px;
}

/* Background */
.stApp {
  background:
    radial-gradient(1200px 600px at 10% 0%, rgba(99,102,241,.14), transparent 55%),
    radial-gradient(900px 500px at 90% 10%, rgba(59,130,246,.12), transparent 50%),
    linear-gradient(180deg, var(--bg0) 0%, var(--bg1) 100%);
}

/* Font */
html, body, [class*="css"] { font-family: 'Inter', sans-serif; color: var(--text); }

/* Main padding */
/* ✅ Fix: chừa khoảng trống tránh dính chữ vào header Streamlit */
div.block-container { 
  padding-top: 2.8rem; 
  padding-bottom: 2.2rem; 
}

/* (optional) header trong suốt cho đẹp */
header[data-testid="stHeader"]{
  background: transparent;
}

/* Sidebar */
section[data-testid="stSidebar"]{
  background: rgba(255,255,255,.92);
  border-right: 1px solid rgba(148,163,184,.28);
}

/* Card wrapper: st.container(border=True) */
@keyframes fadeInUp { from {opacity:0; transform: translateY(6px);} to {opacity:1; transform: translateY(0);} }
div[data-testid="stVerticalBlockBorderWrapper"]{
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius2);
  box-shadow: var(--shadow);
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
  padding: 14px 16px;
  animation: fadeInUp .32s ease both;
}

/* Badge */
.badge{
  display:inline-block;
  padding: 6px 10px;
  border-radius: 999px;
  border: 1px solid rgba(148,163,184,.35);
  background: rgba(255,255,255,.75);
  color: var(--muted);
  font-size: 12px;
  font-weight: 800;
}
.badge.pulse{ animation: pulse 1.6s ease-in-out infinite; }
@keyframes pulse { 0%{transform:scale(1);} 50%{transform:scale(1.04);} 100%{transform:scale(1);} }

.small-muted{ color: var(--muted); font-size: 13px; }

/* Metric */
div[data-testid="stMetric"]{
  background: var(--card2);
  border: 1px solid rgba(148,163,184,.28);
  border-radius: var(--radius);
  padding: 12px 14px;
  box-shadow: 0 8px 24px rgba(2,6,23,.06);
}
div[data-testid="stMetric"] [data-testid="stMetricLabel"]{
  color: var(--muted);
  font-weight: 650;
}

/* Buttons */
div.stButton > button{
  border: 1px solid rgba(148,163,184,.35) !important;
  border-radius: 14px !important;
  padding: 10px 14px !important;
  font-weight: 800 !important;
  background: linear-gradient(90deg, rgba(99,102,241,.95), rgba(59,130,246,.95)) !important;
  color: #fff !important;
  box-shadow: 0 10px 24px rgba(59,130,246,.18);
  transition: .15s ease;
}
div.stButton > button:hover{
  transform: translateY(-1px);
  box-shadow: 0 16px 32px rgba(59,130,246,.22);
}
button[kind="secondary"]{
  background: rgba(255,255,255,.86) !important;
  color: var(--text) !important;
}

/* Progress */
.stProgress > div > div > div > div{
  background-image: linear-gradient(90deg, rgba(99,102,241,.95), rgba(59,130,246,.95));
}

/* Status pill */
.status-pill{
  display:flex;
  align-items:center;
  gap:8px;
  padding:10px 12px;
  border-radius: 14px;
  border: 1px solid rgba(148,163,184,.35);
  font-weight: 800;
  font-size: 13px;
}
.status-pill.online{
  background: rgba(34,197,94,.12);
  border-color: rgba(34,197,94,.35);
  color: #166534;
}
.status-pill.offline{
  background: rgba(239,68,68,.12);
  border-color: rgba(239,68,68,.35);
  color: #7f1d1d;
}
.status-sub{
  font-weight: 650;
  opacity: .8;
}
</style>
""",
    unsafe_allow_html=True,
)

def bordered_container():
    try:
        return st.container(border=True)
    except TypeError:
        return st.container()

def section_header(title: str, subtitle: str = "", badge: str = ""):
    c1, c2 = st.columns([0.82, 0.18])
    with c1:
        st.markdown(f"### {title}")
        if subtitle:
            st.caption(subtitle)
    with c2:
        if badge:
            st.markdown(f"<div style='text-align:right'><span class='badge'>{badge}</span></div>", unsafe_allow_html=True)
    st.divider()

# ==========================================================
# 3) STORAGE CONFIG
# ==========================================================
FOLDER_NAME = "DiemDanh_AI_Final"
IMG_DIR = "img"
SIDEBAR_LOGO_PATH = os.path.join(IMG_DIR, "logo_sidebar.png")
MAIN_LOGO_PATH = os.path.join(IMG_DIR, "logo_main.png")

drive_paths = [r"G:\My Drive", r"E:\My Drive", r"F:\My Drive"]
ROOT_FOLDER = FOLDER_NAME
MODE = "OFFLINE"

for p in drive_paths:
    if os.path.exists(p):
        ROOT_FOLDER = os.path.join(p, FOLDER_NAME)
        MODE = "ONLINE (Cloud)"
        break

os.makedirs(ROOT_FOLDER, exist_ok=True)

DATA_FILE_NAME = "data_khuon_mat.pickle"
EXCEL_NAME = "DanhSachLop.xlsx"
ROSTER_JSON = "roster.json"
HISTORY_DIR_NAME = "Attendance_History"
SESSIONS_CSV = "sessions.csv"
BAD_IMG_DIR = "Anh_Loi"  # lưu ảnh không detect được mặt

# ==========================================================
# 4) UTILS
# ==========================================================
def get_real_cell(sheet, row, col):
    cell = sheet.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return cell

    # nếu là merged cell -> tìm vùng merge chứa nó
    for rng in sheet.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return sheet.cell(row=rng.min_row, column=rng.min_col)  # top-left cell
    return None
def show_centered_image(path, width=220):
    if os.path.exists(path):
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            st.image(path, width=width)

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except:
        return "localhost"

def generate_qr(url: str):
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    return img.convert("RGB")


def get_list_classes():
    if not os.path.exists(ROOT_FOLDER):
        return []
    return sorted([
        d for d in os.listdir(ROOT_FOLDER)
        if os.path.isdir(os.path.join(ROOT_FOLDER, d))
        and not d.startswith(".")
    ])

def ensure_class_paths(class_name: str):
    class_path = os.path.join(ROOT_FOLDER, class_name)
    img_base = os.path.join(class_path, "Anh_Sinh_Vien")
    bad_base = os.path.join(class_path, BAD_IMG_DIR)
    os.makedirs(class_path, exist_ok=True)
    os.makedirs(img_base, exist_ok=True)
    os.makedirs(bad_base, exist_ok=True)
    return class_path, img_base, bad_base

def load_class_data(class_name):
    path = os.path.join(ROOT_FOLDER, class_name, DATA_FILE_NAME)
    if not os.path.exists(path):
        return {"encodings": [], "ids": []}
    with open(path, "rb") as f:
        return pickle.load(f)

def save_class_data(class_name, data):
    path = os.path.join(ROOT_FOLDER, class_name, DATA_FILE_NAME)
    with open(path, "wb") as f:
        pickle.dump(data, f)

def _cell_to_text(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return str(v).strip()

# ==========================================================
# 5) ROSTER (DANH SÁCH SV TỪ EXCEL) + SYNC FOLDER
# ==========================================================
def find_excel_header_in_sheet(sheet):
    """
    Tìm dòng header và mapping cột theo template IUH.
    Trả về: (header_row, cols)
    cols có thể gồm: stt, mssv, ho, ten, gioitinh, ngaysinh, lophoc
    """
    header_row = None
    cols = {
        "stt": None,
        "mssv": None,
        "ho": None,
        "ten": None,
        "gioitinh": None,
        "ngaysinh": None,
        "lophoc": None,
    }

    for r in range(1, 26):
        for c in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=r, column=c)
            if isinstance(cell, MergedCell):
                continue
            val = str(cell.value).strip().lower() if cell.value else ""

            if val == "stt" or "số thứ tự" in val:
                cols["stt"] = c
            elif ("mã sinh viên" in val) or (val == "mssv") or ("mssv" in val):
                header_row = r
                cols["mssv"] = c
            elif val in ["họ đệm", "họ", "họ và tên", "họ tên"]:
                cols["ho"] = c
            elif val == "tên":
                cols["ten"] = c
            elif "giới tính" in val:
                cols["gioitinh"] = c
            elif "ngày sinh" in val:
                cols["ngaysinh"] = c
            elif val == "lớp học" or val == "lớp":
                cols["lophoc"] = c

        if header_row and cols["mssv"]:
            break

    return header_row, cols

def load_roster_from_excel_path(excel_path: str) -> pd.DataFrame:
    """
    ✅ Nâng cấp: đọc đầy đủ các cột cần cho SV mới:
    - Mã sinh viên, Họ đệm, Tên, Giới tính, Ngày sinh, Lớp học
    """
    if not os.path.exists(excel_path):
        return pd.DataFrame()

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb.active
        header_row, cols = find_excel_header_in_sheet(sheet)
        if not header_row or not cols.get("mssv"):
            return pd.DataFrame()

        rows = []
        for r in range(header_row + 1, sheet.max_row + 1):
            m_cell = sheet.cell(row=r, column=cols["mssv"])
            if isinstance(m_cell, MergedCell):
                continue
            mssv = str(m_cell.value).strip() if m_cell.value else ""
            if not mssv:
                # nếu gặp dòng trống -> coi như kết thúc
                continue
            if "tổng" in mssv.lower():
                break

            ho = _cell_to_text(sheet.cell(row=r, column=cols["ho"]).value) if cols.get("ho") else ""
            ten = _cell_to_text(sheet.cell(row=r, column=cols["ten"]).value) if cols.get("ten") else ""
            gt = _cell_to_text(sheet.cell(row=r, column=cols["gioitinh"]).value) if cols.get("gioitinh") else ""
            ns = sheet.cell(row=r, column=cols["ngaysinh"]).value if cols.get("ngaysinh") else ""
            lh = _cell_to_text(sheet.cell(row=r, column=cols["lophoc"]).value) if cols.get("lophoc") else ""

            ns_text = _cell_to_text(ns)
            full_name = f"{ho} {ten}".strip()

            rows.append({
                "MSSV": mssv.strip(),
                "Họ đệm": ho,
                "Họ": ho,  # alias để không gãy code cũ
                "Tên": ten,
                "Họ tên": full_name,
                "Giới tính": gt,
                "Ngày sinh": ns_text,
                "Lớp học": lh,
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df["MSSV"] = df["MSSV"].astype(str).str.strip()
            df = df.drop_duplicates(subset=["MSSV"]).reset_index(drop=True)
        return df
    except:
        return pd.DataFrame()

def save_roster_json(class_name: str, df_roster: pd.DataFrame):
    class_path = os.path.join(ROOT_FOLDER, class_name)
    os.makedirs(class_path, exist_ok=True)
    path = os.path.join(class_path, ROSTER_JSON)
    try:
        df_roster.to_json(path, orient="records", force_ascii=False, indent=2)
    except:
        pass

def load_roster(class_name: str, autosync_folders: bool = True) -> pd.DataFrame:
    class_path, img_base, _ = ensure_class_paths(class_name)
    excel_path = os.path.join(class_path, EXCEL_NAME)
    roster_path = os.path.join(class_path, ROSTER_JSON)

    df = pd.DataFrame()
    try:
        excel_mtime = os.path.getmtime(excel_path) if os.path.exists(excel_path) else 0
        roster_mtime = os.path.getmtime(roster_path) if os.path.exists(roster_path) else 0
        if os.path.exists(roster_path) and roster_mtime >= excel_mtime:
            df = pd.read_json(roster_path)
        else:
            df = load_roster_from_excel_path(excel_path)
            if not df.empty:
                save_roster_json(class_name, df)
    except:
        df = load_roster_from_excel_path(excel_path)

    # đảm bảo có đủ các cột (nếu roster.json cũ)
    if not df.empty:
        if "Họ tên" not in df.columns:
            ho = df["Họ"].astype(str) if "Họ" in df.columns else ""
            ten = df["Tên"].astype(str) if "Tên" in df.columns else ""
            df["Họ tên"] = (ho + " " + ten).str.strip()

        for col in ["Họ đệm", "Giới tính", "Ngày sinh", "Lớp học"]:
            if col not in df.columns:
                df[col] = ""

        if "Họ" not in df.columns and "Họ đệm" in df.columns:
            df["Họ"] = df["Họ đệm"]

    if autosync_folders and not df.empty:
        for mssv in df["MSSV"].astype(str).tolist():
            os.makedirs(os.path.join(img_base, str(mssv).strip()), exist_ok=True)
            os.makedirs(os.path.join(class_path, BAD_IMG_DIR, str(mssv).strip()), exist_ok=True)

    return df

def sync_roster_folders(class_name: str):
    df = load_roster(class_name, autosync_folders=True)
    return len(df)

# ==========================================================
# 5.1) EXCEL: ADD/REMOVE STUDENT (NEW)
# ==========================================================
def _copy_cell_style(src_cell, dst_cell):
    if not src_cell or not dst_cell:
        return
    try:
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.border = copy(src_cell.border)
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.number_format = copy(src_cell.number_format)
            dst_cell.protection = copy(src_cell.protection)
            dst_cell._style = copy(src_cell._style)
    except:
        pass
def flash(msg, kind="success", seconds=30):
    """
    kind: success | info | warning | error
    """
    box = st.empty()
    if kind == "success":
        box.success(msg)
    elif kind == "info":
        box.info(msg)
    elif kind == "warning":
        box.warning(msg)
    else:
        box.error(msg)

    time.sleep(seconds)
    box.empty()
def _safe_save_wb(wb, file_path: str, class_path: str, prefix: str):
    """
    Lưu workbook. Nếu file đang mở -> lưu bản copy.
    Return: (saved_path or None, warning_or_error or None, is_copy)
    """
    try:
        wb.save(file_path)
        return file_path, None, False
    except Exception:
        copy_path = os.path.join(class_path, f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        try:
            wb.save(copy_path)
            return copy_path, "⚠️ File Excel đang mở. Đã lưu bản copy mới.", True
        except Exception as e:
            return None, f"❌ Không thể lưu Excel: {e}", False

def _renumber_stt(sheet, header_row: int, cols: dict):
    if not cols.get("stt") or not cols.get("mssv"):
        return
    idx = 1
    for r in range(header_row + 1, sheet.max_row + 1):
        m = sheet.cell(row=r, column=cols["mssv"]).value
        mssv = str(m).strip() if m else ""
        if not mssv:
            continue
        if "tổng" in mssv.lower():
            break
        sheet.cell(row=r, column=cols["stt"]).value = idx
        idx += 1

def add_student_to_excel(class_name: str, student: dict):
    """
    Thêm 1 sinh viên vào bảng Excel (IUH template).
    Return: (ok, msg, saved_path)
    """
    class_path = os.path.join(ROOT_FOLDER, class_name)
    file_path = os.path.join(class_path, EXCEL_NAME)

    # check excel exists
    if not os.path.exists(file_path):
        return False, "⚠️ Lớp chưa có file Excel. Excel chưa được cập nhật.", None

    # check lock file (~$DanhSachLop.xlsx)
    lock_path = os.path.join(class_path, "~$" + EXCEL_NAME)
    if os.path.exists(lock_path):
        return False, f"❌ Excel đang mở/LOCK: {lock_path}. Hãy đóng Excel rồi thử lại.", None

    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active

        header_row, cols = find_excel_header_in_sheet(sheet)
        if not header_row or not cols.get("mssv"):
            return False, "❌ Không tìm thấy header/cột MSSV trong Excel.", None

        target_mssv = str(student.get("MSSV", "")).strip()
        if not target_mssv:
            return False, "❌ MSSV rỗng.", None

        # ---------- helpers ----------
        def _merged_master_cell(row: int, col: int):
            """
            Nếu (row,col) nằm trong merge và là MergedCell -> trả về ô master (góc trái-trên).
            Nếu không -> trả về chính nó.
            """
            cell = sheet.cell(row=row, column=col)
            if not isinstance(cell, MergedCell):
                return cell

            for rng in sheet.merged_cells.ranges:
                if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                    return sheet.cell(row=rng.min_row, column=rng.min_col)

            # fallback (hiếm)
            return sheet.cell(row=row, column=col)

        def _safe_set(row: int, col: int, value):
            c = _merged_master_cell(row, col)
            c.value = value

        def _row_has_total(r: int) -> bool:
            # kiểm tra text "tổng" trên toàn dòng, nhưng chỉ đọc ô master để tránh MergedCell
            for c in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=r, column=c)
                if isinstance(cell, MergedCell):
                    continue
                v = str(cell.value).strip().lower() if cell.value is not None else ""
                if "tổng" in v:
                    return True
            return False

        def _is_row_safe_for_insert(r: int) -> bool:
            # tránh dòng merge kiểu "Tổng cộng" hoặc dòng tiêu đề phụ
            if _row_has_total(r):
                return False
            # nếu cell MSSV là merged -> cũng tránh
            mcell = sheet.cell(row=r, column=cols["mssv"])
            if isinstance(mcell, MergedCell):
                return False
            return True

        # ---------- 1) check duplicate ----------
        for r in range(header_row + 1, sheet.max_row + 1):
            if _row_has_total(r):
                break

            mcell = sheet.cell(row=r, column=cols["mssv"])
            if isinstance(mcell, MergedCell):
                continue

            mv = mcell.value
            mssv = str(mv).strip() if mv else ""
            if not mssv:
                continue

            if mssv.upper() == target_mssv.upper():
                return False, f"⚠️ MSSV {target_mssv} đã tồn tại trong Excel.", None

        # ---------- 2) find insert_row ----------
        insert_row = None
        total_row = None

        # tìm total_row trước
        for r in range(header_row + 1, sheet.max_row + 1):
            if _row_has_total(r):
                total_row = r
                break

        # ưu tiên: nếu có dòng trống trước total_row và “an toàn” -> chèn vào đó
        scan_end = total_row if total_row else (sheet.max_row + 1)
        for r in range(header_row + 1, scan_end):
            if not _is_row_safe_for_insert(r):
                continue

            mv = sheet.cell(row=r, column=cols["mssv"]).value
            mssv = str(mv).strip() if mv else ""
            if not mssv:
                insert_row = r
                break

        # nếu không có dòng trống hợp lệ:
        if insert_row is None:
            if total_row:
                insert_row = total_row
                sheet.insert_rows(insert_row, 1)  # chèn ngay trước tổng
                # total_row bị đẩy xuống 1 dòng
                total_row = total_row + 1
            else:
                insert_row = sheet.max_row + 1

        # ---------- 3) copy style từ template row ----------
        template_row = max(header_row + 1, insert_row - 1)

        try:
            sheet.row_dimensions[insert_row].height = sheet.row_dimensions[template_row].height
        except:
            pass

        for c in range(1, sheet.max_column + 1):
            src = sheet.cell(row=template_row, column=c)
            dst = sheet.cell(row=insert_row, column=c)

            # nếu dst là mergedcell thì skip (không copy vào được)
            if isinstance(dst, MergedCell) or isinstance(src, MergedCell):
                continue

            _copy_cell_style(src, dst)
            dst.value = None

        # ---------- 4) set values (SAFE) ----------
        def _set(col_key, value):
            if cols.get(col_key):
                _safe_set(insert_row, cols[col_key], value)

        # STT
        if cols.get("stt"):
            prev_cell = sheet.cell(row=template_row, column=cols["stt"])
            prev_val = prev_cell.value
            if isinstance(prev_val, (int, float)):
                stt_val = int(prev_val) + 1
            else:
                # fallback count until total
                stt_val = 1
                end_r = total_row if total_row else (sheet.max_row + 1)
                for rr in range(header_row + 1, end_r):
                    mv = sheet.cell(row=rr, column=cols["mssv"]).value
                    ms = str(mv).strip() if mv else ""
                    if ms:
                        stt_val += 1
            _set("stt", stt_val)

        _set("mssv", target_mssv)
        _set("ho", student.get("Họ đệm", ""))
        _set("ten", student.get("Tên", ""))
        _set("gioitinh", student.get("Giới tính", ""))

        ns = student.get("Ngày sinh", "")
        if isinstance(ns, (date, datetime)):
            _set("ngaysinh", ns if isinstance(ns, date) else ns.date())
        else:
            _set("ngaysinh", str(ns))

        _set("lophoc", student.get("Lớp học", ""))

        # renumber STT
        _renumber_stt(sheet, header_row, cols)

        # ---------- 5) update "Tổng cộng" (nếu có) ----------
        try:
            # tìm lại total_row (vì có thể bị shift)
            total_row2 = None
            for r in range(header_row + 1, sheet.max_row + 1):
                if _row_has_total(r):
                    total_row2 = r
                    break

            if total_row2:
                # đếm SV từ header tới trước total
                count_sv = 0
                for rr in range(header_row + 1, total_row2):
                    mv = sheet.cell(row=rr, column=cols["mssv"]).value
                    ms = str(mv).strip() if mv else ""
                    if ms:
                        count_sv += 1

                # tìm ô số (ưu tiên ô có number/formula trong dòng tổng)
                placed = False
                for cc in range(1, sheet.max_column + 1):
                    c = sheet.cell(row=total_row2, column=cc)
                    if isinstance(c, MergedCell):
                        continue
                    v = c.value
                    if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()) or (isinstance(v, str) and v.startswith("=")):
                        c.value = count_sv
                        placed = True
                        break

                # nếu không tìm được ô số thì bỏ qua (không crash)
        except:
            pass

        # ---------- 6) save ----------
        saved_path, warn, is_copy = _safe_save_wb(wb, file_path, class_path, "DanhSachLop_UPDATED")
        if saved_path is None:
            return False, warn or "❌ Không thể lưu Excel.", None

        if warn:
            return True, f"✅ Đã thêm SV vào Excel (bản copy). {warn}", saved_path

        return True, "✅ Đã thêm SV vào Excel.", saved_path

    except Exception as e:
        return False, f"❌ Lỗi khi cập nhật Excel: {e}", None


        # =========================
        # 6) SAVE
        # =========================
        saved_path, warn, is_copy = _safe_save_wb(wb, file_path, class_path, "DanhSachLop_UPDATED")

        # ✅ LOG ra terminal (để bắt bệnh)
        try:
            print("[ADD_SV] saved_path =", saved_path, "| warn =", warn, "| is_copy =", is_copy)
        except:
            pass

        if saved_path is None:
            return False, warn or "❌ Không thể lưu Excel.", None

        if warn:
            return True, f"✅ Đã thêm SV vào Excel (bản copy). {warn}", saved_path

        return True, "✅ Đã thêm SV vào Excel.", saved_path

    except Exception as e:
        return False, f"❌ Lỗi khi cập nhật Excel: {e}", None

def remove_student_from_excel(class_name: str, mssv: str):
    """
    Xóa sinh viên khỏi bảng Excel theo MSSV.
    Return: (ok, msg, saved_path)
    """
    class_path = os.path.join(ROOT_FOLDER, class_name)
    file_path = os.path.join(class_path, EXCEL_NAME)
    if not os.path.exists(file_path):
        return False, "⚠️ Lớp chưa có file Excel. Đã xóa SV khỏi app (roster.json), nhưng Excel chưa được cập nhật.", None

    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        header_row, cols = find_excel_header_in_sheet(sheet)
        if not header_row or not cols.get("mssv"):
            return False, "❌ Không tìm thấy header/cột MSSV trong Excel.", None

        target = str(mssv).strip().upper()
        row_to_del = None
        for r in range(header_row + 1, sheet.max_row + 1):
            mv = sheet.cell(row=r, column=cols["mssv"]).value
            ms = str(mv).strip() if mv else ""
            if not ms:
                continue
            if "tổng" in ms.lower():
                break
            if ms.upper() == target:
                row_to_del = r
                break

        if row_to_del is None:
            return False, f"⚠️ Không tìm thấy MSSV {mssv} trong Excel.", None

        sheet.delete_rows(row_to_del, 1)

        # renumber STT cho đẹp
        _renumber_stt(sheet, header_row, cols)

        saved_path, warn, is_copy = _safe_save_wb(wb, file_path, class_path, "DanhSachLop_UPDATED")
        if saved_path is None:
            return False, warn or "❌ Không thể lưu Excel.", None
        if warn:
            return True, f"✅ Đã xóa SV trong Excel (bản copy). {warn}", saved_path
        return True, "✅ Đã xóa SV trong Excel.", saved_path

    except Exception as e:
        return False, f"❌ Lỗi khi xóa SV trong Excel: {e}", None

def add_student_to_roster(class_name: str, student: dict):
    """
    Add student to roster.json (ưu tiên), không phụ thuộc excel.
    Return: (ok, msg)
    """
    df = load_roster(class_name, autosync_folders=False)
    mssv = str(student.get("MSSV", "")).strip()
    if not mssv:
        return False, "Vui lòng nhập MSSV."

    if df is None or df.empty:
        df = pd.DataFrame(columns=["MSSV", "Họ đệm", "Họ", "Tên", "Họ tên", "Giới tính", "Ngày sinh", "Lớp học"])

    exists = False
    try:
        exists = (df["MSSV"].astype(str).str.strip().str.upper() == mssv.upper()).any()
    except:
        exists = False

    if exists:
        return False, f"MSSV {mssv} đã tồn tại trong danh sách."

    # normalize
    ho_dem = str(student.get("Họ đệm", "")).strip()
    ten = str(student.get("Tên", "")).strip()
    full_name = f"{ho_dem} {ten}".strip()

    row = {
        "MSSV": mssv,
        "Họ đệm": ho_dem,
        "Họ": ho_dem,
        "Tên": ten,
        "Họ tên": full_name,
        "Giới tính": str(student.get("Giới tính", "")).strip(),
        "Ngày sinh": _cell_to_text(student.get("Ngày sinh", "")),
        "Lớp học": str(student.get("Lớp học", "")).strip(),
    }

    df2 = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df2["MSSV"] = df2["MSSV"].astype(str).str.strip()

    save_roster_json(class_name, df2)

    # create folders now
    class_path, img_base, _ = ensure_class_paths(class_name)
    os.makedirs(os.path.join(img_base, mssv), exist_ok=True)
    os.makedirs(os.path.join(class_path, BAD_IMG_DIR, mssv), exist_ok=True)

    return True, f"✅ Đã thêm SV {mssv} vào roster (app)."

def remove_student_from_roster(class_name: str, mssv: str):
    df = load_roster(class_name, autosync_folders=False)
    if df is None or df.empty:
        return False, "Danh sách rỗng."
    target = str(mssv).strip().upper()
    df2 = df[df["MSSV"].astype(str).str.strip().str.upper() != target].copy()
    if len(df2) == len(df):
        return False, f"Không tìm thấy MSSV {mssv} trong roster."
    save_roster_json(class_name, df2)
    return True, f"✅ Đã xóa SV {mssv} khỏi roster (app)."

# ==========================================================
# 6) CLASS STATS (ĐÚNG THEO EXCEL + COVERAGE TRAIN)
# ==========================================================
def count_student_folders(img_base: str):
    folders = []
    if os.path.exists(img_base):
        for d in os.listdir(img_base):
            p = os.path.join(img_base, d)
            if os.path.isdir(p):
                folders.append(d)
    return folders

def class_stats(class_name: str):
    class_path, img_base, bad_base = ensure_class_paths(class_name)

    # roster
    df_roster = load_roster(class_name, autosync_folders=False)
    roster_count = int(len(df_roster)) if df_roster is not None else 0

    # folders & images
    folders = count_student_folders(img_base)
    folder_count = len(folders)

    with_images = 0
    img_count = 0
    for d in folders:
        p = os.path.join(img_base, d)
        imgs = [f for f in os.listdir(p) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
        if len(imgs) > 0:
            with_images += 1
            img_count += len(imgs)

    bad_img_count = 0
    if os.path.exists(bad_base):
        for d in os.listdir(bad_base):
            p = os.path.join(bad_base, d)
            if os.path.isdir(p):
                bad_img_count += len([f for f in os.listdir(p) if f.lower().endswith((".png", ".jpg", ".jpeg"))])

    data = load_class_data(class_name)
    enc_count = len(data.get("encodings", []))
    uniq_ids = len(set([str(x) for x in data.get("ids", [])]))

    data_path = os.path.join(class_path, DATA_FILE_NAME)
    last_train = "—"
    if os.path.exists(data_path):
        last_train = datetime.fromtimestamp(os.path.getmtime(data_path)).strftime("%d/%m/%Y %H:%M")

    excel_path = os.path.join(class_path, EXCEL_NAME)
    has_excel = os.path.exists(excel_path)

    # sessions
    sessions_csv = os.path.join(class_path, HISTORY_DIR_NAME, SESSIONS_CSV)
    sessions_count = 0
    if os.path.exists(sessions_csv):
        try:
            sessions_count = len(pd.read_csv(sessions_csv, encoding="utf-8-sig"))
        except:
            sessions_count = 0

    return {
        "roster_count": roster_count,           # Tổng SV theo Excel/roster
        "folder_count": folder_count,           # Số folder hiện có
        "with_images": with_images,             # SV có ảnh (valid)
        "img_count": img_count,                 # Tổng ảnh valid
        "bad_img_count": bad_img_count,         # Ảnh lỗi (no-face)
        "enc_count": enc_count,                 # encodings (≈ ảnh valid)
        "uniq_ids": uniq_ids,                   # unique MSSV trong encodings
        "last_train": last_train,
        "has_excel": has_excel,
        "excel_path": excel_path,
        "sessions_count": sessions_count,
    }

# ==========================================================
# 7) DELETE (giữ folder để không gãy logic lớp)
# ==========================================================
def delete_class(class_name):
    path = os.path.join(ROOT_FOLDER, class_name)
    if not os.path.exists(path):
        return True

    try:
        shutil.rmtree(path)
        return True
    except Exception as e:
        print("[DELETE_CLASS] FAIL:", e)
        return False


def delete_student(class_name, mssv_to_delete: str, remove_folder: bool = False):
    """
    Xóa dữ liệu ảnh + encodings của SV (không xử lý roster/excel).
    """
    mssv_to_delete = str(mssv_to_delete).strip()

    # xóa encodings
    data = load_class_data(class_name)
    new_enc = []
    new_ids = []
    for enc, mssv in zip(data.get("encodings", []), data.get("ids", [])):
        if str(mssv).strip() != mssv_to_delete:
            new_enc.append(enc)
            new_ids.append(mssv)
    data["encodings"] = new_enc
    data["ids"] = new_ids
    save_class_data(class_name, data)

    class_path, img_base, bad_base = ensure_class_paths(class_name)
    sv_folder = os.path.join(img_base, mssv_to_delete)
    sv_bad = os.path.join(bad_base, mssv_to_delete)

    # Xóa ảnh trong folder (không xóa folder nếu remove_folder=False)
    if os.path.exists(sv_folder):
        try:
            for f in os.listdir(sv_folder):
                fp = os.path.join(sv_folder, f)
                if os.path.isfile(fp):
                    os.remove(fp)
        except:
            pass
    if os.path.exists(sv_bad):
        try:
            for f in os.listdir(sv_bad):
                fp = os.path.join(sv_bad, f)
                if os.path.isfile(fp):
                    os.remove(fp)
        except:
            pass

    if remove_folder:
        try:
            shutil.rmtree(sv_folder, ignore_errors=True)
        except:
            pass
        try:
            shutil.rmtree(sv_bad, ignore_errors=True)
        except:
            pass

    return True

# ==========================================================
# 8) TRAINING (ZIP / SINGLE) - GIỮ ẢNH LỖI
# ==========================================================
def _save_image_and_encode(img: Image.Image, save_path: str):
    # Save
    img.save(save_path)
    # Encode
    encs = face_recognition.face_encodings(face_recognition.load_image_file(save_path))
    if encs:
        return encs[0], None
    return None, "NO_FACE"

def train_single_student(class_name, mssv, uploaded_images):
    class_path, img_base, bad_base = ensure_class_paths(class_name)
    mssv = str(mssv).strip()

    sv_folder = os.path.join(img_base, mssv)
    sv_bad = os.path.join(bad_base, mssv)
    os.makedirs(sv_folder, exist_ok=True)
    os.makedirs(sv_bad, exist_ok=True)

    data = load_class_data(class_name)

    ok_count = 0
    bad_count = 0

    my_bar = st.progress(0, text="Đang xử lý ảnh...")
    total = len(uploaded_images)

    for i, file in enumerate(uploaded_images):
        try:
            img = Image.open(file)
            try:
                img = ImageOps.exif_transpose(img)
            except:
                pass
            if img.mode != "RGB":
                img = img.convert("RGB")
            if max(img.size) > 1200:
                img.thumbnail((1200, 1200))

            new_filename = f"{uuid.uuid4().hex}.jpg"
            final_path = os.path.join(sv_folder, new_filename)

            enc, err = _save_image_and_encode(img, final_path)
            if enc is not None:
                data["encodings"].append(enc)
                data["ids"].append(mssv)
                ok_count += 1
            else:
                # chuyển sang Anh_Loi
                bad_path = os.path.join(sv_bad, new_filename)
                try:
                    shutil.move(final_path, bad_path)
                except:
                    pass
                bad_count += 1

        except:
            bad_count += 1

        my_bar.progress((i + 1) / total, text=f"Đang xử lý ảnh {i+1}/{total}")

    my_bar.empty()
    save_class_data(class_name, data)

    if ok_count > 0:
        return True, f"✅ {mssv}: {ok_count} ảnh hợp lệ • {bad_count} ảnh lỗi (đã chuyển vào {BAD_IMG_DIR})"
    return False, f"⚠️ {mssv}: Không tìm thấy khuôn mặt trong ảnh! ({bad_count} ảnh lỗi)"

def _derive_mssv_from_relpath(rel_path: str, roster_set: set):
    parts = [p for p in rel_path.replace("\\", "/").split("/") if p and not p.startswith(".")]
    if not parts:
        return None

    # Nếu có roster_set: tìm phần nào match
    for p in parts[:-1]:
        if p.strip() in roster_set:
            return p.strip()

    # Fallback: lấy folder cấp 1 nếu không phải __MACOSX
    if len(parts) >= 2:
        if parts[0] != "__MACOSX":
            return parts[0].strip()

    # Fallback cuối: tên file
    return os.path.splitext(parts[-1])[0].strip()

def train_class_with_zip(class_name, excel_file, zip_file, reset_encodings: bool = False):
    class_path, img_base, bad_base = ensure_class_paths(class_name)

    # save excel
    if excel_file:
        with open(os.path.join(class_path, EXCEL_NAME), "wb") as f:
            f.write(excel_file.getbuffer())

    # sync roster folders theo Excel (nếu có)
    df_roster = load_roster(class_name, autosync_folders=True)
    roster_set = set(df_roster["MSSV"].astype(str).str.strip().tolist()) if not df_roster.empty else set()

    # load or reset data
    data = load_class_data(class_name)
    if reset_encodings:
        data = {"encodings": [], "ids": []}

    if not zip_file:
        save_class_data(class_name, data)
        roster_n = len(df_roster) if not df_roster.empty else 0
        return True, f"✅ Đã cập nhật Excel. Tổng SV theo Excel: {roster_n}", {
            "total_images": 0, "valid": 0, "no_face": 0, "error": 0, "unknown": 0, "roster": roster_n
        }

    extract_path = os.path.join(class_path, "Temp_Extract")
    if os.path.exists(extract_path):
        shutil.rmtree(extract_path, ignore_errors=True)
    os.makedirs(extract_path, exist_ok=True)

    try:
        with zipfile.ZipFile(zip_file, "r") as z:
            z.extractall(extract_path)
    except:
        return False, "❌ File ZIP lỗi!", None

    # collect images
    all_images = []
    for root, _, files in os.walk(extract_path):
        for file in files:
            if file.lower().endswith((".png", ".jpg", ".jpeg")):
                # bỏ __MACOSX
                if "__macosx" in root.lower():
                    continue
                all_images.append(os.path.join(root, file))

    total = len(all_images)
    valid = 0
    no_face = 0
    err_cnt = 0
    unknown_cnt = 0

    zip_bar = st.progress(0, text="Đang giải nén & training...")

    for i, file_path in enumerate(all_images):
        try:
            rel = os.path.relpath(file_path, extract_path)
            mssv = _derive_mssv_from_relpath(rel, roster_set)
            if not mssv:
                mssv = "UNKNOWN"
            mssv = str(mssv).strip()

            if roster_set and mssv not in roster_set:
                unknown_cnt += 1

            sv_folder = os.path.join(img_base, mssv)
            sv_bad = os.path.join(bad_base, mssv)
            os.makedirs(sv_folder, exist_ok=True)
            os.makedirs(sv_bad, exist_ok=True)

            img = Image.open(file_path)
            try:
                img = ImageOps.exif_transpose(img)
            except:
                pass
            if img.mode != "RGB":
                img = img.convert("RGB")
            if max(img.size) > 1200:
                img.thumbnail((1200, 1200))

            new_filename = f"{uuid.uuid4().hex}.jpg"
            final_path = os.path.join(sv_folder, new_filename)

            enc, err = _save_image_and_encode(img, final_path)
            if enc is not None:
                data["encodings"].append(enc)
                data["ids"].append(mssv)
                valid += 1
            else:
                bad_path = os.path.join(sv_bad, new_filename)
                try:
                    shutil.move(final_path, bad_path)
                except:
                    pass
                no_face += 1

        except:
            err_cnt += 1

        if total > 0:
            zip_bar.progress((i + 1) / total, text=f"Training {i+1}/{total}")

    zip_bar.empty()
    shutil.rmtree(extract_path, ignore_errors=True)
    save_class_data(class_name, data)

    roster_n = len(df_roster) if not df_roster.empty else 0
    msg = f"✅ Training xong! Ảnh: {total} • Hợp lệ: {valid} • Ảnh lỗi: {no_face} • Lỗi đọc: {err_cnt}"
    if roster_set:
        msg += f" • Không khớp Excel: {unknown_cnt}"
    return True, msg, {
        "total_images": total,
        "valid": valid,
        "no_face": no_face,
        "error": err_cnt,
        "unknown": unknown_cnt,
        "roster": roster_n
    }

# ==========================================================
# 9) EXCEL ATTENDANCE (FIX WIDTH + STYLE ĐẸP)
# ==========================================================
def _get_col_width(sheet, col_letter: str):
    try:
        dim = sheet.column_dimensions.get(col_letter)
        if dim and dim.width:
            return float(dim.width)
    except:
        pass
    # defaultColWidth (nếu có)
    try:
        if sheet.sheet_format and sheet.sheet_format.defaultColWidth:
            return float(sheet.sheet_format.defaultColWidth)
    except:
        pass
    return None

def _replicate_vertical_merge(sheet, template_col: int, target_col: int, header_row: int):
    # Nếu cột template đang nằm trong merge theo chiều dọc ở header_row, replicate cho target_col
    for rng in sheet.merged_cells.ranges:
        if rng.min_col <= template_col <= rng.max_col and rng.min_row <= header_row <= rng.max_row:
            # replicate vertical span, 1 column
            if rng.max_row > rng.min_row:
                try:
                    sheet.merge_cells(
                        start_row=rng.min_row,
                        start_column=target_col,
                        end_row=rng.max_row,
                        end_column=target_col,
                    )
                except:
                    pass
            return

def _find_or_create_attendance_col(sheet, header_row: int, mssv_col: int, code: str):
    # 1) tìm tất cả cột DD
    dd_cols = []
    last_header_col = mssv_col
    for c in range(1, sheet.max_column + 1):
        cell = sheet.cell(row=header_row, column=c)
        if isinstance(cell, MergedCell):
            continue
        val = str(cell.value).strip() if cell.value is not None else ""
        if val:
            last_header_col = max(last_header_col, c)
        if val.startswith("DD "):
            dd_cols.append(c)

    # 2) nếu đã có cột cho code
    for c in dd_cols:
        cell = sheet.cell(row=header_row, column=c)
        if str(cell.value).strip() == code:
            template_col = dd_cols[-1] if dd_cols else (c - 1)
            return c, template_col, False

    # 3) tạo mới: ưu tiên sau cột DD cuối, nếu chưa có thì sau last_header_col
    start = dd_cols[-1] if dd_cols else last_header_col
    target_col = None
    template_col = dd_cols[-1] if dd_cols else start

    # scan sang phải tìm ô trống
    for c in range(start + 1, sheet.max_column + 30):
        cell = sheet.cell(row=header_row, column=c)
        if isinstance(cell, MergedCell):
            continue
        val = cell.value
        if val is None or str(val).strip() == "":
            target_col = c
            cell.value = code
            break

    if target_col is None:
        target_col = sheet.max_column + 1
        sheet.cell(row=header_row, column=target_col).value = code

    return target_col, template_col, True

def process_attendance_excel(class_name, present_ids, attendance_date=None, save_copy_if_open=True):
    class_path = os.path.join(ROOT_FOLDER, class_name)
    file_path = os.path.join(class_path, EXCEL_NAME)
    if not os.path.exists(file_path):
        return None, "Thiếu file Excel!", 0, [], [], None

    if attendance_date is None:
        attendance_date = datetime.now().date()

    code = f"DD {attendance_date.strftime('%d/%m/%y')}"
    present_norm = set([str(x).strip().upper() for x in present_ids])

    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active

        header_row, cols = find_excel_header_in_sheet(sheet)
        if not header_row or not cols.get("mssv"):
            return None, "Không tìm thấy cột Mã SV!", 0, [], [], None

        # tạo/tìm cột DD
        target_col, template_col, created = _find_or_create_attendance_col(sheet, header_row, cols["mssv"], code)
        t_letter = get_column_letter(target_col)

        # Lấy cột mẫu bên trái để copy style (Cột Ghi chú hoặc Ký tên)
        # Tìm lùi về trước đến khi gặp cột có độ rộng > 0
        left_col_idx = max(1, target_col - 1)
        
        # === 1. FIX ĐỘ RỘNG (WIDTH) ===
        # Lấy độ rộng của cột bên trái áp vào cột mới
        left_letter = get_column_letter(left_col_idx)
        if left_letter in sheet.column_dimensions:
            w = sheet.column_dimensions[left_letter].width
            if w:
                sheet.column_dimensions[t_letter].width = w
            else:
                sheet.column_dimensions[t_letter].width = 15.0 # Mặc định nếu ko tìm thấy
        else:
            sheet.column_dimensions[t_letter].width = 15.0

        # === 2. FIX CHIỀU CAO HEADER (GỘP Ô) ===
        # Kiểm tra xem ô bên trái có đang merge dọc không (ví dụ Header chiếm 2 dòng)
        # Nếu có -> Merge cột điểm danh y hệt vậy
        is_merged = False
        for rng in sheet.merged_cells.ranges:
            # Nếu cột bên trái nằm trong 1 vùng merge và vùng đó dính dáng tới header_row
            if rng.min_col <= left_col_idx <= rng.max_col and rng.min_row <= header_row <= rng.max_row:
                # Nếu vùng merge này cao hơn 1 dòng -> Copy merge sang cột mới
                if rng.max_row > rng.min_row:
                    try:
                        sheet.merge_cells(
                            start_row=rng.min_row, 
                            start_column=target_col, 
                            end_row=rng.max_row, 
                            end_column=target_col
                        )
                        is_merged = True
                    except:
                        pass
                break # Chỉ cần tìm thấy 1 vùng merge chứa header là đủ

        # Định nghĩa viền
        thin = Side(style="thin", color="000000")
        full_border = Border(top=thin, left=thin, right=thin, bottom=thin)

        # === 3. COPY STYLE HEADER ===
        head_cell = sheet.cell(row=header_row, column=target_col)
        # Lấy ô master của cột bên trái (trong trường hợp nó bị merge) để lấy style chuẩn
        left_head_cell = get_real_cell(sheet, header_row, left_col_idx)

        # Copy Font & Alignment y chang ô bên cạnh
        if left_head_cell:
            if left_head_cell.font:
                head_cell.font = copy(left_head_cell.font)
            if left_head_cell.alignment:
                head_cell.alignment = copy(left_head_cell.alignment)
            if left_head_cell.fill:
                head_cell.fill = copy(left_head_cell.fill)
        else:
            # Fallback nếu không copy được
            head_cell.font = Font(name="Arial", bold=True, size=10)
            head_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Ghi đè lại nội dung và viền
        head_cell.value = code
        head_cell.border = full_border
        
        # Nếu không có màu nền -> tô xanh nhạt cho dễ nhìn
        if not (head_cell.fill and head_cell.fill.patternType):
             head_cell.fill = PatternFill("solid", fgColor="E8F0FE")

        # mark attendance
        real_total = 0
        ls_present = []
        ls_absent = []

        # === VÒNG LẶP DUYỆT TỪNG SINH VIÊN ===
        for r in range(header_row + 1, sheet.max_row + 1):
            # Kiểm tra dòng này có phải header phụ hay merge không để né
            if is_merged and r <= (sheet.merged_cells.ranges[0].max_row if sheet.merged_cells.ranges else header_row):
                 continue

            mssv_cell = sheet.cell(row=r, column=cols["mssv"])
            # Nếu ô MSSV bị merge (dòng rỗng do merge) -> skip
            if isinstance(mssv_cell, MergedCell):
                continue

            mssv_val = str(mssv_cell.value).strip() if mssv_cell.value else ""
            if not mssv_val:
                continue
            if "tổng" in mssv_val.lower():
                break

            real_total += 1
            ho_val = sheet.cell(row=r, column=cols["ho"]).value if cols.get("ho") else ""
            ten_val = sheet.cell(row=r, column=cols["ten"]).value if cols.get("ten") else ""
            full_name = f"{str(ho_val).strip()} {str(ten_val).strip()}".strip()

            res_cell = sheet.cell(row=r, column=target_col)
            if isinstance(res_cell, MergedCell):
                continue

            # Kẻ ô
            res_cell.border = full_border 

            # === 4. COPY STYLE DÒNG DỮ LIỆU ===
            # Lấy ô bên trái làm mẫu
            left_cell = sheet.cell(row=r, column=left_col_idx)
            
            # Copy Alignment (Căn giữa/trái/phải y hệt ô bên cạnh)
            if left_cell.alignment:
                res_cell.alignment = copy(left_cell.alignment)
            else:
                res_cell.alignment = Alignment(horizontal="center", vertical="center")

            # Lấy Font mẫu để giữ Size và Font Name
            base_font_name = "Arial"
            base_font_size = 11
            if left_cell.font:
                base_font_name = left_cell.font.name
                base_font_size = left_cell.font.size

            if mssv_val.upper() in present_norm:
                res_cell.value = "✓"
                # Tạo font mới dựa trên font cũ nhưng đổi màu
                res_cell.font = Font(name=base_font_name, size=base_font_size, bold=True, color="008000")
                res_cell.fill = PatternFill("solid", fgColor="ECFDF3") 
                ls_present.append(f"- {mssv_val} - {full_name}")
            else:
                res_cell.value = "X"
                res_cell.font = Font(name=base_font_name, size=base_font_size, bold=True, color="B91C1C")
                res_cell.fill = PatternFill("solid", fgColor="FEF2F2") 
                ls_absent.append(f"- {mssv_val} - {full_name}")

        # save workbook
        saved_path = file_path
        try:
            wb.save(file_path)
        except Exception:
            if save_copy_if_open:
                saved_path = os.path.join(class_path, f"DanhSachLop_UPDATED_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                try:
                    wb.save(saved_path)
                except Exception as e:
                    return None, f"❌ Không thể lưu Excel: {e}", 0, [], [], None
                return saved_path, "⚠️ File Excel đang mở. Đã lưu bản copy mới.", real_total, ls_present, ls_absent, code
            return None, "❌ File Excel đang mở!", 0, [], [], None

        return saved_path, None, real_total, ls_present, ls_absent, code

    except Exception as e:
        return None, str(e), 0, [], [], None

# ==========================================================
# 10) ATTENDANCE SCAN + HISTORY + STATS
# ==========================================================
def scan_images(class_name, uploaded_files, tolerance: float):
    data = load_class_data(class_name)
    if not data.get("encodings"):
        return None, "Chưa có dữ liệu training cho lớp này!"

    known_enc = np.array(data["encodings"])
    known_ids = [str(x) for x in data["ids"]]

    found = set()
    hits = {}
    min_dist = {}
    faces_total = 0

    bar = st.progress(0, text="Đang phân tích ảnh...")
    total = len(uploaded_files)

    for i, f in enumerate(uploaded_files):
        try:
            img = Image.open(f)
            try:
                img = ImageOps.exif_transpose(img)
            except:
                pass
            img = img.convert("RGB")
            arr = np.array(img)

            encs = face_recognition.face_encodings(arr)
            faces_total += len(encs)

            for e in encs:
                dists = face_recognition.face_distance(known_enc, e)
                best_i = int(np.argmin(dists))
                best_dist = float(dists[best_i])
                best_id = known_ids[best_i]

                if best_dist <= float(tolerance):
                    found.add(best_id)
                    hits[best_id] = hits.get(best_id, 0) + 1
                    min_dist[best_id] = min(min_dist.get(best_id, 999.0), best_dist)

        except:
            pass

        if total > 0:
            bar.progress((i + 1) / total, text=f"Đang xử lý {i+1}/{total}")

    bar.empty()
    payload = {
        "class_name": class_name,
        "tolerance": float(tolerance),
        "scanned_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "images_count": total,
        "faces_total": faces_total,
        "found_ids": sorted(list(found)),
        "hits": hits,
        "min_dist": min_dist,
    }
    return payload, None

def ensure_history_dir(class_name):
    hist_dir = os.path.join(ROOT_FOLDER, class_name, HISTORY_DIR_NAME)
    os.makedirs(hist_dir, exist_ok=True)
    return hist_dir

def save_session_history(class_name, session_payload: dict):
    hist_dir = ensure_history_dir(class_name)
    session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_payload = dict(session_payload)
    session_payload["session_id"] = session_id

    # json detail
    json_path = os.path.join(hist_dir, f"{session_id}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(session_payload, f, ensure_ascii=False, indent=2)

    # csv summary
    csv_path = os.path.join(hist_dir, SESSIONS_CSV)
    row = {
        "session_id": session_id,
        "datetime": session_payload.get("scanned_at", ""),
        "class": class_name,
        "tolerance": session_payload.get("tolerance", ""),
        "date_mark": session_payload.get("date_mark", ""),
        "images": session_payload.get("images_count", 0),
        "faces": session_payload.get("faces_total", 0),
        "recognized_unique": len(session_payload.get("found_ids", [])),
        "present_final": len(session_payload.get("final_ids", [])),
        "excel_updated": session_payload.get("excel_updated", False),
    }
    df = pd.DataFrame([row])
    if os.path.exists(csv_path):
        df.to_csv(csv_path, mode="a", index=False, header=False, encoding="utf-8-sig")
    else:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    return json_path, csv_path

def load_sessions_summary(class_name):
    csv_path = os.path.join(ROOT_FOLDER, class_name, HISTORY_DIR_NAME, SESSIONS_CSV)
    if not os.path.exists(csv_path):
        return pd.DataFrame()
    try:
        return pd.read_csv(csv_path, encoding="utf-8-sig")
    except:
        return pd.DataFrame()

def compute_attendance_rate(class_name: str):
    hist_dir = os.path.join(ROOT_FOLDER, class_name, HISTORY_DIR_NAME)
    files = sorted(glob.glob(os.path.join(hist_dir, "*.json")))
    if not files:
        return pd.DataFrame(), 0

    df_roster = load_roster(class_name, autosync_folders=False)
    if df_roster.empty:
        return pd.DataFrame(), 0

    ids = df_roster["MSSV"].astype(str).tolist()
    cnt = {m: 0 for m in ids}
    total_sessions = 0

    for fp in files:
        try:
            with open(fp, "r", encoding="utf-8") as f:
                payload = json.load(f)
            final_ids = payload.get("final_ids", [])
            if final_ids is None:
                continue
            total_sessions += 1
            for m in final_ids:
                m = str(m).strip()
                if m in cnt:
                    cnt[m] += 1
        except:
            pass

    out = df_roster[["MSSV", "Họ tên"]].copy()
    out["Có mặt"] = out["MSSV"].map(cnt).fillna(0).astype(int)
    out["Tổng buổi"] = total_sessions
    out["Tỷ lệ"] = out["Có mặt"] / total_sessions if total_sessions > 0 else np.nan
    out = out.sort_values(["Tỷ lệ", "Có mặt"], ascending=[True, True]).reset_index(drop=True)
    return out, total_sessions

# ==========================================================
# 11) BÁO CÁO - ĐỌC EXCEL CHO ĐẸP
# ==========================================================
def read_excel_for_report(class_name):
    class_path = os.path.join(ROOT_FOLDER, class_name)
    path = os.path.join(class_path, EXCEL_NAME)
    if not os.path.exists(path):
        return None, "Chưa có file Excel."

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        header_row, _ = find_excel_header_in_sheet(sheet)
        if not header_row:
            df = pd.read_excel(path, engine="openpyxl")
            df = df.dropna(how="all").dropna(how="all", axis=1)
            return df, None

        df = pd.read_excel(path, engine="openpyxl", header=header_row - 1)
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]
        df = df.dropna(how="all")
        return df, None
    except Exception as e:
        return None, str(e)

# ==========================================================
# 12) SIDEBAR (UI: Online xanh, Offline đỏ + thời gian có giây)
# ==========================================================
with st.sidebar:
    if os.path.exists(SIDEBAR_LOGO_PATH):
        c1, c2, c3 = st.columns([1, 1.2, 1])
        with c2:
            st.image(SIDEBAR_LOGO_PATH, width=64)

    st.markdown("### IUH Smart Attendance")
    st.caption("Điểm danh bằng nhận diện khuôn mặt")

    labels = list(PAGE_LABELS.values())
    current_label = PAGE_LABELS.get(st.session_state.menu_selection, "🏠 Trang Chủ")
    selected_label = st.radio("Điều hướng", labels, index=labels.index(current_label), key="menu_selection_label")
    st.session_state.menu_selection = LABEL_TO_PAGE[selected_label]
    selected = st.session_state.menu_selection

    st.markdown("---")

    # ✅ status pill online/offline
    if "ONLINE" in MODE:
        st.markdown(
            "<div class='status-pill online'>🟢 ONLINE <span class='status-sub'>(Cloud)</span></div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<div class='status-pill offline'>🔴 OFFLINE <span class='status-sub'>(Local)</span></div>",
            unsafe_allow_html=True,
        )

    # ✅ time with seconds
    st.markdown(
        f"<span class='badge pulse'>📅 {datetime.now().strftime('%d/%m/%Y')} • ⏱ {datetime.now().strftime('%H:%M:%S')}</span>",
        unsafe_allow_html=True,
    )

# ==========================================================
# 13) PAGES
# ==========================================================
if selected == "Trang Chủ":
    with bordered_container():
        c1, c2 = st.columns([0.42, 0.58])
        with c1:
            if os.path.exists(MAIN_LOGO_PATH):
                show_centered_image(MAIN_LOGO_PATH, width=230)
        with c2:
            st.markdown("## IUH SMART ATTENDANCE")
            st.markdown("<div class='small-muted'>Điểm danh nhận diện khuôn mặt • IUH • Phiên bản nâng cấp UI + Logic</div>", unsafe_allow_html=True)
            st.write("")
            st.markdown(
                "<span class='badge'>✅ Roster theo Excel</span> "
                "<span class='badge'>🗂️ Lưu lịch sử</span> "
                "<span class='badge'>✨ UI animation</span>",
                unsafe_allow_html=True,
            )

    st.write("")
    classes = get_list_classes()
    total_cl = len(classes)

    # ✅ add icons to metrics
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("🏫 Tổng lớp", total_cl)
    with m2:
        st.metric("📅 Hôm nay", datetime.now().strftime("%d/%m"))
    with m3:
        st.metric("☁️ Chế độ", "Cloud" if "ONLINE" in MODE else "Offline")
    with m4:
        st.metric("✅ Trạng thái", "OK")

    st.write("")
    with bordered_container():
        section_header("🚀 Truy cập nhanh", "Chọn tác vụ để thao tác nhanh", "Quick")
        q1, q2 = st.columns(2)
        with q1:
            st.markdown("**📦 Tạo lớp / cập nhật dữ liệu**")
            st.caption("Upload Excel và/hoặc ZIP ảnh sinh viên.")
            st.button("Tạo / cập nhật lớp", on_click=navigate_to, args=("Quản Lý Lớp & SV",))
        with q2:
            st.markdown("**📸 Bắt đầu điểm danh**")
            st.caption("Upload ảnh lớp/camera → quét → chỉnh → lưu Excel.")
            st.button("Điểm danh ngay", on_click=navigate_to, args=("Điểm Danh",))

    if classes:
        st.write("")
        with bordered_container():
            section_header("📌 Tổng quan lớp", "Đúng theo Excel/roster + coverage training", "Overview")
            rows = []
            for c in classes:
                s = class_stats(c)
                missing_train = max(0, s["roster_count"] - s["with_images"]) if s["roster_count"] else None
                rows.append({
                    "Lớp": c,
                    "SV (Roster)": s["roster_count"],
                    "SV có ảnh": s["with_images"],
                    "Thiếu ảnh": missing_train if missing_train is not None else "—",
                    "Encodings": s["enc_count"],
                    "Ảnh lỗi": s["bad_img_count"],
                    "Buổi DD": s["sessions_count"],
                    "Excel": "✅" if s["has_excel"] else "❌",
                    "Cập nhật train": s["last_train"],
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, height=320)

elif selected == "Quản Lý Lớp & SV":
    st.markdown("## 📂 Quản Lý Dữ Liệu")

    tab1, tab2 = st.tabs(["📦 Lớp Học", "👤 Sinh Viên"])

    with tab1:
        with bordered_container():
            section_header("Khởi tạo / cập nhật lớp", "Upload Excel và/hoặc ZIP ảnh sinh viên", "DATA")
            c1, c2 = st.columns(2)
            with c1:
                n_class = st.text_input("Mã lớp", placeholder="VD: 420300305601", key="input_create_class_name")
            with c2:
                f_excel = st.file_uploader("File Excel (.xlsx)", type=["xlsx"], key="input_create_class_excel")
            f_zip = st.file_uploader("File ảnh (ZIP)", type=["zip"], key="input_create_class_zip")

            c3, c4 = st.columns(2)
            with c3:
                reset_enc = st.checkbox("Reset encodings trước khi train (tránh trùng)", value=False)
            with c4:
                st.caption("Ảnh không nhận diện được sẽ chuyển vào thư mục **Anh_Loi** để xem lại.")

            if st.button("🚀 Xử lý dữ liệu", type="primary"):
                if not n_class:
                    st.info("Vui lòng nhập mã lớp.")
                else:
                    with st.status("Đang xử lý dữ liệu lớp...", expanded=True) as status:
                        status.update(label="Bước 1/3: Lưu Excel & đồng bộ roster...", state="running")
                        ok, msg, detail = train_class_with_zip(n_class, f_excel, f_zip, reset_encodings=reset_enc)
                        if ok:
                            status.update(label="Bước 2/3: Hoàn tất training & lưu dữ liệu...", state="running")
                            time.sleep(30)
                            status.update(label="Bước 3/3: Done ✅", state="complete")
                            st.toast("Xử lý xong!", icon="✅")
                            st.success(msg)
                            if detail:
                                a1, a2, a3, a4, a5 = st.columns(5)
                                with a1:
                                    st.metric("SV (Excel)", detail.get("roster", 0))
                                with a2:
                                    st.metric("Ảnh trong ZIP", detail.get("total_images", 0))
                                with a3:
                                    st.metric("Ảnh hợp lệ", detail.get("valid", 0))
                                with a4:
                                    st.metric("Ảnh lỗi", detail.get("no_face", 0))
                                with a5:
                                    st.metric("Không khớp Excel", detail.get("unknown", 0))
                            st.balloons()
                            time.sleep(30)
                            for k in ["input_create_class_name", "input_create_class_excel", "input_create_class_zip"]:
                                if k in st.session_state:
                                    del st.session_state[k]
                            st.rerun()
                        else:
                            status.update(label="Lỗi ❌", state="error")
                            st.error(msg)

        with bordered_container():
            section_header("Đồng bộ folder theo roster", "Tạo đủ folder SV theo danh sách (không làm mất ảnh)", "SYNC")
            cl = get_list_classes()
            if not cl:
                st.info("Chưa có lớp.")
            else:
                sclass = st.selectbox("Chọn lớp", cl, key="sync_class")
                if st.button("🔄 Đồng bộ ngay", type="primary"):
                    n = sync_roster_folders(sclass)
                    st.toast("Đồng bộ xong!", icon="✅")
                    st.success(f"✅ Đã đồng bộ folder theo roster. Tổng SV: {n}")
                    time.sleep(30)
                    st.rerun()

        with bordered_container():
            section_header("Xóa lớp", "Xóa toàn bộ dữ liệu lớp (không thể khôi phục)", "DANGER")
            l_del = get_list_classes()
            if l_del:
                cd = st.selectbox("Chọn lớp cần xóa", l_del, key="del_class_sel")
                if st.button("🗑️ Xóa vĩnh viễn", type="secondary"):
                    ok = delete_class(cd)
                    if ok:
                        st.toast("Đã xóa lớp!", icon="🗑️")
                        st.session_state.clear()
                        st.rerun()
                else:
                        st.error("Không thể xóa lớp (folder còn tồn tại / đang bị khóa).")

            else:
                st.info("Hiện chưa có lớp nào.")

    with tab2:
        cl = get_list_classes()
        if not cl:
            with bordered_container():
                section_header("Sinh viên", "Chưa có lớp để quản lý", "EMPTY")
                st.info("Bạn hãy tạo lớp trước ở tab **Lớp Học**.")
        else:
            sl = st.selectbox("Chọn lớp", cl, key="student_class")
            df_roster = load_roster(sl, autosync_folders=True)
            s = class_stats(sl)

            m1, m2, m3, m4, m5 = st.columns(5)
            with m1:
                st.metric("👥 Sĩ số (Roster)", s["roster_count"])
            with m2:
                st.metric("🖼️ SV có ảnh", s["with_images"])
            with m3:
                st.metric("🧬 Encodings", s["enc_count"])
            with m4:
                st.metric("⚠️ Ảnh lỗi", s["bad_img_count"])
            with m5:
                st.metric("📌 Buổi DD", s["sessions_count"])

            # ✅ thêm chế độ thêm/xóa SV
            mode = st.radio(
                "Chế độ",
                ["➕/➖ Thêm & Xóa SV", "Quản lý ảnh theo SV", "Danh sách Excel"],
                horizontal=True
            )

            if mode == "➕/➖ Thêm & Xóa SV":
                with bordered_container():
                    section_header("➕ Thêm sinh viên mới", "Thêm vào roster + tạo folder + (nếu có) cập nhật Excel", "ADD")
                    with st.form("add_student_form", clear_on_submit=False):
                        cA, cB = st.columns(2)
                        with cA:
                            new_mssv = st.text_input("Mã sinh viên (MSSV)*", placeholder="VD: SV041")
                            new_ho = st.text_input("Họ đệm*", placeholder="VD: Nguyễn Văn")
                            new_gt = st.selectbox("Giới tính*", ["Nam", "Nữ", "Khác"])
                        with cB:
                            new_ten = st.text_input("Tên*", placeholder="VD: A")
                            new_ns = st.date_input("Ngày sinh*", value=date(2005, 1, 1))
                            new_lh = st.text_input("Lớp học*", placeholder="VD: DHKHMT19ATT")

                        new_imgs = st.file_uploader("Ảnh khuôn mặt (tùy chọn, để train ngay)", accept_multiple_files=True, key="new_sv_imgs")

                        submitted = st.form_submit_button("✅ Thêm sinh viên", type="primary")
                        if submitted:
                            # validate
                            if not new_mssv.strip() or not new_ho.strip() or not new_ten.strip() or not new_lh.strip():
                                st.error("Vui lòng nhập đủ: MSSV, Họ đệm, Tên, Lớp học.")
                            else:
                                stu = {
                                    "MSSV": new_mssv.strip(),
                                    "Họ đệm": new_ho.strip(),
                                    "Tên": new_ten.strip(),
                                    "Giới tính": new_gt.strip(),
                                    "Ngày sinh": new_ns,
                                    "Lớp học": new_lh.strip(),
                                }

                                ok1, msg1 = add_student_to_roster(sl, stu)
                                if not ok1:
                                    st.error(msg1)
                                else:
                                    # update excel if exists
                                    ok2, msg2, saved_excel = add_student_to_excel(sl, stu)
                                    if ok2:
                                        st.success(msg1)
                                        st.info(msg2)
                                        if saved_excel and os.path.exists(saved_excel) and saved_excel != os.path.join(ROOT_FOLDER, sl, EXCEL_NAME):
                                            with open(saved_excel, "rb") as f:
                                                st.download_button("⬇️ Tải Excel đã cập nhật (bản copy)", data=f, file_name=os.path.basename(saved_excel))
                                    else:
                                        # vẫn ok vì roster đã thêm
                                        st.success(msg1)
                                        st.warning(msg2)

                                    # train images if provided
                                    if new_imgs:
                                        ok_tr, msg_tr = train_single_student(sl, new_mssv.strip(), new_imgs)
                                        if ok_tr:
                                            st.toast("Train ảnh xong!", icon="✅")
                                            st.success(msg_tr)
                                        else:
                                            st.warning(msg_tr)

                                    st.toast("Đã thêm sinh viên!", icon="🎉")
                                    time.sleep(30)
                                    st.rerun()

                with bordered_container():
                    section_header("🗑️ Xóa sinh viên", "Xóa khỏi roster + (nếu có) xóa dòng trong Excel + xóa dữ liệu ảnh/encodings", "DELETE")

                    if df_roster.empty:
                        st.info("Chưa có danh sách sinh viên.")
                    else:
                        roster_ids = df_roster["MSSV"].astype(str).tolist()
                        del_mssv = st.selectbox("Chọn MSSV cần xóa", roster_ids, key="del_sv_sel")
                        # show detail
                        row = df_roster[df_roster["MSSV"].astype(str) == str(del_mssv)].head(1)
                        if not row.empty:
                            info = row.iloc[0].to_dict()
                            st.markdown(
                                f"**Thông tin:** {info.get('MSSV','')} • {info.get('Họ tên','')} • {info.get('Giới tính','')} • {info.get('Ngày sinh','')} • {info.get('Lớp học','')}"
                            )

                        c1, c2 = st.columns(2)
                        with c1:
                            remove_folder = st.checkbox("Xóa luôn folder ảnh (không khuyến nghị)", value=False, key="rm_folder_del_chk")
                        with c2:
                            confirm = st.checkbox("Xác nhận xóa", value=False)

                        if st.button("🗑️ Xóa sinh viên", type="secondary"):
                            if not confirm:
                                st.warning("Hãy tick **Xác nhận xóa** trước.")
                            else:
                                # 1) remove from roster
                                okr, msgr = remove_student_from_roster(sl, del_mssv)
                                if okr:
                                    st.success(msgr)
                                else:
                                    st.warning(msgr)

                                # 2) remove from excel
                                oke, msge, saved_excel = remove_student_from_excel(sl, del_mssv)
                                if oke:
                                    st.info(msge)
                                    if saved_excel and os.path.exists(saved_excel) and saved_excel != os.path.join(ROOT_FOLDER, sl, EXCEL_NAME):
                                        with open(saved_excel, "rb") as f:
                                            st.download_button("⬇️ Tải Excel đã cập nhật (bản copy)", data=f, file_name=os.path.basename(saved_excel))
                                else:
                                    st.warning(msge)

                                # 3) remove images + encodings
                                delete_student(sl, del_mssv, remove_folder=remove_folder)

                                st.toast("Đã xóa sinh viên!", icon="🗑️")
                                time.sleep(30)
                                st.rerun()

            elif mode == "Quản lý ảnh theo SV":
                with bordered_container():
                    section_header("Quản lý ảnh", "Xem / thêm / xóa ảnh (folder luôn tồn tại theo roster)", "IMAGES")
                    with st.expander("➕ Thêm sinh viên mới (ghi vào Excel)", expanded=False):
                        c1, c2, c3 = st.columns([0.3, 0.35, 0.35])
                    with c1:
                        new_mssv = st.text_input("MSSV mới", key="add_sv_manage_mssv")
                    with c2:
                        new_ho = st.text_input("Họ / Họ đệm", key="add_sv_manage_ho")
                    with c3:
                        new_ten = st.text_input("Tên", key="add_sv_manage_ten")

            save_copy_if_open2 = st.checkbox("Nếu Excel đang mở → lưu bản copy", value=True, key="add_sv_manage_savecopy")

            if st.button("✅ Tạo sinh viên", type="primary", key="btn_add_sv_manage"):
                ok, msg, saved = add_student_to_excel(
                    sl, new_mssv, new_ho, new_ten, save_copy_if_open=save_copy_if_open2
                )

                # ✅ DEBUG: xem app lưu file nào
                st.write("📌 File Excel app vừa lưu:", saved)
                if saved and os.path.exists(saved):
                    st.write("🕒 Modified:", datetime.fromtimestamp(os.path.getmtime(saved)).strftime("%d/%m/%Y %H:%M:%S"))
                    st.write("📦 Size:", os.path.getsize(saved), "bytes")

                if ok:
                    st.success(msg)
                    st.toast("Đã thêm sinh viên!", icon="✅")
                    time.sleep(30)
                    st.rerun()
                else:
                    st.warning(msg)
                    df_roster = load_roster(sl, autosync_folders=True)
                    roster_ids = df_roster["MSSV"].astype(str).tolist() if not df_roster.empty else []
                    base_path = os.path.join(ROOT_FOLDER, sl, "Anh_Sinh_Vien")
                    
                    if roster_ids:
                        tar = st.selectbox("Chọn sinh viên (MSSV)", roster_ids, key="manage_sv_sel")
                    else:
                        fds = sorted([f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))])
                        tar = st.selectbox("Chọn sinh viên (MSSV)", fds, key="manage_sv_sel2")

                    sv_folder = os.path.join(base_path, tar)
                    bad_folder = os.path.join(ROOT_FOLDER, sl, BAD_IMG_DIR, tar)
                    os.makedirs(sv_folder, exist_ok=True)
                    os.makedirs(bad_folder, exist_ok=True)

                    imgs = [f for f in os.listdir(sv_folder) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
                    bads = [f for f in os.listdir(bad_folder) if f.lower().endswith((".png", ".jpg", ".jpeg"))]

                    data = load_class_data(sl)
                    enc_by_id = 0
                    try:
                        enc_by_id = sum(1 for x in data.get("ids", []) if str(x).strip() == str(tar).strip())
                    except:
                        enc_by_id = 0

                    cA, cB, cC = st.columns(3)
                    with cA:
                        st.metric("Ảnh hợp lệ", len(imgs))
                    with cB:
                        st.metric("Encodings", enc_by_id)
                    with cC:
                        st.metric("Ảnh lỗi", len(bads))

                    if imgs:
                        with st.expander("🖼️ Ảnh hợp lệ", expanded=True):
                            cols = st.columns(5)
                            for i, f in enumerate(imgs):
                                with cols[i % 5]:
                                    st.image(os.path.join(sv_folder, f), use_container_width=True)
                                    if st.button("Xóa", key=f"del_ok_{tar}_{f}", type="secondary"):
                                        try:
                                            os.remove(os.path.join(sv_folder, f))
                                        except:
                                            pass
                                        st.rerun()
                    else:
                        st.info("Chưa có ảnh hợp lệ.")

                    if bads:
                        with st.expander("⚠️ Ảnh lỗi (không thấy mặt)", expanded=False):
                            cols = st.columns(5)
                            for i, f in enumerate(bads):
                                with cols[i % 5]:
                                    st.image(os.path.join(bad_folder, f), use_container_width=True)
                                    if st.button("Xóa", key=f"del_bad_{tar}_{f}", type="secondary"):
                                        try:
                                            os.remove(os.path.join(bad_folder, f))
                                        except:
                                            pass
                                        st.rerun()

                    st.write("")
                    add_more = st.file_uploader("Thêm ảnh mới", accept_multiple_files=True, key="add_more_imgs")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        if st.button("⬆️ Upload & Train", type="primary") and add_more:
                            train_single_student(sl, tar, add_more)
                            st.toast("Đã thêm ảnh!", icon="✅")
                            time.sleep(30)
                            st.rerun()
                    with c2:
                        remove_folder = st.checkbox("Xóa luôn folder (không khuyến nghị)", value=False, key="rm_folder_chk")
                    with c3:
                        if st.button("🗑️ Xóa dữ liệu SV (ảnh+encodings)", type="secondary"):
                            delete_student(sl, tar, remove_folder=remove_folder)
                            st.toast("Đã xóa dữ liệu SV!", icon="🗑️")
                            time.sleep(30)
                            st.rerun()

            else:  # Danh sách Excel / roster
                with bordered_container():
                    section_header("Danh sách sinh viên (Roster)", "Tìm kiếm & xem SV thiếu ảnh để bổ sung", "ROSTER")
                    df_roster = load_roster(sl, autosync_folders=True)
                    if df_roster.empty:
                        st.warning("Chưa đọc được danh sách. Hãy upload file Excel ở tab Lớp Học hoặc thêm SV thủ công.")
                    else:
                        q = st.text_input("Tìm MSSV / Họ tên", placeholder="VD: SV010 hoặc Nguyễn...")
                        df_show = df_roster.copy()
                        if q:
                            ql = q.strip().lower()
                            df_show = df_show[df_show.apply(lambda r: ql in str(r.get("MSSV","")).lower() or ql in str(r.get("Họ tên","")).lower(), axis=1)]

                        img_base = os.path.join(ROOT_FOLDER, sl, "Anh_Sinh_Vien")
                        def _has_img(m):
                            p = os.path.join(img_base, str(m).strip())
                            if not os.path.exists(p):
                                return False
                            return any(f.lower().endswith((".png", ".jpg", ".jpeg")) for f in os.listdir(p))

                        df_show["Có ảnh"] = df_show["MSSV"].apply(_has_img)

                        # ✅ hiển thị đầy đủ cột quan trọng
                        cols_show = ["MSSV", "Họ tên", "Giới tính", "Ngày sinh", "Lớp học", "Có ảnh"]
                        cols_show = [c for c in cols_show if c in df_show.columns]
                        st.dataframe(df_show[cols_show], use_container_width=True, height=420)

                        miss = int((~df_roster["MSSV"].apply(_has_img)).sum())
                        if miss > 0:
                            st.warning(f"⚠️ Có {miss} sinh viên chưa có ảnh training. Nên bổ sung để điểm danh chính xác hơn!")
                
elif selected == "Điểm Danh":
    st.markdown("## 📸 Điểm Danh")

    cl = get_list_classes()
    if not cl:
        with bordered_container():
            section_header("Điểm danh", "Chưa có lớp để điểm danh", "EMPTY")
            st.info("Bạn cần tạo lớp và train dữ liệu trước.")
    else:
        left, right = st.columns([0.95, 1.05])

        with left:
            with bordered_container():
                section_header("Thiết lập", "Chọn lớp, độ nhạy, ngày điểm danh", "SET")
                sl = st.selectbox("Lớp", cl, key="attendance_class")
                tol = st.slider("Độ nhạy (tolerance)", 0.3, 0.6, 0.5, key="attendance_tol")
                dd_date = st.date_input("Ngày điểm danh", value=datetime.now().date(), key="attendance_date")
                st.caption("Tolerance thấp → khắt khe hơn, ít nhận nhầm hơn.")

                df_roster = load_roster(sl, autosync_folders=True)
                stats = class_stats(sl)
                roster_n = stats["roster_count"]
                missing_train = max(0, roster_n - stats["with_images"]) if roster_n else 0

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("👥 Sĩ số", roster_n)
                with c2:
                    st.metric("🖼️ SV có ảnh", stats["with_images"])
                with c3:
                    st.metric("⚠️ Thiếu ảnh", missing_train)

                if missing_train > 0:
                    st.warning("Một số SV chưa có ảnh training → dễ bị đánh vắng. Hãy bổ sung ảnh ở mục Quản Lý SV.")

                st.write("")
                st.caption("QR để truy cập nhanh từ thiết bị khác:")
                st.image(generate_qr(f"http://{get_local_ip()}:8501"), width=180)

        with right:
            with bordered_container():
                section_header("Upload ảnh → Quét", "Quét xong sẽ được chỉnh tay trước khi lưu Excel", sl)
                ups = st.file_uploader(
                    "Chọn ảnh để điểm danh",
                    accept_multiple_files=True,
                    key=f"attendance_upload_{st.session_state.attendance_key}",
                )

                if ups:
                    with st.expander("👀 Xem nhanh ảnh đã chọn", expanded=False):
                        cols = st.columns(5)
                        for i, f in enumerate(ups[:10]):
                            with cols[i % 5]:
                                try:
                                    img = Image.open(f)
                                    st.image(img, use_container_width=True)
                                except:
                                    pass
                        if len(ups) > 10:
                            st.caption(f"Hiển thị 10/{len(ups)} ảnh.")

                a1, a2 = st.columns(2)
                with a1:
                    if st.button("🔍 Quét ngay", type="primary"):
                        if not ups:
                            st.info("Vui lòng chọn ảnh trước!")
                        else:
                            with st.status("Đang quét ảnh...", expanded=True) as status:
                                status.update(label="Bước 1/2: Nhận diện khuôn mặt", state="running")
                                scan_payload, err = scan_images(sl, ups, tol)
                                if err:
                                    status.update(label="Lỗi ❌", state="error")
                                    st.error(err)
                                else:
                                    status.update(label="Bước 2/2: Done ✅", state="complete")
                                    st.session_state.scan_result = scan_payload
                                    st.session_state.save_result = None
                                    st.toast("Quét xong! Kéo xuống để xác nhận.", icon="✅")
                                    st.rerun()

                with a2:
                    if st.button("🔄 Làm mới uploader", type="secondary"):
                        st.session_state.attendance_key += 1
                        st.rerun()

        # ===== REVIEW & SAVE =====
        if st.session_state.scan_result and st.session_state.scan_result.get("class_name") == st.session_state.get("attendance_class"):
            sr = st.session_state.scan_result

            st.write("")
            with bordered_container():
                section_header("✅ Xác nhận kết quả", "Chỉnh danh sách có mặt trước khi lưu", "REVIEW")

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("📷 Ảnh", sr.get("images_count", 0))
                with c2:
                    st.metric("🙂 Khuôn mặt", sr.get("faces_total", 0))
                with c3:
                    st.metric("🆔 Nhận diện", len(sr.get("found_ids", [])))
                with c4:
                    st.metric("⏱ Thời gian", sr.get("scanned_at", ""))

                # bảng chỉnh tay theo roster
                df_roster = load_roster(sl, autosync_folders=True)
                recognized = set([str(x).strip() for x in sr.get("found_ids", [])])
                hits = sr.get("hits", {})
                min_dist = sr.get("min_dist", {})

                if df_roster.empty:
                    st.warning("Không đọc được roster. Dùng danh sách nhận diện để lưu.")
                    final_ids = st.multiselect("Danh sách MSSV có mặt", options=sorted(list(recognized)), default=sorted(list(recognized)))
                else:
                    df_edit = df_roster[["MSSV", "Họ tên"]].copy()
                    df_edit["Có mặt"] = df_edit["MSSV"].astype(str).apply(lambda x: str(x).strip() in recognized)
                    df_edit["Hits"] = df_edit["MSSV"].astype(str).apply(lambda x: hits.get(str(x).strip(), 0))
                    df_edit["MinDist"] = df_edit["MSSV"].astype(str).apply(lambda x: round(min_dist.get(str(x).strip(), np.nan), 4) if str(x).strip() in min_dist else np.nan)

                    extra = []
                    roster_set = set(df_roster["MSSV"].astype(str).str.strip().tolist())
                    for m in recognized:
                        if m not in roster_set:
                            extra.append({
                                "MSSV": m,
                                "Họ tên": "(Không có trong roster)",
                                "Có mặt": True,
                                "Hits": hits.get(m, 0),
                                "MinDist": round(min_dist.get(m, np.nan), 4) if m in min_dist else np.nan,
                            })
                    if extra:
                        df_edit = pd.concat([df_edit, pd.DataFrame(extra)], ignore_index=True)

                    st.caption("Tip: tick/untick cột **Có mặt** để chỉnh tay. Hits/MinDist giúp kiểm tra độ tin cậy.")
                    edited = st.data_editor(df_edit, use_container_width=True, height=420)
                    final_ids = edited[edited["Có mặt"] == True]["MSSV"].astype(str).tolist()

                st.write("")
                s1, s2, s3 = st.columns([0.34, 0.33, 0.33])
                with s1:
                    do_save_excel = st.checkbox("Cập nhật Excel", value=True)
                with s2:
                    do_save_history = st.checkbox("Lưu lịch sử phiên", value=True)
                with s3:
                    save_copy_if_open = st.checkbox("Nếu Excel đang mở → lưu bản copy", value=True)

                st.write("")
                bsave, bclear = st.columns([0.6, 0.4])

                with bsave:
                    if st.button("✅ Lưu điểm danh", type="primary"):
                        excel_updated = False
                        real_total = 0
                        lp, la = [], []
                        err = None
                        saved_excel_path = None

                        if do_save_excel:
                            saved_excel_path, err, real_total, lp, la, dd_code = process_attendance_excel(
                                sl, final_ids, attendance_date=dd_date, save_copy_if_open=save_copy_if_open
                            )
                            if err:
                                st.warning(err)
                            else:
                                excel_updated = True
                                st.toast("Đã cập nhật Excel!", icon="📄")

                        if do_save_history:
                            payload = dict(sr)
                            payload["final_ids"] = list(final_ids)
                            payload["excel_updated"] = bool(excel_updated)
                            payload["present_count"] = len(lp) if lp else None
                            payload["absent_count"] = len(la) if la else None
                            payload["date_mark"] = dd_date.strftime("%d/%m/%Y")
                            save_session_history(sl, payload)
                            st.toast("Đã lưu lịch sử phiên!", icon="🗂️")

                        st.session_state.save_result = {
                            "excel_updated": excel_updated,
                            "real_total": real_total,
                            "present": lp,
                            "absent": la,
                            "final_ids": final_ids,
                            "saved_excel_path": saved_excel_path,
                        }
                        st.balloons()
                        st.rerun()

                with bclear:
                    if st.button("🧹 Xóa kết quả quét", type="secondary"):
                        st.session_state.scan_result = None
                        st.session_state.save_result = None
                        st.rerun()

        # ===== RESULT AFTER SAVE =====
        if st.session_state.save_result and st.session_state.scan_result:
            res = st.session_state.save_result
            st.write("")
            with bordered_container():
                section_header("📌 Kết quả phiên", "Tóm tắt sau khi lưu", "RESULT")

                if res.get("excel_updated"):
                    real_total = int(res.get("real_total", 0))
                    present = res.get("present", [])
                    absent = res.get("absent", [])

                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.metric("👥 Tổng SV", real_total)
                    with c2:
                        st.metric("✅ Có mặt", len(present))
                    with c3:
                        st.metric("❌ Vắng", max(0, real_total - len(present)))

                    t1, t2 = st.columns(2)
                    with t1:
                        st.markdown("**✅ Danh sách có mặt**")
                        st.text("\n".join(present[:200]) if present else "—")
                        if len(present) > 200:
                            st.caption(f"... +{len(present)-200} dòng")

                    with t2:
                        st.markdown("**❌ Danh sách vắng**")
                        st.text("\n".join(absent[:200]) if absent else "—")
                        if len(absent) > 200:
                            st.caption(f"... +{len(absent)-200} dòng")

                    saved_excel_path = res.get("saved_excel_path")
                    if saved_excel_path and os.path.exists(saved_excel_path):
                        st.write("")
                        with open(saved_excel_path, "rb") as f:
                            st.download_button(
                                "⬇️ Tải Excel đã cập nhật",
                                data=f,
                                file_name=os.path.basename(saved_excel_path),
                            )
                else:
                    st.warning("Excel chưa được cập nhật (thiếu Excel hoặc Excel đang mở). Bạn vẫn có thể xem lịch sử phiên nếu đã bật lưu lịch sử.")

elif selected == "Báo Cáo":
    st.markdown("## 📊 Báo Cáo")

    cl = get_list_classes()
    if not cl:
        with bordered_container():
            section_header("Báo cáo", "Chưa có lớp để hiển thị", "EMPTY")
            st.info("Bạn cần tạo lớp trước.")
    else:
        c = st.selectbox("Chọn lớp", cl, key="report_class")
        stats = class_stats(c)

        with bordered_container():
            section_header("Tổng quan", "Roster + coverage + số buổi điểm danh", c)
            a1, a2, a3, a4, a5 = st.columns(5)
            with a1:
                st.metric("👥 Sĩ số", stats["roster_count"])
            with a2:
                st.metric("🖼️ SV có ảnh", stats["with_images"])
            with a3:
                st.metric("🧬 Encodings", stats["enc_count"])
            with a4:
                st.metric("⚠️ Ảnh lỗi", stats["bad_img_count"])
            with a5:
                st.metric("📌 Buổi DD", stats["sessions_count"])

        with bordered_container():
            section_header("🗂️ Lịch sử phiên", "Tự động lưu trong Attendance_History", "HISTORY")
            df_sessions = load_sessions_summary(c)
            if df_sessions.empty:
                st.info("Chưa có phiên điểm danh nào.")
            else:
                if "session_id" in df_sessions.columns:
                    df_sessions = df_sessions.sort_values("session_id", ascending=False)
                st.dataframe(df_sessions, use_container_width=True, height=300)

                csv_path = os.path.join(ROOT_FOLDER, c, HISTORY_DIR_NAME, SESSIONS_CSV)
                if os.path.exists(csv_path):
                    with open(csv_path, "rb") as f:
                        st.download_button("⬇️ Tải sessions.csv", data=f, file_name=f"{c}_sessions.csv")

        with bordered_container():
            section_header("📈 Thống kê theo sinh viên", "Tỷ lệ có mặt dựa trên lịch sử phiên", "STATS")
            df_rate, total_sess = compute_attendance_rate(c)
            if total_sess == 0 or df_rate.empty:
                st.info("Chưa đủ dữ liệu lịch sử để thống kê.")
            else:
                st.caption(f"Tổng số buổi đã lưu: {total_sess}")
                st.dataframe(df_rate, use_container_width=True, height=420)
                st.caption("Danh sách được sắp theo SV có tỷ lệ thấp (dễ vắng) lên trên.")

        with bordered_container():
            section_header("📄 Xem Excel", "Đọc đúng dòng header, hạn chế Unnamed", "EXCEL")
            df, err = read_excel_for_report(c)
            if err:
                st.error(err)
            elif df is None:
                st.info("Chưa có file Excel.")
            else:
                st.dataframe(df, use_container_width=True, height=520)
