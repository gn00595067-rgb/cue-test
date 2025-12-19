import streamlit as st
import math
import io
import os
import shutil
import tempfile
import subprocess
import re
from datetime import timedelta, datetime, date
from copy import copy

import requests
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# =========================================================
# 0. 基礎設定 & 強力自動載入
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v61.5 (Dual-Mode Load)")

# 您的檔案 ID
GOOGLE_DRIVE_FILE_ID = "11R1SA_hpFD5O_MGmYeh4BdtcUhK2bPta"
DEFAULT_FILENAME = "1209-Cue表相關資料.xlsx"

@st.cache_resource(ttl=600)
def load_default_template():
    status_msg = []
    
    # 1. 雲端下載嘗試
    if GOOGLE_DRIVE_FILE_ID:
        # 模式 A: 針對上傳的 Excel 檔案 (Binary Download)
        url_binary = f"https://drive.google.com/uc?export=download&id={GOOGLE_DRIVE_FILE_ID}"
        # 模式 B: 針對原生 Google Sheets (Export)
        url_export = f"https://docs.google.com/spreadsheets/d/{GOOGLE_DRIVE_FILE_ID}/export?format=xlsx"
        
        urls_to_try = [
            ("Binary Mode", url_binary),
            ("Export Mode", url_export)
        ]

        for mode, url in urls_to_try:
            try:
                r = requests.get(url, timeout=10)
                if r.status_code == 200:
                    # 檢查檔頭：Excel 檔通常以 'PK' 開頭 (Zip)，HTML 則是 '<!DOCTYPE'
                    content_start = r.content[:50]
                    if b"<!DOCTYPE" in content_start or b"<html" in content_start:
                        status_msg.append(f"❌ {mode}: 下載到網頁而非檔案 (請確認權限是否為「知道連結者均可檢視」)")
                    else:
                        return r.content, f"雲端硬碟 ({mode})", status_msg
                else:
                    status_msg.append(f"❌ {mode}: HTTP {r.status_code}")
            except Exception as e:
                status_msg.append(f"❌ {mode}: 連線錯誤 {e}")

    # 2. 本地備援
    if os.path.exists(DEFAULT_FILENAME):
        try:
            with open(DEFAULT_FILENAME, "rb") as f:
                return f.read(), "系統主機 (Local)", status_msg
        except: pass
    
    return None, None, status_msg

st.markdown("""
<style>
  .stApp { background-color: #f0f2f6; font-family: "Microsoft JhengHei", "Segoe UI", sans-serif; }
  .block-container { padding-top: 1.5rem; max-width: 98% !important; }
  
  .preview-wrapper { 
      background: white; 
      padding: 20px; 
      border: 1px solid #ccc; 
      box-shadow: 0 4px 12px rgba(0,0,0,0.1); 
      margin-bottom: 20px; 
      overflow: auto; 
      max-height: 800px; 
  }
  
  table.cue-table { 
      border-collapse: separate; 
      border-spacing: 0; 
      width: max-content; 
      min-width: 100%; 
      font-size: 13px; 
      color: #000;
      border: 1px solid #a0a0a0;
  }
  
  .cue-table th, .cue-table td { 
      border-right: 1px solid #a0a0a0;
      border-bottom: 1px solid #a0a0a0;
      padding: 6px 8px; 
      text-align: center; 
      vertical-align: middle; 
      white-space: nowrap; 
      line-height: 1.4;
  }
  
  .cue-table thead { position: sticky; top: 0; z-index: 100; }
  .cue-table thead th { position: sticky; top: 0; z-index: 100; border-bottom: 2px solid #000; }
  .cue-table thead tr:nth-child(2) th { top: 35px; z-index: 99; } 

  .sticky-col { position: sticky; left: 0; z-index: 101; background-color: #fff; border-right: 2px solid #000 !important; }
  .col-1 { left: 0px; z-index: 105; width: 140px; min-width: 140px; }
  .col-2 { left: 140px; z-index: 104; width: 120px; min-width: 120px; }
  .col-3 { left: 260px; z-index: 103; width: 80px; min-width: 80px; }
  .col-4 { left: 340px; z-index: 102; width: 100px; min-width: 100px; }
  .col-5 { left: 440px; z-index: 101; width: 80px; min-width: 80px; }
  
  .header-dw { background-color: #4472C4; color: white; font-weight: bold; }
  .header-sh { background-color: #BDD7EE; color: black; font-weight: bold; border-color: #999; }
  
  .we-dw { background-color: #FFD966; color: black; } 
  .total-row { background-color: #FFF2CC; font-weight: bold; border-top: 2px solid #000 !important; }
  .num-cell { text-align: right !important; font-family: "Consolas", "Arial", sans-serif; }
  .txt-left { text-align: left !important; padding-left: 10px; }
  
  .station-cell { 
      background-color: #fff; 
      font-weight: bold; 
      vertical-align: middle;
      border-right: 2px solid #000;
  }

  .remarks-box {
      margin-top: 20px;
      padding: 15px;
      background: #fdfdfd;
      border: 1px dashed #ccc;
      font-size: 13px;
      line-height: 1.8;
  }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 1. 資料庫 (2026 新制)
# =========================================================
STORE_COUNTS_RAW = {
    "全省": "4,437店",
    "北區": "1,649店", "桃竹苗": "779店", "中區": "839店", "雲嘉南": "499店", "高屏": "490店", "東區": "181店",
    "新鮮視_全省": "3,124面",
    "新鮮視_北區": "1,127面", "新鮮視_桃竹苗": "616面", "新鮮視_中區": "528面",
    "新鮮視_雲嘉南": "365面", "新鮮視_高屏": "405面", "新鮮視_東區": "83面",
    "家樂福_量販": "67店", "家樂福_超市": "250店"
}

def parse_count_to_int(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return int(x)
    s = str(x)
    m = re.findall(r"[\d,]+", s)
    if not m: return 0
    return int(m[0].replace(",", ""))

STORE_COUNTS_NUM = {k: parse_count_to_int(v) for k, v in STORE_COUNTS_RAW.items()}
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

PRICING_DB = {
    "全家廣播": { "Std_Spots": 480, "Day_Part": "00:00-24:00", "Base_Unit": 30,
        "全省": [400000, 320000], "北區": [250000, 200000], "桃竹苗": [150000, 120000], "中區": [150000, 120000],
        "雲嘉南": [100000, 80000], "高屏": [100000, 80000], "東區": [62500, 50000] },
    "新鮮視": { "Std_Spots": 504, "Day_Part": "07:00-22:00", "Base_Unit": 10,
        "全省": [150000, 120000], "北區": [150000, 120000], "桃竹苗": [120000, 96000], "中區": [90000, 72000],
        "雲嘉南": [75000, 60000], "高屏": [75000, 60000], "東區": [45000, 36000] },
    "家樂福": { "Base_Unit": 20,
        "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420, "Day_Part": "09:00-23:00"},
        "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720, "Day_Part": "00:00-24:00"} }
}

SEC_FACTORS = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5, 5: 0.25},
    "新鮮視": {30: 3.0, 20: 2.0, 15: 1.5, 10: 1.0, 5: 0.5},
    "家樂福": {30: 1.5, 20: 1.0, 15: 0.85, 10: 0.65, 5: 0.35}
}

def get_sec_factor(media_type, seconds): return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    if days <= 0: return []
    if total_spots % 2 != 0: total_spots += 1
    half_spots = total_spots // 2
    base, rem = divmod(half_spots, days)
    half_schedule = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in half_schedule]

def get_remarks_text(sign_deadline, billing_month, payment_date):
    d_str = sign_deadline.strftime("%Y/%m/%d (%a) %H:%M") if sign_deadline else "____/__/__ (__) 12:00"
    p_str = payment_date.strftime("%Y/%m/%d") if payment_date else "____/__/__"
    return [
        f"1.請於 {d_str}前 回簽及進單，方可順利上檔。",
        "2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇時段滿檔，上檔時間挪後或更換至同級時段。",
        "3.通路店鋪數與開機率開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。",
        f"6.付款兌現日期：{p_str}"
    ]

REGION_DISPLAY_6 = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市",
}
def region_display(region: str) -> str: return REGION_DISPLAY_6.get(region, region)

# =========================================================
# 2. openpyxl 工具 (強化格式)
# =========================================================
def _get_master_cell(ws, cell):
    if not isinstance(cell, MergedCell): return cell
    r, c = cell.row, cell.column
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
            return ws.cell(row=mr.min_row, column=mr.min_col)
    return None

def safe_write(ws, addr: str, value):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master: master.value = value
        return
    cell.value = value

def safe_write_rc(ws, row: int, col: int, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master: master.value = value
        return
    cell.value = value

def apply_center_style(cell):
    """ 強制置中 + 自動換行 """
    existing_align = cell.alignment
    cell.alignment = Alignment(
        horizontal="center", 
        vertical="center", 
        wrap_text=True, 
        indent=existing_align.indent if existing_align else 0
    )

def copy_row_with_style_fix(ws, src_row, dst_row, max_col):
    """ 精準複製列樣式 """
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    row_shift = dst_row - src_row
    for c in range(1, max_col + 1):
        sc = ws.cell(src_row, c)
        dc = ws.cell(dst_row, c)
        if sc.has_style:
            dc.font = copy(sc.font)
            dc.border = copy(sc.border)
            dc.fill = copy(sc.fill)
            dc.number_format = sc.number_format
            dc.protection = copy(sc.protection)
            dc.alignment = copy(sc.alignment)
        
        v = sc.value
        if isinstance(v, str) and v.startswith("="):
            try: dc.value = Translator(v, origin=sc.coordinate).translate_formula(row_shift=row_shift, col_shift=0)
            except: dc.value = v
        else:
            dc.value = v

def force_center_columns_range(ws, col_letters, start_row, end_row):
    if start_row is None or end_row is None: return
    for r in range(start_row, end_row + 1):
        for col in col_letters:
            addr = f"{col}{r}"
            cell = ws[addr]
            if isinstance(cell, MergedCell):
                master = _get_master_cell(ws, cell)
                if master: cell = master
                else: continue
            apply_center_style(cell)

# =========================================================
# 3. 模板與輸出邏輯
# =========================================================
SHEET_META = {
    "Dongwu": {
        "sheet_name": "東吳-格式", "date_start_cell": "I7", "schedule_start_col": "I",
        "max_days": 31, "total_col": "AN",
        "anchors": {"全家廣播": "通路廣播廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "header_cells": {"client": "C3", "product": "C4", "period": "C5", "medium": "C6", "month": "I6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "rate": "G", "pkg": "H"},
        "header_override": {"G7": "rate\n(List)", "H7": "Package-cost\n(List)"},
        "station_merge": True, "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": ["E", "F", "G", "H"], 
    },
    "Shenghuo": {
        "sheet_name": "聲活-格式", "date_start_cell": "G7", "schedule_start_col": "G",
        "max_days": 23, "total_col": "AD",
        "anchors": {"全家廣播": "廣播通路廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "header_cells": {"client": "C5", "product": "C6", "month": "G6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "proj_price": "AF"},
        "header_override": {"AF7": "專案價\n(List)"}, 
        "station_merge": False, "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": [],
    }
}

def hide_unused_sheets(wb, keep_sheet_names, mode="veryHidden"):
    for sh in wb.worksheets:
        if sh.title not in keep_sheet_names: sh.sheet_state = mode

def find_first_row_contains(ws, col_letter, keyword):
    col_idx = column_index_from_string(col_letter)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if isinstance(v, str) and keyword in v: return r
    return None

def find_cell_exact(ws, text):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == text: return cell.row, cell.column
    return None

def unmerge_col_overlap(ws, col_letter, start_row, end_row):
    st_col = column_index_from_string(col_letter)
    to_unmerge = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col == st_col and mr.max_col == st_col:
            if not (mr.max_row < start_row or mr.min_row > end_row):
                to_unmerge.append(str(mr))
    for s in set(to_unmerge):
        try: ws.unmerge_cells(s)
        except: pass

def set_schedule(ws, row, start_col_letter, max_days, schedule_list):
    start_col = column_index_from_string(start_col_letter)
    for i in range(max_days):
        v = schedule_list[i] if (schedule_list and i < len(schedule_list)) else None
        safe_write_rc(ws, row, start_col + i, v)

def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    if meta["sheet_name"] not in wb.sheetnames: raise ValueError(f"缺少分頁：{meta['sheet_name']}")
    ws = wb[meta["sheet_name"]]
    hide_unused_sheets(wb, [meta["sheet_name"]])

    hc = meta["header_cells"]
    if "client" in hc: safe_write(ws, hc["client"], client_name)
    if "product" in hc: safe_write(ws, hc["product"], product_display_str)
    if "period" in hc: safe_write(ws, hc["period"], f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y.%m. %d')}")
    if "medium" in hc: safe_write(ws, hc["medium"], " ".join(sorted(set([r["media_type"] for r in rows]))))
    if "month" in hc: safe_write(ws, hc["month"], f" {start_dt.month}月")
    safe_write(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))

    for addr, text in meta.get("header_override", {}).items():
        safe_write(ws, addr, text)

    total_cell = find_cell_exact(ws, meta["total_label"])
    if not total_cell: raise ValueError("找不到 Total")
    total_row = total_cell[0]

    cols = meta["cols"]
    sec_start = {}
    for m_key, kw in meta["anchors"].items():
        r0 = find_first_row_contains(ws, cols["station"], kw)
        if r0: sec_start[m_key] = r0
    
    sec_order = sorted(sec_start.items(), key=lambda x: x[1])
    sec_ranges = []
    for i, (k, sr) in enumerate(sec_order):
        next_start = sec_order[i + 1][1] if i + 1 < len(sec_order) else total_row
        sec_ranges.append((k, sr, next_start - 1))

    reg_map = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}
    def sort_key(x): return (x["seconds"], reg_map.get(x["region"], 999))
    
    grouped = {
        "全家廣播": sorted([r for r in rows if r["media_type"] == "全家廣播"], key=sort_key),
        "新鮮視": sorted([r for r in rows if r["media_type"] == "新鮮視"], key=sort_key),
        "家樂福": sorted([r for r in rows if r["media_type"] == "家樂福"], key=sort_key),
    }

    for k, sr, er in sorted(sec_ranges, key=lambda x: x[1], reverse=True):
        data = grouped.get(k, [])
        needed = len(data)
        if needed <= 0: continue
        existing = er - sr + 1
        if needed > existing:
            ws.insert_rows(er + 1, amount=needed - existing)
            for rr in range(er + 1, er + 1 + needed - existing):
                copy_row_with_style_fix(ws, sr, rr, ws.max_column)

    total_row = find_cell_exact(ws, meta["total_label"])[0]
    sec_start = {}
    for m_key, kw in meta["anchors"].items():
        r0 = find_first_row_contains(ws, cols["station"], kw)
        if r0: sec_start[m_key] = r0
    sec_order = sorted(sec_start.items(), key=lambda x: x[1])
    sec_ranges = []
    for i, (k, sr) in enumerate(sec_order):
        next_start = sec_order[i + 1][1] if i + 1 < len(sec_order) else total_row
        sec_ranges.append((k, sr, next_start - 1))

    def station_title(m):
        prefix = "全家便利商店\n" if m != "家樂福" else ""
        name = "通路廣播廣告" if m == "全家廣播" else "新鮮視廣告" if m == "新鮮視" else "家樂福"
        if format_type == "Shenghuo" and m == "全家廣播": name = "廣播通路廣告"
        return prefix + name

    written_rows = []
    for m, sr, er in sec_ranges:
        data = grouped.get(m, [])
        if not data: continue
        
        if meta["station_merge"]:
            unmerge_col_overlap(ws, cols["station"], sr, er)
            merge_rng = f"{cols['station']}{sr}:{cols['station']}{sr + len(data) - 1}"
            ws.merge_cells(merge_rng)
            top_cell = ws[f"{cols['station']}{sr}"]
            top_cell.value = station_title(m)
            apply_center_style(top_cell)

        row_ptr = sr
        for r in data:
            if not meta["station_merge"]: 
                cell = ws[f"{cols['station']}{row_ptr}"]
                cell.value = station_title(m)
                apply_center_style(cell)

            safe_write(ws, f"{cols['location']}{row_ptr}", region_display(r["region"]))
            prog_val = r.get("program_num", parse_count_to_int(r.get("program", 0)))
            safe_write(ws, f"{cols['program']}{row_ptr}", int(prog_val))

            if format_type == "Dongwu":
                safe_write(ws, f"{cols['daypart']}{row_ptr}", r["daypart"])
                if m == "家樂福": safe_write(ws, f"{cols['seconds']}{row_ptr}", f"{r['seconds']}秒")
                else: safe_write(ws, f"{cols['seconds']}{row_ptr}", int(r["seconds"]))
                safe_write(ws, f"{cols['rate']}{row_ptr}", r["rate_list"])
                safe_write(ws, f"{cols['pkg']}{row_ptr}", r["pkg_display_val"])
            else:
                safe_write(ws, f"{cols['daypart']}{row_ptr}", r["daypart"])
                safe_write(ws, f"{cols['seconds']}{row_ptr}", f"{r['seconds']}秒廣告")
                safe_write(ws, f"{cols['proj_price']}{row_ptr}", r["pkg_display_val"] if isinstance(r["pkg_display_val"], int) else 0)

            set_schedule(ws, row_ptr, meta["schedule_start_col"], meta["max_days"], r["schedule"])
            spot_sum = sum(r["schedule"][:meta["max_days"]])
            safe_write(ws, f"{meta['total_col']}{row_ptr}", spot_sum)
            written_rows.append(row_ptr)
            row_ptr += 1

    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [sum([x["schedule"][d] for x in rows if d < len(x["schedule"])]) for d in range(eff_days)]
    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))
    
    total_pkg = sum([x["pkg_display_val"] for x in rows if isinstance(x["pkg_display_val"], int)])
    pkg_col = cols.get("pkg") or cols.get("proj_price")
    safe_write(ws, f"{pkg_col}{total_row}", total_pkg)

    lbl = meta["footer_labels"]
    def write_footer(key, val):
        pos = find_cell_exact(ws, lbl.get(key, ""))
        if pos: safe_write_rc(ws, pos[0], pos[1]+1, int(val))

    make_fee = 10000 
    pos_make = find_cell_exact(ws, lbl["make"])
    if pos_make:
        v = ws.cell(pos_make[0], pos_make[1]+1).value
        if isinstance(v, (int, float)) and v > 0: make_fee = int(v)
        else: safe_write_rc(ws, pos_make[0], pos_make[1]+1, make_fee)
    
    vat = int(round((total_pkg + make_fee) * 0.05))
    write_footer("vat", vat)
    write_footer("grand", total_pkg + make_fee + vat)

    rem_pos = find_cell_exact(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_write_rc(ws, rem_pos[0] + 1 + i, rem_pos[1], rm)

    if format_type == "Dongwu" and written_rows:
        min_r, max_r = min(written_rows), total_row
        force_center_columns_range(ws, meta["force_center_cols"], min_r, max_r)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 5. HTML to PDF via WeasyPrint
# =========================================================
def html_to_pdf_weasyprint(html_str):
    try: 
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
    except ImportError: 
        return None, "WeasyPrint not installed"
    except Exception as e:
        return None, f"WeasyPrint setup error: {str(e)}"
    
    font_path = "NotoSansTC-Regular.ttf"
    if not os.path.exists(font_path):
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/TTF/TraditionalChinese/NotoSansTC-Regular.ttf"
        try:
            r = requests.get(url, timeout=15)
            if r.status_code == 200: 
                with open(font_path, "wb") as f: f.write(r.content)
        except: pass
        
    font_config = FontConfiguration()
    css_str = f"""
    @font-face {{ font-family: 'NotoSansTC'; src: url(file://{os.path.abspath(font_path)}); }}
    body, table, th, td {{ font-family: 'NotoSansTC', sans-serif !important; font-size: 10px; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ border: 1px solid #000; padding: 4px; }}
    tr {{ page-break-inside: avoid; }}
    """
    
    try:
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[CSS(string=css_str)], font_config=font_config)
        return pdf_bytes, ""
    except Exception as e:
        return None, f"WeasyPrint render error: {str(e)}"

# =========================================================
# 6. HTML Preview
# =========================================================
def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks):
    header_cls = "header-dw" if format_type == "Dongwu" else "header-sh"
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    
    eff_days = min(days_cnt, 31)
    
    date_headers_1 = ""
    date_headers_2 = ""
    curr = start_dt
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for i in range(eff_days):
        wd = curr.weekday()
        bg_cls = "we-dw" if (format_type == "Dongwu" and wd >= 5) else ""
        date_headers_1 += f"<th class='{header_cls} {bg_cls}'>{curr.day}</th>"
        date_headers_2 += f"<th class='{header_cls} {bg_cls}'>{weekdays[wd]}</th>"
        curr += timedelta(days=1)

    if format_type == "Dongwu":
        cols_def = [
            ("Station", "sticky-col col-1"), ("Location", "sticky-col col-2"), 
            ("Program", "sticky-col col-3"), ("Day-part", "sticky-col col-4"), 
            ("Size", "sticky-col col-5"), ("rate<br>(List)", ""), ("Package<br>(List)", "")
        ]
    else:
        cols_def = [
            ("頻道", "sticky-col col-1"), ("播出地區", "sticky-col col-2"), 
            ("播出店數", "sticky-col col-3"), ("播出時間", "sticky-col col-4"), 
            ("秒數<br>規格", "sticky-col col-5"), ("專案價", "")
        ]
    
    th_fixed = "".join([f"<th class='{header_cls} {c[1]}' rowspan='2'>{c[0]}</th>" for c in cols_def])
    
    rows_sorted = sorted(rows, key=lambda x: (
        media_order.get(x["media_type"], 99), 
        x["seconds"], 
        REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 99
    ))
    
    body_html = ""
    media_group_counts = {}
    for r in rows_sorted:
        m = r["media_type"]
        media_group_counts[m] = media_group_counts.get(m, 0) + 1
        
    media_printed = {m: False for m in media_group_counts}
    
    for r in rows_sorted:
        m = r["media_type"]
        body_html += "<tr>"
        
        if not media_printed[m]:
            rowspan = media_group_counts[m]
            display_name = "全家便利商店<br>通路廣播廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
            if format_type == "Shenghuo" and m == "全家廣播": display_name = "全家便利商店<br>廣播通路廣告"
            
            if format_type == "Shenghuo":
                body_html += f"<td class='station-cell sticky-col col-1'>{display_name}</td>"
            else:
                body_html += f"<td class='station-cell sticky-col col-1' rowspan='{rowspan}'>{display_name}</td>"
                media_printed[m] = True
        elif format_type == "Shenghuo":
             display_name = "全家便利商店<br>廣播通路廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
             body_html += f"<td class='station-cell sticky-col col-1'>{display_name}</td>"
            
        body_html += f"<td class='sticky-col col-2'>{region_display(r['region'])}</td>"
        prog = r.get('program_num', '')
        body_html += f"<td class='sticky-col col-3 num-cell'>{prog}</td>"
        body_html += f"<td class='sticky-col col-4'>{r['daypart']}</td>"
        
        sec_txt = f"{r['seconds']}秒" if format_type=="Dongwu" and m=="家樂福" else f"{r['seconds']}" if format_type=="Dongwu" else f"{r['seconds']}秒廣告"
        body_html += f"<td class='sticky-col col-5'>{sec_txt}</td>"
        
        rate = f"{r['rate_list']:,}" if isinstance(r['rate_list'], int) else r['rate_list']
        pkg = f"{r['pkg_display_val']:,}" if isinstance(r['pkg_display_val'], int) else r['pkg_display_val']
        
        if format_type == "Dongwu":
            body_html += f"<td class='num-cell'>{rate}</td>"
            body_html += f"<td class='num-cell'>{pkg}</td>"
        else:
            body_html += f"<td class='num-cell'>{pkg}</td>"
        
        for d in r['schedule'][:eff_days]:
            body_html += f"<td class='num-cell'>{d}</td>"
            
        body_html += f"<td class='total-row num-cell'>{sum(r['schedule'])}</td>"
        body_html += "</tr>"

    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]
    total_spots = sum(totals)
    total_pkg = sum([r["pkg_display_val"] for r in rows if isinstance(r["pkg_display_val"], int)])
    
    colspan = 5 if format_type == "Dongwu" else 5
    empty = "<td></td>" if format_type == "Dongwu" else ""
    
    footer_html = f"<tr class='total-row'><td colspan='{colspan}' class='txt-left sticky-col col-1'>Total</td>{empty}<td class='num-cell'>{total_pkg:,}</td>"
    for t in totals:
        footer_html += f"<td class='num-cell'>{t}</td>"
    footer_html += f"<td class='num-cell'>{total_spots}</td></tr>"

    return f"""
    <div class='preview-container'>
        <div class='info-bar'>
            <div><b>客戶：</b>{c_name}</div>
            <div><b>產品：</b>{p_display}</div>
            <div style="color:#666;">期間：{start_dt} ~ {end_dt}</div>
        </div>
        <table class='excel-table'>
            <thead>
                <tr>
                    {th_fixed}
                    {date_headers_1}
                    <th class='{header_cls} total' rowspan='2'>檔次</th>
                </tr>
                <tr>
                    {date_headers_2}
                </tr>
            </thead>
            <tbody>
                {body_html}
                {footer_html}
            </tbody>
        </table>
        <div class="remarks-box">
            <b>Remarks：</b><br>
            {"<br>".join(remarks)}
        </div>
    </div>
    """

# =========================================================
# 7. UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v61.5)")

auto_tpl, source, msgs = load_default_template()
template_bytes = None

if auto_tpl:
    st.success(f"✅ 已自動載入系統公版 ({source})")
    template_bytes = auto_tpl
    with st.expander("🛠️ 進階：上傳其他版本模板"):
        tpl = st.file_uploader("上傳 Excel 模板", type=["xlsx"])
        if tpl: template_bytes = tpl.read()
else:
    st.warning("⚠️ 尚未偵測到公版檔案，請手動上傳")
    if msgs:
        with st.expander("🔍 下載失敗原因診斷"):
            for m in msgs: st.write(m)
            
    tpl = st.file_uploader("上傳 Excel 模板 (1209-Cue表相關資料.xlsx)", type=["xlsx"])
    if tpl: template_bytes = tpl.read()

# [New] Format Selection
st.markdown("### 1. 選擇格式")
format_type = st.radio("", ["Dongwu", "Shenghuo"], horizontal=True, label_visibility="collapsed")

# 1. 基本資料
with st.container():
    st.markdown("### 2. 基本資料設定")
    with st.expander("📝 點擊展開/收合基本資料", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1: client_name = st.text_input("客戶名稱", "萬國通路")
        with c2: product_name = st.text_input("產品名稱", "統一布丁")
        with c3: total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)
        
        c4, c5 = st.columns(2)
        with c4: start_date = st.date_input("開始日", datetime(2026, 1, 1))
        with c5: end_date = st.date_input("結束日", datetime(2026, 1, 31))
        
        days_count = (end_date - start_date).days + 1
        st.info(f"📅 走期共 **{days_count}** 天")

with st.expander("📝 備註欄位設定 (Remarks)", expanded=False):
    rc1, rc2, rc3 = st.columns(3)
    sign_deadline = rc1.date_input("回簽截止日", datetime.now() + timedelta(days=3))
    billing_month = rc2.text_input("請款月份", "2026年2月")
    payment_date = rc3.date_input("付款兌現日", datetime(2026, 3, 31))

# 2. 媒體設定
st.markdown("### 3. 媒體投放設定")
m1, m2, m3 = st.columns(3)
config = {}
rem_budget = 100

with m1:
    if st.checkbox("全家廣播", True):
        is_nat = st.checkbox("全省聯播", True, key="rad_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_reg")
        secs = st.multiselect("秒數", DURATIONS, [20], key="rad_sec")
        share = st.slider("預算 %", 0, 100, 60, key="rad_share")
        rem_budget -= share
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in sorted(secs)[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"rs_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs: sec_shares[secs[0]] = 100
        config["全家廣播"] = {"is_national": is_nat, "regions": regs, "seconds": sorted(secs), "share": share, "sec_shares": sec_shares}

with m2:
    if st.checkbox("新鮮視", True):
        is_nat = st.checkbox("全省聯播", False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
        secs = st.multiselect("秒數", DURATIONS, [10], key="fv_sec")
        share = st.slider("預算 %", 0, rem_budget, min(20, rem_budget), key="fv_share")
        rem_budget -= share
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in sorted(secs)[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"fs_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs: sec_shares[secs[0]] = 100
        config["新鮮視"] = {"is_national": is_nat, "regions": regs, "seconds": sorted(secs), "share": share, "sec_shares": sec_shares}

with m3:
    if st.checkbox("家樂福", True):
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.info(f"剩餘預算: {rem_budget}%")
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in sorted(secs)[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"cs_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs: sec_shares[secs[0]] = 100
        config["家樂福"] = {"regions": ["全省"], "seconds": sorted(secs), "share": rem_budget, "sec_shares": sec_shares}

# 3. 計算引擎
rows = []
debug_logs = []

if config:
    for m, cfg in config.items():
        m_budget = total_budget_input * (cfg["share"] / 100.0)
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0: continue
            
            factor = get_sec_factor(m, sec)
            
            if m in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m]
                calc_regs = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                display_regs = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                unit_net_sum = 0
                for r in calc_regs:
                    unit_net_sum += (db[r][1] / db["Std_Spots"]) * factor
                
                if unit_net_sum == 0: continue
                
                spots_init = math.ceil(s_budget / unit_net_sum)
                penalty = 1.1 if spots_init < db["Std_Spots"] else 1.0
                spots_final = math.ceil(s_budget / (unit_net_sum * penalty))
                if spots_final % 2 != 0: spots_final += 1
                if spots_final == 0: spots_final = 2
                
                sch = calculate_schedule(spots_final, days_count)
                
                debug_logs.append({
                    "media": m, "sec": sec, "budget": s_budget, 
                    "unit_cost": unit_net_sum * penalty, "spots": spots_final, 
                    "std": db["Std_Spots"], "factor": factor, 
                    "status": "未達標" if penalty > 1 else "達標",
                    "reason": f"懲罰 x1.1" if penalty > 1 else "費率正常"
                })
                
                for r in display_regs:
                    rate_list = int((db[r][0] / db["Std_Spots"]) * factor)
                    pkg_list = rate_list * spots_final
                    is_start = (cfg["is_national"] and r == "北區")
                    
                    rows.append({
                        "media_type": m, "region": r, 
                        "program_num": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec,
                        "spots": spots_final, "schedule": sch,
                        "rate_list": rate_list, "pkg_display_val": pkg_list,
                        "is_pkg_start": is_start, "is_pkg_member": cfg["is_national"]
                    })

            elif m == "家樂福":
                db = PRICING_DB["家樂福"]
                base_std = db["量販_全省"]["Std_Spots"]
                unit_net = (db["量販_全省"]["Net"] / base_std) * factor
                
                spots_init = math.ceil(s_budget / unit_net)
                penalty = 1.1 if spots_init < base_std else 1.0
                spots_final = math.ceil(s_budget / (unit_net * penalty))
                if spots_final % 2 != 0: spots_final += 1
                
                sch_h = calculate_schedule(spots_final, days_count)
                
                debug_logs.append({
                    "media": m, "sec": sec, "budget": s_budget, 
                    "unit_cost": unit_net * penalty, "spots": spots_final, 
                    "std": base_std, "factor": factor,
                    "status": "未達標" if penalty > 1 else "達標",
                    "reason": f"懲罰 x1.1" if penalty > 1 else "費率正常"
                })
                
                rate_h = int((db["量販_全省"]["List"] / base_std) * factor)
                rows.append({
                    "media_type": m, "region": "全省量販", 
                    "program_num": STORE_COUNTS_NUM["家樂福_量販"],
                    "daypart": db["量販_全省"]["Day_Part"], "seconds": sec,
                    "spots": spots_final, "schedule": sch_h,
                    "rate_list": rate_h, "pkg_display_val": rate_h * spots_final,
                    "is_pkg_start": False, "is_pkg_member": False
                })
                
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({
                    "media_type": m, "region": "全省超市", 
                    "program_num": STORE_COUNTS_NUM["家樂福_超市"],
                    "daypart": db["超市_全省"]["Day_Part"], "seconds": sec,
                    "spots": spots_s, "schedule": sch_s,
                    "rate_list": "計量販", "pkg_display_val": "計量販",
                    "is_pkg_start": False, "is_pkg_member": False
                })

# Output
st.divider()
p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}" if rows else ""
rem = get_remarks_text(sign_deadline, billing_month, payment_date)

with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
    st.markdown("#### 1. 本次預算分配 (Waterfall)")
    for log in debug_logs:
        color = "green" if log["status"] == "達標" else "red"
        st.markdown(f"""
        * **{log['media']} ({log['sec']}秒)**: 
            * 分配預算: `${log['budget']:,.0f}`
            * 實收單檔成本 (Net/Std × Factor): `${log['unit_cost']:.2f}` (含 {log['factor']}x 係數)
            * 試算檔次: `{log['spots']}` (基準: {log['std']})
            * 狀態: <span style='color:{color}'><b>{log['status']}</b></span> ({log['reason']})
        """, unsafe_allow_html=True)

    st.markdown("#### 2. 通用規則備註")
    st.markdown("""
    * **優先順序**：廣播 -> 新鮮視 -> 家樂福 (餘額全包)
    * **未達標加價**：若計算檔次 < 基準，成本(Net) 自動 **x 1.1**
    * **偶數修正**：所有檔次無條件進位並 **強制轉為偶數**
    * **Excel 顯示**：Rate 與 Package-cost 皆顯示 **牌價 (List Price)** 以凸顯折扣
    """)

if rows:
    st.components.v1.html(generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem), height=600, scrolling=True)
    
    if template_bytes:
        try:
            xlsx = generate_excel_from_template(format_type, start_date, end_date, client_name, p_str, rows, rem, template_bytes)
            st.download_button("下載 Excel", xlsx, f"Cue_{client_name}.xlsx")
            
            pdf_bytes, err = html_to_pdf_weasyprint(generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem))
            if pdf_bytes:
                st.download_button("下載 PDF", pdf_bytes, f"Cue_{client_name}.pdf")
                st.caption("PDF 來源: WeasyPrint (HTML->PDF)")
            else:
                st.error(f"PDF 產出失敗 (請確保 requirements/packages 已更新): {err}")
        except Exception as e:
            st.error(f"Excel 產出錯誤: {e}")
    else:
        st.warning("請上傳模板以啟用下載。")
