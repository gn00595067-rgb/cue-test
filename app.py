import streamlit as st
import math
import io
import os
import shutil
import tempfile
import subprocess
import re
import base64
from datetime import timedelta, datetime, date
from copy import copy

import requests
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# =========================================================
# 0. 基礎工具
# =========================================================
def parse_count_to_int(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return int(x)
    s = str(x)
    m = re.findall(r"[\d,]+", s)
    if not m: return 0
    return int(m[0].replace(",", ""))

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def html_escape(s):
    if s is None: return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")

# =========================================================
# 1. 頁面設定 & 自動載入
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v66.1 (Value Anchor)")

GOOGLE_DRIVE_FILE_ID = "11R1SA_hpFD5O_MGmYeh4BdtcUhK2bPta"
DEFAULT_FILENAME = "1209-Cue表相關資料.xlsx"

@st.cache_resource(ttl=600)
def load_default_template():
    status_msg = []
    if GOOGLE_DRIVE_FILE_ID:
        url = f"https://drive.google.com/uc?export=download&id={GOOGLE_DRIVE_FILE_ID}"
        try:
            r = requests.get(url, timeout=20, allow_redirects=True)
            if r.status_code == 200:
                if b"<!DOCTYPE html>" in r.content[:500]:
                    status_msg.append("⚠️ Drive 下載權限受限")
                else:
                    return r.content, "雲端硬碟 (Google Drive)", status_msg
        except Exception as e:
            status_msg.append(f"❌ 連線錯誤: {e}")

    if os.path.exists(DEFAULT_FILENAME):
        try:
            with open(DEFAULT_FILENAME, "rb") as f:
                return f.read(), "系統主機 (Local)", status_msg
        except: pass
    
    return None, None, status_msg

# =========================================================
# 2. GPT 核心引擎：Excel 轉 PDF
# =========================================================
def find_soffice_path():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice: return soffice
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in candidates:
            if os.path.exists(p): return p
    return None

def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    if os.name == "nt":
        try:
            import win32com.client
            with tempfile.TemporaryDirectory() as tmp:
                xlsx_path = os.path.join(tmp, "cue.xlsx")
                pdf_path = os.path.join(tmp, "cue.pdf")
                with open(xlsx_path, "wb") as f: f.write(xlsx_bytes)

                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = None
                try:
                    wb = excel.Workbooks.Open(xlsx_path)
                    wb.ExportAsFixedFormat(0, pdf_path)
                except: pass
                finally:
                    if wb: 
                        try: wb.Close(False)
                        except: pass
                    try: excel.Quit()
                    except: pass

                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f: return f.read(), "Excel App (Local)", ""
        except: pass

    soffice = find_soffice_path()
    if soffice:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                xlsx_path = os.path.join(tmp, "cue.xlsx")
                with open(xlsx_path, "wb") as f: f.write(xlsx_bytes)

                subprocess.run(
                    [soffice, "--headless", "--nologo", "--convert-to", "pdf", "--outdir", tmp, xlsx_path],
                    capture_output=True, timeout=60
                )
                
                pdf_path = os.path.join(tmp, "cue.pdf")
                if not os.path.exists(pdf_path):
                    for fn in os.listdir(tmp):
                        if fn.endswith(".pdf"): pdf_path = os.path.join(tmp, fn); break
                
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f: return f.read(), "LibreOffice", ""
                
                return None, "Fail", "LibreOffice 轉檔無輸出"
        except Exception as e:
            return None, "Fail", str(e)

    return None, "Fail", "無可用的 Excel 轉檔引擎"

# =========================================================
# 3. WeasyPrint Fallback
# =========================================================
def html_to_pdf_fallback(html_str, font_b64):
    try: 
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        font_config = FontConfiguration()
        css_str = """
        @page { size: A4 landscape; margin: 0.5cm; }
        body { font-family: 'NotoSansTC', sans-serif !important; font-size: 8pt; }
        table { width: 100%; border-collapse: collapse; table-layout: fixed; }
        th, td { border: 0.5pt solid #000; padding: 2px; text-align: center; white-space: nowrap; overflow: hidden; }
        .bg-dw-head { background-color: #4472C4; color: white; -webkit-print-color-adjust: exact; }
        .bg-sh-head { background-color: #BDD7EE; color: black; -webkit-print-color-adjust: exact; }
        .bg-weekend { background-color: #FFD966; -webkit-print-color-adjust: exact; }
        .bg-total   { background-color: #FFF2CC; -webkit-print-color-adjust: exact; }
        .left { text-align: left !important; }
        .right { text-align: right !important; }
        tr { page-break-inside: avoid; }
        """
        if font_b64:
            css_str = f"@font-face {{ font-family: 'NotoSansTC'; src: url(data:font/ttf;base64,{font_b64}) format('truetype'); }} " + css_str
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[CSS(string=css_str)], font_config=font_config)
        return pdf_bytes, ""
    except Exception as e:
        return None, str(e)

# =========================================================
# 4. 資料庫 (2026 核心)
# =========================================================
STORE_COUNTS_RAW = {
    "全省": "4,437店", "北區": "1,649店", "桃竹苗": "779店", "中區": "839店", "雲嘉南": "499店", "高屏": "490店", "東區": "181店",
    "新鮮視_全省": "3,124面", "新鮮視_北區": "1,127面", "新鮮視_桃竹苗": "616面", "新鮮視_中區": "528面",
    "新鮮視_雲嘉南": "365面", "新鮮視_高屏": "405面", "新鮮視_東區": "83面",
    "家樂福_量販": "67店", "家樂福_超市": "250店"
}
STORE_COUNTS_NUM = {k: parse_count_to_int(v) for k, v in STORE_COUNTS_RAW.items()}
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

PRICING_DB = {
    "全家廣播": { "Std_Spots": 480, "Day_Part": "00:00-24:00", 
        "全省": [400000, 320000], 
        "北區": [250000, 200000], "桃竹苗": [150000, 120000], "中區": [150000, 120000],
        "雲嘉南": [100000, 80000], "高屏": [100000, 80000], "東區": [62500, 50000] },
    "新鮮視": { "Std_Spots": 504, "Day_Part": "07:00-22:00", 
        "全省": [150000, 120000], 
        "北區": [150000, 120000], "桃竹苗": [120000, 96000], "中區": [90000, 72000],
        "雲嘉南": [75000, 60000], "高屏": [75000, 60000], "東區": [45000, 36000] },
    "家樂福": {
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
# 5. Excel 生成模組 (Dynamic Rebuild)
# =========================================================
def _get_master_cell(ws, cell):
    if not isinstance(cell, MergedCell): return cell
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= cell.row <= mr.max_row and mr.min_col <= cell.column <= mr.max_col:
            return ws.cell(row=mr.min_row, column=mr.min_col)
    return None

def safe_write(ws, addr, value):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master: master.value = value
    else: cell.value = value

def safe_write_rc(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master: master.value = value
    else: cell.value = value

def apply_center_style(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, indent=al.indent)

def copy_row_with_style_fix(ws, src_row, dst_row, max_col):
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
        else: dc.value = v

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

def find_cell_exact(ws, text):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == text: return cell.row, cell.column
    return None

def find_first_row_contains(ws, col_letter, keyword):
    col_idx = column_index_from_string(col_letter)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if isinstance(v, str) and keyword in v: return r
    return None

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

def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    target_sheet = meta["sheet_name"]
    if target_sheet not in wb.sheetnames: raise ValueError(f"缺少分頁：{target_sheet}")
    
    for s in list(wb.sheetnames):
        if s != target_sheet: del wb[s]
    ws = wb[target_sheet]

    hc = meta["header_cells"]
    if "client" in hc: safe_write(ws, hc["client"], client_name)
    if "product" in hc: safe_write(ws, hc["product"], product_display_str)
    if "period" in hc: safe_write(ws, hc["period"], f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y.%m. %d')}")
    if "medium" in hc: safe_write(ws, hc["medium"], " ".join(sorted(set([r["media_type"] for r in rows]))))
    if "month" in hc: safe_write(ws, hc["month"], f" {start_dt.month}月")
    safe_write(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))
    for addr, text in meta.get("header_override", {}).items(): safe_write(ws, addr, text)

    total_cell = find_cell_exact(ws, meta["total_label"])
    if not total_cell: raise ValueError("找不到 Total")
    total_row_orig = total_cell[0]
    cols = meta["cols"]
    
    sec_start = {}
    for m_key, kw in meta["anchors"].items():
        r0 = find_first_row_contains(ws, cols["station"], kw)
        if r0: sec_start[m_key] = r0
    
    sec_order = sorted(sec_start.items(), key=lambda x: x[1], reverse=True)
    written_ranges = [] 

    reg_map = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}
    def sort_key(x): return (x["seconds"], reg_map.get(x["region"], 999))
    grouped_data = {
        "全家廣播": sorted([r for r in rows if r["media_type"] == "全家廣播"], key=sort_key),
        "新鮮視": sorted([r for r in rows if r["media_type"] == "新鮮視"], key=sort_key),
        "家樂福": sorted([r for r in rows if r["media_type"] == "家樂福"], key=sort_key),
    }
    
    current_end_marker = total_row_orig - 1
    
    def station_title(m):
        prefix = "全家便利商店\n" if m != "家樂福" else ""
        name = "通路廣播廣告" if m == "全家廣播" else "新鮮視廣告" if m == "新鮮視" else "家樂福"
        if format_type == "Shenghuo" and m == "全家廣播": name = "廣播通路廣告"
        return prefix + name

    for i, (m_key, start_row_orig) in enumerate(sec_order):
        style_source_row = start_row_orig + 1
        rows_to_delete = max(0, current_end_marker - style_source_row)
        
        data = grouped_data.get(m_key, [])
        needed = len(data)
        
        if needed == 0:
            if rows_to_delete > 0: ws.delete_rows(style_source_row + 1, amount=rows_to_delete)
            for c in range(1, ws.max_column+1): safe_write_rc(ws, style_source_row, c, None)
            current_end_marker = start_row_orig - 1
            continue

        if rows_to_delete > 0: ws.delete_rows(style_source_row + 1, amount=rows_to_delete)
        if needed > 1:
            ws.insert_rows(style_source_row + 1, amount=needed - 1)
            for r_idx in range(style_source_row + 1, style_source_row + 1 + needed - 1):
                copy_row_with_style_fix(ws, style_source_row, r_idx, ws.max_column)
        
        curr_row = style_source_row
        
        if meta["station_merge"] and needed > 0:
            unmerge_col_overlap(ws, cols["station"], curr_row, curr_row + needed - 1)
            merge_rng = f"{cols['station']}{curr_row}:{cols['station']}{curr_row + needed - 1}"
            ws.merge_cells(merge_rng)
            top_cell = ws[f"{cols['station']}{curr_row}"]
            top_cell.value = station_title(m_key)
            apply_center_style(top_cell)

        for idx, r_data in enumerate(data):
            if not meta["station_merge"]:
                cell = ws[f"{cols['station']}{curr_row}"]
                cell.value = station_title(m_key)
                apply_center_style(cell)
            
            safe_write(ws, f"{cols['location']}{curr_row}", region_display(r_data["region"]))
            prog_val = r_data.get("program_num", parse_count_to_int(r_data.get("program", 0)))
            safe_write(ws, f"{cols['program']}{curr_row}", int(prog_val))

            if format_type == "Dongwu":
                safe_write(ws, f"{cols['daypart']}{curr_row}", r_data["daypart"])
                if m_key == "家樂福": safe_write(ws, f"{cols['seconds']}{curr_row}", f"{r_data['seconds']}秒")
                else: safe_write(ws, f"{cols['seconds']}{curr_row}", int(r_data["seconds"]))
                
                # 🌟 [關鍵修正]：每一列都顯示分區定價 (Value Anchor)
                safe_write(ws, f"{cols['rate']}{curr_row}", r_data["rate_list"])
                safe_write(ws, f"{cols['pkg']}{curr_row}", r_data["pkg_display_val"])
            else:
                safe_write(ws, f"{cols['daypart']}{curr_row}", r_data["daypart"])
                safe_write(ws, f"{cols['seconds']}{curr_row}", f"{r_data['seconds']}秒廣告")
                safe_write(ws, f"{cols['proj_price']}{curr_row}", r_data["pkg_display_val"] if isinstance(r_data["pkg_display_val"], int) else 0)

            set_schedule(ws, curr_row, meta["schedule_start_col"], meta["max_days"], r_data["schedule"])
            spot_sum = sum(r_data["schedule"][:meta["max_days"]])
            safe_write(ws, f"{meta['total_col']}{curr_row}", spot_sum)
            curr_row += 1
            
        written_ranges.append((curr_row - needed, curr_row - 1))
        current_end_marker = start_row_orig - 1

    total_cell = find_cell_exact(ws, meta["total_label"])
    if not total_cell: raise ValueError("找不到 Total")
    total_row = total_cell[0]

    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [sum([x["schedule"][d] for x in rows if d < len(x["schedule"])]) for d in range(eff_days)]
    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))
    
    # 🌟 [總金額覆蓋邏輯]：全省聯播時，Total 用 National Package Price
    total_pkg = 0
    # 先計算家樂福的部分 (它沒有全省/分區之分，直接累加)
    # 然後處理廣播/新鮮視：如果是全省聯播，直接加上全省總價
    
    # 為了簡化，我們重新掃描 Rows 來計算
    # 這裡需要一個 flag 來避免重複計算全省的總價
    processed_national_media = set()
    
    for r in rows:
        m = r["media_type"]
        val = r["pkg_display_val"] if isinstance(r["pkg_display_val"], int) else 0
        
        if r.get("is_national_display"):
            if m not in processed_national_media:
                # 找到對應的全省定價 (List)
                # 從 r 裡面反推有點麻煩，直接從 PRICING_DB 拿最準
                factor = SEC_FACTORS[m][r["seconds"]]
                std = PRICING_DB[m]["Std_Spots"]
                nat_list_price = PRICING_DB[m]["全省"][0] # 400k or 150k
                
                # 計算全省總價 = 單檔定價 * 總檔次
                # 這裡的 "總檔次" 是 "全省檔次" (1766)，不是 6 區加總 (10596)
                # rows 裡的 spots 已經是 1766
                # Rate = 400k / 480 * Factor
                # Total = Rate * 1766
                nat_rate = int((nat_list_price / std) * factor)
                nat_total = nat_rate * r["spots"]
                
                total_pkg += nat_total
                processed_national_media.add(m)
        else:
            total_pkg += val

    pkg_col = cols.get("pkg") or cols.get("proj_price")
    safe_write(ws, f"{pkg_col}{total_row}", total_pkg)

    lbl = meta["footer_labels"]
    make_fee = 10000 
    pos_make = find_cell_exact(ws, lbl["make"])
    if pos_make:
        v = ws.cell(pos_make[0], pos_make[1]+1).value
        if isinstance(v, (int, float)) and v > 0: make_fee = int(v)
        else: safe_write_rc(ws, pos_make[0], pos_make[1]+1, make_fee)
    
    vat = int(round((total_pkg + make_fee) * 0.05))
    pos_vat = find_cell_exact(ws, lbl["vat"])
    if pos_vat: safe_write_rc(ws, pos_vat[0], pos_vat[1]+1, vat)
    
    pos_grand = find_cell_exact(ws, lbl["grand"])
    if pos_grand: safe_write_rc(ws, pos_grand[0], pos_grand[1]+1, total_pkg + make_fee + vat)

    rem_pos = find_cell_exact(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_write_rc(ws, rem_pos[0] + 1 + i, rem_pos[1], rm)

    if format_type == "Dongwu":
        force_center_columns_range(ws, meta["force_center_cols"], min(r[0] for r in written_ranges) if written_ranges else total_row, total_row)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 6. HTML Preview
# =========================================================
def load_font_base64():
    font_path = "NotoSansTC-Regular.ttf"
    if os.path.exists(font_path):
        with open(font_path, "rb") as f: return base64.b64encode(f.read()).decode("utf-8")
    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/TTF/TraditionalChinese/NotoSansTC-Regular.ttf"
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            with open(font_path, "wb") as f: f.write(r.content)
            return base64.b64encode(r.content).decode("utf-8")
    except: pass
    return None

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks):
    header_cls = "bg-dw-head" if format_type == "Dongwu" else "bg-sh-head"
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    eff_days = min(days_cnt, 31)
    
    st.markdown(f"""<style>
    .excel-container {{ overflow-x: auto; }}
    .excel-table {{ width: 100%; border-collapse: collapse; min-width: 1200px; font-family: Arial, sans-serif; font-size: 12px; }}
    .excel-table th, .excel-table td {{ border: 1px solid #999; padding: 4px; text-align: center; white-space: nowrap; height: 24px; }}
    .bg-dw-head {{ background-color: #4472C4; color: white; font-weight: bold; }}
    .bg-sh-head {{ background-color: #BDD7EE; color: black; font-weight: bold; }}
    .bg-weekend {{ background-color: #FFD966; color: black; }}
    .bg-total   {{ background-color: #FFF2CC; font-weight: bold; }}
    .left {{ text-align: left !important; padding-left: 5px; }}
    .right {{ text-align: right !important; padding-right: 5px; font-family: Consolas, monospace; }}
    .remarks {{ margin-top: 15px; font-size: 13px; text-align: left; line-height: 1.5; }}
    </style>""", unsafe_allow_html=True)

    date_th1, date_th2 = "", ""
    curr = start_dt
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for i in range(eff_days):
        wd = curr.weekday()
        bg = "bg-weekend" if (format_type == "Dongwu" and wd >= 5) else header_cls
        if format_type == "Shenghuo": bg = header_cls 
        date_th1 += f"<th class='{bg} col_day'>{curr.day}</th>"
        date_th2 += f"<th class='{bg} col_day'>{weekdays[wd]}</th>"
        curr += timedelta(days=1)

    if format_type == "Dongwu":
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(List)", "Package<br>(List)"]
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "專案價"]
    th_fixed = "".join([f"<th rowspan='2'>{c}</th>" for c in cols_def])
    
    rows_sorted = sorted(rows, key=lambda x: (media_order.get(x["media_type"], 99), x["seconds"], REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 99))
    tbody = ""
    media_counts = {}
    for r in rows_sorted: media_counts[r["media_type"]] = media_counts.get(r["media_type"], 0) + 1
    media_printed = {m: False for m in media_counts}

    for idx, r in enumerate(rows_sorted):
        m = r["media_type"]
        tbody += "<tr>"
        if not media_printed[m]:
            rowspan = media_counts[m]
            display_name = "全家便利商店<br>通路廣播廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
            if format_type == "Shenghuo" and m == "全家廣播": display_name = "全家便利商店<br>廣播通路廣告"
            if format_type == "Shenghuo": tbody += f"<td class='left'>{display_name}</td>"
            else: tbody += f"<td class='left' rowspan='{rowspan}'>{display_name}</td>"; media_printed[m] = True
        elif format_type == "Shenghuo":
             display_name = "全家便利商店<br>廣播通路廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
             tbody += f"<td class='left'>{display_name}</td>"

        tbody += f"<td>{region_display(r['region'])}</td><td class='right'>{r.get('program_num','')}</td><td>{r['daypart']}</td>"
        sec_txt = f"{r['seconds']}秒" if format_type=="Dongwu" and m=="家樂福" else f"{r['seconds']}" if format_type=="Dongwu" else f"{r['seconds']}秒廣告"
        tbody += f"<td>{sec_txt}</td>"
        
        # 顯示邏輯：全部顯示 (Value Anchor)
        rate = f"{r['rate_list']:,}" if isinstance(r['rate_list'], int) else r['rate_list']
        pkg = f"{r['pkg_display_val']:,}" if isinstance(r['pkg_display_val'], int) else r['pkg_display_val']
        
        if format_type == "Dongwu": tbody += f"<td class='right'>{rate}</td><td class='right'>{pkg}</td>"
        else: tbody += f"<td class='right'>{pkg}</td>"
        
        for d in r['schedule'][:eff_days]: tbody += f"<td>{d}</td>"
        tbody += f"<td class='bg-total'>{sum(r['schedule'])}</td></tr>"

    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]
    
    # 總金額計算 (Total Override)
    total_pkg = 0
    processed_national = set()
    for r in rows:
        m = r["media_type"]
        val = r["pkg_display_val"] if isinstance(r["pkg_display_val"], int) else 0
        if r.get("is_national_display"):
            if m not in processed_national:
                factor = SEC_FACTORS[m][r["seconds"]]
                std = PRICING_DB[m]["Std_Spots"]
                nat_list = PRICING_DB[m]["全省"][0]
                total_pkg += int((nat_list / std) * factor) * r["spots"]
                processed_national.add(m)
        else:
            total_pkg += val
            
    colspan = 5; empty_td = "<td></td>" if format_type == "Dongwu" else ""
    tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td>{empty_td}<td class='right'>{total_pkg:,}</td>"
    for t in totals: tfoot += f"<td>{t}</td>"
    tfoot += f"<td>{sum(totals)}</td></tr>"

    return f"""<div class="excel-container"><div style="margin-bottom:10px;"><b>客戶：</b>{c_name} &nbsp; <b>產品：</b>{p_display}<br><span style="color:#666;">走期：{start_dt} ~ {end_dt}</span></div><table class="excel-table"><thead><tr>{th_fixed}{date_th1}<th class='{header_cls}' rowspan='2'>檔次</th></tr><tr>{date_th2}</tr></thead><tbody>{tbody}{tfoot}</tbody></table><div class="remarks"><b>Remarks：</b><br>{"<br>".join(remarks)}</div></div>"""

# =========================================================
# 7. UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v66.1: Value Anchor)")

auto_tpl, source, msgs = load_default_template()
template_bytes = auto_tpl

if auto_tpl:
    st.success(f"✅ 已載入系統公版 ({source})")
else:
    st.warning("⚠️ 無法載入公版，請手動上傳")
    tpl = st.file_uploader("上傳 Excel 模板", type=["xlsx"])
    if tpl: template_bytes = tpl.read()

st.markdown("### 1. 選擇格式")
format_type = st.radio("", ["Dongwu", "Shenghuo"], horizontal=True, label_visibility="collapsed")

st.markdown("### 2. 基本資料設定")
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

st.markdown("### 3. 媒體投放設定")

# 1. 狀態初始化
if "rad_share" not in st.session_state: st.session_state.rad_share = 100
if "fv_share" not in st.session_state: st.session_state.fv_share = 0
if "cf_share" not in st.session_state: st.session_state.cf_share = 0

# 2. 自動平衡 Callback
def on_media_change():
    active = []
    if st.session_state.get("cb_rad"): active.append("rad_share")
    if st.session_state.get("cb_fv"): active.append("fv_share")
    if st.session_state.get("cb_cf"): active.append("cf_share")
    if not active: return
    share = 100 // len(active)
    for key in active: st.session_state[key] = share
    rem = 100 - sum([st.session_state[k] for k in active])
    st.session_state[active[0]] += rem

def on_slider_change(changed_key):
    active = []
    if st.session_state.get("cb_rad"): active.append("rad_share")
    if st.session_state.get("cb_fv"): active.append("fv_share")
    if st.session_state.get("cb_cf"): active.append("cf_share")
    others = [k for k in active if k != changed_key]
    if not others: st.session_state[changed_key] = 100
    elif len(others) == 1:
        val = st.session_state[changed_key]
        st.session_state[others[0]] = max(0, 100 - val)
    elif len(others) == 2:
        val = st.session_state[changed_key]
        rem = max(0, 100 - val)
        k1, k2 = others[0], others[1]
        sum_others = st.session_state[k1] + st.session_state[k2]
        if sum_others == 0: st.session_state[k1] = rem // 2; st.session_state[k2] = rem - st.session_state[k1]
        else:
            ratio = st.session_state[k1] / sum_others
            st.session_state[k1] = int(rem * ratio)
            st.session_state[k2] = rem - st.session_state[k1]

# 3. 媒體勾選區
st.write("請勾選要投放的媒體：")
col_cb1, col_cb2, col_cb3 = st.columns(3)
with col_cb1: is_rad = st.checkbox("全家廣播", value=True, key="cb_rad", on_change=on_media_change)
with col_cb2: is_fv = st.checkbox("新鮮視", value=False, key="cb_fv", on_change=on_media_change)
with col_cb3: is_cf = st.checkbox("家樂福", value=False, key="cb_cf", on_change=on_media_change)

m1, m2, m3 = st.columns(3)
config = {}

if is_rad:
    with m1:
        st.markdown("#### 📻 全家廣播")
        is_nat = st.checkbox("全省聯播", True, key="rad_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_reg")
        secs = st.multiselect("秒數", DURATIONS, [20], key="rad_sec")
        st.slider("預算 %", 0, 100, key="rad_share", on_change=on_slider_change, args=("rad_share",))
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in sorted(secs)[:-1]: v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"rs_{s}"); sec_shares[s] = v; ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs: sec_shares[secs[0]] = 100
        config["全家廣播"] = {"is_national": is_nat, "regions": regs, "seconds": sorted(secs), "share": st.session_state.rad_share, "sec_shares": sec_shares}

if is_fv:
    with m2:
        st.markdown("#### 📺 新鮮視")
        is_nat = st.checkbox("全省聯播", False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
        secs = st.multiselect("秒數", DURATIONS, [10], key="fv_sec")
        st.slider("預算 %", 0, 100, key="fv_share", on_change=on_slider_change, args=("fv_share",))
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in sorted(secs)[:-1]: v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"fs_{s}"); sec_shares[s] = v; ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs: sec_shares[secs[0]] = 100
        config["新鮮視"] = {"is_national": is_nat, "regions": regs, "seconds": sorted(secs), "share": st.session_state.fv_share, "sec_shares": sec_shares}

if is_cf:
    with m3:
        st.markdown("#### 🛒 家樂福")
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.slider("預算 %", 0, 100, key="cf_share", on_change=on_slider_change, args=("cf_share",))
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in sorted(secs)[:-1]: v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"cs_{s}"); sec_shares[s] = v; ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs: sec_shares[secs[0]] = 100
        config["家樂福"] = {"regions": ["全省"], "seconds": sorted(secs), "share": st.session_state.cf_share, "sec_shares": sec_shares}

# ---------------------------------------------------------
# 計算引擎
# ---------------------------------------------------------
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
                
                if cfg["is_national"]:
                    calc_regs = ["全省"]
                    display_regs = REGIONS_ORDER # 展開6區
                else:
                    calc_regs = cfg["regions"]
                    display_regs = cfg["regions"]
                
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
                
                for i, r in enumerate(display_regs):
                    is_national_display = cfg["is_national"]
                    is_primary_pricing_row = (i == 0)
                    
                    # 顯示邏輯：每列都顯示該區的 List Rate (Anchor)
                    rate_list = int((db[r][0] / db["Std_Spots"]) * factor)
                    pkg_list = rate_list * spots_final
                    
                    rows.append({
                        "media_type": m, "region": r, 
                        "program_num": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec,
                        "spots": spots_final, "schedule": sch,
                        "rate_list": rate_list, "pkg_display_val": pkg_list,
                        "is_national_display": is_national_display,
                        "is_primary_pricing_row": is_primary_pricing_row
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
                rows.append({"media_type": m, "region": "全省量販", "program_num": STORE_COUNTS_NUM["家樂福_量販"], "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch_h, "rate_list": rate_h, "pkg_display_val": rate_h * spots_final})
                
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({"media_type": m, "region": "全省超市", "program_num": STORE_COUNTS_NUM["家樂福_超市"], "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "spots": spots_s, "schedule": sch_s, "rate_list": "計量販", "pkg_display_val": "計量販"})

p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}" if rows else ""
rem = get_remarks_text(sign_deadline, billing_month, payment_date)

with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
    st.markdown("#### 1. 本次預算分配 (Waterfall)")
    for log in debug_logs:
        color = "green" if log["status"] == "達標" else "red"
        st.markdown(f"**{log['media']} ({log['sec']}秒)**: 預算${log['budget']:,.0f} | 執行{log['spots']}檔 -> <span style='color:{color}'><b>{log['status']}</b></span>", unsafe_allow_html=True)

if rows:
    font_b64 = load_font_base64()
    html = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem)
    st.components.v1.html(html, height=700, scrolling=True)
    
    if template_bytes:
        try:
            xlsx = generate_excel_from_template(format_type, start_date, end_date, client_name, p_str, rows, rem, template_bytes)
            st.download_button("下載 Excel", xlsx, f"Cue_{client_name}.xlsx")
            
            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx)
            
            if pdf_bytes:
                st.download_button(f"下載 PDF ({method})", pdf_bytes, f"Cue_{client_name}.pdf")
            else:
                st.warning(f"Excel 轉 PDF 失敗 ({method}: {err})，切換至備用渲染引擎 (HTML)...")
                pdf_bytes, err = html_to_pdf_fallback(html, font_b64)
                if pdf_bytes: st.download_button("下載 PDF (Fallback)", pdf_bytes, f"Cue_{client_name}.pdf")
                else: st.error(f"PDF 產出失敗: {err}")
                
        except Exception as e: st.error(f"Excel 產出錯誤: {e}")
    else: st.warning("請上傳模板以啟用下載。")
