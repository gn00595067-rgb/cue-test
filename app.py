# -*- coding: utf-8 -*-
"""
Cue Sheet Pro v66.0
- Block Renderer (Dongwu/Shenghuo)
- Value Anchor Logic aligned (Master Spec v1.01+)
- Fix: Total schedule MUST NOT multiply by region rows (national buy expands rows)
- Dongwu: G=Rate(Net header but LIST-based anchor total), H=Package-cost(Net actual)
- Total Row: G uses National List override for national-buys; H sums Net totals
- Footer polish: VAT block align + contract (甲乙方) spacing tidy

Run:
  streamlit run freshvision_closer.py
"""

# -----------------------------
# 0) Imports (with friendly error)
# -----------------------------
try:
    import streamlit as st
except ModuleNotFoundError:
    raise SystemExit(
        "找不到 streamlit。\n\n"
        "請先安裝：\n"
        "  py -m pip install streamlit\n\n"
        "並用下列方式啟動（不要用 IDLE 直接跑）：\n"
        "  streamlit run freshvision_closer.py\n"
    )

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
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font

# =========================================================
# 1) Master Configuration (align to your spec)
# =========================================================
STORE_COUNTS_CONFIG = {
    "全省": 4437,
    "北區": 1649, "桃竹苗": 779, "中區": 839, "雲嘉南": 499, "高屏": 490, "東區": 181,
    "新鮮視_全省": 3124,
    "新鮮視_北區": 1127, "新鮮視_桃竹苗": 616, "新鮮視_中區": 528,
    "新鮮視_雲嘉南": 365, "新鮮視_高屏": 405, "新鮮視_東區": 83,
    "家樂福_量販": 68,   # 更新
    "家樂福_超市": 249,   # 更新
}

PRICING_CONFIG = {
    "全家廣播": {
        "Std_Spots": 480,
        "Day_Part": "00:00-24:00",
        "Prices": {
            "全省":   {"List": 400000, "Net": 320000},
            "北區":   {"List": 250000, "Net": 200000},
            "桃竹苗": {"List": 150000, "Net": 120000},
            "中區":   {"List": 150000, "Net": 120000},
            "雲嘉南": {"List": 100000, "Net": 80000},
            "高屏":   {"List": 100000, "Net": 80000},
            "東區":   {"List": 62500,  "Net": 50000},
        }
    },
    "新鮮視": {
        "Std_Spots": 504,
        "Day_Part": "07:00-22:00",
        "Prices": {
            "全省":   {"List": 150000, "Net": 120000},
            "北區":   {"List": 150000, "Net": 120000},
            "桃竹苗": {"List": 120000, "Net": 96000},
            "中區":   {"List": 90000,  "Net": 72000},
            "雲嘉南": {"List": 75000,  "Net": 60000},
            "高屏":   {"List": 75000,  "Net": 60000},
            "東區":   {"List": 45000,  "Net": 36000},
        }
    },
    "家樂福": {
        "Configs": {
            "量販_全省": {"List": 300000, "Net": 250000, "Std": 420, "Time": "09:00-23:00"},
            "超市_全省": {"List": 100000, "Net": 80000,  "Std": 720, "Time": "00:00-24:00"},
        }
    }
}

SEC_FACTORS_CONFIG = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5, 5: 0.25},
    "新鮮視":   {30: 3.0, 20: 2.0,  15: 1.5,  10: 1.0, 5: 0.5},
    "家樂福":   {30: 1.5, 20: 1.0,  15: 0.85, 10: 0.65, 5: 0.35},
}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

REGION_DISPLAY_6 = {
    "北區": "北區-北北基",
    "桃竹苗": "桃區-桃竹苗",
    "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南",
    "高屏": "高屏區-高屏",
    "東區": "東區-宜花東",
    "全省量販": "全省量販",
    "全省超市": "全省超市",
}

# =========================================================
# 2) 基礎工具
# =========================================================
def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def html_escape(s):
    if s is None:
        return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")

def norm_ws(s: str) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def region_display(region: str) -> str:
    return REGION_DISPLAY_6.get(region, region)

def get_sec_factor(media_type, seconds):
    return SEC_FACTORS_CONFIG.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    """平均分配到每天；同你原本邏輯（偶數修正由外面做）"""
    if days <= 0:
        return []
    if total_spots % 2 != 0:
        total_spots += 1
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
        "3.通路店舖數與開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。",
        f"6.付款兌現日期：{p_str}"
    ]

# =========================================================
# 3) 模板載入
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v66.0 (Value Anchor Aligned)")

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
        except:
            pass

    return None, None, status_msg

# =========================================================
# 4) Excel 轉 PDF
# =========================================================
def find_soffice_path():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        return soffice
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    return None

def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    # 1) Windows Excel COM（最像）
    if os.name == "nt":
        try:
            import win32com.client
            with tempfile.TemporaryDirectory() as tmp:
                xlsx_path = os.path.join(tmp, "cue.xlsx")
                pdf_path = os.path.join(tmp, "cue.pdf")
                with open(xlsx_path, "wb") as f:
                    f.write(xlsx_bytes)

                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = None
                try:
                    wb = excel.Workbooks.Open(xlsx_path)
                    wb.ExportAsFixedFormat(0, pdf_path)
                finally:
                    if wb:
                        try:
                            wb.Close(False)
                        except:
                            pass
                    try:
                        excel.Quit()
                    except:
                        pass

                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        return f.read(), "Excel App (Local)", ""
        except:
            pass

    # 2) LibreOffice
    soffice = find_soffice_path()
    if soffice:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                xlsx_path = os.path.join(tmp, "cue.xlsx")
                with open(xlsx_path, "wb") as f:
                    f.write(xlsx_bytes)

                subprocess.run(
                    [soffice, "--headless", "--nologo", "--convert-to", "pdf", "--outdir", tmp, xlsx_path],
                    capture_output=True, timeout=60
                )

                pdf_path = os.path.join(tmp, "cue.pdf")
                if not os.path.exists(pdf_path):
                    for fn in os.listdir(tmp):
                        if fn.endswith(".pdf"):
                            pdf_path = os.path.join(tmp, fn)
                            break

                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        return f.read(), "LibreOffice", ""

                return None, "Fail", "LibreOffice 轉檔無輸出"
        except Exception as e:
            return None, "Fail", str(e)

    return None, "Fail", "無可用的 Excel 轉檔引擎"

def html_to_pdf_fallback(html_str, font_b64):
    try:
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        font_config = FontConfiguration()
        css_str = """
        @page { size: A4 landscape; margin: 0.5cm; }
        body { font-family: 'NotoSansTC', sans-serif !important; font-size: 8pt; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 0.5pt solid #555; padding: 2px; text-align: center; white-space: nowrap; }
        .bg-dw-head { background-color: #4472C4; color: white; }
        .bg-sh-head { background-color: #BDD7EE; color: black; }
        .bg-weekend { background-color: #FFD966; }
        .bg-total   { background-color: #FFF2CC; }
        tr { page-break-inside: avoid; }
        """
        if font_b64:
            css_str = f"@font-face {{ font-family: 'NotoSansTC'; src: url(data:font/ttf;base64,{font_b64}) format('truetype'); }} " + css_str
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[CSS(string=css_str)], font_config=font_config)
        return pdf_bytes, ""
    except Exception as e:
        return None, str(e)

# =========================================================
# 5) Excel 安全寫入 / 複製樣式 / 解除合併
# =========================================================
def _get_master_cell(ws, cell):
    if not isinstance(cell, MergedCell):
        return cell
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= cell.row <= mr.max_row and mr.min_col <= cell.column <= mr.max_col:
            return ws.cell(row=mr.min_row, column=mr.min_col)
    return None

def safe_write(ws, addr, value):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master:
            master.value = value
    else:
        cell.value = value

def safe_write_rc(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master:
            master.value = value
    else:
        cell.value = value

def apply_center_style(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, indent=al.indent)

def apply_right_style(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True, indent=al.indent)

def apply_left_style(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True, indent=al.indent)

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
            try:
                dc.value = Translator(v, origin=sc.coordinate).translate_formula(row_shift=row_shift, col_shift=0)
            except:
                dc.value = v
        else:
            dc.value = v

def force_center_columns_range(ws, col_letters, start_row, end_row):
    if start_row is None or end_row is None:
        return
    for r in range(start_row, end_row + 1):
        for col in col_letters:
            addr = f"{col}{r}"
            cell = ws[addr]
            if isinstance(cell, MergedCell):
                master = _get_master_cell(ws, cell)
                if master:
                    cell = master
                else:
                    continue
            apply_center_style(cell)

def set_schedule(ws, row, start_col_letter, max_days, schedule_list):
    start_col = column_index_from_string(start_col_letter)
    for i in range(max_days):
        v = schedule_list[i] if (schedule_list and i < len(schedule_list)) else None
        safe_write_rc(ws, row, start_col + i, v)

# =========================================================
# 6) Block Renderer helpers
# =========================================================
def find_row_contains_any(ws, col_letters, keyword):
    for col in col_letters:
        col_idx = column_index_from_string(col)
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, col_idx).value
            if isinstance(v, str) and keyword in v:
                return r
    return None

def find_cell_exact_in_ws(ws, text):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == text:
                return cell.row, cell.column
    return None

def unmerge_overlaps_by_cols(ws, start_row, end_row, col_letters):
    if start_row is None or end_row is None or start_row > end_row:
        return
    col_set = set(column_index_from_string(c) for c in col_letters)
    min_c, max_c = min(col_set), max(col_set)

    to_unmerge = []
    for mr in list(ws.merged_cells.ranges):
        if mr.max_row < start_row or mr.min_row > end_row:
            continue
        if mr.max_col < min_c or mr.min_col > max_c:
            continue
        to_unmerge.append(str(mr))

    for s in set(to_unmerge):
        try:
            ws.unmerge_cells(s)
        except:
            pass

def detect_blocks(ws, format_type, meta):
    total_pos = find_cell_exact_in_ws(ws, meta["total_label"])
    if not total_pos:
        raise ValueError("找不到 Total")
    total_row = total_pos[0]

    cols_station = meta["cols"]["station"]
    search_cols = meta.get("anchor_search_cols", [cols_station, "C", "D"])

    rad_kw = meta["anchors"]["全家廣播"]
    rad_start = find_row_contains_any(ws, search_cols, rad_kw)

    fv_kw = meta["anchors"]["新鮮視"]
    fv_start = find_row_contains_any(ws, search_cols, fv_kw)

    cf_kw = meta["anchors"]["家樂福"]
    cf_anchor = find_row_contains_any(ws, search_cols, cf_kw)
    cf_offset = meta.get("carrefour_data_offset", 0)
    cf_start = (cf_anchor + cf_offset) if cf_anchor else None

    starts = []
    if rad_start:
        starts.append(("全家廣播", rad_start))
    if fv_start:
        starts.append(("新鮮視", fv_start))
    if cf_start:
        starts.append(("家樂福", cf_start))

    starts.sort(key=lambda x: x[1])

    blocks = []
    for i, (m, srow) in enumerate(starts):
        if i + 1 < len(starts):
            erow = starts[i + 1][1] - 1
        else:
            erow = total_row - 1
        if erow >= srow:
            blocks.append({"media": m, "start": srow, "end": erow})

    if not blocks:
        raise ValueError("找不到任何媒體區段 blocks（anchors 可能跟模板不一致）")

    return blocks, total_row

def station_title(format_type, media_key):
    if media_key == "全家廣播":
        return "全家便利商店\n通路廣播廣告" if format_type == "Dongwu" else "全家便利商店\n廣播通路廣告"
    if media_key == "新鮮視":
        return "全家便利商店\n新鮮視廣告"
    return "家樂福"

def resize_block(ws, start_row, end_row, needed_rows, style_row, max_col):
    height = end_row - start_row + 1
    if needed_rows == height:
        return end_row

    if needed_rows <= 0:
        ws.delete_rows(start_row, amount=height)
        return start_row - 1

    if needed_rows > height:
        ins = needed_rows - height
        ws.insert_rows(end_row + 1, amount=ins)
        for r in range(end_row + 1, end_row + 1 + ins):
            copy_row_with_style_fix(ws, style_row, r, max_col)
        return end_row + ins
    else:
        del_amt = height - needed_rows
        ws.delete_rows(start_row + needed_rows, amount=del_amt)
        return end_row - del_amt

def build_groups(rows):
    """
    media -> seconds -> group
    group has:
      - rows (items)
      - daypart
      - net_total (H merge)
      - list_total_for_total_row (used for Total Row G)
    """
    reg_map = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}

    grouped = {}
    for r in rows:
        m = r["media_type"]
        sec = r["seconds"]
        grouped.setdefault(m, {}).setdefault(sec, []).append(r)

    media_groups = {}
    for m, sec_map in grouped.items():
        sec_groups = []
        for sec, items in sorted(sec_map.items(), key=lambda x: x[0]):
            items_sorted = sorted(items, key=lambda x: reg_map.get(x["region"], 999))
            g0 = items_sorted[0]
            sec_groups.append({
                "seconds": sec,
                "daypart": g0.get("daypart", ""),
                "rows": items_sorted,
                "net_total": g0.get("group_net_total", 0),
                "list_total_for_total": g0.get("group_list_total", 0),
            })
        media_groups[m] = sec_groups
    return media_groups

# =========================================================
# 7) Dongwu / Shenghuo renderers
# =========================================================
def render_block_dongwu(ws, meta, format_type, blk, sec_groups):
    """
    Dongwu:
      - B: merge per media block
      - E/F/H: merge per seconds group
      - C/D/G: row-by-row
      - Total spots: per row
    """
    cols = meta["cols"]
    s, e = blk["start"], blk["end"]
    max_col = ws.max_column
    style_row = s

    # flatten lines
    lines = []
    group_ranges = []  # (start_idx, end_idx, group)
    cur = 0
    for g in sec_groups:
        g_start = cur
        for item in g["rows"]:
            lines.append({"group": g, "item": item})
            cur += 1
        group_ranges.append((g_start, cur - 1, g))

    new_end = resize_block(ws, s, e, len(lines), style_row, max_col)
    if new_end < s:
        return

    # clean merges
    unmerge_overlaps_by_cols(ws, s, new_end, ["B", "E", "F", "H"])

    # merge station(B) across media block
    ws.merge_cells(f"{cols['station']}{s}:{cols['station']}{new_end}")
    top_station = ws[f"{cols['station']}{s}"]
    top_station.value = station_title(format_type, blk["media"])
    apply_center_style(top_station)

    # write row-by-row fields
    for i, line in enumerate(lines):
        r = s + i
        item = line["item"]

        safe_write(ws, f"{cols['location']}{r}", region_display(item["region"]))
        safe_write(ws, f"{cols['program']}{r}", int(item.get("program_num", 0)))

        # Rate (Net header but LIST-based anchor total)
        safe_write(ws, f"{cols['rate']}{r}", item.get("rate_anchor_total"))

        # schedule + total spots
        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)

    # merge E/F/H per seconds group
    for g_start, g_end, g in group_ranges:
        r1 = s + g_start
        r2 = s + g_end

        # Daypart (E)
        ws.merge_cells(f"{cols['daypart']}{r1}:{cols['daypart']}{r2}")
        cE = ws[f"{cols['daypart']}{r1}"]
        cE.value = g["daypart"]
        apply_center_style(cE)

        # Seconds (F)
        ws.merge_cells(f"{cols['seconds']}{r1}:{cols['seconds']}{r2}")
        cF = ws[f"{cols['seconds']}{r1}"]
        cF.value = int(g["seconds"])
        apply_center_style(cF)

        # Package-cost (H) => NET total for the group (charged)
        ws.merge_cells(f"{cols['pkg']}{r1}:{cols['pkg']}{r2}")
        cH = ws[f"{cols['pkg']}{r1}"]
        cH.value = int(g["net_total"]) if isinstance(g["net_total"], (int, float)) else g["net_total"]
        apply_center_style(cH)

def render_block_shenghuo(ws, meta, format_type, blk, sec_groups):
    """
    Shenghuo:
      - station per row
      - proj_price column: use group_net_total per group? (here keep row pkg_display_val)
    """
    cols = meta["cols"]
    s, e = blk["start"], blk["end"]
    max_col = ws.max_column
    style_row = s

    lines = []
    for g in sec_groups:
        for item in g["rows"]:
            lines.append({"group": g, "item": item})

    new_end = resize_block(ws, s, e, len(lines), style_row, max_col)
    if new_end < s:
        return

    for i, line in enumerate(lines):
        r = s + i
        g = line["group"]
        item = line["item"]

        safe_write(ws, f"{cols['station']}{r}", station_title(format_type, blk["media"]))
        apply_center_style(ws[f"{cols['station']}{r}"])

        safe_write(ws, f"{cols['location']}{r}", region_display(item["region"]))
        safe_write(ws, f"{cols['program']}{r}", int(item.get("program_num", 0)))
        safe_write(ws, f"{cols['daypart']}{r}", g["daypart"])
        safe_write(ws, f"{cols['seconds']}{r}", f"{g['seconds']}秒廣告")

        proj_col = cols.get("proj_price")
        if proj_col:
            v = item.get("pkg_display_val", 0)
            safe_write(ws, f"{proj_col}{r}", int(v) if isinstance(v, (int, float)) else v)

        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)

# =========================================================
# 8) Footer polish (VAT + 合約甲乙方)
# =========================================================
def polish_footer(ws, format_type, meta, client_name):
    lbl = meta["footer_labels"]

    # (A) VAT/製作/Grand Total 對齊美化
    for key in ["make", "vat", "grand"]:
        pos = find_cell_exact_in_ws(ws, lbl.get(key, ""))
        if not pos:
            continue
        r, c = pos
        label_cell = ws.cell(r, c)
        val_cell = ws.cell(r, c + 1)

        apply_left_style(label_cell)
        apply_right_style(val_cell)

        # number format
        if isinstance(val_cell.value, (int, float)):
            val_cell.number_format = "#,##0"

        # grand bold
        if key == "grand":
            label_cell.font = copy(label_cell.font) if label_cell.font else Font(bold=True)
            val_cell.font = copy(val_cell.font) if val_cell.font else Font(bold=True)

        # 小幅調整欄寬（讓金額更不擠）
        try:
            col_lbl = get_column_letter(c)
            col_val = get_column_letter(c + 1)
            if ws.column_dimensions[col_lbl].width is None or ws.column_dimensions[col_lbl].width < 12:
                ws.column_dimensions[col_lbl].width = 12
            if ws.column_dimensions[col_val].width is None or ws.column_dimensions[col_val].width < 16:
                ws.column_dimensions[col_val].width = 16
        except:
            pass

    # (B) Dongwu 合約區「甲方/乙方」文字整理
    if format_type == "Dongwu":
        # 甲方在 B33:D33 (merged)，乙方 label 在 I33:K33，乙方名稱在 L33(通常公式=C3)
        # 我們不硬拆 merge（避免版面崩），但把空白與冒號格式整理到最好看
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str):
                    continue

                # 甲方
                if ("甲" in v) and ("方" in v) and ("：" in v or ":" in v):
                    cleaned = re.sub(r"\s+", "", v)
                    # 常見：甲方：東吳...
                    cleaned = cleaned.replace(":", "：")
                    # 保留公司名
                    # 可能是「甲方：東吳廣告股股份有限公司」或中間夾空白
                    cleaned = cleaned.replace("甲方：", "甲方：")
                    cell.value = cleaned
                    apply_left_style(cell)

                # 乙方 label
                if ("乙" in v) and ("方" in v) and ("：" in v or ":" in v) and len(v) <= 12:
                    cell.value = "乙方："
                    apply_right_style(cell)

        # 把乙方名稱公式（如 =$C$3）直接改成文字，避免轉檔字距跑掉
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    if "$C$3" in cell.value.replace(" ", ""):
                        cell.value = client_name
                        apply_left_style(cell)

# =========================================================
# 9) Sheet Meta
# =========================================================
SHEET_META = {
    "Dongwu": {
        "sheet_name": "東吳-格式",
        "date_start_cell": "I7",
        "schedule_start_col": "I",
        "max_days": 31,
        "total_col": "AN",
        "rate_total_col": "G",  # Total Row 用（List override total）
        "pkg_total_col": "H",   # Total Row 用（Net total）
        "anchors": {"全家廣播": "通路廣播廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "anchor_search_cols": ["B", "C", "D"],
        "carrefour_data_offset": 0,
        "header_cells": {"client": "C3", "product": "C4", "period": "C5", "medium": "C6", "month": "I6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "rate": "G", "pkg": "H"},
        # 你指出表頭要改成 Net（但值仍走 Value Anchor）
        "header_override": {"G7": "rate\n(Net)", "H7": "Package-cost\n(Net)"},
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": ["E", "F", "G", "H"],
    },
    "Shenghuo": {
        "sheet_name": "聲活-格式",
        "date_start_cell": "G7",
        "schedule_start_col": "G",
        "max_days": 23,
        "total_col": "AD",
        "anchors": {"全家廣播": "廣播通路廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "anchor_search_cols": ["B", "C", "D"],
        "carrefour_data_offset": 0,
        "header_cells": {"client": "C5", "product": "C6", "month": "G6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "proj_price": "AF"},
        "header_override": {"AF7": "專案價\n(List)"},
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": [],
    }
}

# =========================================================
# 10) Business Logic (Calculation + Value Anchor rows)
# =========================================================
def _calc_spots_with_penalty(budget_allocated, unit_cost, std_spots):
    spots_init = math.ceil(budget_allocated / unit_cost) if unit_cost > 0 else 0
    penalty = 1.1 if spots_init < std_spots else 1.0
    spots_final = math.ceil(budget_allocated / (unit_cost * penalty)) if (unit_cost * penalty) > 0 else 0
    if spots_final % 2 != 0:
        spots_final += 1
    if spots_final <= 0:
        spots_final = 2
    return spots_init, spots_final, penalty

def build_rows_for_familymart_or_freshview(media, seconds, days_count, budget_allocated, is_national, selected_regions):
    """
    returns: rows(list[dict]), debug(dict)
    Rows expansion:
      - if is_national: display 6 regions
      - else: display selected_regions
    Critical:
      - schedule must be counted ONCE in totals (avoid x6)
      - so we mark only first row is_schedule_master=True
    """
    cfg = PRICING_CONFIG[media]
    std = cfg["Std_Spots"]
    daypart = cfg["Day_Part"]
    factor = get_sec_factor(media, seconds)

    # Base Net for calculation
    if is_national:
        base_net = cfg["Prices"]["全省"]["Net"]
        list_for_total_override = cfg["Prices"]["全省"]["List"]
    else:
        base_net = sum(cfg["Prices"][r]["Net"] for r in selected_regions)
        list_for_total_override = sum(cfg["Prices"][r]["List"] for r in selected_regions)

    unit_cost = (base_net / std) * factor if std > 0 else 0
    spots_init, spots_final, penalty = _calc_spots_with_penalty(budget_allocated, unit_cost, std)

    schedule = calculate_schedule(spots_final, days_count)

    # Group net total (charged) based on base_net
    group_net_total = int(round((base_net / std) * spots_final * factor))

    # Group list total for TOTAL ROW override:
    # - national buy: use national list
    # - non-national: sum selected list (same as override var above)
    group_list_total = int(round((list_for_total_override / std) * spots_final * factor))

    display_regions = REGIONS_ORDER if is_national else list(selected_regions)

    rows = []
    for idx, r in enumerate(display_regions):
        # Value Anchor: per region list total (this is what client "sees" as expensive)
        list_price = cfg["Prices"][r]["List"]
        rate_anchor_total = int(round((list_price / std) * spots_final * factor))

        prog_key = (f"新鮮視_{r}" if media == "新鮮視" else r)
        program_num = STORE_COUNTS_CONFIG.get(prog_key, 0)

        rows.append({
            "media_type": media,
            "region": r,
            "program_num": program_num,
            "daypart": daypart,
            "seconds": seconds,
            "spots": spots_final,
            "schedule": schedule,
            # columns:
            "rate_anchor_total": rate_anchor_total,      # Dongwu G (anchor)
            # group totals:
            "group_net_total": group_net_total,          # Dongwu H (charged net, merged)
            "group_list_total": group_list_total,        # Total row G override
            # total schedule count guard:
            "is_schedule_master": (idx == 0),
        })

    debug = {
        "media": media, "sec": seconds, "budget": budget_allocated,
        "unit_cost": unit_cost * penalty,
        "spots_init": spots_init,
        "spots": spots_final,
        "std": std,
        "factor": factor,
        "status": "未達標" if penalty > 1 else "達標",
        "reason": "懲罰 x1.1" if penalty > 1 else "費率正常",
        "group_net_total": group_net_total,
        "group_list_total": group_list_total,
    }
    return rows, debug

def build_rows_for_carrefour(seconds, days_count, budget_allocated):
    """
    Carrefour:
      - Hyper + Super created
      - schedules counted separately (both rows is_schedule_master=True)
      - value anchor: keep hyper list total on G; super can show "計量販"
      - net total charged: based on hyper base (same as your old logic)
    """
    factor = get_sec_factor("家樂福", seconds)
    db = PRICING_CONFIG["家樂福"]["Configs"]
    hype = db["量販_全省"]
    sup  = db["超市_全省"]

    std_h = hype["Std"]
    unit_net = (hype["Net"] / std_h) * factor
    spots_init, spots_final, penalty = _calc_spots_with_penalty(budget_allocated, unit_net, std_h)
    sch_h = calculate_schedule(spots_final, days_count)

    # hyper list anchor total
    rate_anchor_total_h = int(round((hype["List"] / std_h) * spots_final * factor))
    group_net_total = int(round((hype["Net"] / std_h) * spots_final * factor))
    group_list_total = int(round((hype["List"] / std_h) * spots_final * factor))

    # supermarket derived spots
    spots_s = int(round(spots_final * (sup["Std"] / std_h)))
    if spots_s % 2 != 0:
        spots_s += 1
    sch_s = calculate_schedule(spots_s, days_count)

    rows = [
        {
            "media_type": "家樂福",
            "region": "全省量販",
            "program_num": STORE_COUNTS_CONFIG.get("家樂福_量販", 0),
            "daypart": hype["Time"],
            "seconds": seconds,
            "spots": spots_final,
            "schedule": sch_h,
            "rate_anchor_total": rate_anchor_total_h,
            "group_net_total": group_net_total,
            "group_list_total": group_list_total,
            "is_schedule_master": True,
        },
        {
            "media_type": "家樂福",
            "region": "全省超市",
            "program_num": STORE_COUNTS_CONFIG.get("家樂福_超市", 0),
            "daypart": sup["Time"],
            "seconds": seconds,
            "spots": spots_s,
            "schedule": sch_s,
            "rate_anchor_total": "計量販",
            "group_net_total": group_net_total,     # 同一 group 合併顯示（可調）
            "group_list_total": group_list_total,   # Total row override（可調）
            "is_schedule_master": True,
        }
    ]

    debug = {
        "media": "家樂福", "sec": seconds, "budget": budget_allocated,
        "unit_cost": unit_net * penalty,
        "spots_init": spots_init,
        "spots": spots_final,
        "std": std_h,
        "factor": factor,
        "status": "未達標" if penalty > 1 else "達標",
        "reason": "懲罰 x1.1" if penalty > 1 else "費率正常",
        "group_net_total": group_net_total,
        "group_list_total": group_list_total,
    }
    return rows, debug

# =========================================================
# 11) Excel Generate (Block Renderer + Total + Footer)
# =========================================================
def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if meta["sheet_name"] not in wb.sheetnames:
        raise ValueError(f"缺少分頁：{meta['sheet_name']}")

    # only keep target sheet
    target_sheet = meta["sheet_name"]
    for sname in list(wb.sheetnames):
        if sname != target_sheet:
            del wb[sname]
    ws = wb[target_sheet]

    # Header
    hc = meta["header_cells"]
    if "client" in hc:
        safe_write(ws, hc["client"], client_name)
    if "product" in hc:
        safe_write(ws, hc["product"], product_display_str)
    if "period" in hc:
        safe_write(ws, hc["period"], f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y.%m. %d')}")
    if "medium" in hc:
        safe_write(ws, hc["medium"], " ".join(sorted(set([r["media_type"] for r in rows]))))
    if "month" in hc:
        safe_write(ws, hc["month"], f" {start_dt.month}月")

    safe_write(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))
    for addr, text in meta.get("header_override", {}).items():
        safe_write(ws, addr, text)

    # find blocks + total row
    blocks, total_row_orig = detect_blocks(ws, format_type, meta)

    # unmerge data area to protect Total/Footer
    data_start = min(b["start"] for b in blocks)
    data_end = total_row_orig - 1
    if format_type == "Dongwu":
        unmerge_overlaps_by_cols(ws, data_start, data_end, ["B", "E", "F", "H"])

    # group
    media_groups = build_groups(rows)

    # render blocks bottom-up
    for blk in sorted(blocks, key=lambda b: b["start"], reverse=True):
        m = blk["media"]
        sec_groups = media_groups.get(m, [])

        if not sec_groups:
            ws.delete_rows(blk["start"], amount=(blk["end"] - blk["start"] + 1))
            continue

        if format_type == "Dongwu":
            render_block_dongwu(ws, meta, format_type, blk, sec_groups)
        else:
            render_block_shenghuo(ws, meta, format_type, blk, sec_groups)

    # relocate Total row after shifts
    total_pos = find_cell_exact_in_ws(ws, meta["total_label"])
    if not total_pos:
        raise ValueError("渲染後找不到 Total（可能 Total 被模板 merge/公式影響）")
    total_row = total_pos[0]

    # ---- Total schedule MUST NOT multiply by expanded region rows ----
    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [0] * eff_days
    for x in rows:
        if not x.get("is_schedule_master", False):
            continue
        sch = x.get("schedule", [])
        for d in range(min(eff_days, len(sch))):
            daily_sums[d] += sch[d]

    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))

    # ---- Total Row: Dongwu has two money columns (G list total, H net total) ----
    if format_type == "Dongwu":
        # sum group totals once per (media, seconds) group:
        seen = set()
        total_list = 0
        total_net = 0
        for x in rows:
            gid = (x["media_type"], x["seconds"])
            if gid in seen:
                continue
            seen.add(gid)
            total_list += int(x.get("group_list_total", 0) or 0)
            total_net  += int(x.get("group_net_total", 0) or 0)

        safe_write(ws, f"{meta['rate_total_col']}{total_row}", total_list)
        safe_write(ws, f"{meta['pkg_total_col']}{total_row}", total_net)

    else:
        # Shenghuo: only one price column (AF)
        cols = meta["cols"]
        pkg_col = cols.get("proj_price")
        total_pkg = sum([x["pkg_display_val"] for x in rows if isinstance(x.get("pkg_display_val"), (int, float))])
        if pkg_col:
            safe_write(ws, f"{pkg_col}{total_row}", int(total_pkg))

    # Footer make/vat/grand based on NET total (Dongwu H / Shenghuo total_pkg)
    lbl = meta["footer_labels"]
    make_fee = 10000
    pos_make = find_cell_exact_in_ws(ws, lbl["make"])
    if pos_make:
        v = ws.cell(pos_make[0], pos_make[1] + 1).value
        if isinstance(v, (int, float)) and v > 0:
            make_fee = int(v)
        else:
            safe_write_rc(ws, pos_make[0], pos_make[1] + 1, make_fee)

    if format_type == "Dongwu":
        # net base is Total Row H
        net_total_val = ws[f"{meta['pkg_total_col']}{total_row}"].value
        net_total_val = int(net_total_val) if isinstance(net_total_val, (int, float)) else 0
        vat = int(round((net_total_val + make_fee) * 0.05))
        grand = net_total_val + make_fee + vat
    else:
        vat = int(round((total_pkg + make_fee) * 0.05))
        grand = total_pkg + make_fee + vat

    pos_vat = find_cell_exact_in_ws(ws, lbl["vat"])
    if pos_vat:
        safe_write_rc(ws, pos_vat[0], pos_vat[1] + 1, vat)

    pos_grand = find_cell_exact_in_ws(ws, lbl["grand"])
    if pos_grand:
        safe_write_rc(ws, pos_grand[0], pos_grand[1] + 1, grand)

    # Remarks
    rem_pos = find_cell_exact_in_ws(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_write_rc(ws, rem_pos[0] + 1 + i, rem_pos[1], rm)

    # force center (Dongwu)
    if format_type == "Dongwu":
        min_data_row = min(b["start"] for b in blocks)
        force_center_columns_range(ws, meta["force_center_cols"], min_data_row, total_row)

    # polish footer
    polish_footer(ws, format_type, meta, client_name)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 12) HTML Preview (use same rows; totals also must avoid x6)
# =========================================================
def load_font_base64():
    font_path = "NotoSansTC-Regular.ttf"
    if os.path.exists(font_path):
        with open(font_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/TTF/TraditionalChinese/NotoSansTC-Regular.ttf"
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            with open(font_path, "wb") as f:
                f.write(r.content)
            return base64.b64encode(r.content).decode("utf-8")
    except:
        pass
    return None

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks):
    header_cls = "bg-dw-head" if format_type == "Dongwu" else "bg-sh-head"
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    eff_days = min(days_cnt, 31)

    st.markdown(f"""<style>
    .bg-dw-head {{ background-color: #4472C4; color: white; font-weight: bold; }}
    .bg-sh-head {{ background-color: #BDD7EE; color: black; font-weight: bold; }}
    .bg-weekend {{ background-color: #FFD966; color: black; }}
    .bg-total   {{ background-color: #FFF2CC; font-weight: bold; }}
    .col_day {{ min-width: 25px; }}
    .left {{ text-align:left; }}
    .right {{ text-align:right; }}
    </style>""", unsafe_allow_html=True)

    date_th1, date_th2 = "", ""
    curr = start_dt
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for _ in range(eff_days):
        wd = curr.weekday()
        bg = "bg-weekend" if (format_type == "Dongwu" and wd >= 5) else header_cls
        if format_type == "Shenghuo":
            bg = header_cls
        date_th1 += f"<th class='{bg} col_day'>{curr.day}</th>"
        date_th2 += f"<th class='{bg} col_day'>{weekdays[wd]}</th>"
        curr += timedelta(days=1)

    if format_type == "Dongwu":
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate (Net)", "Package-cost (Net)"]
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數規格", "專案價"]

    th_fixed = "".join([f"<th rowspan='2'>{c}</th>" for c in cols_def])

    rows_sorted = sorted(
        rows,
        key=lambda x: (media_order.get(x["media_type"], 99), x["seconds"], REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 99)
    )

    tbody = ""
    # station rowspans (Dongwu only)
    media_counts = {}
    for r in rows_sorted:
        media_counts[r["media_type"]] = media_counts.get(r["media_type"], 0) + 1
    media_printed = {m: False for m in media_counts}

    for r in rows_sorted:
        m = r["media_type"]
        tbody += "<tr>"
        if not media_printed[m]:
            rowspan = media_counts[m]
            display_name = station_title(format_type, m).replace("\n", "<br>")
            if format_type == "Shenghuo":
                tbody += f"<td class='left'>{display_name}</td>"
            else:
                tbody += f"<td class='left' rowspan='{rowspan}'>{display_name}</td>"
                media_printed[m] = True
        elif format_type == "Shenghuo":
            display_name = station_title(format_type, m).replace("\n", "<br>")
            tbody += f"<td class='left'>{display_name}</td>"

        tbody += f"<td>{region_display(r['region'])}</td>"
        tbody += f"<td class='right'>{r.get('program_num','')}</td>"
        tbody += f"<td>{r.get('daypart','')}</td>"
        tbody += f"<td>{r['seconds']}</td>"

        if format_type == "Dongwu":
            g = r.get("rate_anchor_total", "")
            h = r.get("group_net_total", "")
            # show G per row; H only show on schedule_master row to avoid repetition in HTML
            g_txt = f"{g:,}" if isinstance(g, int) else str(g)
            h_txt = f"{h:,}" if (isinstance(h, int) and r.get("is_schedule_master")) else ("" if not r.get("is_schedule_master") else str(h))
            tbody += f"<td class='right'>{g_txt}</td><td class='right'>{h_txt}</td>"
        else:
            v = r.get("pkg_display_val", "")
            v_txt = f"{v:,}" if isinstance(v, int) else str(v)
            tbody += f"<td class='right'>{v_txt}</td>"

        for d in r.get("schedule", [])[:eff_days]:
            tbody += f"<td>{d}</td>"
        tbody += f"<td class='bg-total'>{sum(r.get('schedule', []))}</td></tr>"

    # totals: schedule must sum only masters
    totals = [0] * eff_days
    for r in rows:
        if not r.get("is_schedule_master"):
            continue
        sch = r.get("schedule", [])
        for i in range(min(eff_days, len(sch))):
            totals[i] += sch[i]

    if format_type == "Dongwu":
        seen = set()
        total_list = 0
        total_net = 0
        for r in rows:
            gid = (r["media_type"], r["seconds"])
            if gid in seen:
                continue
            seen.add(gid)
            total_list += int(r.get("group_list_total", 0) or 0)
            total_net += int(r.get("group_net_total", 0) or 0)

        colspan = 5
        tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td>"
        tfoot += f"<td class='right'>{total_list:,}</td><td class='right'>{total_net:,}</td>"
    else:
        total_pkg = sum([r.get("pkg_display_val", 0) for r in rows if isinstance(r.get("pkg_display_val"), int)])
        colspan = 5
        tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td><td class='right'>{total_pkg:,}</td>"

    for t in totals:
        tfoot += f"<td>{t}</td>"
    tfoot += f"<td>{sum(totals)}</td></tr>"

    return f"""<div class="excel-container">
    <div style="margin-bottom:10px;"><b>客戶：</b>{html_escape(c_name)} &nbsp; <b>產品：</b>{html_escape(p_display)}<br><span style="color:#666;">走期：{start_dt} ~ {end_dt}</span></div>
    <table class="excel-table">
    <thead><tr>{th_fixed}{date_th1}<th class='{header_cls}' rowspan='2'>檔次</th></tr><tr>{date_th2}</tr></thead>
    <tbody>{tbody}{tfoot}</tbody></table>
    <div class="remarks"><b>Remarks：</b><br>{"<br>".join([html_escape(x) for x in remarks])}</div></div>"""

# =========================================================
# 13) UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v66.0)")

auto_tpl, source, msgs = load_default_template()
template_bytes = auto_tpl

if auto_tpl:
    st.success(f"✅ 已載入系統公版 ({source})")
else:
    st.warning("⚠️ 無法載入公版，請手動上傳")
    tpl = st.file_uploader("上傳 Excel 模板", type=["xlsx"])
    if tpl:
        template_bytes = tpl.read()

if msgs:
    for m in msgs:
        st.info(m)

st.markdown("### 1. 選擇格式")
format_type = st.radio("", ["Dongwu", "Shenghuo"], horizontal=True, label_visibility="collapsed")

st.markdown("### 2. 基本資料設定")
c1, c2, c3 = st.columns(3)
with c1:
    client_name = st.text_input("客戶名稱", "萬國通路")
with c2:
    product_name = st.text_input("產品名稱", "統一布丁")
with c3:
    total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)

c4, c5 = st.columns(2)
with c4:
    start_date = st.date_input("開始日", datetime(2026, 1, 1))
with c5:
    end_date = st.date_input("結束日", datetime(2026, 1, 31))

days_count = (end_date - start_date).days + 1
st.info(f"📅 走期共 **{days_count}** 天")

with st.expander("📝 備註欄位設定 (Remarks)", expanded=False):
    rc1, rc2, rc3 = st.columns(3)
    sign_deadline = rc1.date_input("回簽截止日", datetime.now() + timedelta(days=3))
    billing_month = rc2.text_input("請款月份", "2026年2月")
    payment_date = rc3.date_input("付款兌現日", datetime(2026, 3, 31))

st.markdown("### 3. 媒體投放設定")

# state init
if "rad_share" not in st.session_state:
    st.session_state.rad_share = 100
if "fv_share" not in st.session_state:
    st.session_state.fv_share = 0
if "cf_share" not in st.session_state:
    st.session_state.cf_share = 0

def on_media_change():
    active = []
    if st.session_state.get("cb_rad"):
        active.append("rad_share")
    if st.session_state.get("cb_fv"):
        active.append("fv_share")
    if st.session_state.get("cb_cf"):
        active.append("cf_share")
    if not active:
        return
    share = 100 // len(active)
    for key in active:
        st.session_state[key] = share
    rem = 100 - sum([st.session_state[k] for k in active])
    st.session_state[active[0]] += rem

def on_slider_change(changed_key):
    active = []
    if st.session_state.get("cb_rad"):
        active.append("rad_share")
    if st.session_state.get("cb_fv"):
        active.append("fv_share")
    if st.session_state.get("cb_cf"):
        active.append("cf_share")

    others = [k for k in active if k != changed_key]
    if not others:
        st.session_state[changed_key] = 100
    elif len(others) == 1:
        val = st.session_state[changed_key]
        st.session_state[others[0]] = max(0, 100 - val)
    elif len(others) == 2:
        val = st.session_state[changed_key]
        rem = max(0, 100 - val)
        k1, k2 = others[0], others[1]
        sum_others = st.session_state[k1] + st.session_state[k2]
        if sum_others == 0:
            st.session_state[k1] = rem // 2
            st.session_state[k2] = rem - st.session_state[k1]
        else:
            ratio = st.session_state[k1] / sum_others
            st.session_state[k1] = int(rem * ratio)
            st.session_state[k2] = rem - st.session_state[k1]

st.write("請勾選要投放的媒體：")
col_cb1, col_cb2, col_cb3 = st.columns(3)
with col_cb1:
    is_rad = st.checkbox("全家廣播", value=True, key="cb_rad", on_change=on_media_change)
with col_cb2:
    is_fv = st.checkbox("新鮮視", value=False, key="cb_fv", on_change=on_media_change)
with col_cb3:
    is_cf = st.checkbox("家樂福", value=False, key="cb_cf", on_change=on_media_change)

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
            for s in sorted(secs)[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls / 2), key=f"rs_{s}")
                sec_shares[s] = v
                ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs:
            sec_shares[secs[0]] = 100
        config["全家廣播"] = {
            "is_national": is_nat,
            "regions": (REGIONS_ORDER if is_nat else regs),
            "seconds": sorted(secs),
            "share": st.session_state.rad_share,
            "sec_shares": sec_shares
        }

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
            for s in sorted(secs)[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls / 2), key=f"fs_{s}")
                sec_shares[s] = v
                ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs:
            sec_shares[secs[0]] = 100
        config["新鮮視"] = {
            "is_national": is_nat,
            "regions": (REGIONS_ORDER if is_nat else regs),
            "seconds": sorted(secs),
            "share": st.session_state.fv_share,
            "sec_shares": sec_shares
        }

if is_cf:
    with m3:
        st.markdown("#### 🛒 家樂福")
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.slider("預算 %", 0, 100, key="cf_share", on_change=on_slider_change, args=("cf_share",))
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in sorted(secs)[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls / 2), key=f"cs_{s}")
                sec_shares[s] = v
                ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs:
            sec_shares[secs[0]] = 100
        config["家樂福"] = {
            "regions": ["全省"],
            "seconds": sorted(secs),
            "share": st.session_state.cf_share,
            "sec_shares": sec_shares
        }

# ------------------------------
# Calculation engine
# ------------------------------
rows = []
debug_logs = []

if config:
    for media, cfg in config.items():
        m_budget = total_budget_input * (cfg["share"] / 100.0)

        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0:
                continue

            if media in ["全家廣播", "新鮮視"]:
                is_nat = bool(cfg.get("is_national", False))
                selected_regions = list(cfg.get("regions", REGIONS_ORDER))
                rws, dbg = build_rows_for_familymart_or_freshview(
                    media=media,
                    seconds=sec,
                    days_count=days_count,
                    budget_allocated=s_budget,
                    is_national=is_nat,
                    selected_regions=selected_regions
                )
                rows.extend(rws)
                debug_logs.append(dbg)

            elif media == "家樂福":
                rws, dbg = build_rows_for_carrefour(
                    seconds=sec,
                    days_count=days_count,
                    budget_allocated=s_budget
                )
                rows.extend(rws)
                debug_logs.append(dbg)

# Product display string
p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}" if rows else ""
rem = get_remarks_text(sign_deadline, billing_month, payment_date)

with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
    st.markdown("#### 1) 本次預算分配與檔次計算（Spots）")
    for log in debug_logs:
        color = "green" if log["status"] == "達標" else "red"
        st.markdown(
            f"**{log['media']} ({log['sec']}秒)**: 預算 ${log['budget']:,.0f} | "
            f"Spots_init={log['spots_init']} -> Spots_final={log['spots']} | "
            f"<span style='color:{color}'><b>{log['status']}</b></span>（{log['reason']}）<br>"
            f"Group Net={log['group_net_total']:,} | Group List(Override)={log['group_list_total']:,}",
            unsafe_allow_html=True
        )
    st.caption("⚠️ Total 的每日檔次只會加總每個 seconds-group 的一次（不會因為展開 6 區而乘 6）。")

if rows:
    font_b64 = load_font_base64()
    html = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem)
    st.components.v1.html(html, height=700, scrolling=True)

    if template_bytes:
        try:
            xlsx = generate_excel_from_template(format_type, start_date, end_date, client_name, p_str, rows, rem, template_bytes)
            st.download_button("下載 Excel", xlsx, f"Cue_{safe_filename(client_name)}.xlsx")

            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx)
            if pdf_bytes:
                st.download_button(f"下載 PDF ({method})", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
            else:
                st.warning(f"Excel 轉 PDF 失敗 ({method}: {err})，切換至備用渲染引擎 (HTML)...")
                pdf_bytes, err2 = html_to_pdf_fallback(html, font_b64)
                if pdf_bytes:
                    st.download_button("下載 PDF (Fallback)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                else:
                    st.error(f"PDF 產出失敗: {err2}")

        except Exception as e:
            st.error(f"Excel 產出錯誤: {e}")
    else:
        st.warning("請上傳模板以啟用下載。")
else:
    st.info("請先勾選媒體並設定秒數/區域，產生資料列後即可預覽與下載。")
