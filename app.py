# -*- coding: utf-8 -*-
"""
Cue Sheet Pro v65.3 (Value Anchor + Block Renderer + Footer Polish)

重點修正：
1) Rate(Net) = 分區總價（以 List 算出單位價後 * Spots）
2) Package-cost(Net):
   - 勾全省聯播：分區列不做 *1.1（避免數字太怪），Total 列顯示「全省打包價」（未達標可 *1.1）
   - 未勾全省：分區列 Package = Rate * (1.1 if penalty else 1.0)，Total 列 Package = 分區 Package 加總
3) Total 列 Rate = 所有分區列 Rate 加總（欺敵視覺）
4) 修正家樂福店數 68/249
5) HTML/PDF 不再噴 err（容錯格式化）
"""

# ---------------------------------------------------------
# 0) Imports (with friendly Streamlit hint)
# ---------------------------------------------------------
try:
    import streamlit as st
except ModuleNotFoundError as e:
    raise ModuleNotFoundError(
        "找不到 streamlit。\n\n"
        "請先安裝：\n"
        "  pip install streamlit\n\n"
        "並用以下方式啟動（不要用 python 直接跑）：\n"
        "  streamlit run freshvision_closer.py\n"
    ) from e

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
from openpyxl.styles import Alignment, Font


# =========================================================
# 1) Master Configuration v2.0 (你規格書的「憲法」)
# =========================================================

# 店數（家樂福重要更新：68/249）
STORE_COUNTS_CONFIG = {
    "全省": 4437,
    "北區": 1649, "桃竹苗": 779, "中區": 839, "雲嘉南": 499, "高屏": 490, "東區": 181,
    "新鮮視_全省": 3124,
    "新鮮視_北區": 1127, "新鮮視_桃竹苗": 616, "新鮮視_中區": 528,
    "新鮮視_雲嘉南": 365, "新鮮視_高屏": 405, "新鮮視_東區": 83,
    "家樂福_量販": 68, "家樂福_超市": 249
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
    "家樂福":   {30: 1.5, 20: 1.0,  15: 0.85, 10: 0.65, 5: 0.35}
}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]


# =========================================================
# 2) 基礎工具
# =========================================================
def parse_count_to_int(x):
    if x is None:
        return 0
    if isinstance(x, (int, float)):
        return int(x)
    s = str(x)
    m = re.findall(r"[\d,]+", s)
    if not m:
        return 0
    return int(m[0].replace(",", ""))

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def html_escape(s):
    if s is None:
        return ""
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#39;"))

def norm_ws(s: str) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def fmt_int(v):
    """HTML/顯示用：避免 err"""
    try:
        if isinstance(v, (int, float)):
            return f"{int(round(v)):,}"
        return str(v)
    except:
        return str(v)

def get_sec_factor(media_type, seconds: int) -> float:
    """
    規格書只定義 5/10/15/20/30。
    你 UI 允許 25/35/45...：這裡採「向下就近」+「大於最大秒數用最大」策略。
    例如 35/45/60 -> 用 30 的 factor；25 -> 用 20 的 factor。
    """
    mp = SEC_FACTORS_CONFIG.get(media_type, {})
    if not mp:
        return 1.0
    if seconds in mp:
        return mp[seconds]
    keys = sorted(mp.keys())
    if seconds > keys[-1]:
        return mp[keys[-1]]
    # closest lower
    lower = [k for k in keys if k <= seconds]
    return mp[lower[-1]] if lower else mp[keys[0]]

def calculate_schedule(total_spots, days):
    if days <= 0:
        return []
    # 確保偶數
    if total_spots % 2 != 0:
        total_spots += 1
    half_spots = total_spots // 2
    base, rem = divmod(half_spots, days)
    half_schedule = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in half_schedule]

REGION_DISPLAY_6 = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市",
}
def region_display(region: str) -> str:
    return REGION_DISPLAY_6.get(region, region)


# =========================================================
# 3) Streamlit Page 設定 & Template Auto-load
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v65.3 (Value Anchor + Bundle Price)")

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
# 4) Excel 轉 PDF (Local 擬真 + WeasyPrint Fallback)
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
    # 1) Windows Excel COM
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
        .bg-dw-head { background-color: #4472C4; color: white; font-weight: bold; }
        .bg-sh-head { background-color: #BDD7EE; color: black; font-weight: bold; }
        .bg-weekend { background-color: #FFD966; }
        .bg-total   { background-color: #FFF2CC; font-weight: bold; }
        tr { page-break-inside: avoid; }
        .left { text-align:left; }
        .right { text-align:right; }
        """
        if font_b64:
            css_str = f"@font-face {{ font-family: 'NotoSansTC'; src: url(data:font/ttf;base64,{font_b64}) format('truetype'); }} " + css_str
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[CSS(string=css_str)], font_config=font_config)
        return pdf_bytes, ""
    except Exception as e:
        return None, str(e)


# =========================================================
# 5) Remarks
# =========================================================
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
# 6) Excel 輔助：安全寫入 / 複製樣式 / 解除合併
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
# 7) Block Renderer：偵測 blocks / resize / render
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

    rad_start = find_row_contains_any(ws, search_cols, meta["anchors"]["全家廣播"])
    fv_start  = find_row_contains_any(ws, search_cols, meta["anchors"]["新鮮視"])
    cf_anchor = find_row_contains_any(ws, search_cols, meta["anchors"]["家樂福"])
    cf_start  = (cf_anchor + meta.get("carrefour_data_offset", 0)) if cf_anchor else None

    starts = []
    if rad_start: starts.append(("全家廣播", rad_start))
    if fv_start:  starts.append(("新鮮視", fv_start))
    if cf_start:  starts.append(("家樂福", cf_start))
    starts.sort(key=lambda x: x[1])

    blocks = []
    for i, (m, srow) in enumerate(starts):
        erow = (starts[i + 1][1] - 1) if (i + 1 < len(starts)) else (total_row - 1)
        if erow >= srow:
            blocks.append({"media": m, "start": srow, "end": erow})

    if not blocks:
        raise ValueError("找不到任何媒體區段 blocks（anchors 可能跟模板不一致）")

    return blocks, total_row

def build_groups(rows):
    """media -> seconds -> items（排序）"""
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
            pkg_sum = sum([x["pkg_display_val"] for x in items_sorted if isinstance(x["pkg_display_val"], int)])
            sec_groups.append({
                "seconds": sec,
                "daypart": items_sorted[0].get("daypart", ""),
                "rows": items_sorted,
                "pkg_sum": pkg_sum
            })
        media_groups[m] = sec_groups
    return media_groups

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

def render_block_dongwu(ws, meta, format_type, blk, sec_groups):
    cols = meta["cols"]
    s, e = blk["start"], blk["end"]
    max_col = ws.max_column
    style_row = s

    # flatten
    lines = []
    group_ranges = []
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

    unmerge_overlaps_by_cols(ws, s, new_end, ["B", "E", "F", "H"])

    # merge station (B) across block
    ws.merge_cells(f"{cols['station']}{s}:{cols['station']}{new_end}")
    top_station = ws[f"{cols['station']}{s}"]
    top_station.value = station_title(format_type, blk["media"])
    apply_center_style(top_station)

    for i, line in enumerate(lines):
        r = s + i
        item = line["item"]
        safe_write(ws, f"{cols['location']}{r}", region_display(item["region"]))
        safe_write(ws, f"{cols['program']}{r}", int(item.get("program_num", 0) or 0))

        # ✅ Rate(Net) = 分區總價（int or string）
        safe_write(ws, f"{cols['rate']}{r}", item.get("rate_list"))

        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)

    # merge Daypart(E)/Seconds(F)/Package(H) per seconds group
    for g_start, g_end, g in group_ranges:
        r1 = s + g_start
        r2 = s + g_end

        ws.merge_cells(f"{cols['daypart']}{r1}:{cols['daypart']}{r2}")
        cE = ws[f"{cols['daypart']}{r1}"]
        cE.value = g["daypart"]
        apply_center_style(cE)

        ws.merge_cells(f"{cols['seconds']}{r1}:{cols['seconds']}{r2}")
        cF = ws[f"{cols['seconds']}{r1}"]
        cF.value = f"{g['seconds']}秒" if blk["media"] == "家樂福" else int(g["seconds"])
        apply_center_style(cF)

        ws.merge_cells(f"{cols['pkg']}{r1}:{cols['pkg']}{r2}")
        cH = ws[f"{cols['pkg']}{r1}"]
        cH.value = int(g["pkg_sum"]) if isinstance(g["pkg_sum"], int) else g["pkg_sum"]
        apply_center_style(cH)

def render_block_shenghuo(ws, meta, format_type, blk, sec_groups):
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
        safe_write(ws, f"{cols['program']}{r}", int(item.get("program_num", 0) or 0))
        safe_write(ws, f"{cols['daypart']}{r}", g["daypart"])
        safe_write(ws, f"{cols['seconds']}{r}", f"{g['seconds']}秒廣告")

        # Shenghuo 的金額欄（用 proj_price）
        proj_col = cols.get("proj_price")
        if proj_col:
            v = item.get("pkg_display_val", 0)
            safe_write(ws, f"{proj_col}{r}", int(v) if isinstance(v, int) else v)

        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)


# =========================================================
# 8) Footer 美化（VAT + 合約甲乙方）
# =========================================================
def polish_footer(ws, format_type, meta, client_name):
    lbl = meta["footer_labels"]

    # 只處理底部 60 列，避免誤傷
    scan_start = max(1, ws.max_row - 60)
    scan_end = ws.max_row

    def find_in_bottom(text):
        for r in range(scan_start, scan_end + 1):
            for c in range(1, ws.max_column + 1):
                if ws.cell(r, c).value == text:
                    return r, c
        return None

    # VAT / 製作 / Grand Total：label 左、金額右、千分位
    for key in ["make", "vat", "grand"]:
        pos = find_in_bottom(lbl.get(key, ""))
        if not pos:
            continue
        r, c = pos
        label_cell = ws.cell(r, c)
        val_cell = ws.cell(r, c + 1)

        apply_left_style(label_cell)
        apply_right_style(val_cell)

        if isinstance(val_cell.value, (int, float)):
            val_cell.number_format = "#,##0"

        if key == "grand":
            label_cell.font = copy(label_cell.font) if label_cell.font else Font(bold=True)
            val_cell.font = copy(val_cell.font) if val_cell.font else Font(bold=True)

        # 小幅調整欄寬（如果可行）
        try:
            col_label = openpyxl.utils.get_column_letter(c)
            col_val = openpyxl.utils.get_column_letter(c + 1)
            ws.column_dimensions[col_label].width = max(ws.column_dimensions[col_label].width or 0, 10)
            ws.column_dimensions[col_val].width = max(ws.column_dimensions[col_val].width or 0, 14)
        except:
            pass

    # Dongwu 合約區：清理「甲       方：」「乙       方：」與日期
    if format_type == "Dongwu":
        for r in range(scan_start, scan_end + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                v = cell.value
                if not isinstance(v, str):
                    continue

                v2 = norm_ws(v)

                # 甲方 / 乙方 / 日期 清理
                if ("甲" in v2) and ("方" in v2) and ("：" in v2):
                    cell.value = "甲方："
                    apply_right_style(cell)
                if ("乙" in v2) and ("方" in v2) and ("：" in v2):
                    cell.value = "乙方："
                    apply_right_style(cell)
                if ("日期" in v2) and ("：" in v2):
                    cell.value = "日期："
                    apply_right_style(cell)

                # 把公式 =$C$3 直接換成客戶名，避免 PDF 跑版
                if v2.startswith("=") and ("$C$3" in v2):
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
        "anchors": {"全家廣播": "通路廣播廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "anchor_search_cols": ["B", "C", "D"],
        "carrefour_data_offset": 0,
        "header_cells": {"client": "C3", "product": "C4", "period": "C5", "medium": "C6", "month": "I6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "rate": "G", "pkg": "H"},
        "header_override": {"G7": "Rate\n(Net)", "H7": "Package-cost\n(Net)"},
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
        "header_override": {"AF7": "專案價\n(Net)"},
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": [],
    }
}


# =========================================================
# 10) 核心運算：Spots + Value Anchor (Rate/Package)
# =========================================================
def calc_spots_for_radio_tv(media, sec, budget_allocated, is_national, regions_selected):
    cfg = PRICING_CONFIG[media]
    std = cfg["Std_Spots"]
    factor = get_sec_factor(media, sec)

    if is_national:
        base_net = cfg["Prices"]["全省"]["Net"]
    else:
        base_net = sum(cfg["Prices"][r]["Net"] for r in regions_selected)

    unit_cost = (base_net / std) * factor
    if unit_cost <= 0:
        return 0, False

    spots_init = math.ceil(budget_allocated / unit_cost)
    penalty = spots_init < std

    if penalty:
        spots_final = math.ceil(budget_allocated / (unit_cost * 1.1))
    else:
        spots_final = spots_init

    if spots_final % 2 != 0:
        spots_final += 1
    if spots_final <= 0:
        spots_final = 2

    return spots_final, penalty

def calc_bundle_package_total(media, sec, spots_final, penalty):
    """勾全省聯播時，Total 列的 Package-cost 用全省 List 打包價（可 *1.1）"""
    cfg = PRICING_CONFIG[media]
    std = cfg["Std_Spots"]
    factor = get_sec_factor(media, sec)
    unit_list = int(round((cfg["Prices"]["全省"]["List"] / std) * factor))
    base = unit_list * spots_final
    if penalty:
        return int(round(base * 1.1))
    return int(base)

def build_rows_and_totals(config, total_budget_input, days_count):
    """
    產出：
    - rows：給 Excel/HTML
    - pricing_groups：每個 group 的 bundle package（只有勾全省才有）
    """
    rows = []
    debug_logs = []
    pricing_groups = {}  # group_id -> dict

    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}

    for media, cfg in config.items():
        m_budget = total_budget_input * (cfg["share"] / 100.0)

        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0:
                continue

            group_id = f"{media}__{sec}__{int(cfg.get('is_national', False))}"

            # ---- Radio / TV ----
            if media in ["全家廣播", "新鮮視"]:
                is_national = bool(cfg.get("is_national", False))
                regions_selected = (REGIONS_ORDER[:] if is_national else cfg["regions"][:])

                spots_final, penalty = calc_spots_for_radio_tv(media, sec, s_budget, is_national, regions_selected)
                factor = get_sec_factor(media, sec)
                std = PRICING_CONFIG[media]["Std_Spots"]
                daypart = PRICING_CONFIG[media]["Day_Part"]

                debug_logs.append({
                    "media": media, "sec": sec, "budget": s_budget,
                    "spots": spots_final, "std": std, "factor": factor,
                    "status": "未達標" if penalty else "達標",
                    "reason": "懲罰 x1.1" if penalty else "費率正常",
                    "is_national": is_national
                })

                # 勾全省時，Total 列 Package-cost 用打包價（只記一次）
                if is_national:
                    pricing_groups[group_id] = {
                        "is_national": True,
                        "bundle_pkg_total": calc_bundle_package_total(media, sec, spots_final, penalty),
                        "penalty": penalty
                    }
                else:
                    pricing_groups[group_id] = {"is_national": False, "bundle_pkg_total": None, "penalty": penalty}

                sch = calculate_schedule(spots_final, days_count)

                for r in regions_selected:
                    list_price = PRICING_CONFIG[media]["Prices"][r]["List"]
                    unit_list = int(round((list_price / std) * factor))
                    rate_total = int(unit_list * spots_final)  # ✅ Rate(Net)=分區總價（不含*1.1）

                    # ✅ Package-cost(Net)
                    # - 勾全省：分區列不做 *1.1（避免誇張）
                    # - 未勾全省：分區列做 *1.1（若未達標）
                    if is_national:
                        pkg_total = rate_total
                    else:
                        pkg_total = int(round(rate_total * (1.1 if penalty else 1.0)))

                    prog_key = r if media == "全家廣播" else f"新鮮視_{r}"
                    rows.append({
                        "media_type": media,
                        "region": r,
                        "program_num": STORE_COUNTS_CONFIG.get(prog_key, 0),
                        "daypart": daypart,
                        "seconds": sec,
                        "spots": spots_final,
                        "schedule": sch,
                        # 沿用舊 key，避免你 Excel renderer 改很大
                        "rate_list": rate_total,          # ✅ Rate(Net)=分區總價
                        "pkg_display_val": pkg_total,     # ✅ Package-cost(Net)=分區列顯示
                        "group_id": group_id,
                        "is_national": is_national,
                        "penalty": penalty
                    })

            # ---- Carrefour ----
            elif media == '家樂福':
                factor = get_sec_factor("家樂福", sec)
                cfg_cf = PRICING_CONFIG["家樂福"]["Configs"]
                h = cfg_cf["量販_全省"]
                s = cfg_cf["超市_全省"]

                base_std = h["Std"]
                unit_net = (h["Net"] / base_std) * factor
                if unit_net <= 0:
                    continue

                spots_init = math.ceil(s_budget / unit_net)
                penalty = spots_init < base_std
                if penalty:
                    spots_final = math.ceil(s_budget / (unit_net * 1.1))
                else:
                    spots_final = spots_init
                if spots_final % 2 != 0:
                    spots_final += 1
                if spots_final <= 0:
                    spots_final = 2

                debug_logs.append({
                    "media": media, "sec": sec, "budget": s_budget,
                    "spots": spots_final, "std": base_std, "factor": factor,
                    "status": "未達標" if penalty else "達標",
                    "reason": "懲罰 x1.1" if penalty else "費率正常",
                    "is_national": True
                })

                sch_h = calculate_schedule(spots_final, days_count)

                # 量販：Rate=分區總價（不含*1.1），Package=含*1.1（若未達標）
                unit_list_h = int(round((h["List"] / base_std) * factor))
                rate_total_h = int(unit_list_h * spots_final)
                pkg_total_h = int(round(rate_total_h * (1.1 if penalty else 1.0)))

                rows.append({
                    "media_type": media, "region": "全省量販",
                    "program_num": STORE_COUNTS_CONFIG.get("家樂福_量販", 0),
                    "daypart": h["Time"],
                    "seconds": sec,
                    "spots": spots_final,
                    "schedule": sch_h,
                    "rate_list": rate_total_h,
                    "pkg_display_val": pkg_total_h,
                    "group_id": f"{media}__{sec}__1",
                    "is_national": True,
                    "penalty": penalty
                })

                # 超市：依比例換算 spots
                spots_s = int(round(spots_final * (s["Std"] / base_std)))
                if spots_s % 2 != 0:
                    spots_s += 1
                sch_s = calculate_schedule(spots_s, days_count)

                rows.append({
                    "media_type": media, "region": "全省超市",
                    "program_num": STORE_COUNTS_CONFIG.get("家樂福_超市", 0),
                    "daypart": s["Time"],
                    "seconds": sec,
                    "spots": spots_s,
                    "schedule": sch_s,
                    "rate_list": "計量販",
                    "pkg_display_val": "計量販",
                    "group_id": f"{media}__{sec}__1",
                    "is_national": True,
                    "penalty": penalty
                })

    # 排序：媒體 -> seconds -> region order
    def region_rank(x):
        if x["region"] in REGIONS_ORDER:
            return REGIONS_ORDER.index(x["region"])
        if x["region"] == "全省量販":
            return 90
        if x["region"] == "全省超市":
            return 91
        return 999

    rows.sort(key=lambda x: (media_order.get(x["media_type"], 99), x["seconds"], region_rank(x)))
    return rows, debug_logs, pricing_groups


# =========================================================
# 11) Excel 產生：整合 Block Renderer + Total 規則 + Footer
# =========================================================
def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str,
                                rows, remarks_list, template_bytes, pricing_groups):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if meta["sheet_name"] not in wb.sheetnames:
        raise ValueError(f"缺少分頁：{meta['sheet_name']}")

    # 只保留目標分頁
    target_sheet = meta["sheet_name"]
    for sname in list(wb.sheetnames):
        if sname != target_sheet:
            del wb[sname]
    ws = wb[target_sheet]

    # Header
    hc = meta["header_cells"]
    if "client" in hc:  safe_write(ws, hc["client"], client_name)
    if "product" in hc: safe_write(ws, hc["product"], product_display_str)
    if "period" in hc:  safe_write(ws, hc["period"], f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y.%m. %d')}")
    if "medium" in hc:  safe_write(ws, hc["medium"], " ".join(sorted(set([r["media_type"] for r in rows]))))
    if "month" in hc:   safe_write(ws, hc["month"], f" {start_dt.month}月")

    safe_write(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))
    for addr, text in meta.get("header_override", {}).items():
        safe_write(ws, addr, text)

    # 找 blocks + total row
    blocks, total_row_orig = detect_blocks(ws, format_type, meta)

    # 先解除資料區 merge（避免 Total/Footer 被吞）
    data_start = min(b["start"] for b in blocks)
    data_end = total_row_orig - 1
    if format_type == "Dongwu":
        unmerge_overlaps_by_cols(ws, data_start, data_end, ["B", "E", "F", "H"])

    # 分群
    media_groups = build_groups(rows)

    # 逐 block 渲染（由下往上）
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

    # 重新定位 Total row
    total_pos = find_cell_exact_in_ws(ws, meta["total_label"])
    if not total_pos:
        raise ValueError("渲染後找不到 Total（可能 Total 被模板 merge/公式影響）")
    total_row = total_pos[0]

    # Total schedule（照你原本：顯示列的加總）
    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [0] * eff_days
    for x in rows:
        sch = x.get("schedule", [])
        for d in range(min(eff_days, len(sch))):
            daily_sums[d] += sch[d]
    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))

    # ✅ Total 金額規則（你要的最終邏輯）
    # - Total Rate = 所有 row 的 Rate（int）加總（欺敵）
    total_rate = sum([x["rate_list"] for x in rows if isinstance(x.get("rate_list"), int)])

    # - Total Package:
    #   group 若 is_national -> 用 bundle_pkg_total（只加一次）
    #   group 若 not national -> 用 rows 的 pkg_display_val 加總（int）
    group_seen = set()
    total_pkg = 0
    for x in rows:
        gid = x.get("group_id")
        if not gid:
            continue
        if gid in group_seen:
            continue
        group_seen.add(gid)
        ginfo = pricing_groups.get(gid, {})
        if ginfo.get("is_national") and isinstance(ginfo.get("bundle_pkg_total"), int):
            total_pkg += int(ginfo["bundle_pkg_total"])

    # 非全省 group 的 package：直接加 rows 的 package
    #（避免重複加：這裡只加那些 is_national=False 的列）
    for x in rows:
        if x.get("is_national"):
            continue
        if isinstance(x.get("pkg_display_val"), int):
            total_pkg += x["pkg_display_val"]

    # 寫入 Total row 的金額欄
    cols = meta["cols"]
    if format_type == "Dongwu":
        safe_write(ws, f"{cols['rate']}{total_row}", total_rate)
        safe_write(ws, f"{cols['pkg']}{total_row}", total_pkg)
    else:
        proj_col = cols.get("proj_price")
        if proj_col:
            safe_write(ws, f"{proj_col}{total_row}", total_pkg)

    # Remarks
    rem_pos = find_cell_exact_in_ws(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_write_rc(ws, rem_pos[0] + 1 + i, rem_pos[1], rm)

    # 欄位置中
    if format_type == "Dongwu":
        min_data_row = min(b["start"] for b in blocks)
        force_center_columns_range(ws, meta["force_center_cols"], min_data_row, total_row)

    # Footer (make/vat/grand) 先算，後 polish 對齊
    lbl = meta["footer_labels"]
    make_fee = 10000
    pos_make = find_cell_exact_in_ws(ws, lbl["make"])
    if pos_make:
        v = ws.cell(pos_make[0], pos_make[1] + 1).value
        if isinstance(v, (int, float)) and v > 0:
            make_fee = int(v)
        else:
            safe_write_rc(ws, pos_make[0], pos_make[1] + 1, make_fee)

    vat = int(round((total_pkg + make_fee) * 0.05))
    pos_vat = find_cell_exact_in_ws(ws, lbl["vat"])
    if pos_vat:
        safe_write_rc(ws, pos_vat[0], pos_vat[1] + 1, vat)

    pos_grand = find_cell_exact_in_ws(ws, lbl["grand"])
    if pos_grand:
        safe_write_rc(ws, pos_grand[0], pos_grand[1] + 1, total_pkg + make_fee + vat)

    # Footer 美化（VAT/Grand 對齊 + 合約區）
    polish_footer(ws, format_type, meta, client_name)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================================================
# 12) HTML Preview（與 Excel 同變數 rows）
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

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks,
                          pricing_groups, total_rate, total_pkg):
    header_cls = "bg-dw-head" if format_type == "Dongwu" else "bg-sh-head"
    eff_days = min(days_cnt, 31)

    st.markdown(f"""<style>
    .bg-dw-head {{ background-color: #4472C4; color: white; font-weight: bold; }}
    .bg-sh-head {{ background-color: #BDD7EE; color: black; font-weight: bold; }}
    .bg-weekend {{ background-color: #FFD966; color: black; }}
    .bg-total   {{ background-color: #FFF2CC; font-weight: bold; }}
    .col_day {{ min-width: 25px; }}
    .left {{ text-align:left; }}
    .right {{ text-align:right; }}
    table {{ border-collapse: collapse; width: 100%; }}
    th, td {{ border: 1px solid #555; padding: 2px 4px; }}
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
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "Rate<br>(Net)", "Package-cost<br>(Net)"]
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "專案價<br>(Net)"]

    th_fixed = "".join([f"<th rowspan='2'>{c}</th>" for c in cols_def])

    # Station 合併顯示（Dongwu）
    media_counts = {}
    for r in rows:
        media_counts[r["media_type"]] = media_counts.get(r["media_type"], 0) + 1
    media_printed = {m: False for m in media_counts}

    tbody = ""
    for r in rows:
        m = r["media_type"]
        tbody += "<tr>"

        display_name = ("全家便利商店<br>通路廣播廣告" if m == "全家廣播"
                        else "全家便利商店<br>新鮮視廣告" if m == "新鮮視"
                        else "家樂福")
        if format_type == "Shenghuo" and m == "全家廣播":
            display_name = "全家便利商店<br>廣播通路廣告"

        if format_type == "Dongwu":
            if not media_printed[m]:
                tbody += f"<td class='left' rowspan='{media_counts[m]}'>{display_name}</td>"
                media_printed[m] = True
        else:
            tbody += f"<td class='left'>{display_name}</td>"

        tbody += f"<td>{html_escape(region_display(r['region']))}</td>"
        tbody += f"<td class='right'>{fmt_int(r.get('program_num',''))}</td>"
        tbody += f"<td>{html_escape(r.get('daypart',''))}</td>"

        sec_txt = f"{r['seconds']}秒" if format_type == "Dongwu" and m == "家樂福" else f"{r['seconds']}" if format_type == "Dongwu" else f"{r['seconds']}秒廣告"
        tbody += f"<td>{html_escape(sec_txt)}</td>"

        if format_type == "Dongwu":
            tbody += f"<td class='right'>{fmt_int(r.get('rate_list',''))}</td>"
            tbody += f"<td class='right'>{fmt_int(r.get('pkg_display_val',''))}</td>"
        else:
            tbody += f"<td class='right'>{fmt_int(r.get('pkg_display_val',''))}</td>"

        for d in r.get('schedule', [])[:eff_days]:
            tbody += f"<td>{fmt_int(d)}</td>"
        tbody += f"<td class='bg-total'>{fmt_int(sum(r.get('schedule', [])[:eff_days]))}</td></tr>"

    # Total row
    colspan = 5
    empty_td = "<td></td>" if format_type == "Dongwu" else ""
    if format_type == "Dongwu":
        tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td>{empty_td}<td class='right'>{fmt_int(total_rate)}</td><td class='right'>{fmt_int(total_pkg)}</td>"
    else:
        tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td><td class='right'>{fmt_int(total_pkg)}</td>"

    totals = [0]*eff_days
    for d in range(eff_days):
        totals[d] = sum([(rr.get("schedule", [])[d] if d < len(rr.get("schedule", [])) else 0) for rr in rows])

    for t in totals:
        tfoot += f"<td>{fmt_int(t)}</td>"
    tfoot += f"<td>{fmt_int(sum(totals))}</td></tr>"

    return f"""<div class="excel-container">
    <div style="margin-bottom:10px;">
      <b>客戶：</b>{html_escape(c_name)} &nbsp;
      <b>產品：</b>{html_escape(p_display)}<br>
      <span style="color:#666;">走期：{start_dt} ~ {end_dt}</span>
    </div>
    <table class="excel-table">
    <thead><tr>{th_fixed}{date_th1}<th class='{header_cls}' rowspan='2'>檔次</th></tr>
    <tr>{date_th2}</tr></thead>
    <tbody>{tbody}{tfoot}</tbody></table>
    <div class="remarks"><b>Remarks：</b><br>{"<br>".join([html_escape(x) for x in remarks])}</div></div>"""


# =========================================================
# 13) UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v65.3)")

auto_tpl, source, msgs = load_default_template()
template_bytes = auto_tpl

if auto_tpl:
    st.success(f"✅ 已載入系統公版 ({source})")
else:
    st.warning("⚠️ 無法載入公版，請手動上傳")
    tpl = st.file_uploader("上傳 Excel 模板", type=["xlsx"])
    if tpl:
        template_bytes = tpl.read()

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

# 狀態初始化
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
        regs = (["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_reg"))
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
            "regions": (REGIONS_ORDER[:] if is_nat else regs),
            "seconds": sorted(secs),
            "share": st.session_state.rad_share,
            "sec_shares": sec_shares
        }

if is_fv:
    with m2:
        st.markdown("#### 📺 新鮮視")
        is_nat = st.checkbox("全省聯播", False, key="fv_nat")
        regs = (["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg"))
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
            "regions": (REGIONS_ORDER[:] if is_nat else regs),
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
            "is_national": True,
            "regions": ["全省"],
            "seconds": sorted(secs),
            "share": st.session_state.cf_share,
            "sec_shares": sec_shares
        }

# ---------------------------------------------------------
# 計算 & 輸出
# ---------------------------------------------------------
rem = get_remarks_text(sign_deadline, billing_month, payment_date)

if config:
    rows, debug_logs, pricing_groups = build_rows_and_totals(config, total_budget_input, days_count)

    # 產品顯示字串：用使用者選到的秒數去列（去重）
    secs_all = sorted(set([r["seconds"] for r in rows]))
    p_str = f"{'、'.join([f'{s}秒' for s in secs_all])} {product_name}" if rows else ""

    with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
        st.markdown("#### 1. 本次預算分配 / 達標狀態")
        for log in debug_logs:
            color = "green" if log["status"] == "達標" else "red"
            nat_txt = "全省" if log.get("is_national") else "分區"
            st.markdown(
                f"**{log['media']} ({log['sec']}秒 / {nat_txt})**: 預算${log['budget']:,.0f} | "
                f"執行 {log['spots']} 檔 -> "
                f"<span style='color:{color}'><b>{log['status']}</b></span> ({log['reason']})",
                unsafe_allow_html=True
            )

    if rows:
        # 先算 Total Rate / Total Package（跟 Excel 完全同規則）
        total_rate = sum([x["rate_list"] for x in rows if isinstance(x.get("rate_list"), int)])

        group_seen = set()
        total_pkg = 0
        for x in rows:
            gid = x.get("group_id")
            if not gid or gid in group_seen:
                continue
            group_seen.add(gid)
            ginfo = pricing_groups.get(gid, {})
            if ginfo.get("is_national") and isinstance(ginfo.get("bundle_pkg_total"), int):
                total_pkg += int(ginfo["bundle_pkg_total"])
        for x in rows:
            if x.get("is_national"):
                continue
            if isinstance(x.get("pkg_display_val"), int):
                total_pkg += x["pkg_display_val"]

        font_b64 = load_font_base64()
        html = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str,
                                     format_type, rem, pricing_groups, total_rate, total_pkg)
        st.components.v1.html(html, height=700, scrolling=True)

        st.markdown("### 4. 下載")
        prefer_weasy = st.checkbox("強制使用 WeasyPrint 產 PDF（避免 Excel/LibreOffice 跑版）", value=False)

        if template_bytes:
            try:
                xlsx = generate_excel_from_template(format_type, start_date, end_date, client_name, p_str,
                                                    rows, rem, template_bytes, pricing_groups)
                st.download_button("下載 Excel", xlsx, f"Cue_{safe_filename(client_name)}.xlsx")

                if prefer_weasy:
                    pdf_bytes, err2 = html_to_pdf_fallback(html, font_b64)
                    if pdf_bytes:
                        st.download_button("下載 PDF (WeasyPrint)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                    else:
                        st.error(f"PDF 產出失敗(WeasyPrint): {err2}")
                else:
                    pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx)
                    if pdf_bytes:
                        st.download_button(f"下載 PDF ({method})", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                    else:
                        st.warning(f"Excel 轉 PDF 失敗 ({method}: {err})，改用 WeasyPrint...")
                        pdf_bytes, err2 = html_to_pdf_fallback(html, font_b64)
                        if pdf_bytes:
                            st.download_button("下載 PDF (WeasyPrint Fallback)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                        else:
                            st.error(f"PDF 產出失敗: {err2}")

            except Exception as e:
                st.error(f"Excel 產出錯誤: {e}")
        else:
            st.warning("請上傳模板以啟用下載。")
    else:
        st.info("請先勾選媒體並設定秒數/區域，產生資料列後即可預覽與下載。")
else:
    st.info("請先勾選媒體並設定秒數/區域，產生資料列後即可預覽與下載。")
