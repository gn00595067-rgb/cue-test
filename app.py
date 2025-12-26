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
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font

# =========================================================
# 0. 基礎工具
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
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")

def norm_ws(s: str) -> str:
    """把多重空白壓成單一空白，避免 LibreOffice/PDF 轉檔時空白跑版"""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def ceil_div(a, b):
    return int(math.ceil(a / b)) if b else 0

# =========================================================
# 1. 頁面設定 & 自動載入
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v66.0 (Value Anchor + True Layout)")

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
# 2. Excel 轉 PDF (Local 方案)
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
            import win32com.client  # pywin32
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
                    capture_output=True, timeout=90
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

# =========================================================
# 3. WeasyPrint Fallback (雲端擬真)
# =========================================================
def html_to_pdf_fallback(html_str, font_b64):
    try:
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        font_config = FontConfiguration()
        css_str = """
        @page { size: A4 landscape; margin: 0.5cm; }
        body { font-family: 'NotoSansTC', sans-serif !important; font-size: 8pt; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 0.5pt solid #555; padding: 2px 3px; text-align: center; white-space: nowrap; }
        .bg-dw-head { background-color: #4472C4; color: white; font-weight: bold; }
        .bg-sh-head { background-color: #BDD7EE; color: black; font-weight: bold; }
        .bg-weekend { background-color: #FFD966; color: black; }
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
# 4. Master Configuration（依你規格）
# =========================================================
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

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

def get_sec_factor(media_type, seconds):
    return SEC_FACTORS_CONFIG.get(media_type, {}).get(int(seconds), 1.0)

REGION_DISPLAY_6 = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市",
}
def region_display(region: str) -> str:
    return REGION_DISPLAY_6.get(region, region)

def calculate_schedule(total_spots, days):
    if days <= 0:
        return []
    total_spots = int(total_spots)
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
# 5. Excel 輔助：安全寫入 / 樣式 / merge
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

# =========================================================
# 5.x Block Renderer（資料區動態、合併儲存格維持）
# =========================================================
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

def build_groups(rows):
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
            # 重要：Dongwu 的 H 欄是同秒數群組「只顯示一次」的 Package-cost(Net)
            pkg_net_total = items_sorted[0].get("pkg_net_total", 0)
            sec_groups.append({
                "seconds": sec,
                "daypart": items_sorted[0].get("daypart", ""),
                "rows": items_sorted,
                "pkg_net_total": pkg_net_total
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

    # 避免舊 merge 吞掉 Total/Footer
    unmerge_overlaps_by_cols(ws, s, new_end, ["B", "E", "F", "H"])

    # Station(B) 合併整段
    ws.merge_cells(f"{cols['station']}{s}:{cols['station']}{new_end}")
    top_station = ws[f"{cols['station']}{s}"]
    top_station.value = station_title(format_type, blk["media"])
    apply_center_style(top_station)

    # 逐列寫入
    for i, line in enumerate(lines):
        r = s + i
        item = line["item"]

        safe_write(ws, f"{cols['location']}{r}", region_display(item["region"]))
        safe_write(ws, f"{cols['program']}{r}", int(item.get("program_num", 0)))
        safe_write(ws, f"{cols['rate']}{r}", int(item.get("rate_display", 0)))

        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)

    # 同秒數群組：E/F/H 合併（擬真）
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
        cH.value = int(g.get("pkg_net_total", 0))
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
        safe_write(ws, f"{cols['program']}{r}", int(item.get("program_num", 0)))
        safe_write(ws, f"{cols['daypart']}{r}", g["daypart"])
        safe_write(ws, f"{cols['seconds']}{r}", f"{g['seconds']}秒廣告")

        proj_col = cols.get("proj_price")
        if proj_col:
            v = item.get("proj_price_val", 0)
            safe_write(ws, f"{proj_col}{r}", int(v) if isinstance(v, (int, float)) else v)

        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)

# =========================================================
# 5.y Footer 美化（甲乙方/統編拆分 + VAT 對齊）
# =========================================================
def _unmerge_if_merged(ws, cell):
    if not isinstance(cell, MergedCell):
        # 也可能是 master cell 但屬於 merged range
        pass
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= cell.row <= mr.max_row and mr.min_col <= cell.column <= mr.max_col:
            try:
                ws.unmerge_cells(str(mr))
            except:
                pass

def _split_label_value_in_cell(ws, cell, label_text, value_text, value_merge_to_col=None):
    """
    把「甲 方：XXX」這種文字拆成：
    label cell = '甲方：'
    value cell(右邊一格起) = 'XXX' 並可合併到指定欄
    """
    # 找到該 cell 若在 merged range，先取得左上 master
    master = _get_master_cell(ws, cell) if isinstance(cell, MergedCell) else cell
    if master is None:
        return

    # 若 master 本身在 merge 範圍，先解除
    _unmerge_if_merged(ws, master)

    r = master.row
    c = master.column
    lbl_cell = ws.cell(r, c)
    lbl_cell.value = label_text
    apply_right_style(lbl_cell)

    v_start = c + 1
    v_end = value_merge_to_col if value_merge_to_col and value_merge_to_col >= v_start else v_start
    try:
        if v_end > v_start:
            ws.merge_cells(start_row=r, start_column=v_start, end_row=r, end_column=v_end)
    except:
        pass

    val_cell = ws.cell(r, v_start)
    val_cell.value = value_text
    apply_left_style(val_cell)

def polish_footer(ws, format_type, meta, client_name):
    # --- (A) VAT/製作/Grand Total：對齊 + 千分位 ---
    lbl = meta["footer_labels"]
    for key in ["make", "vat", "grand"]:
        pos = find_cell_exact_in_ws(ws, lbl.get(key, ""))
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

    # 欄寬微調（避免 VAT 區塊擠）
    try:
        ws.column_dimensions["G"].width = max(ws.column_dimensions["G"].width or 0, 14)
        ws.column_dimensions["H"].width = max(ws.column_dimensions["H"].width or 0, 18)
    except:
        pass

    # --- (B) Dongwu：甲乙方/統編 拆分，避免「空白排版很醜」 ---
    if format_type != "Dongwu":
        return

    # 甲方：找包含「甲」「方」「：」且後面還有文字
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str):
                continue
            s = v.replace(" ", "").replace("\u3000", "")
            m = re.match(r"甲方[:：](.+)", s)
            if m:
                company = m.group(1)
                master = _get_master_cell(ws, cell) if isinstance(cell, MergedCell) else cell
                # 以原本 merged 範圍的最右欄做 value 合併終點（若有）
                merge_end = master.column
                for mr in list(ws.merged_cells.ranges):
                    if mr.min_row <= master.row <= mr.max_row and mr.min_col <= master.column <= mr.max_col:
                        merge_end = mr.max_col
                        break
                _split_label_value_in_cell(ws, master, "甲方：", company, value_merge_to_col=merge_end)
                break

    # 統一編號：同理拆分（甲方那行常把數字塞一起）
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str):
                continue
            s = v.replace(" ", "").replace("\u3000", "")
            m = re.match(r"統一編號[:：]([0-9]+)", s)
            if m:
                taxid = m.group(1)
                master = _get_master_cell(ws, cell) if isinstance(cell, MergedCell) else cell
                merge_end = master.column
                for mr in list(ws.merged_cells.ranges):
                    if mr.min_row <= master.row <= mr.max_row and mr.min_col <= master.column <= mr.max_col:
                        merge_end = mr.max_col
                        break
                _split_label_value_in_cell(ws, master, "統一編號：", taxid, value_merge_to_col=merge_end)
                break

    # 乙方：把「乙       方：」改乾淨；若同列有 =$C$3 或客戶名，直接覆蓋成純文字避免跑版
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str):
                continue
            s = v.replace(" ", "").replace("\u3000", "")
            if s.startswith("乙方"):
                # 清 label
                master = _get_master_cell(ws, cell) if isinstance(cell, MergedCell) else cell
                _unmerge_if_merged(ws, master)
                master.value = "乙方："
                apply_right_style(master)

                # 嘗試找同列客戶名稱 cell（常見 =$C$3）
                rr = master.row
                found = None
                for cc in range(master.column + 1, min(master.column + 20, ws.max_column) + 1):
                    vv = ws.cell(rr, cc).value
                    if isinstance(vv, str) and vv.startswith("=") and "$C$3" in vv:
                        found = (rr, cc)
                        break
                    if isinstance(vv, str) and client_name in vv:
                        found = (rr, cc)
                        break
                if found:
                    ws.cell(found[0], found[1]).value = client_name
                    apply_left_style(ws.cell(found[0], found[1]))
                break

# =========================================================
# 5.z Sheet Meta（盡量不覆蓋模板表頭文字）
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
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": [],
    }
}

# =========================================================
# 5.main Excel 產生（Block Renderer + Value Anchor Total）
# =========================================================
def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes, total_pkg_override=None):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if meta["sheet_name"] not in wb.sheetnames:
        raise ValueError(f"缺少分頁：{meta['sheet_name']}")

    # 1) 只保留目標分頁
    target_sheet = meta["sheet_name"]
    for sname in list(wb.sheetnames):
        if sname != target_sheet:
            del wb[sname]
    ws = wb[target_sheet]

    # 2) Header
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

    # 3) 找 blocks + total row
    blocks, total_row_orig = detect_blocks(ws, format_type, meta)

    # 4) 先解 merge（避免 Total/Footer 被吞）
    data_start = min(b["start"] for b in blocks)
    data_end = total_row_orig - 1
    if format_type == "Dongwu":
        unmerge_overlaps_by_cols(ws, data_start, data_end, ["B", "E", "F", "H"])

    # 5) 分群
    media_groups = build_groups(rows)

    # 6) 逐 block 渲染（由下往上）
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

    # 7) 重新定位 Total row
    total_pos = find_cell_exact_in_ws(ws, meta["total_label"])
    if not total_pos:
        raise ValueError("渲染後找不到 Total（可能 Total 被模板 merge/公式影響）")
    total_row = total_pos[0]

    # 8) Total 檔次計算（每日加總）
    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [0] * eff_days
    for x in rows:
        sch = x.get("schedule", [])
        for d in range(min(eff_days, len(sch))):
            daily_sums[d] += sch[d]

    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))

    # 9) Total Row Package-cost（Value Anchor 覆蓋）
    cols = meta["cols"]
    pkg_col = cols.get("pkg") or cols.get("proj_price")
    if pkg_col:
        pkg_total = int(total_pkg_override) if isinstance(total_pkg_override, (int, float)) else None
        if pkg_total is None:
            # fallback：避免把每列重複的 group 值加爆，改用 unique group key
            seen = set()
            ssum = 0
            for r in rows:
                k = (r.get("media_type"), r.get("seconds"))
                if k in seen:
                    continue
                seen.add(k)
                ssum += int(r.get("pkg_net_total", 0))
            pkg_total = ssum
        safe_write(ws, f"{pkg_col}{total_row}", int(pkg_total))

    # 10) Remarks
    rem_pos = find_cell_exact_in_ws(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_write_rc(ws, rem_pos[0] + 1 + i, rem_pos[1], rm)

    # 11) Dongwu 欄位置中
    if format_type == "Dongwu":
        min_data_row = min(b["start"] for b in blocks)
        force_center_columns_range(ws, meta["force_center_cols"], min_data_row, total_row)

    # 12) Footer（VAT/甲乙方）
    polish_footer(ws, format_type, meta, client_name)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 6. HTML Preview（Dongwu row-span 擬真）
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

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks, total_pkg_override=None):
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
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(Net)", "Package-cost<br>(Net)"]
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "專案價"]
    th_fixed = "".join([f"<th rowspan='2'>{c}</th>" for c in cols_def])

    # 排序
    rows_sorted = sorted(
        rows,
        key=lambda x: (
            media_order.get(x["media_type"], 99),
            int(x.get("seconds", 0)),
            REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 99
        )
    )

    # 計 rowspan
    media_counts = {}
    sec_counts = {}  # (media, sec) -> count
    for r in rows_sorted:
        media_counts[r["media_type"]] = media_counts.get(r["media_type"], 0) + 1
        sec_counts[(r["media_type"], r["seconds"])] = sec_counts.get((r["media_type"], r["seconds"]), 0) + 1

    tbody = ""
    printed_media = set()
    printed_sec = set()

    for r in rows_sorted:
        m = r["media_type"]
        sec = r["seconds"]
        key_ms = (m, sec)

        tbody += "<tr>"

        # Station rowspan（Dongwu 才合併）
        if format_type == "Dongwu":
            if m not in printed_media:
                printed_media.add(m)
                rowspan = media_counts[m]
                display_name = "全家便利商店<br>通路廣播廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
                tbody += f"<td class='left' rowspan='{rowspan}'>{display_name}</td>"
        else:
            display_name = "全家便利商店<br>廣播通路廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
            tbody += f"<td class='left'>{display_name}</td>"

        # Location / Program
        tbody += f"<td>{region_display(r['region'])}</td>"
        tbody += f"<td class='right'>{r.get('program_num','')}</td>"

        if format_type == "Dongwu":
            # Day-part / Size / Package-cost：同媒體同秒數 group 做 rowspan 擬真
            if key_ms not in printed_sec:
                printed_sec.add(key_ms)
                rowspan_s = sec_counts[key_ms]
                tbody += f"<td rowspan='{rowspan_s}'>{html_escape(r.get('daypart',''))}</td>"
                tbody += f"<td rowspan='{rowspan_s}'>{int(sec)}</td>"
                # rate（每列）
                tbody += f"<td class='right'>{int(r.get('rate_display', 0)):,}</td>"
                # package（只顯示一次）
                pkg_show = int(r.get("pkg_net_total", 0))
                tbody += f"<td class='right' rowspan='{rowspan_s}'>{pkg_show:,}</td>"
            else:
                # rate（每列）
                tbody += f"<td class='right'>{int(r.get('rate_display', 0)):,}</td>"
        else:
            tbody += f"<td>{html_escape(r.get('daypart',''))}</td>"
            tbody += f"<td>{int(r.get('seconds',0))}秒廣告</td>"
            tbody += f"<td class='right'>{int(r.get('proj_price_val',0)):,}</td>"

        # 日檔次
        for d in r['schedule'][:eff_days]:
            tbody += f"<td>{d}</td>"

        tbody += f"<td class='bg-total'>{sum(r['schedule'][:eff_days])}</td></tr>"

    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]

    # Total Row Package-cost（用 override）
    if total_pkg_override is None:
        seen = set()
        ssum = 0
        for r in rows:
            k = (r.get("media_type"), r.get("seconds"))
            if k in seen:
                continue
            seen.add(k)
            ssum += int(r.get("pkg_net_total", 0))
        total_pkg_override = ssum

    if format_type == "Dongwu":
        colspan = 6
        tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td><td class='right'>{int(total_pkg_override):,}</td>"
    else:
        colspan = 5
        tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td><td class='right'>{int(total_pkg_override):,}</td>"

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
# 7. 商業計算核心（依你規格修正）
# =========================================================
def calc_spots_and_group_pkg(media, sec, budget_alloc, is_national, regions):
    """
    回傳：
    - spots_final
    - penalty_applied(bool)
    - unit_cost_net（含 sec factor / 未含 penalty）
    - group_pkg_net_total（Package-cost(Net) 要顯示的整包金額）
    - display_regions（要展開顯示的區域）
    """
    sec = int(sec)
    factor = get_sec_factor(media, sec)

    if media in ("全家廣播", "新鮮視"):
        cfg = PRICING_CONFIG[media]
        std = cfg["Std_Spots"]
        prices = cfg["Prices"]

        if is_national:
            base_net = prices["全省"]["Net"]
            display_regions = REGIONS_ORDER[:]  # 固定展開六區
        else:
            # 分區：用選定區域 net 加總
            base_net = sum(prices[r]["Net"] for r in regions)
            display_regions = regions[:]

        unit_cost_net = (base_net / std) * factor
        spots_init = ceil_div(budget_alloc, unit_cost_net)
        penalty = 1.1 if spots_init < std else 1.0
        spots_final = ceil_div(budget_alloc, unit_cost_net * penalty)

        if spots_final % 2 != 0:
            spots_final += 1
        if spots_final <= 0:
            spots_final = 2

        # Package-cost(Net)：用「含 penalty 後」的單檔成本 * spots（這樣總額會貼近預算邏輯）
        group_pkg_net_total = int(round((unit_cost_net * penalty) * spots_final))

        return spots_final, (penalty > 1.0), unit_cost_net, group_pkg_net_total, display_regions, std, factor

    # 家樂福：以量販為基準算 spots，超市按比例推 spots
    if media == "家樂福":
        conf = PRICING_CONFIG["家樂福"]["Configs"]
        base = conf["量販_全省"]
        std = base["Std"]
        unit_cost_net = (base["Net"] / std) * factor
        spots_init = ceil_div(budget_alloc, unit_cost_net)
        penalty = 1.1 if spots_init < std else 1.0
        spots_final = ceil_div(budget_alloc, unit_cost_net * penalty)
        if spots_final % 2 != 0:
            spots_final += 1
        if spots_final <= 0:
            spots_final = 2
        group_pkg_net_total = int(round((unit_cost_net * penalty) * spots_final))
        return spots_final, (penalty > 1.0), unit_cost_net, group_pkg_net_total, ["全省"], std, factor

    return 0, False, 0, 0, [], 0, factor

# =========================================================
# 8. UI Main
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
            for s in sorted(secs)[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls / 2), key=f"fs_{s}")
                sec_shares[s] = v
                ls -= v
            sec_shares[sorted(secs)[-1]] = ls
        elif secs:
            sec_shares[secs[0]] = 100
        config["新鮮視"] = {"is_national": is_nat, "regions": regs, "seconds": sorted(secs), "share": st.session_state.fv_share, "sec_shares": sec_shares}

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
        config["家樂福"] = {"regions": ["全省"], "seconds": sorted(secs), "share": st.session_state.cf_share, "sec_shares": sec_shares}

# ---------------------------------------------------------
# 計算引擎（rows + group total）
# ---------------------------------------------------------
rows = []
debug_logs = []
total_pkg_override = 0

if config:
    for m, cfg in config.items():
        m_budget = total_budget_input * (cfg["share"] / 100.0)
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0:
                continue

            if m in ["全家廣播", "新鮮視"]:
                is_nat = bool(cfg.get("is_national"))
                sel_regions = cfg["regions"]
                spots_final, penalty_applied, unit_cost_net, group_pkg_net_total, display_regions, std, factor = \
                    calc_spots_and_group_pkg(m, sec, s_budget, is_nat, sel_regions if not is_nat else REGIONS_ORDER)

                sch = calculate_schedule(spots_final, days_count)

                # Total Row 的 Package-cost(Net) 用「每個群組」加總
                total_pkg_override += int(group_pkg_net_total)

                debug_logs.append({
                    "media": m, "sec": int(sec), "budget": s_budget,
                    "unit_cost_net": unit_cost_net,
                    "spots": spots_final,
                    "std": std,
                    "factor": factor,
                    "status": "未達標" if penalty_applied else "達標",
                    "reason": "懲罰 x1.1" if penalty_applied else "費率正常",
                    "pkg_net_total": int(group_pkg_net_total),
                    "mode": "全省" if is_nat else "分區"
                })

                prices = PRICING_CONFIG[m]["Prices"]
                daypart = PRICING_CONFIG[m]["Day_Part"]

                # 展開顯示列：全省 -> 固定六區；分區 -> 選定區域
                for rgn in (REGIONS_ORDER if is_nat else sel_regions):
                    list_price = prices[rgn]["List"]
                    rate_display = int((list_price / std) * factor)  # 欺敵：用 List 換算單價

                    program_key = f"新鮮視_{rgn}" if m == "新鮮視" else rgn
                    program_num = int(STORE_COUNTS_CONFIG.get(program_key, 0))

                    rows.append({
                        "media_type": m,
                        "region": rgn,
                        "program_num": program_num,
                        "daypart": daypart,
                        "seconds": int(sec),
                        "schedule": sch,
                        "rate_display": rate_display,
                        "pkg_net_total": int(group_pkg_net_total),  # 同群組一致（Excel/HTML 會合併顯示）
                    })

            elif m == "家樂福":
                # 家樂福：以量販計算 spots，再推超市 spots
                is_nat = True
                spots_final, penalty_applied, unit_cost_net, group_pkg_net_total, _, std, factor = \
                    calc_spots_and_group_pkg(m, sec, s_budget, is_nat, ["全省"])

                sch_h = calculate_schedule(spots_final, days_count)

                total_pkg_override += int(group_pkg_net_total)

                debug_logs.append({
                    "media": m, "sec": int(sec), "budget": s_budget,
                    "unit_cost_net": unit_cost_net,
                    "spots": spots_final,
                    "std": std,
                    "factor": factor,
                    "status": "未達標" if penalty_applied else "達標",
                    "reason": "懲罰 x1.1" if penalty_applied else "費率正常",
                    "pkg_net_total": int(group_pkg_net_total),
                    "mode": "鎖定"
                })

                # 顯示：量販 + 超市（超市依 Std 比例推 spots）
                conf = PRICING_CONFIG["家樂福"]["Configs"]
                base_std = conf["量販_全省"]["Std"]
                std_s = conf["超市_全省"]["Std"]

                spots_s = int(spots_final * (std_s / base_std))
                sch_s = calculate_schedule(spots_s, days_count)

                # rate_display：家樂福在 Dongwu 的 G 欄仍是「List 換算單價」的參考
                list_unit_h = int((conf["量販_全省"]["List"] / base_std) * factor)
                list_unit_s = int((conf["超市_全省"]["List"] / std_s) * factor)

                rows.append({
                    "media_type": m,
                    "region": "全省量販",
                    "program_num": int(STORE_COUNTS_CONFIG.get("家樂福_量販", 0)),
                    "daypart": conf["量販_全省"]["Time"],
                    "seconds": int(sec),
                    "schedule": sch_h,
                    "rate_display": list_unit_h,
                    "pkg_net_total": int(group_pkg_net_total),
                })
                rows.append({
                    "media_type": m,
                    "region": "全省超市",
                    "program_num": int(STORE_COUNTS_CONFIG.get("家樂福_超市", 0)),
                    "daypart": conf["超市_全省"]["Time"],
                    "seconds": int(sec),
                    "schedule": sch_s,
                    "rate_display": list_unit_s,
                    "pkg_net_total": int(group_pkg_net_total),
                })

p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}" if rows else ""
rem = get_remarks_text(sign_deadline, billing_month, payment_date)

with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
    st.markdown("#### 1. 本次預算分配 / 檔次 / Package-cost(Net)")
    for log in debug_logs:
        color = "green" if log["status"] == "達標" else "red"
        st.markdown(
            f"**{log['media']} ({log['sec']}秒 / {log['mode']})**: "
            f"預算 ${log['budget']:,.0f} | 檔次 {log['spots']} "
            f"| Package(Net) ${log['pkg_net_total']:,.0f} -> "
            f"<span style='color:{color}'><b>{log['status']}</b></span>",
            unsafe_allow_html=True
        )

# ---------------------------------------------------------
# 輸出
# ---------------------------------------------------------
if rows:
    font_b64 = load_font_base64()
    html = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem, total_pkg_override=total_pkg_override)
    st.components.v1.html(html, height=720, scrolling=True)

    if template_bytes:
        try:
            xlsx = generate_excel_from_template(
                format_type, start_date, end_date, client_name, p_str, rows, rem, template_bytes,
                total_pkg_override=total_pkg_override
            )
            st.download_button("下載 Excel", xlsx, f"Cue_{safe_filename(client_name)}.xlsx")

            # PDF：先試 Excel/LO，失敗再用 WeasyPrint（雲端擬真）
            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx)
            if pdf_bytes:
                st.download_button(f"下載 PDF ({method})", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
            else:
                st.warning(f"Excel 轉 PDF 失敗 ({method}: {err})，切換至 WeasyPrint 擬真輸出 (HTML)...")
                pdf_bytes, err2 = html_to_pdf_fallback(html, font_b64)
                if pdf_bytes:
                    st.download_button("下載 PDF (WeasyPrint)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                else:
                    st.error(f"PDF 產出失敗: {err2}")

        except Exception as e:
            st.error(f"Excel 產出錯誤: {e}")
    else:
        st.warning("請上傳模板以啟用下載。")
else:
    st.info("請先勾選媒體並設定秒數/區域，產生資料列後即可預覽與下載。")
