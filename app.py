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

# =========================================================
# 1. 頁面設定 & 自動載入
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v65.2 (Master Spec v2.0 Aligned)")

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
# 2. Excel 轉 PDF (最擬真方案)
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

    # 2) LibreOffice (Cloud 首選)
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

# =========================================================
# 3. WeasyPrint Fallback (備案)
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
# 4. MASTER CONFIGURATION v2.0（規格書：絕對真理）
# =========================================================
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
            "東區":   {"List": 62500,  "Net": 50000}
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
            "東區":   {"List": 45000,  "Net": 36000}
        }
    },
    "家樂福": {
        "Configs": {
            "量販_全省": {"List": 300000, "Net": 250000, "Std": 420, "Time": "09:00-23:00"},
            "超市_全省": {"List": 100000, "Net": 80000,  "Std": 720, "Time": "00:00-24:00"}
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

def get_sec_factor(media_type, seconds):
    return SEC_FACTORS_CONFIG.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
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

REGION_DISPLAY_6 = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市",
}
def region_display(region: str) -> str:
    return REGION_DISPLAY_6.get(region, region)

# -------------------------
# v2.0 計算工具（Spot 計算嚴格按規格）
# -------------------------
def compute_spots_final(budget: float, base_net: float, std_spots: int, factor: float):
    """
    規格書 v2.0：
    Net_Unit_Cost = (Base_Net / Std_Spots) * Sec_Factor
    Spots_Init = ceil(Budget / Net_Unit_Cost)
    if Spots_Init < Std_Spots => Net_Unit_Cost *= 1.1 => Spots_Final = ceil(Budget / Net_Unit_Cost)
    else Spots_Final = Spots_Init
    奇數補偶數
    """
    if std_spots <= 0 or base_net <= 0 or factor <= 0:
        return 0, 0, 0, 0.0, 1.0

    net_unit = (base_net / std_spots) * factor
    if net_unit <= 0:
        return 0, 0, 0, 0.0, 1.0

    spots_init = int(math.ceil(budget / net_unit))
    penalty = 1.1 if spots_init < std_spots else 1.0
    net_unit2 = net_unit * penalty
    spots_final = int(math.ceil(budget / net_unit2))

    if spots_final % 2 != 0:
        spots_final += 1
    if spots_final <= 0:
        spots_final = 2

    return spots_init, spots_final, std_spots, net_unit2, penalty

def display_rate_list(media: str, region: str, seconds: int):
    """
    規格書 v2.0：
    欄位雖叫 Net，但要填 List 的換算：
    Display_Rate = (Region_List_Price / Std_Spots) * Sec_Factor
    取 int 以符合你原本表格的整數風格
    """
    if media not in PRICING_CONFIG:
        return 0
    factor = get_sec_factor(media, seconds)
    if media in ["全家廣播", "新鮮視"]:
        std = PRICING_CONFIG[media]["Std_Spots"]
        lp = PRICING_CONFIG[media]["Prices"][region]["List"]
        return int((lp / std) * factor)
    return 0

def compute_sheet_total_pkg(rows):
    """
    規格書 v2.0 Total Row：
    - 分區購買：Total = sum(所有 Row_Total)
    - 全省聯播：Total 不能加總 rows，要用 全省 List 回推：
      Grand_List_Total = ((全省 List / Std_Spots) * Sec_Factor) * Spots_Final
    支援同一張表同時有「分區」與「全省」混合、以及多媒體/多秒數。
    """
    total = 0
    # group by (media, seconds, is_national_buy)
    groups = {}
    for r in rows:
        key = (r.get("media_type"), r.get("seconds"), bool(r.get("is_national_buy", False)))
        groups.setdefault(key, []).append(r)

    for (media, sec, is_nat), items in groups.items():
        if not media or sec is None:
            continue

        if media in ["全家廣播", "新鮮視"]:
            if is_nat:
                # 用「全省 List」回推（不加總六區列）
                factor = get_sec_factor(media, sec)
                std = PRICING_CONFIG[media]["Std_Spots"]
                nat_list = PRICING_CONFIG[media]["Prices"]["全省"]["List"]
                spots_final = int(items[0].get("spots", 0))
                nat_rate = int((nat_list / std) * factor)
                total += int(nat_rate * spots_final)
            else:
                # 分區：加總每列 row_total
                for it in items:
                    v = it.get("pkg_display_val")
                    if isinstance(v, int):
                        total += v

        elif media == "家樂福":
            # 你原本的商務邏輯：超市列為「計量販」，Total 以可加總者為主
            for it in items:
                v = it.get("pkg_display_val")
                if isinstance(v, int):
                    total += v

    return int(total)

# =========================================================
# 5. Excel 輔助：安全寫入 / 複製樣式 / 解除合併
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
# 5.x Excel 動態表格渲染：Block Renderer
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

    ws.merge_cells(f"{cols['station']}{s}:{cols['station']}{new_end}")
    top_station = ws[f"{cols['station']}{s}"]
    top_station.value = station_title(format_type, blk["media"])
    apply_center_style(top_station)

    for i, line in enumerate(lines):
        r = s + i
        item = line["item"]

        safe_write(ws, f"{cols['location']}{r}", region_display(item["region"]))
        prog_val = item.get("program_num", 0)
        safe_write(ws, f"{cols['program']}{r}", int(prog_val))

        safe_write(ws, f"{cols['rate']}{r}", item.get("rate_list"))

        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)

    for g_start, g_end, g in group_ranges:
        r1 = s + g_start
        r2 = s + g_end

        ws.merge_cells(f"{cols['daypart']}{r1}:{cols['daypart']}{r2}")
        cE = ws[f"{cols['daypart']}{r1}"]
        cE.value = g["daypart"]
        apply_center_style(cE)

        ws.merge_cells(f"{cols['seconds']}{r1}:{cols['seconds']}{r2}")
        cF = ws[f"{cols['seconds']}{r1}"]
        if blk["media"] == "家樂福":
            cF.value = f"{g['seconds']}秒"
        else:
            cF.value = int(g["seconds"])
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
        prog_val = item.get("program_num", 0)
        safe_write(ws, f"{cols['program']}{r}", int(prog_val))

        safe_write(ws, f"{cols['daypart']}{r}", g["daypart"])
        safe_write(ws, f"{cols['seconds']}{r}", f"{g['seconds']}秒廣告")

        proj_col = cols.get("proj_price")
        if proj_col:
            v = item.get("pkg_display_val", 0)
            safe_write(ws, f"{proj_col}{r}", int(v) if isinstance(v, int) else 0)

        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)

# =========================================================
# 5.y Footer 美化（甲乙方/日期 + VAT 區塊對齊）
# =========================================================
def polish_footer(ws, format_type, meta, client_name):
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

        try:
            if isinstance(val_cell.value, (int, float)):
                val_cell.number_format = "#,##0"
        except:
            pass

        if key == "grand":
            label_cell.font = copy(label_cell.font) if label_cell.font else Font(bold=True)
            val_cell.font = copy(val_cell.font) if val_cell.font else Font(bold=True)

    if format_type == "Dongwu":
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and ("甲" in v) and ("方" in v) and ("：" in v):
                    cell.value = "甲方："
                    apply_right_style(cell)
                    try:
                        cB = ws.cell(cell.row, column_index_from_string("B"))
                        if isinstance(cB, MergedCell):
                            cB = _get_master_cell(ws, cB) or cB
                        apply_left_style(cB)
                    except:
                        pass

                if isinstance(v, str) and ("乙" in v) and ("方" in v) and ("：" in v):
                    cell.value = "乙方："
                    apply_right_style(cell)

        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("=") and "$C$3" in cell.value:
                    cell.value = client_name
                    apply_left_style(cell)

        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and "日期" in v and "：" in v:
                    cell.value = "日期："
                    apply_right_style(cell)

# =========================================================
# 5.z Sheet Meta
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
        "header_override": {"G7": "rate\n(List)", "H7": "Package-cost\n(List)"},
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
# 5.main Excel 產生（整合 Block Renderer + Footer Polish）
# =========================================================
def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if meta["sheet_name"] not in wb.sheetnames:
        raise ValueError(f"缺少分頁：{meta['sheet_name']}")

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

    # blocks + total row
    blocks, total_row_orig = detect_blocks(ws, format_type, meta)

    # 解除資料區 merge
    data_start = min(b["start"] for b in blocks)
    data_end = total_row_orig - 1
    if format_type == "Dongwu":
        unmerge_overlaps_by_cols(ws, data_start, data_end, ["B", "E", "F", "H"])

    # 分群 & 渲染
    media_groups = build_groups(rows)
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

    # Total schedule
    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [0] * eff_days
    for x in rows:
        sch = x.get("schedule", [])
        for d in range(min(eff_days, len(sch))):
            daily_sums[d] += sch[d]
    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))

    # Total package cost（v2.0：National Buy 不可加總）
    cols = meta["cols"]
    pkg_col = cols.get("pkg") or cols.get("proj_price")
    total_pkg = compute_sheet_total_pkg(rows)
    if pkg_col:
        safe_write(ws, f"{pkg_col}{total_row}", total_pkg)

    # Footer values (make/vat/grand)
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

    # Remarks
    rem_pos = find_cell_exact_in_ws(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_write_rc(ws, rem_pos[0] + 1 + i, rem_pos[1], rm)

    # Dongwu 欄位置中
    if format_type == "Dongwu":
        min_data_row = min(b["start"] for b in blocks)
        force_center_columns_range(ws, meta["force_center_cols"], min_data_row, total_row)

    # Footer 美化
    polish_footer(ws, format_type, meta, client_name)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 6. HTML Preview
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
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(List)", "Package<br>(List)"]
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "專案價"]
    th_fixed = "".join([f"<th rowspan='2'>{c}</th>" for c in cols_def])

    rows_sorted = sorted(rows, key=lambda x: (media_order.get(x["media_type"], 99), x["seconds"], REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 99))
    tbody = ""
    media_counts = {}
    for r in rows_sorted:
        media_counts[r["media_type"]] = media_counts.get(r["media_type"], 0) + 1
    media_printed = {m: False for m in media_counts}

    for r in rows_sorted:
        m = r["media_type"]
        tbody += "<tr>"
        if not media_printed[m]:
            rowspan = media_counts[m]
            display_name = "全家便利商店<br>通路廣播廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
            if format_type == "Shenghuo" and m == "全家廣播":
                display_name = "全家便利商店<br>廣播通路廣告"
            if format_type == "Shenghuo":
                tbody += f"<td class='left'>{display_name}</td>"
            else:
                tbody += f"<td class='left' rowspan='{rowspan}'>{display_name}</td>"
                media_printed[m] = True
        elif format_type == "Shenghuo":
            display_name = "全家便利商店<br>廣播通路廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
            tbody += f"<td class='left'>{display_name}</td>"

        tbody += f"<td>{region_display(r['region'])}</td><td class='right'>{r.get('program_num','')}</td><td>{r['daypart']}</td>"
        sec_txt = f"{r['seconds']}秒" if format_type == "Dongwu" and m == "家樂福" else f"{r['seconds']}" if format_type == "Dongwu" else f"{r['seconds']}秒廣告"
        tbody += f"<td>{sec_txt}</td>"
        rate = f"{r['rate_list']:,}" if isinstance(r['rate_list'], int) else r['rate_list']
        pkg = f"{r['pkg_display_val']:,}" if isinstance(r['pkg_display_val'], int) else r['pkg_display_val']
        if format_type == "Dongwu":
            tbody += f"<td class='right'>{rate}</td><td class='right'>{pkg}</td>"
        else:
            tbody += f"<td class='right'>{pkg}</td>"

        for d in r['schedule'][:eff_days]:
            tbody += f"<td>{d}</td>"
        tbody += f"<td class='bg-total'>{sum(r['schedule'])}</td></tr>"

    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]
    total_pkg = compute_sheet_total_pkg(rows)  # v2.0：National Buy 不可加總 rows
    colspan = 5
    empty_td = "<td></td>" if format_type == "Dongwu" else ""
    tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='left'>Total</td>{empty_td}<td class='right'>{total_pkg:,}</td>"
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
# 7. UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v65.2 - Master Spec v2.0 aligned)")

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
# 計算引擎（v2.0 aligned）
# ---------------------------------------------------------
rows = []
debug_logs = []

if config:
    for media, cfg in config.items():
        m_budget = total_budget_input * (cfg["share"] / 100.0)

        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0:
                continue

            factor = get_sec_factor(media, sec)

            if media in ["全家廣播", "新鮮視"]:
                std = PRICING_CONFIG[media]["Std_Spots"]
                daypart = PRICING_CONFIG[media]["Day_Part"]

                # v2.0：Base_Net 決定 Spots 計算
                if cfg["is_national"]:
                    # 全省聯播：Base_Net = 全省 Net
                    base_net = PRICING_CONFIG[media]["Prices"]["全省"]["Net"]
                    calc_regions = ["全省"]
                    display_regions = REGIONS_ORDER  # 保持你原本「顯示六區」的方式
                else:
                    # 分區購買：Base_Net = 選定區域 Net 加總
                    calc_regions = cfg["regions"]
                    display_regions = cfg["regions"]
                    base_net = sum(PRICING_CONFIG[media]["Prices"][r]["Net"] for r in calc_regions)

                spots_init, spots_final, std_spots, unit_net_cost, penalty = compute_spots_final(
                    budget=s_budget, base_net=base_net, std_spots=std, factor=factor
                )
                sch = calculate_schedule(spots_final, days_count)

                debug_logs.append({
                    "media": media,
                    "sec": sec,
                    "budget": s_budget,
                    "base_net": base_net,
                    "unit_cost": unit_net_cost,
                    "spots_init": spots_init,
                    "spots": spots_final,
                    "std": std_spots,
                    "factor": factor,
                    "status": "未達標" if penalty > 1 else "達標",
                    "reason": "懲罰 x1.1" if penalty > 1 else "費率正常",
                    "mode": "全省聯播" if cfg["is_national"] else "分區購買",
                })

                for r in display_regions:
                    rate_list = display_rate_list(media, r, sec)
                    pkg_list = int(rate_list * spots_final)

                    # store count
                    if media == "新鮮視":
                        prog_num = STORE_COUNTS_CONFIG.get(f"新鮮視_{r}", 0)
                    else:
                        prog_num = STORE_COUNTS_CONFIG.get(r, 0)

                    rows.append({
                        "media_type": media,
                        "region": r,
                        "program_num": prog_num,
                        "daypart": daypart,
                        "seconds": sec,
                        "spots": spots_final,
                        "schedule": sch,
                        "rate_list": rate_list,
                        "pkg_display_val": pkg_list,
                        "is_national_buy": bool(cfg["is_national"]),
                    })

            elif media == "家樂福":
                cfgs = PRICING_CONFIG["家樂福"]["Configs"]
                base_std = cfgs["量販_全省"]["Std"]
                base_net = cfgs["量販_全省"]["Net"]

                spots_init, spots_final, std_spots, unit_net_cost, penalty = compute_spots_final(
                    budget=s_budget, base_net=base_net, std_spots=base_std, factor=factor
                )

                sch_h = calculate_schedule(spots_final, days_count)

                debug_logs.append({
                    "media": media,
                    "sec": sec,
                    "budget": s_budget,
                    "base_net": base_net,
                    "unit_cost": unit_net_cost,
                    "spots_init": spots_init,
                    "spots": spots_final,
                    "std": std_spots,
                    "factor": factor,
                    "status": "未達標" if penalty > 1 else "達標",
                    "reason": "懲罰 x1.1" if penalty > 1 else "費率正常",
                    "mode": "家樂福(量販為主)",
                })

                # 量販列（可計價）
                rate_h = int((cfgs["量販_全省"]["List"] / base_std) * factor)
                rows.append({
                    "media_type": media,
                    "region": "全省量販",
                    "program_num": STORE_COUNTS_CONFIG.get("家樂福_量販", 0),
                    "daypart": cfgs["量販_全省"]["Time"],
                    "seconds": sec,
                    "spots": spots_final,
                    "schedule": sch_h,
                    "rate_list": rate_h,
                    "pkg_display_val": int(rate_h * spots_final),
                    "is_national_buy": False,
                })

                # 超市列（維持你原本「計量販」的顯示邏輯）
                spots_s = int(spots_final * (cfgs["超市_全省"]["Std"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({
                    "media_type": media,
                    "region": "全省超市",
                    "program_num": STORE_COUNTS_CONFIG.get("家樂福_超市", 0),
                    "daypart": cfgs["超市_全省"]["Time"],
                    "seconds": sec,
                    "spots": spots_s,
                    "schedule": sch_s,
                    "rate_list": "計量販",
                    "pkg_display_val": "計量販",
                    "is_national_buy": False,
                })

p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}" if rows else ""
rem = get_remarks_text(sign_deadline, billing_month, payment_date)

with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
    st.markdown("#### 1. 本次預算分配 (v2.0)")
    for log in debug_logs:
        color = "green" if log["status"] == "達標" else "red"
        st.markdown(
            f"**{log['media']} / {log['mode']} ({log['sec']}秒)**: 預算${log['budget']:,.0f} | "
            f"Init={log['spots_init']} -> Final={log['spots']} (Std={log['std']}) -> "
            f"<span style='color:{color}'><b>{log['status']}</b></span>",
            unsafe_allow_html=True
        )

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
