import streamlit as st
import math
import io
import os
import re
import base64
import tempfile
import subprocess
import shutil
from datetime import timedelta, datetime, date

import requests
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment
from copy import copy


# =========================================================
# 0) 基礎工具
# =========================================================
def parse_count_to_int(x):
    """將含有逗號的字串或數字轉為整數（例如：'4,437店' -> 4437）"""
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
    s = str(s)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


# =========================================================
# 1) Streamlit 頁面設定
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v63.0 (Excel/PDF 擬真強化版)")


# =========================================================
# 2) 資料庫（你原邏輯保留）
# =========================================================
STORE_COUNTS_RAW = {
    "全省": "4,437店",
    "北區": "1,649店",
    "桃竹苗": "779店",
    "中區": "839店",
    "雲嘉南": "499店",
    "高屏": "490店",
    "東區": "181店",
    "新鮮視_全省": "3,124面",
    "新鮮視_北區": "1,127面",
    "新鮮視_桃竹苗": "616面",
    "新鮮視_中區": "528面",
    "新鮮視_雲嘉南": "365面",
    "新鮮視_高屏": "405面",
    "新鮮視_東區": "83面",
    "家樂福_量販": "67店",
    "家樂福_超市": "250店",
}
STORE_COUNTS_NUM = {k: parse_count_to_int(v) for k, v in STORE_COUNTS_RAW.items()}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]  # 你說 2026 改 6 區：OK，這裡就是 6 區
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

PRICING_DB = {
    "全家廣播": {
        "Std_Spots": 480,
        "Day_Part": "00:00-24:00",
        "全省": [400000, 320000],
        "北區": [250000, 200000],
        "桃竹苗": [150000, 120000],
        "中區": [150000, 120000],
        "雲嘉南": [100000, 80000],
        "高屏": [100000, 80000],
        "東區": [62500, 50000],
    },
    "新鮮視": {
        "Std_Spots": 504,
        "Day_Part": "07:00-22:00",
        "全省": [150000, 120000],
        "北區": [150000, 120000],
        "桃竹苗": [120000, 96000],
        "中區": [90000, 72000],
        "雲嘉南": [75000, 60000],
        "高屏": [75000, 60000],
        "東區": [45000, 36000],
    },
    "家樂福": {
        "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420, "Day_Part": "09:00-23:00"},
        "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720, "Day_Part": "00:00-24:00"},
    },
}

SEC_FACTORS = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5, 5: 0.25},
    "新鮮視": {30: 3.0, 20: 2.0, 15: 1.5, 10: 1.0, 5: 0.5},
    "家樂福": {30: 1.5, 20: 1.0, 15: 0.85, 10: 0.65, 5: 0.35},
}


def get_sec_factor(media_type, seconds):
    return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)


def calculate_schedule(total_spots, days):
    """維持你目前做法：每格是偶數（*2）分配"""
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
        "3.通路店舖數與開機率開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。",
        f"6.付款兌現日期：{p_str}",
    ]


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


def region_display(region: str) -> str:
    return REGION_DISPLAY_6.get(region, region)


# =========================================================
# 3) 模板載入（本機 / Drive）
# =========================================================
GOOGLE_DRIVE_FILE_ID = "11R1SA_hpFD5O_MGmYeh4BdtcUhK2bPta"  # 你原本那個
DEFAULT_FILENAME = "1209-Cue表相關資料.xlsx"


@st.cache_data(ttl=600)
def load_template_bytes():
    msgs = []
    # 1) Drive
    if GOOGLE_DRIVE_FILE_ID:
        url = f"https://drive.google.com/uc?export=download&id={GOOGLE_DRIVE_FILE_ID}"
        try:
            r = requests.get(url, timeout=30, allow_redirects=True)
            if r.status_code == 200 and r.content and (b"PK" in r.content[:20]):  # xlsx 檔頭常見 PK
                return r.content, "Google Drive", msgs
            else:
                msgs.append("⚠️ Google Drive 下載失敗或權限不足（可能需要公開分享）")
        except Exception as e:
            msgs.append(f"⚠️ Google Drive 連線錯誤：{e}")

    # 2) Local file
    if os.path.exists(DEFAULT_FILENAME):
        try:
            with open(DEFAULT_FILENAME, "rb") as f:
                return f.read(), "Local", msgs
        except PermissionError:
            # Windows 常見：檔案被 Excel 開著會鎖住
            msgs.append("⚠️ 本機模板檔可能被 Excel 開啟鎖住，請先關閉 Excel 後再試。")
        except Exception as e:
            msgs.append(f"⚠️ 本機讀取失敗：{e}")

    return None, None, msgs


# =========================================================
# 4) Excel 生成（openpyxl）— 強化：對齊、刪除不必要分頁
# =========================================================
def _get_master_cell(ws, cell):
    if not isinstance(cell, MergedCell):
        return cell
    r, c = cell.row, cell.column
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
            return ws.cell(row=mr.min_row, column=mr.min_col)
    return None


def safe_set_value(ws, row: int, col: int, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master is not None:
            master.value = value
        return
    cell.value = value


def safe_set_value_addr(ws, addr: str, value):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master is not None:
            master.value = value
        return
    cell.value = value


def set_center(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
        indent=al.indent,
        text_rotation=al.text_rotation,
        shrink_to_fit=al.shrink_to_fit,
    )


def set_left(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True,
        indent=al.indent,
        text_rotation=al.text_rotation,
        shrink_to_fit=al.shrink_to_fit,
    )


def apply_center_addr(ws, addr: str):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master is None:
            return
        cell = master
    set_center(cell)


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
            except Exception:
                dc.value = v
        else:
            dc.value = v


def find_cell_exact(ws, text):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == text:
                return cell.row, cell.column
    return None


def find_first_row_contains(ws, col_letter, keyword):
    col_idx = column_index_from_string(col_letter)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if isinstance(v, str) and keyword in v:
            return r
    return None


def unmerge_col_overlap(ws, col_letter, start_row, end_row):
    st_col = column_index_from_string(col_letter)
    to_unmerge = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col == st_col and mr.max_col == st_col:
            if not (mr.max_row < start_row or mr.min_row > end_row):
                to_unmerge.append(str(mr))
    for s in set(to_unmerge):
        try:
            ws.unmerge_cells(s)
        except Exception:
            pass


def set_schedule(ws, row, start_col_letter, max_days, schedule_list):
    start_col = column_index_from_string(start_col_letter)
    for i in range(max_days):
        v = schedule_list[i] if (schedule_list and i < len(schedule_list)) else None
        safe_set_value(ws, row, start_col + i, v)
        # 每日格：強制置中（避免模板有右對齊殘留）
        addr = f"{openpyxl.utils.get_column_letter(start_col + i)}{row}"
        apply_center_addr(ws, addr)


def detect_referenced_sheets(ws):
    """掃描目前工作表中所有公式，抓出像 'SomeSheet'!A1 這種引用，避免你刪掉必要分頁"""
    refs = set()
    pat1 = re.compile(r"'([^']+)'!")
    pat2 = re.compile(r"\b([A-Za-z0-9_ \-]+)!")
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                for m in pat1.findall(v):
                    refs.add(m)
                # 無引號的引用容易誤判，所以只加「確實是 workbook 的 sheet 名稱」才算
                for m in pat2.findall(v):
                    refs.add(m.strip())
    return refs


def keep_only_needed_sheets(wb, keep_names: set):
    """直接移除不需要的分頁（比 hidden 更乾淨，避免 PDF/輸出帶出別頁）"""
    for sh in list(wb.worksheets):
        if sh.title not in keep_names:
            wb.remove(sh)


SHEET_META = {
    "Dongwu": {
        "sheet_name": "東吳-格式",
        "date_start_cell": "I7",
        "schedule_start_col": "I",
        "max_days": 31,
        "total_col": "AN",
        "anchors": {"全家廣播": "通路廣播廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "header_cells": {"client": "C3", "product": "C4", "period": "C5", "medium": "C6", "month": "I6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "rate": "G", "pkg": "H"},
        "header_override": {"G7": "rate\n(List)", "H7": "Package-cost\n(List)"},
        "station_merge": True,
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        # 你反映 Day-part / Size / rate / Package-cost 沒置中：強制置中這幾欄
        "force_center_cols": ["E", "F", "G", "H"],
    },
    "Shenghuo": {
        "sheet_name": "聲活-格式",
        "date_start_cell": "G7",
        "schedule_start_col": "G",
        "max_days": 23,
        "total_col": "AD",
        "anchors": {"全家廣播": "廣播通路廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "header_cells": {"client": "C5", "product": "C6", "month": "G6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "proj_price": "AF"},
        "header_override": {"AF7": "專案價\n(List)"},
        "station_merge": False,
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": ["E", "F", "AF"],
    },
}


def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if meta["sheet_name"] not in wb.sheetnames:
        raise ValueError(f"缺少分頁：{meta['sheet_name']}")

    ws = wb[meta["sheet_name"]]

    # ====== 先填表頭 ======
    hc = meta["header_cells"]
    if "client" in hc:
        safe_set_value_addr(ws, hc["client"], client_name)
    if "product" in hc:
        safe_set_value_addr(ws, hc["product"], product_display_str)
    if "period" in hc:
        safe_set_value_addr(ws, hc["period"], f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y.%m. %d')}")
    if "medium" in hc:
        safe_set_value_addr(ws, hc["medium"], " ".join(sorted(set([r["media_type"] for r in rows]))))
    if "month" in hc:
        safe_set_value_addr(ws, hc["month"], f" {start_dt.month}月")

    safe_set_value_addr(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))

    for addr, text in meta.get("header_override", {}).items():
        safe_set_value_addr(ws, addr, text)
        apply_center_addr(ws, addr)

    # ====== 找 Total 行（用來決定插入列區間） ======
    total_cell = find_cell_exact(ws, meta["total_label"])
    if not total_cell:
        raise ValueError("找不到 Total")
    total_row = total_cell[0]

    cols = meta["cols"]

    # ====== 找各媒體區塊起始行 ======
    sec_start = {}
    for m_key, kw in meta["anchors"].items():
        r0 = find_first_row_contains(ws, cols["station"], kw)
        if r0:
            sec_start[m_key] = r0

    sec_order = sorted(sec_start.items(), key=lambda x: x[1])
    sec_ranges = []
    for i, (k, sr) in enumerate(sec_order):
        next_start = sec_order[i + 1][1] if i + 1 < len(sec_order) else total_row
        sec_ranges.append((k, sr, next_start - 1))

    # ====== 排序資料 ======
    reg_map = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}

    def sort_key(x):
        return (x["seconds"], reg_map.get(x["region"], 999))

    grouped = {
        "全家廣播": sorted([r for r in rows if r["media_type"] == "全家廣播"], key=sort_key),
        "新鮮視": sorted([r for r in rows if r["media_type"] == "新鮮視"], key=sort_key),
        "家樂福": sorted([r for r in rows if r["media_type"] == "家樂福"], key=sort_key),
    }

    # ====== 若資料行數大於模板區塊行數，插入列並複製樣式 ======
    for k, sr, er in sorted(sec_ranges, key=lambda x: x[1], reverse=True):
        data = grouped.get(k, [])
        needed = len(data)
        if needed <= 0:
            continue
        existing = er - sr + 1
        if needed > existing:
            ws.insert_rows(er + 1, amount=needed - existing)
            for rr in range(er + 1, er + 1 + (needed - existing)):
                copy_row_with_style_fix(ws, sr, rr, ws.max_column)

    # ====== 插入列後重新定位 Total/區塊範圍 ======
    total_row = find_cell_exact(ws, meta["total_label"])[0]
    sec_start = {}
    for m_key, kw in meta["anchors"].items():
        r0 = find_first_row_contains(ws, cols["station"], kw)
        if r0:
            sec_start[m_key] = r0
    sec_order = sorted(sec_start.items(), key=lambda x: x[1])
    sec_ranges = []
    for i, (k, sr) in enumerate(sec_order):
        next_start = sec_order[i + 1][1] if i + 1 < len(sec_order) else total_row
        sec_ranges.append((k, sr, next_start - 1))

    def station_title(m):
        prefix = "全家便利商店\n" if m != "家樂福" else ""
        name = "通路廣播廣告" if m == "全家廣播" else ("新鮮視廣告" if m == "新鮮視" else "家樂福")
        if format_type == "Shenghuo" and m == "全家廣播":
            name = "廣播通路廣告"
        return prefix + name

    # ====== 寫入資料 ======
    written_rows = []

    # 對齊強制置中欄（解你 Day-part/Size/rate/Package-cost 沒置中的問題）
    force_center_cols = meta.get("force_center_cols", [])

    def force_center_row(row_idx: int):
        for col_letter in force_center_cols:
            addr = f"{col_letter}{row_idx}"
            apply_center_addr(ws, addr)

    for m, sr, er in sec_ranges:
        data = grouped.get(m, [])
        if not data:
            continue

        # Station 欄合併（東吳格式）
        if meta["station_merge"]:
            unmerge_col_overlap(ws, cols["station"], sr, er)
            merge_rng = f"{cols['station']}{sr}:{cols['station']}{sr + len(data) - 1}"
            ws.merge_cells(merge_rng)
            top_addr = f"{cols['station']}{sr}"
            safe_set_value_addr(ws, top_addr, station_title(m))
            apply_center_addr(ws, top_addr)

        row_ptr = sr
        for r in data:
            if not meta["station_merge"]:
                addr = f"{cols['station']}{row_ptr}"
                safe_set_value_addr(ws, addr, station_title(m))
                apply_center_addr(ws, addr)

            # Location
            addr_loc = f"{cols['location']}{row_ptr}"
            safe_set_value_addr(ws, addr_loc, region_display(r["region"]))
            apply_center_addr(ws, addr_loc)

            # Program（店數/面數）
            prog_val = r.get("program_num", 0)
            addr_prog = f"{cols['program']}{row_ptr}"
            safe_set_value_addr(ws, addr_prog, int(prog_val))
            apply_center_addr(ws, addr_prog)

            # Day-part / Seconds / 價格欄位
            addr_day = f"{cols['daypart']}{row_ptr}"
            safe_set_value_addr(ws, addr_day, r["daypart"])
            apply_center_addr(ws, addr_day)

            if format_type == "Dongwu":
                addr_sec = f"{cols['seconds']}{row_ptr}"
                if m == "家樂福":
                    safe_set_value_addr(ws, addr_sec, f"{r['seconds']}秒")
                else:
                    safe_set_value_addr(ws, addr_sec, int(r["seconds"]))
                apply_center_addr(ws, addr_sec)

                addr_rate = f"{cols['rate']}{row_ptr}"
                safe_set_value_addr(ws, addr_rate, r["rate_list"])
                apply_center_addr(ws, addr_rate)

                addr_pkg = f"{cols['pkg']}{row_ptr}"
                safe_set_value_addr(ws, addr_pkg, r["pkg_display_val"])
                apply_center_addr(ws, addr_pkg)
            else:
                addr_sec = f"{cols['seconds']}{row_ptr}"
                safe_set_value_addr(ws, addr_sec, f"{r['seconds']}秒廣告")
                apply_center_addr(ws, addr_sec)

                addr_price = f"{cols['proj_price']}{row_ptr}"
                val = r["pkg_display_val"] if isinstance(r["pkg_display_val"], int) else 0
                safe_set_value_addr(ws, addr_price, val)
                apply_center_addr(ws, addr_price)

            # Schedule
            set_schedule(ws, row_ptr, meta["schedule_start_col"], meta["max_days"], r["schedule"])

            # Total（檔次）
            spot_sum = sum(r["schedule"][: meta["max_days"]])
            addr_total = f"{meta['total_col']}{row_ptr}"
            safe_set_value_addr(ws, addr_total, spot_sum)
            apply_center_addr(ws, addr_total)

            force_center_row(row_ptr)
            written_rows.append(row_ptr)
            row_ptr += 1

    # ====== Total Row（每日加總 + 總檔次 + 總金額） ======
    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [sum([x["schedule"][d] for x in rows if d < len(x["schedule"])]) for d in range(eff_days)]
    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_set_value_addr(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))
    apply_center_addr(ws, f"{meta['total_col']}{total_row}")

    total_pkg = sum([x["pkg_display_val"] for x in rows if isinstance(x["pkg_display_val"], int)])
    pkg_col = cols.get("pkg") or cols.get("proj_price")
    safe_set_value_addr(ws, f"{pkg_col}{total_row}", total_pkg)
    apply_center_addr(ws, f"{pkg_col}{total_row}")

    force_center_row(total_row)

    # ====== Footer（製作 / VAT / Grand Total） ======
    lbl = meta["footer_labels"]

    def write_footer(key, val):
        pos = find_cell_exact(ws, lbl.get(key, ""))
        if pos:
            safe_set_value(ws, pos[0], pos[1] + 1, int(val))
            apply_center_addr(ws, f"{openpyxl.utils.get_column_letter(pos[1]+1)}{pos[0]}")

    make_fee = 10000
    pos_make = find_cell_exact(ws, lbl["make"])
    if pos_make:
        v = ws.cell(pos_make[0], pos_make[1] + 1).value
        if isinstance(v, (int, float)) and v > 0:
            make_fee = int(v)
        else:
            safe_set_value(ws, pos_make[0], pos_make[1] + 1, make_fee)
            apply_center_addr(ws, f"{openpyxl.utils.get_column_letter(pos_make[1]+1)}{pos_make[0]}")

    vat = int(round((total_pkg + make_fee) * 0.05))
    write_footer("vat", vat)
    write_footer("grand", total_pkg + make_fee + vat)

    # ====== Remarks ======
    rem_pos = find_cell_exact(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_set_value(ws, rem_pos[0] + 1 + i, rem_pos[1], rm)

    # ====== 只保留必要分頁（避免你說的「多出不用分頁」） ======
    referenced = detect_referenced_sheets(ws)
    keep = {meta["sheet_name"]}
    # 只加「確實存在於 workbook 的分頁」避免誤判
    keep |= {name for name in referenced if name in wb.sheetnames}
    keep_only_needed_sheets(wb, keep)
    # active sheet 指定回主表
    wb.active = 0

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================================================
# 5) HTML Preview（更像 Excel：固定欄寬、table-layout fixed）
# =========================================================
def build_preview_css(format_type: str, for_pdf: bool, font_base64: str | None):
    # PDF 用 pt 比較穩，Preview 用 px 比較舒服
    if for_pdf:
        unit = "pt"
        border = "0.6pt"
        pad = "2.0pt"
        fs = "8.2pt"
        head_fs = "8.2pt"
        row_h = "16pt"
        day_w = "14pt"
        fixed_w_station = "90pt"
        fixed_w_loc = "78pt"
        fixed_w_prog = "52pt"
        fixed_w_daypart = "70pt"
        fixed_w_size = "36pt"
        fixed_w_rate = "66pt"
        fixed_w_pkg = "74pt"
        fixed_w_total = "40pt"
    else:
        unit = "px"
        border = "1px"
        pad = "4px"
        fs = "12px"
        head_fs = "12px"
        row_h = "24px"
        day_w = "22px"
        fixed_w_station = "140px"
        fixed_w_loc = "120px"
        fixed_w_prog = "72px"
        fixed_w_daypart = "96px"
        fixed_w_size = "52px"
        fixed_w_rate = "92px"
        fixed_w_pkg = "110px"
        fixed_w_total = "60px"

    head_bg = "#4472C4" if format_type == "Dongwu" else "#BDD7EE"
    head_fg = "#FFFFFF" if format_type == "Dongwu" else "#000000"
    weekend_bg = "#FFD966"
    total_bg = "#FFF2CC"

    font_face = ""
    font_family = "NotoSansTC, Arial, sans-serif"
    if font_base64:
        # 內嵌字型：雲端不會亂碼、字寬一致降低跑版
        font_face = f"""
        @font-face {{
            font-family: 'NotoSansTC';
            src: url(data:font/ttf;base64,{font_base64}) format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
        """
        font_family = "NotoSansTC, Arial, sans-serif"

    sticky = "" if for_pdf else "position: sticky; top: 0; z-index: 10;"

    return f"""
    {font_face}
    body {{
        margin: 0;
        font-family: {font_family};
        font-size: {fs};
        color: #000;
    }}
    .wrap {{
        background: #fff;
        padding: {('8pt' if for_pdf else '12px')};
    }}
    table {{
        border-collapse: collapse;
        width: 100%;
        table-layout: fixed;
    }}
    th, td {{
        border: {border} solid #000;
        padding: {pad};
        height: {row_h};
        vertical-align: middle;
        text-align: center;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: clip;
    }}
    thead th {{
        {sticky}
        font-size: {head_fs};
        font-weight: 700;
        background: {head_bg};
        color: {head_fg};
    }}
    .wknd {{
        background: {weekend_bg} !important;
        color: #000 !important;
    }}
    .total {{
        background: {total_bg};
        font-weight: 700;
    }}
    .left {{
        text-align: left !important;
        padding-left: {('4pt' if for_pdf else '6px')} !important;
        white-space: pre-line;
    }}
    .right {{
        text-align: center !important; /* 你要求數字也置中 */
        font-variant-numeric: tabular-nums;
    }}
    .remarks {{
        margin-top: {('10pt' if for_pdf else '14px')};
        font-size: {('8pt' if for_pdf else '12px')};
        line-height: 1.5;
    }}
    .remarks b {{
        text-decoration: underline;
    }}
    .meta {{
        margin-bottom: {('6pt' if for_pdf else '10px')};
        font-size: {('9pt' if for_pdf else '13px')};
    }}
    .meta .label {{
        font-weight: 700;
    }}
    .box {{
        border: 1px solid #bbb;
        box-shadow: {('none' if for_pdf else '0 2px 6px rgba(0,0,0,0.08)')};
    }}
    .scroll {{
        overflow-x: auto;
    }}
    .col_station {{ width: {fixed_w_station}; }}
    .col_loc     {{ width: {fixed_w_loc}; }}
    .col_prog    {{ width: {fixed_w_prog}; }}
    .col_daypart  {{ width: {fixed_w_daypart}; }}
    .col_size     {{ width: {fixed_w_size}; }}
    .col_rate     {{ width: {fixed_w_rate}; }}
    .col_pkg      {{ width: {fixed_w_pkg}; }}
    .col_day      {{ width: {day_w}; }}
    .col_total    {{ width: {fixed_w_total}; }}
    """


def generate_html(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks, for_pdf: bool, font_base64: str | None):
    eff_days = min(days_cnt, 31 if format_type == "Dongwu" else 23)
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]

    # 欄位定義
    if format_type == "Dongwu":
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(List)", "Package<br>(List)"]
        has_rate_pkg = True
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "專案價"]
        has_rate_pkg = False

    # 日期表頭
    date_th1 = []
    date_th2 = []
    curr = start_dt
    for i in range(eff_days):
        wd = curr.weekday()
        is_weekend = (format_type == "Dongwu" and wd >= 5)
        cls = "wknd" if is_weekend else ""
        date_th1.append(f"<th class='{cls} col_day'>{curr.day}</th>")
        date_th2.append(f"<th class='{cls} col_day'>{weekdays[wd]}</th>")
        curr += timedelta(days=1)

    # 排序與 group
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    rows_sorted = sorted(
        rows,
        key=lambda x: (
            media_order.get(x["media_type"], 99),
            x["seconds"],
            REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 99,
        ),
    )
    media_counts = {}
    for r in rows_sorted:
        media_counts[r["media_type"]] = media_counts.get(r["media_type"], 0) + 1
    media_printed = {m: False for m in media_counts}

    # colgroup（固定欄寬，讓 PDF/Preview 更像 Excel）
    if format_type == "Dongwu":
        colgroup = """
        <colgroup>
          <col class="col_station">
          <col class="col_loc">
          <col class="col_prog">
          <col class="col_daypart">
          <col class="col_size">
          <col class="col_rate">
          <col class="col_pkg">
        """
    else:
        colgroup = """
        <colgroup>
          <col class="col_station">
          <col class="col_loc">
          <col class="col_prog">
          <col class="col_daypart">
          <col class="col_size">
          <col class="col_pkg">
        """
    for _ in range(eff_days):
        colgroup += '<col class="col_day">'
    colgroup += '<col class="col_total"></colgroup>'

    # tbody
    tbody = []
    for r in rows_sorted:
        m = r["media_type"]
        tr = ["<tr>"]

        # Station 欄：Dongwu 合併（rowspan），Shenghuo 每列顯示
        display_name = (
            "全家便利商店<br>通路廣播廣告"
            if m == "全家廣播"
            else ("全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福")
        )
        if format_type == "Shenghuo" and m == "全家廣播":
            display_name = "全家便利商店<br>廣播通路廣告"

        if format_type == "Dongwu":
            if not media_printed[m]:
                tr.append(f"<td class='left' rowspan='{media_counts[m]}'>{display_name}</td>")
                media_printed[m] = True
        else:
            tr.append(f"<td class='left'>{display_name}</td>")

        # Location / Program / Day-part / Size
        tr.append(f"<td>{html_escape(region_display(r['region']))}</td>")
        tr.append(f"<td class='right'>{html_escape(r.get('program_num',''))}</td>")
        tr.append(f"<td>{html_escape(r['daypart'])}</td>")

        if format_type == "Dongwu":
            sec_txt = f"{r['seconds']}秒" if m == "家樂福" else f"{r['seconds']}"
        else:
            sec_txt = f"{r['seconds']}秒廣告"
        tr.append(f"<td>{html_escape(sec_txt)}</td>")

        # rate/pkg or proj price（你要數字置中，所以用 right class 但 CSS 設成 center）
        pkg = r["pkg_display_val"]
        pkg_disp = f"{pkg:,}" if isinstance(pkg, int) else html_escape(pkg)

        if has_rate_pkg:
            rate = r["rate_list"]
            rate_disp = f"{rate:,}" if isinstance(rate, int) else html_escape(rate)
            tr.append(f"<td class='right'>{rate_disp}</td>")
            tr.append(f"<td class='right'>{pkg_disp}</td>")
        else:
            tr.append(f"<td class='right'>{pkg_disp}</td>")

        # schedule
        for d in r["schedule"][:eff_days]:
            tr.append(f"<td>{d}</td>")

        # total
        tr.append(f"<td class='total'>{sum(r['schedule'][:eff_days])}</td>")
        tr.append("</tr>")
        tbody.append("".join(tr))

    # Total row
    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]
    total_spots = sum(totals)
    total_pkg = sum([r["pkg_display_val"] for r in rows if isinstance(r["pkg_display_val"], int)])

    if format_type == "Dongwu":
        fixed_cols = 5  # Station~Size
        t = [f"<tr class='total'><td colspan='{fixed_cols}' class='left'>Total</td>"]
        t.append("<td></td>")  # rate 欄留空（你原來就是這樣）
        t.append(f"<td class='right'>{total_pkg:,}</td>")
    else:
        fixed_cols = 5
        t = [f"<tr class='total'><td colspan='{fixed_cols}' class='left'>Total</td>"]
        t.append(f"<td class='right'>{total_pkg:,}</td>")

    for x in totals:
        t.append(f"<td>{x}</td>")
    t.append(f"<td class='total'>{total_spots}</td></tr>")
    tfoot = "".join(t)

    css = build_preview_css(format_type, for_pdf=for_pdf, font_base64=font_base64)
    meta = f"""
    <div class="meta">
      <span class="label">客戶：</span>{html_escape(c_name)}
      &nbsp;&nbsp;
      <span class="label">產品：</span>{html_escape(p_display)}
      <br>
      <span style="color:#555">走期：{html_escape(start_dt)} ~ {html_escape(end_dt)}</span>
    </div>
    """

    rem = "<div class='remarks'><b>Remarks：</b><br>" + "<br>".join([html_escape(x) for x in remarks]) + "</div>"

    outer_open = "<div class='wrap box'>" if for_pdf else "<div class='wrap box scroll'>"

    return f"""
    <html><head><meta charset="utf-8"><style>{css}</style></head>
    <body>
      {outer_open}
        {meta}
        <table>
          {colgroup}
          <thead>
            <tr>
              {''.join([f"<th rowspan='2'>{c}</th>" for c in cols_def])}
              {''.join(date_th1)}
              <th rowspan='2'>檔次</th>
            </tr>
            <tr>
              {''.join(date_th2)}
            </tr>
          </thead>
          <tbody>
            {''.join(tbody)}
            {tfoot}
          </tbody>
        </table>
        {rem}
      </div>
    </body></html>
    """


# =========================================================
# 6) PDF 產出：優先 soffice（Excel->PDF），否則 WeasyPrint
# =========================================================
def which_soffice():
    # Linux/Cloud: usually `soffice`
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p
    # Windows 常見安裝路徑（本機測試）
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def excel_bytes_to_pdf_via_soffice(xlsx_bytes: bytes) -> tuple[bytes | None, str]:
    soffice = which_soffice()
    if not soffice:
        return None, "此環境未偵測到 LibreOffice/soffice"

    with tempfile.TemporaryDirectory() as td:
        xlsx_path = os.path.join(td, "cue.xlsx")
        pdf_path = os.path.join(td, "cue.pdf")
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_bytes)

        # --headless convert
        cmd = [
            soffice,
            "--headless",
            "--nologo",
            "--nolockcheck",
            "--nodefault",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            td,
            xlsx_path,
        ]
        try:
            p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
            if p.returncode != 0:
                return None, f"LibreOffice 轉檔失敗：{p.stderr.decode('utf-8', errors='ignore')[:500]}"
            if not os.path.exists(pdf_path):
                # 有些版本輸出檔名會是 cue.pdf 或 cue.xlsx.pdf
                alt = os.path.join(td, "cue.xlsx.pdf")
                if os.path.exists(alt):
                    pdf_path = alt
                else:
                    files = os.listdir(td)
                    return None, f"找不到輸出 PDF，目錄內容：{files}"
            with open(pdf_path, "rb") as f:
                return f.read(), ""
        except Exception as e:
            return None, f"LibreOffice 執行錯誤：{e}"


def load_font_base64():
    """
    雲端 PDF 失真/亂碼關鍵：內嵌字型讓 WeasyPrint 不依賴系統字型。
    優先用 repo 同層的 NotoSansTC-Regular.ttf；沒有才嘗試下載。
    """
    font_path = "NotoSansTC-Regular.ttf"
    if os.path.exists(font_path):
        with open(font_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    # 下載（注意：Cloud 有時 outbound 會慢/失敗，所以我們 cache）
    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/TTF/TraditionalChinese/NotoSansTC-Regular.ttf"
    try:
        r = requests.get(url, timeout=30)
        if r.status_code == 200 and len(r.content) > 100000:
            with open(font_path, "wb") as f:
                f.write(r.content)
            return base64.b64encode(r.content).decode("utf-8")
    except Exception:
        pass
    return None


@st.cache_data(ttl=3600)
def cached_font_b64():
    return load_font_base64()


def html_to_pdf_weasyprint(html_str: str, font_b64: str | None) -> tuple[bytes | None, str]:
    try:
        from weasyprint import HTML, CSS
    except ImportError:
        return None, "WeasyPrint not installed"

    # PDF 專用 CSS（用 pt，並關閉 sticky 影響）
    css_pdf = build_preview_css("Dongwu", for_pdf=True, font_base64=font_b64)  # 先放 dummy；真正格式已 inline
    try:
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[CSS(string=css_pdf)])
        return pdf_bytes, ""
    except Exception as e:
        return None, f"WeasyPrint Render Error: {e}"


# =========================================================
# 7) UI
# =========================================================
st.title("📺 媒體 Cue 表生成器（v63.0｜Excel/PDF 擬真強化版）")

tpl_bytes, tpl_source, tpl_msgs = load_template_bytes()
if tpl_bytes:
    st.success(f"✅ 已載入系統公版（{tpl_source}）")
else:
    st.warning("⚠️ 未能自動載入公版，請手動上傳")
    up = st.file_uploader("上傳 Excel 模板", type=["xlsx"])
    if up:
        tpl_bytes = up.read()

if tpl_msgs:
    for m in tpl_msgs:
        st.info(m)

st.markdown("### 1) 選擇格式")
format_type = st.radio("", ["Dongwu", "Shenghuo"], horizontal=True, label_visibility="collapsed")

st.markdown("### 2) 基本資料設定")
c1, c2, c3 = st.columns(3)
with c1:
    client_name = st.text_input("客戶名稱", "萬國通路")
with c2:
    product_name = st.text_input("產品名稱", "統一布丁")
with c3:
    total_budget_input = st.number_input("總預算（未稅 Net）", value=1000000, step=10000)

d1, d2 = st.columns(2)
with d1:
    start_date = st.date_input("開始日", datetime(2026, 1, 1))
with d2:
    end_date = st.date_input("結束日", datetime(2026, 1, 31))

days_count = (end_date - start_date).days + 1
st.info(f"📅 走期共 **{days_count}** 天")

with st.expander("📝 備註欄位設定（Remarks）", expanded=False):
    rc1, rc2, rc3 = st.columns(3)
    sign_deadline = rc1.date_input("回簽截止日", datetime.now() + timedelta(days=3))
    billing_month = rc2.text_input("請款月份", "2026年2月")
    payment_date = rc3.date_input("付款兌現日", datetime(2026, 3, 31))

st.markdown("### 3) 媒體投放設定")
m1, m2, m3 = st.columns(3)
config = {}
rem_budget = 100

with m1:
    if st.checkbox("全家廣播", True):
        is_nat = st.checkbox("全省聯播", True, key="rad_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域（可複選）", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_reg")
        secs = st.multiselect("秒數", DURATIONS, [20], key="rad_sec")
        share = st.slider("預算 %", 0, 100, 60, key="rad_share")
        rem_budget -= share

        sec_shares = {}
        secs_sorted = sorted(secs)
        if len(secs_sorted) > 1:
            st.caption("秒數預算分配")
            left = 100
            for s in secs_sorted[:-1]:
                v = st.slider(f"{s}秒 %", 0, left, int(left / 2), key=f"rad_secshare_{s}")
                sec_shares[s] = v
                left -= v
            sec_shares[secs_sorted[-1]] = left
            st.write(f"🔹 {secs_sorted[-1]}秒: {left}%")
        elif secs_sorted:
            sec_shares[secs_sorted[0]] = 100

        config["全家廣播"] = {"is_national": is_nat, "regions": regs, "seconds": secs_sorted, "share": share, "sec_shares": sec_shares}

with m2:
    if st.checkbox("新鮮視", True):
        is_nat = st.checkbox("全省聯播", False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域（可複選）", REGIONS_ORDER, default=["北區"], key="fv_reg")
        secs = st.multiselect("秒數", DURATIONS, [10], key="fv_sec")
        share = st.slider("預算 %", 0, max(rem_budget, 0), min(20, max(rem_budget, 0)), key="fv_share")
        rem_budget -= share

        sec_shares = {}
        secs_sorted = sorted(secs)
        if len(secs_sorted) > 1:
            st.caption("秒數預算分配")
            left = 100
            for s in secs_sorted[:-1]:
                v = st.slider(f"{s}秒 %", 0, left, int(left / 2), key=f"fv_secshare_{s}")
                sec_shares[s] = v
                left -= v
            sec_shares[secs_sorted[-1]] = left
            st.write(f"🔹 {secs_sorted[-1]}秒: {left}%")
        elif secs_sorted:
            sec_shares[secs_sorted[0]] = 100

        config["新鮮視"] = {"is_national": is_nat, "regions": regs, "seconds": secs_sorted, "share": share, "sec_shares": sec_shares}

with m3:
    if st.checkbox("家樂福", True):
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.info(f"剩餘預算: {max(rem_budget, 0)}%")

        sec_shares = {}
        secs_sorted = sorted(secs)
        if len(secs_sorted) > 1:
            st.caption("秒數預算分配")
            left = 100
            for s in secs_sorted[:-1]:
                v = st.slider(f"{s}秒 %", 0, left, int(left / 2), key=f"cf_secshare_{s}")
                sec_shares[s] = v
                left -= v
            sec_shares[secs_sorted[-1]] = left
            st.write(f"🔹 {secs_sorted[-1]}秒: {left}%")
        elif secs_sorted:
            sec_shares[secs_sorted[0]] = 100

        config["家樂福"] = {"regions": ["全省"], "seconds": secs_sorted, "share": max(rem_budget, 0), "sec_shares": sec_shares}

rows = []
debug_logs = []

if config:
    for m, cfg in config.items():
        m_budget = total_budget_input * (cfg["share"] / 100.0)
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0:
                continue
            factor = get_sec_factor(m, sec)

            if m in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m]

                # 全省聯播：計算用 6 區，顯示也用 6 區（你說 2026 改 6 區 OK）
                if cfg["is_national"]:
                    calc_regs = REGIONS_ORDER
                    display_regs = REGIONS_ORDER
                else:
                    # 非全省：用使用者選的區域
                    # 若使用者沒選，避免除以 0
                    calc_regs = cfg["regions"] if cfg["regions"] else ["北區"]
                    display_regs = calc_regs

                unit_net_sum = 0
                for r in calc_regs:
                    unit_net_sum += (db[r][1] / db["Std_Spots"]) * factor
                if unit_net_sum == 0:
                    continue

                spots_init = math.ceil(s_budget / unit_net_sum)
                penalty = 1.1 if spots_init < db["Std_Spots"] else 1.0
                spots_final = math.ceil(s_budget / (unit_net_sum * penalty))
                if spots_final % 2 != 0:
                    spots_final += 1
                if spots_final == 0:
                    spots_final = 2

                sch = calculate_schedule(spots_final, days_count)

                debug_logs.append(
                    {
                        "media": m,
                        "sec": sec,
                        "budget": s_budget,
                        "unit_cost": unit_net_sum * penalty,
                        "spots": spots_final,
                        "std": db["Std_Spots"],
                        "status": "未達標" if penalty > 1 else "達標",
                        "reason": "懲罰 x1.1" if penalty > 1 else "費率正常",
                    }
                )

                for r in display_regs:
                    rate_list = int((db[r][0] / db["Std_Spots"]) * factor)
                    pkg_list = rate_list * spots_final
                    rows.append(
                        {
                            "media_type": m,
                            "region": r,
                            "program_num": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m == "新鮮視" else r, 0),
                            "daypart": db["Day_Part"],
                            "seconds": sec,
                            "spots": spots_final,
                            "schedule": sch,
                            "rate_list": rate_list,
                            "pkg_display_val": pkg_list,
                        }
                    )

            elif m == "家樂福":
                db = PRICING_DB["家樂福"]
                base_std = db["量販_全省"]["Std_Spots"]
                unit_net = (db["量販_全省"]["Net"] / base_std) * factor
                spots_init = math.ceil(s_budget / unit_net)
                penalty = 1.1 if spots_init < base_std else 1.0
                spots_final = math.ceil(s_budget / (unit_net * penalty))
                if spots_final % 2 != 0:
                    spots_final += 1

                sch_h = calculate_schedule(spots_final, days_count)

                debug_logs.append(
                    {
                        "media": m,
                        "sec": sec,
                        "budget": s_budget,
                        "unit_cost": unit_net * penalty,
                        "spots": spots_final,
                        "std": base_std,
                        "status": "未達標" if penalty > 1 else "達標",
                        "reason": "懲罰 x1.1" if penalty > 1 else "費率正常",
                    }
                )

                rate_h = int((db["量販_全省"]["List"] / base_std) * factor)
                rows.append(
                    {
                        "media_type": m,
                        "region": "全省量販",
                        "program_num": STORE_COUNTS_NUM["家樂福_量販"],
                        "daypart": db["量販_全省"]["Day_Part"],
                        "seconds": sec,
                        "spots": spots_final,
                        "schedule": sch_h,
                        "rate_list": rate_h,
                        "pkg_display_val": rate_h * spots_final,
                    }
                )

                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append(
                    {
                        "media_type": m,
                        "region": "全省超市",
                        "program_num": STORE_COUNTS_NUM["家樂福_超市"],
                        "daypart": db["超市_全省"]["Day_Part"],
                        "seconds": sec,
                        "spots": spots_s,
                        "schedule": sch_s,
                        "rate_list": "計量販",
                        "pkg_display_val": "計量販",
                    }
                )

p_str = ""
if rows:
    all_secs = sorted(list(set([r["seconds"] for r in rows])))
    p_str = f"{'、'.join([f'{s}秒' for s in all_secs])} {product_name}"

rem = get_remarks_text(sign_deadline, billing_month, payment_date)

with st.expander("💡 系統運算邏輯說明（Debug Panel）", expanded=False):
    for log in debug_logs:
        color = "green" if log["status"] == "達標" else "red"
        st.markdown(
            f"**{log['media']} ({log['sec']}秒)**：預算 ${log['budget']:,.0f} | 執行 {log['spots']} 檔 "
            f"-> <span style='color:{color}'><b>{log['status']}</b></span>（{log['reason']}）",
            unsafe_allow_html=True,
        )

st.divider()
st.markdown("### 4) 預覽與下載")

if not rows:
    st.warning("目前沒有任何投放資料（可能秒數未選 / 預算比例為 0）。")
    st.stop()

# Preview HTML（更擬真）
font_b64 = cached_font_b64()
html_preview = generate_html(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem, for_pdf=False, font_base64=font_b64)
st.components.v1.html(html_preview, height=720, scrolling=True)

if not tpl_bytes:
    st.warning("請上傳模板（或確認 Drive 可下載）以啟用 Excel 下載。")
    st.stop()

try:
    xlsx_bytes = generate_excel_from_template(format_type, start_date, end_date, client_name, p_str, rows, rem, tpl_bytes)
    out_name = safe_filename(f"Cue_{client_name}_{format_type}.xlsx")
    st.download_button("📥 下載 Excel（模板擬真）", xlsx_bytes, out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # PDF：優先 LibreOffice（最高擬真），否則 WeasyPrint（內嵌字型降低失真）
    prefer_soffice = st.checkbox("PDF 優先用 LibreOffice/soffice（最像 Excel）", value=True)
    pdf_bytes = None
    pdf_err = ""

    if prefer_soffice:
        pdf_bytes, pdf_err = excel_bytes_to_pdf_via_soffice(xlsx_bytes)

    if not pdf_bytes:
        # WeasyPrint fallback：用 for_pdf=True 的 HTML（關 sticky + pt unit）
        html_pdf = generate_html(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem, for_pdf=True, font_base64=font_b64)
        pdf_bytes, pdf_err = html_to_pdf_weasyprint(html_pdf, font_b64)

    if pdf_bytes:
        pdf_name = safe_filename(f"Cue_{client_name}_{format_type}.pdf")
        st.download_button("📄 下載 PDF", pdf_bytes, pdf_name, mime="application/pdf")
        st.caption("若你在 Streamlit Cloud 上仍覺得 PDF 與 Excel 有落差：請安裝 LibreOffice（下面有做法）。")
    else:
        st.error(f"PDF 產出失敗：{pdf_err}")

except Exception as e:
    st.error(f"產檔失敗：{e}")


# =========================================================
# 8) 部署提示
# =========================================================
with st.expander("🚀 Streamlit Cloud：如何讓 PDF 更像 Excel（安裝 LibreOffice）", expanded=False):
    st.markdown(
        """
在 Streamlit Community Cloud，要讓「PDF 幾乎 100% 跟 Excel 一樣」，建議使用 **LibreOffice/soffice 做 Excel→PDF**（上面程式已支援偵測並自動使用）。

**做法（放在你的 repo 根目錄）：**
1) 新增 `packages.txt`（安裝系統套件）  
   - `libreoffice`  
   - `fonts-noto-cjk`  
   - `fontconfig`

2) `requirements.txt`（Python 套件）至少要有：  
   - `streamlit`  
   - `openpyxl`  
   - `requests`  
   - （可選）`weasyprint`（當 soffice 沒裝時 fallback 用）

Streamlit 官方文件有說明 `packages.txt` 會用 apt-get 安裝外部依賴。  
（如果你加了 packages.txt 之後 build 卡住，通常重新部署一次就會好。）
"""
    )
