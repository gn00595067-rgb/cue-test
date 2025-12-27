import streamlit as st
import pandas as pd
import math
import io
import os
import shutil
import tempfile
import subprocess
import re
import requests
import base64
from datetime import timedelta, datetime, date
from copy import copy
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.formula.translate import Translator

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
# 1. 頁面設定
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v76.1")

# =========================================================
# 2. PDF 策略
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
    soffice = find_soffice_path()
    if not soffice: 
        return None, "Fail", "無可用的 LibreOffice 引擎"

    try:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, "cue.xlsx")
            with open(xlsx_path, "wb") as f: f.write(xlsx_bytes)
            
            subprocess.run([soffice, "--headless", "--nologo", "--convert-to", "pdf", "--outdir", tmp, xlsx_path], capture_output=True, timeout=60)
            
            pdf_path = os.path.join(tmp, "cue.pdf")
            if not os.path.exists(pdf_path):
                for fn in os.listdir(tmp):
                    if fn.endswith(".pdf"): pdf_path = os.path.join(tmp, fn); break
            
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f: return f.read(), "LibreOffice", ""
            return None, "Fail", "LibreOffice 轉檔無輸出"
    except Exception as e: return None, "Fail", str(e)

def html_to_pdf_weasyprint(html_str):
    try:
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        font_config = FontConfiguration()
        css = CSS(string="@page { size: A4 landscape; margin: 1cm; } body { font-family: sans-serif; }")
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[css], font_config=font_config)
        return pdf_bytes, ""
    except Exception as e: return None, str(e)

# =========================================================
# 3. 核心資料設定 (雲端 Google Sheet 版)
# =========================================================
GSHEET_SHARE_URL = "https://docs.google.com/spreadsheets/d/1bzmG-N8XFsj8m3LUPqA8K70AcIqaK4Qhq1VPWcK0w_s/edit?usp=sharing"

@st.cache_data(ttl=300)
def load_config_from_cloud(share_url):
    try:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", share_url)
        if not match: return None, None, None, None, "連結格式錯誤"
        file_id = match.group(1)
        
        def read_sheet(sheet_name):
            url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
            return pd.read_csv(url)

        # 1. Stores
        df_store = read_sheet("Stores")
        df_store.columns = [c.strip() for c in df_store.columns]
        store_counts = dict(zip(df_store['Key'], df_store['Display_Name']))
        store_counts_num = dict(zip(df_store['Key'], df_store['Count']))

        # 2. Factors
        df_fact = read_sheet("Factors")
        df_fact.columns = [c.strip() for c in df_fact.columns]
        sec_factors = {}
        for _, row in df_fact.iterrows():
            if row['Media'] not in sec_factors: sec_factors[row['Media']] = {}
            sec_factors[row['Media']][int(row['Seconds'])] = float(row['Factor'])

        # 3. Pricing
        df_price = read_sheet("Pricing")
        df_price.columns = [c.strip() for c in df_price.columns]
        pricing_db = {}
        for _, row in df_price.iterrows():
            m = row['Media']
            r = row['Region']
            if m == "家樂福":
                if m not in pricing_db: pricing_db[m] = {}
                pricing_db[m][r] = {
                    "List": int(row['List_Price']),
                    "Net": int(row['Net_Price']),
                    "Std_Spots": int(row['Std_Spots']),
                    "Day_Part": row['Day_Part']
                }
            else:
                if m not in pricing_db:
                    pricing_db[m] = {"Std_Spots": int(row['Std_Spots']), "Day_Part": row['Day_Part']}
                pricing_db[m][r] = [int(row['List_Price']), int(row['Net_Price'])]
            
        return store_counts, store_counts_num, pricing_db, sec_factors, None

    except Exception as e:
        return None, None, None, None, f"讀取失敗: {str(e)}"

with st.spinner("正在連線 Google Sheet 載入最新價格表..."):
    STORE_COUNTS, STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, err_msg = load_config_from_cloud(GSHEET_SHARE_URL)

if err_msg:
    st.error(f"❌ 設定檔載入失敗: {err_msg}")
    st.stop()

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

REGION_DISPLAY_MAP = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市"
}
def region_display(region): return REGION_DISPLAY_MAP.get(region, region)

def get_sec_factor(media_type, seconds): return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    if days <= 0: return []
    if total_spots % 2 != 0: total_spots += 1
    half_spots = total_spots // 2
    base, rem = divmod(half_spots, days)
    sch = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in sch]

def get_remarks_text(sign_deadline, billing_month, payment_date):
    d_str = sign_deadline.strftime("%Y/%m/%d (%a) %H:%M") if sign_deadline else "____/__/__ (__) 12:00"
    p_str = payment_date.strftime("%Y/%m/%d") if payment_date else "____/__/__"
    return [
        f"1.請於 {d_str}前 回簽及進單，方可順利上檔。",
        "2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇時段滿檔，上檔時間挪後或更換至同級時段。",
        "3.通路店鋪數與開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。",
        f"6.付款兌現日期：{p_str}"
    ]

# =========================================================
# 4. 核心計算函式 (Logic v4.4)
# =========================================================
def calculate_plan_data(config, total_budget, days_count):
    rows = []
    total_list_accum = 0
    debug_logs = []

    for m, cfg in config.items():
        m_budget_total = total_budget * (cfg["share"] / 100.0)
        
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget_total * (sec_pct / 100.0)
            if s_budget <= 0: continue
            
            factor = get_sec_factor(m, sec)
            
            if m in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m]
                calc_regs = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regs = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                unit_net_sum = 0
                for r in calc_regs:
                    unit_net_sum += (db[r][1] / db["Std_Spots"]) * factor
                if unit_net_sum == 0: continue
                
                spots_init = math.ceil(s_budget / unit_net_sum)
                is_under_target = spots_init < db["Std_Spots"]
                calc_penalty = 1.1 if is_under_target else 1.0 
                
                if cfg["is_national"]:
                    row_display_penalty = 1.0 
                    total_display_penalty = 1.1 if is_under_target else 1.0
                    status_msg = "全省(分區豁免/總價懲罰)" if is_under_target else "達標"
                else:
                    row_display_penalty = 1.1 if is_under_target else 1.0
                    total_display_penalty = 1.0 
                    status_msg = "未達標 x1.1" if is_under_target else "達標"

                spots_final = math.ceil(s_budget / (unit_net_sum * calc_penalty))
                if spots_final % 2 != 0: spots_final += 1
                if spots_final == 0: spots_final = 2
                
                debug_logs.append({
                    "Media": f"{m} ({sec}s)",
                    "Budget": f"${s_budget:,.0f}",
                    "Formula": f"Net總和(${sum([db[r][1] for r in calc_regs]):,}) / {db['Std_Spots']} * {factor}",
                    "Net_Unit": f"${unit_net_sum:.2f}",
                    "Penalty": f"x{calc_penalty} ({status_msg})",
                    "Final_Cost": f"${unit_net_sum * calc_penalty:.2f}",
                    "Final_Spots": spots_final
                })

                sch = calculate_schedule(spots_final, days_count)

                nat_pkg_display = 0
                if cfg["is_national"]:
                    nat_list = db["全省"][0]
                    nat_unit_price = int((nat_list / db["Std_Spots"]) * factor * total_display_penalty)
                    nat_pkg_display = nat_unit_price * spots_final
                    total_list_accum += nat_pkg_display

                for i, r in enumerate(display_regs):
                    list_price_region = db[r][0]
                    unit_rate_display = int((list_price_region / db["Std_Spots"]) * factor * row_display_penalty)
                    total_rate_display = unit_rate_display * spots_final 
                    row_pkg_display = total_rate_display
                    if not cfg["is_national"]:
                        total_list_accum += row_pkg_display

                    rows.append({
                        "media": m, "region": r,
                        "program_num": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec,
                        "spots": spots_final, "schedule": sch,
                        "rate_display": total_rate_display, 
                        "pkg_display": row_pkg_display,
                        "is_pkg_member": cfg["is_national"],
                        "nat_pkg_display": nat_pkg_display
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
                    "Media": f"家樂福 ({sec}s)",
                    "Budget": f"${s_budget:,.0f}",
                    "Formula": f"Net(${db['量販_全省']['Net']:,}) / {base_std} * {factor}",
                    "Net_Unit": f"${unit_net:.2f}",
                    "Penalty": f"x{penalty}",
                    "Final_Cost": f"${unit_net * penalty:.2f}",
                    "Final_Spots": spots_final
                })
                
                base_list = db["量販_全省"]["List"]
                unit_rate_h = int((base_list / base_std) * factor * penalty)
                total_rate_h = unit_rate_h * spots_final
                total_list_accum += total_rate_h
                
                rows.append({"media": m, "region": "全省量販", "program_num": STORE_COUNTS_NUM["家樂福_量販"], "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch_h, "rate_display": total_rate_h, "pkg_display": total_rate_h, "is_pkg_member": False})
                
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({"media": m, "region": "全省超市", "program_num": STORE_COUNTS_NUM["家樂福_超市"], "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "spots": spots_s, "schedule": sch_s, "rate_display": "計量販", "pkg_display": "計量販", "is_pkg_member": False})

    return rows, total_list_accum, debug_logs

# =========================================================
# 5. OpenPyXL 渲染引擎
# =========================================================
SHEET_META = {
    "Dongwu": {
        "sheet_name": "Sheet1", 
        "date_start_cell": "I7", "schedule_start_col": "I", "max_days": 31, "total_col": "AN",
        "anchors": {"全家廣播": "通路廣播廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "rate": "G", "pkg": "H"},
        "header_cells": {"client": "C3", "product": "C4", "period": "C5", "medium": "C6", "month": "I6"},
        "header_override": {"G7": "rate\n(Net)", "H7": "Package-cost\n(Net)"},
        "station_merge": True, "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": ["E", "F", "G", "H"], 
    },
    "Shenghuo": {
        "sheet_name": "Sheet1",
        "date_start_cell": "G7", "schedule_start_col": "G", "max_days": 23, "total_col": "AD",
        "anchors": {"全家廣播": "廣播通路廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "pkg": "AF"},
        "header_cells": {"client": "C5", "product": "C6", "month": "G6"},
        "station_merge": False, "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
        "force_center_cols": [],
    }
}

def safe_write_rc(ws, row, col, value, center=False):
    if isinstance(col, str): col = column_index_from_string(col)
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                cell = ws.cell(mr.min_row, mr.min_col)
                break
    cell.value = value
    if center:
        if cell.has_style:
            new_align = copy(cell.alignment)
            new_align.horizontal = 'center'
            new_align.vertical = 'center'
            new_align.wrap_text = True
            cell.alignment = new_align
        else:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def safe_write_addr(ws, addr, value):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                cell = ws.cell(mr.min_row, mr.min_col)
                break
    cell.value = value

def copy_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.protection = copy(source_cell.protection)

def find_row_by_content(ws, col_letter, keyword):
    col_idx = column_index_from_string(col_letter)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if isinstance(v, str) and keyword in v: return r
    return None

def copy_row_with_style_fix(ws, src_row, dst_row, max_col):
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for c in range(1, max_col + 1):
        sc = ws.cell(src_row, c)
        dc = ws.cell(dst_row, c)
        copy_style(sc, dc)

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

def find_first_row_contains(ws, col_letter, keyword):
    col_idx = column_index_from_string(col_letter)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if isinstance(v, str) and keyword in v: return r
    return None

def force_center_columns_range(ws, col_letters, start_row, end_row):
    if start_row is None or end_row is None: return
    for r in range(start_row, end_row + 1):
        for col in col_letters:
            safe_col = column_index_from_string(col)
            cell = ws.cell(r, safe_col)
            if isinstance(cell, MergedCell):
                master = _get_master_cell(ws, cell)
                if master: cell = master
                else: continue
            if cell.has_style:
                new_align = copy(cell.alignment)
                new_align.horizontal = 'center'
                new_align.vertical = 'center'
                cell.alignment = new_align

def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes, total_list_accum):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    target_sheet = wb.sheetnames[0] 
    ws = wb[target_sheet]

    hc = meta["header_cells"]
    if "client" in hc: safe_write_addr(ws, hc["client"], client_name)
    if "product" in hc: safe_write_addr(ws, hc["product"], product_display_str)
    if "period" in hc: safe_write_addr(ws, hc["period"], f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y.%m. %d')}")
    if "medium" in hc: safe_write_addr(ws, hc["medium"], " ".join(sorted(set([r["media"] for r in rows]))))
    if "month" in hc: safe_write_addr(ws, hc["month"], f" {start_dt.month}月")
    safe_write_addr(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))
    
    for addr, text in meta.get("header_override", {}).items(): 
        safe_write_addr(ws, addr, text)

    cols = meta["cols"]
    total_cell = find_row_by_content(ws, cols["station"], meta["total_label"])
    if not total_cell: return None
    total_row_orig = total_cell
    
    sec_start = {}
    for m_key, kw in meta["anchors"].items():
        r0 = find_first_row_contains(ws, cols["station"], kw)
        if r0: sec_start[m_key] = r0
    
    sec_order = sorted(sec_start.items(), key=lambda x: x[1], reverse=True)
    reg_map = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}
    def sort_key(x): return (x["seconds"], reg_map.get(x["region"], 999))
    
    grouped_data = {
        "全家廣播": sorted([r for r in rows if r["media"] == "全家廣播"], key=sort_key),
        "新鮮視": sorted([r for r in rows if r["media"] == "新鮮視"], key=sort_key),
        "家樂福": sorted([r for r in rows if r["media"] == "家樂福"], key=sort_key),
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
        
        if rows_to_delete > 0: ws.delete_rows(style_source_row + 1, amount=rows_to_delete)
        if needed > 1:
            ws.insert_rows(style_source_row + 1, amount=needed - 1)
            for r_idx in range(style_source_row + 1, style_source_row + 1 + needed - 1):
                copy_row_with_style_fix(ws, style_source_row, r_idx, ws.max_column)
        
        if needed == 0:
             for c in range(1, ws.max_column+1): safe_write_rc(ws, style_source_row, c, None)
             current_end_marker = start_row_orig - 1
             continue

        curr_row = style_source_row
        
        if meta["station_merge"]:
            unmerge_col_overlap(ws, cols["station"], curr_row, curr_row + needed - 1)
            merge_rng = f"{cols['station']}{curr_row}:{cols['station']}{curr_row + needed - 1}"
            ws.merge_cells(merge_rng)
            safe_write_rc(ws, curr_row, cols["station"], station_title(m_key), center=True)

        if needed > 0 and data[0].get("is_pkg_member", False):
            pkg_col = cols.get("pkg")
            if pkg_col:
                unmerge_col_overlap(ws, pkg_col, curr_row, curr_row + needed - 1)
                merge_pkg = f"{pkg_col}{curr_row}:{pkg_col}{curr_row + needed - 1}"
                ws.merge_cells(merge_pkg)
                safe_write_rc(ws, curr_row, pkg_col, data[0]["nat_pkg_display"], center=True)

        for idx, r_data in enumerate(data):
            if not meta["station_merge"]:
                safe_write_rc(ws, curr_row, cols["station"], station_title(m_key))
            
            safe_write_rc(ws, curr_row, cols["location"], region_display(r_data["region"]))
            prog_val = r_data.get("program_num", 0)
            safe_write_rc(ws, curr_row, cols["program"], int(prog_val))

            if format_type == "Dongwu":
                safe_write_rc(ws, curr_row, cols["daypart"], r_data["daypart"])
                if m_key == "家樂福": safe_write_rc(ws, curr_row, cols["seconds"], f"{r_data['seconds']}秒")
                else: safe_write_rc(ws, curr_row, cols["seconds"], int(r_data["seconds"]))
                
                safe_write_rc(ws, curr_row, cols["rate"], r_data["rate_display"])
                if not r_data.get("is_pkg_member", False):
                    safe_write_rc(ws, curr_row, cols["pkg"], r_data["pkg_display"])
            else:
                safe_write_rc(ws, curr_row, cols["daypart"], r_data["daypart"])
                safe_write_rc(ws, curr_row, cols["seconds"], f"{r_data['seconds']}秒廣告")
                if "pkg" in cols and not r_data.get("is_pkg_member", False):
                    safe_write_rc(ws, curr_row, cols["pkg"], r_data["pkg_display"])

            set_schedule(ws, curr_row, meta["schedule_start_col"], meta["max_days"], r_data["schedule"])
            spot_sum = sum(r_data["schedule"][:meta["max_days"]])
            safe_write_rc(ws, curr_row, meta["total_col"], spot_sum)
            curr_row += 1
            
        current_end_marker = start_row_orig - 1

    total_row = find_row_by_content(ws, meta["cols"]["station"], meta["total_label"])
    if total_row:
        eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
        daily_sums = [sum([x["schedule"][d] for x in rows if d < len(x["schedule"])]) for d in range(eff_days)]
        set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
        safe_write_rc(ws, total_row, meta["total_col"], sum(daily_sums))
        
        pkg_col = cols.get("pkg") or cols.get("proj_price")
        safe_write_rc(ws, total_row, pkg_col, total_list_accum)

        lbl = meta["footer_labels"]
        make_fee = 10000 
        pos_make = find_row_by_content(ws, "B", lbl["make"])
        if pos_make: safe_write_rc(ws, pos_make, pkg_col, make_fee)
        
        vat = int(round((total_list_accum + make_fee) * 0.05))
        pos_vat = find_row_by_content(ws, "B", lbl["vat"])
        if pos_vat: safe_write_rc(ws, pos_vat, pkg_col, vat)
        
        pos_grand = find_row_by_content(ws, "B", lbl["grand"])
        if pos_grand: safe_write_rc(ws, pos_grand, pkg_col, total_list_accum + make_fee + vat)

    rem_pos = find_row_by_content(ws, "B", "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            ws.cell(rem_pos + 1 + i, 2).value = rm

    if format_type == "Dongwu":
        force_center_columns_range(ws, meta["force_center_cols"], 9, total_row)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 6. HTML Preview (Fixed: Missing function added)
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

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks, total_list, grand_total, budget, prod):
    header_cls = "bg-dw-head" if format_type == "Dongwu" else "bg-sh-head"
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    eff_days = min(days_cnt, 31)
    
    font_b64 = load_font_base64()
    font_face = f"@font-face {{ font-family: 'NotoSansTC'; src: url(data:font/ttf;base64,{font_b64}) format('truetype'); }}" if font_b64 else ""

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
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(Net)", "Package-cost<br>(Net)"]
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "專案價<br>(Net)"]
    th_fixed = "".join([f"<th rowspan='2' class='{header_cls}'>{c}</th>" for c in cols_def])
    
    rows_sorted = sorted(rows, key=lambda x: (media_order.get(x["media"], 99), x["seconds"], REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 99))
    tbody = ""
    
    grouped_rows = {}
    for r in rows_sorted:
        key = (r['media'], r['seconds'])
        grouped_rows.setdefault(key, []).append(r)

    for (m, sec), group in grouped_rows.items():
        is_nat = group[0].get('is_pkg_member', False)
        group_size = len(group)
        
        for k, r_data in enumerate(group):
            tbody += "<tr>"
            
            if k == 0:
                display_name = "全家便利商店<br>通路廣播廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
                if format_type == "Shenghuo" and m == "全家廣播": display_name = "全家便利商店<br>廣播通路廣告"
                if format_type == "Shenghuo": tbody += f"<td class='left'>{display_name}</td>"
                else: tbody += f"<td class='left' rowspan='{group_size}'>{display_name}</td>"
            elif format_type == "Shenghuo":
                 display_name = "全家便利商店<br>廣播通路廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
                 tbody += f"<td class='left'>{display_name}</td>"

            loc_txt = region_display(r_data['region'])
            if "北北基" in loc_txt and "廣播" in r_data['media']: loc_txt = "北區-北北基+東"
            tbody += f"<td>{loc_txt}</td><td class='right'>{r_data.get('program_num','')}</td><td>{r_data['daypart']}</td>"
            sec_txt = f"{r_data['seconds']}秒" if format_type=="Dongwu" and m=="家樂福" else f"{r_data['seconds']}" if format_type=="Dongwu" else f"{r_data['seconds']}秒廣告"
            tbody += f"<td>{sec_txt}</td>"
            
            rate = f"{r_data['rate_display']:,}" if isinstance(r_data['rate_display'], int) else r_data['rate_display']
            pkg = f"{r_data['pkg_display']:,}" if isinstance(r_data['pkg_display'], int) else r_data['pkg_display']
            
            if format_type == "Dongwu": 
                tbody += f"<td class='right'>{rate}</td>"
                if is_nat:
                    if k == 0:
                        nat_pkg = f"{r_data['nat_pkg_display']:,}"
                        tbody += f"<td class='right' rowspan='{group_size}'>{nat_pkg}</td>"
                else:
                    tbody += f"<td class='right'>{pkg}</td>"
            else: 
                if is_nat:
                    if k == 0:
                        nat_pkg = f"{r_data['nat_pkg_display']:,}"
                        tbody += f"<td class='right' rowspan='{group_size}'>{nat_pkg}</td>"
                else:
                    tbody += f"<td class='right'>{pkg}</td>"
            
            for d in r_data['schedule'][:eff_days]: tbody += f"<td>{d}</td>"
            tbody += f"<td class='bg-total'>{r_data['spots']}</td></tr>"

    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]
    colspan = 5
    empty_td = "<td></td>" if format_type == "Dongwu" else ""
    tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='right'>Total (List Price)</td>{empty_td}<td class='right'>{total_list:,}</td>"
    for t in totals: tfoot += f"<td>{t}</td>"
    tfoot += f"<td>{sum(totals)}</td></tr>"

    vat = int(round((budget + prod) * 0.05))
    footer_rows = f"<tr><td colspan='6' class='right'>製作</td><td class='right'>{prod:,}</td><td colspan='{eff_days+1}'></td></tr>"
    footer_rows += f"<tr><td colspan='6' class='right'>專案優惠價 (Budget)</td><td class='right' style='color:red; font-weight:bold;'>{budget:,}</td><td colspan='{eff_days+1}'></td></tr>"
    footer_rows += f"<tr><td colspan='6' class='right'>5% VAT</td><td class='right'>{vat:,}</td><td colspan='{eff_days+1}'></td></tr>"
    footer_rows += f"<tr class='bg-grand'><td colspan='6' class='right'>Grand Total</td><td class='right'>{grand_total:,}</td><td colspan='{eff_days+1}'></td></tr>"

    html_content = f"""
    <html><head><style>
    {font_face}
    body {{ font-family: 'NotoSansTC', sans-serif !important; font-size: 10px; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ border: 0.5pt solid #000; padding: 2px; text-align: center; white-space: nowrap; }}
    .bg-dw-head {{ background-color: #4472C4; color: white; -webkit-print-color-adjust: exact; }}
    .bg-sh-head {{ background-color: #BDD7EE; color: black; -webkit-print-color-adjust: exact; }}
    .bg-weekend {{ background-color: #FFD966; -webkit-print-color-adjust: exact; }}
    .bg-total   {{ background-color: #E2EFDA; -webkit-print-color-adjust: exact; }}
    .bg-grand   {{ background-color: #FFC107; -webkit-print-color-adjust: exact; }}
    .left {{ text-align: left; }}
    .right {{ text-align: right; }}
    .remarks {{ margin-top: 10px; font-size: 9px; text-align: left; white-space: pre-wrap; }}
    </style></head><body>
    <div style="margin-bottom:10px;">
        <div style="font-size:16px; font-weight:bold; text-align:center;">Media Schedule</div>
        <b>客戶名稱：</b>{html_escape(c_name)} &nbsp; <b>Product：</b>{html_escape(p_display)}<br>
        <b>Period：</b>{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')} &nbsp; <b>Medium：</b>全家廣播/新鮮視/家樂福
    </div>
    <table>
        <thead><tr>{th_fixed}{date_th1}<th class='{header_cls}' rowspan='2'>檔次</th></tr><tr>{date_th2}</tr></thead>
        <tbody>{tbody}{tfoot}{footer_rows}</tbody>
    </table>
    <div class="remarks"><b>Remarks：</b><br>{"<br>".join([html_escape(x) for x in remarks])}</div>
    </body></html>
    """
    return html_content

# =========================================================
# 7. UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v76.1)")

st.markdown("### 1. 選擇格式")
c1, c2 = st.columns(2)
format_type = c1.radio("", ["Dongwu", "Shenghuo"], horizontal=True)

tpl_file = None
if format_type == "Dongwu":
    tpl_file = c2.file_uploader("上傳【東吳】樣板 (.xlsx)", type=["xlsx"], key="upl_dw")
else:
    tpl_file = c2.file_uploader("上傳【聲活】樣板 (.xlsx)", type=["xlsx"], key="upl_sh")

if tpl_file: template_bytes = tpl_file.read()
else: template_bytes = None

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

if "rad_share" not in st.session_state: st.session_state.rad_share = 100
if "fv_share" not in st.session_state: st.session_state.fv_share = 0
if "cf_share" not in st.session_state: st.session_state.cf_share = 0

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
        effective_is_nat = is_nat
        if not is_nat and len(regs) == 6:
            effective_is_nat = True
            regs = ["全省"]
            st.info("✅ 已選滿6區，自動轉為全省聯播計價")
        secs = st.multiselect("秒數", DURATIONS, [20], key="rad_sec")
        st.slider("預算 %", 0, 100, key="rad_share", on_change=on_slider_change, args=("rad_share",))
        sec_shares = {}
        if len(secs) > 1:
            st.caption("分配秒數佔比")
            rem = 100
            sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1:
                    v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"rs_{s}")
                    sec_shares[s] = v; rem -= v
                else:
                    sec_shares[s] = rem
                    st.markdown(f"🔹 **{s}秒**: {rem}% (自動計算)")
        elif secs: sec_shares[secs[0]] = 100
        config["全家廣播"] = {"is_national": effective_is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.rad_share}

if is_fv:
    with m2:
        st.markdown("#### 📺 新鮮視")
        is_nat = st.checkbox("全省聯播", False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
        effective_is_nat = is_nat
        if not is_nat and len(regs) == 6:
            effective_is_nat = True
            regs = ["全省"]
            st.info("✅ 已選滿6區，自動轉為全省聯播計價")
        secs = st.multiselect("秒數", DURATIONS, [10], key="fv_sec")
        st.slider("預算 %", 0, 100, key="fv_share", on_change=on_slider_change, args=("fv_share",))
        sec_shares = {}
        if len(secs) > 1:
            st.caption("分配秒數佔比")
            rem = 100
            sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1:
                    v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"fs_{s}")
                    sec_shares[s] = v; rem -= v
                else:
                    sec_shares[s] = rem
                    st.markdown(f"🔹 **{s}秒**: {rem}% (自動計算)")
        elif secs: sec_shares[secs[0]] = 100
        config["新鮮視"] = {"is_national": effective_is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.fv_share}

if is_cf:
    with m3:
        st.markdown("#### 🛒 家樂福")
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.slider("預算 %", 0, 100, key="cf_share", on_change=on_slider_change, args=("cf_share",))
        sec_shares = {}
        if len(secs) > 1:
            st.caption("分配秒數佔比")
            rem = 100
            sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1:
                    v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"cs_{s}")
                    sec_shares[s] = v; rem -= v
                else:
                    sec_shares[s] = rem
                    st.markdown(f"🔹 **{s}秒**: {rem}% (自動計算)")
        elif secs: sec_shares[secs[0]] = 100
        config["家樂福"] = {"regions": ["全省"], "sec_shares": sec_shares, "share": st.session_state.cf_share}

if config:
    rows, total_list_accum, logs = calculate_plan_data(config, total_budget_input, days_count)
    
    prod_cost = 10000
    vat = int(round((total_budget_input + prod_cost) * 0.05))
    grand_total = total_budget_input + prod_cost + vat
    p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}"
    rem = get_remarks_text(sign_deadline, billing_month, payment_date)

    html_preview = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem, total_list_accum, grand_total, total_budget_input, prod_cost)
    st.components.v1.html(html_preview, height=700, scrolling=True)

    with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
        for log in logs:
            st.markdown(f"**{log.get('Media', 'N/A')}** - 公式: {log.get('Formula', 'N/A')}")
            st.divider()

    if template_bytes and rows:
        try:
            xlsx = generate_excel_from_template(format_type, start_date, end_date, client_name, p_str, rows, rem, template_bytes, total_list_accum)
            if xlsx:
                st.download_button("📥 下載擬真 Excel", xlsx, f"Cue_{safe_filename(client_name)}.xlsx")
                
                pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx)
                if pdf_bytes:
                    st.download_button(f"📥 下載擬真 PDF ({method})", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                else:
                    st.warning(f"本地轉檔失敗 ({err})，使用網頁渲染版")
                    pdf_bytes, err = html_to_pdf_weasyprint(html_preview)
                    if pdf_bytes: st.download_button("📥 下載 PDF (Web版)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
        except Exception as e:
            st.error(f"Excel 產出錯誤: {e}")
