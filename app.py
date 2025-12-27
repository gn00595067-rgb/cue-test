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
# 1. 頁面設定 & 自動載入
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v72.1")

GOOGLE_DRIVE_FILE_ID = "11R1SA_hpFD5O_MGmYeh4BdtcUhK2bPta"
DEFAULT_FILENAME = "1209-Cue表相關資料.xlsx"

@st.cache_resource(ttl=600)
def load_default_template():
    status_msg = []
    if GOOGLE_DRIVE_FILE_ID:
        url = f"https://drive.google.com/uc?export=download&id={GOOGLE_DRIVE_FILE_ID}"
        try:
            r = requests.get(url, timeout=20, allow_redirects=True)
            if r.status_code == 200 and b"<!DOCTYPE html>" not in r.content[:500]:
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
# 2. PDF 雙重策略 (優先 LibreOffice，備援 WeasyPrint)
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
    # 嘗試使用 LibreOffice (最擬真)
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
    # 備援方案：HTML 轉 PDF (保證雲端可用)
    try:
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        font_config = FontConfiguration()
        # 簡單 CSS
        css = CSS(string="@page { size: A4 landscape; margin: 1cm; } body { font-family: sans-serif; }")
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[css], font_config=font_config)
        return pdf_bytes, ""
    except Exception as e: return None, str(e)

# =========================================================
# 3. 核心資料與邏輯設定
# =========================================================
STORE_COUNTS = {
    "全省": "4,437店", "北區": "1,649店", "桃竹苗": "779店", "中區": "839店", 
    "雲嘉南": "499店", "高屏": "490店", "東區": "181店",
    "新鮮視_全省": "3,124面", "新鮮視_北區": "1,127面", "新鮮視_桃竹苗": "616面", 
    "新鮮視_中區": "528面", "新鮮視_雲嘉南": "365面", "新鮮視_高屏": "405面", "新鮮視_東區": "83面",
    "家樂福_量販": "68店", "家樂福_超市": "249店"
}
STORE_COUNTS_NUM = {k: parse_count_to_int(v) for k, v in STORE_COUNTS.items()}

# [List, Net]
PRICING_DB = {
    "全家廣播": { "Std_Spots": 480, "Day_Part": "00:00-24:00", "全省": [400000, 320000], "北區": [250000, 200000], "桃竹苗": [150000, 120000], "中區": [150000, 120000], "雲嘉南": [100000, 80000], "高屏": [100000, 80000], "東區": [62500, 50000] },
    "新鮮視": { "Std_Spots": 504, "Day_Part": "07:00-22:00", "全省": [150000, 120000], "北區": [150000, 120000], "桃竹苗": [120000, 96000], "中區": [90000, 72000], "雲嘉南": [75000, 60000], "高屏": [75000, 60000], "東區": [45000, 36000] },
    "家樂福": { "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420, "Day_Part": "09:00-23:00"}, "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720, "Day_Part": "00:00-24:00"} }
}
SEC_FACTORS = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5, 5: 0.25},
    "新鮮視": {30: 3.0, 20: 2.0, 15: 1.5, 10: 1.0, 5: 0.5},
    "家樂福": {30: 1.5, 20: 1.0, 15: 0.85, 10: 0.65, 5: 0.35}
}
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]

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

# [補回] 備註生成函式
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
# 4. 商業邏輯運算 (Net Calc / List Display)
# =========================================================
def calculate_plan_data(config, total_budget, days_count):
    rows = []
    total_list_accum = 0
    debug_logs = []

    for m, cfg in config.items():
        m_budget = total_budget * (cfg["share"] / 100.0)
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0: continue
            
            factor = get_sec_factor(m, sec)
            
            if m in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m]
                calc_regs = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regs = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                # Net 算檔次
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
                debug_logs.append(f"{m} {sec}s: Budget {s_budget} -> {spots_final} spots")

                # List 顯示
                for i, r in enumerate(display_regs):
                    list_price_region = db[r][0]
                    rate_display = int((list_price_region / db["Std_Spots"]) * factor)
                    pkg_display = rate_display * spots_final
                    
                    if cfg["is_national"]:
                        if i == 0:
                            nat_list = db["全省"][0]
                            nat_rate = int((nat_list / db["Std_Spots"]) * factor)
                            total_list_accum += nat_rate * spots_final
                    else:
                        total_list_accum += pkg_display

                    rows.append({
                        "media": m, "region": r,
                        "program_num": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec,
                        "spots": spots_final, "schedule": sch,
                        "rate_display": rate_display, "pkg_display": pkg_display,
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
                
                rate_h = int((db["量販_全省"]["List"] / base_std) * factor)
                pkg_h = rate_h * spots_final
                total_list_accum += pkg_h
                rows.append({"media": m, "region": "全省量販", "program_num": STORE_COUNTS_NUM["家樂福_量販"], "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch_h, "rate_display": rate_h, "pkg_display": pkg_h})
                
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({"media": m, "region": "全省超市", "program_num": STORE_COUNTS_NUM["家樂福_超市"], "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "spots": spots_s, "schedule": sch_s, "rate_display": "計量販", "pkg_display": "計量販"})

    return rows, total_list_accum, debug_logs

# =========================================================
# 5. OpenPyXL 渲染引擎 (Template Filling)
# =========================================================
SHEET_META = {
    "Dongwu": {
        "sheet_name": "東吳-格式", "date_start_cell": "I7", "schedule_start_col": "I", "max_days": 31, "total_col": "AN",
        "anchors": {"全家廣播": "通路廣播廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "rate": "G", "pkg": "H"},
        "header_cells": {"client": "C3", "product": "C4", "period": "C5", "medium": "C6", "month": "I6"},
        "station_merge": True, "total_label": "Total",
    },
    "Shenghuo": {
        "sheet_name": "聲活-格式", "date_start_cell": "G7", "schedule_start_col": "G", "max_days": 23, "total_col": "AD",
        "anchors": {"全家廣播": "廣播通路廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "pkg": "AF"},
        "header_cells": {"client": "C5", "product": "C6", "month": "G6"},
        "station_merge": False, "total_label": "Total",
    }
}

def find_row_by_content(ws, col_letter, keyword):
    col_idx = column_index_from_string(col_letter)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if isinstance(v, str) and keyword in v: return r
    return None

def copy_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.protection = copy(source_cell.protection)

def safe_write(ws, row, col_letter, value, center=False):
    col_idx = column_index_from_string(col_letter)
    cell = ws.cell(row, col_idx)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col_idx <= mr.max_col:
                cell = ws.cell(mr.min_row, mr.min_col)
                break
    cell.value = value
    if center:
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def generate_excel(format_type, start_dt, end_dt, client_name, p_display, rows, remarks, template_bytes, total_list):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    if meta["sheet_name"] not in wb.sheetnames: return None
    for s in list(wb.sheetnames):
        if s != meta["sheet_name"]: del wb[s]
    ws = wb[meta["sheet_name"]]

    # Header
    hc = meta["header_cells"]
    if "client" in hc: ws[hc["client"]] = client_name
    if "product" in hc: ws[hc["product"]] = p_display
    if "period" in hc: ws[hc["period"]] = f"{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}"
    if "medium" in hc and rows: ws[hc["medium"]] = " ".join(sorted(list(set(r["media"] for r in rows))))
    if "month" in hc: ws[hc["month"]] = f" {start_dt.month}月"
    ws[meta["date_start_cell"]] = start_dt

    # Blocks
    cols = meta["cols"]
    anchors = {}
    for m, kw in meta["anchors"].items():
        r = find_row_by_content(ws, cols["station"], kw)
        if r: anchors[m] = r
    
    sorted_media = sorted(anchors.keys(), key=lambda k: anchors[k], reverse=True)
    reg_map = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}
    grouped_rows = {}
    for r in rows: grouped_rows.setdefault(r["media"], []).append(r)
    for m in grouped_rows: grouped_rows[m].sort(key=lambda x: (x["seconds"], reg_map.get(x["region"], 999)))

    current_footer_boundary = find_row_by_content(ws, cols["station"], meta["total_label"])
    if not current_footer_boundary: current_footer_boundary = ws.max_row

    for m in sorted_media:
        start_row = anchors[m]
        data = grouped_rows.get(m, [])
        count = len(data)
        
        if count == 0:
            ws[f"{cols['station']}{start_row}"] = ""
            continue

        if count > 1:
            ws.insert_rows(start_row + 1, amount=count - 1)
            for i in range(1, count):
                new_r = start_row + i
                for c in range(1, ws.max_column + 1):
                    copy_style(ws.cell(start_row, c), ws.cell(new_r, c))
        
        if meta["station_merge"]:
            merge_range = f"{cols['station']}{start_row}:{cols['station']}{start_row + count - 1}"
            ws.merge_cells(merge_range)
            display_name = f"全家便利商店\n{meta['anchors'][m]}" if m != "家樂福" else "家樂福"
            safe_write(ws, start_row, cols["station"], display_name, center=True)

        for i, row_data in enumerate(data):
            curr_r = start_row + i
            if not meta["station_merge"]:
                safe_write(ws, curr_r, cols["station"], meta['anchors'][m])
            
            safe_write(ws, curr_r, cols["location"], region_display(row_data["region"]))
            safe_write(ws, curr_r, cols["program"], row_data["program_num"])
            safe_write(ws, curr_r, cols["daypart"], row_data["daypart"])
            
            sec_txt = f"{row_data['seconds']}秒" if m=="家樂福" and format_type=="Dongwu" else row_data["seconds"]
            if format_type == "Shenghuo": sec_txt = f"{row_data['seconds']}秒廣告"
            safe_write(ws, curr_r, cols["seconds"], sec_txt)
            
            if "rate" in cols: safe_write(ws, curr_r, cols["rate"], row_data["rate_display"])
            safe_write(ws, curr_r, cols["pkg"], row_data["pkg_display"])
            
            sch_start = column_index_from_string(meta["schedule_start_col"])
            for d_idx, val in enumerate(row_data["schedule"]):
                if d_idx < meta["max_days"]: ws.cell(curr_r, sch_start + d_idx).value = val
            
            safe_write(ws, curr_r, meta["total_col"], row_data["spots"])

    # Total & Footer
    total_row = find_row_by_content(ws, cols["station"], meta["total_label"])
    if total_row:
        safe_write(ws, total_row, cols.get("pkg", "H"), total_list)
        
        daily_sums = [0] * meta["max_days"]
        for r in rows:
            for d, v in enumerate(r["schedule"]):
                if d < meta["max_days"]: daily_sums[d] += v
        
        sch_start = column_index_from_string(meta["schedule_start_col"])
        for d_idx, val in enumerate(daily_sums):
            ws.cell(total_row, sch_start + d_idx).value = val
        safe_write(ws, total_row, meta["total_col"], sum(daily_sums))

        make_fee = 10000
        make_row = find_row_by_content(ws, cols["rate"] if "rate" in cols else "AF", "製作")
        if make_row: ws.cell(make_row, column_index_from_string(cols.get("pkg", "H"))).value = make_fee
        
        vat_row = find_row_by_content(ws, cols["rate"] if "rate" in cols else "AF", "5% VAT")
        if vat_row: ws.cell(vat_row, column_index_from_string(cols.get("pkg", "H"))).value = int(round((total_list + make_fee) * 0.05))
            
        grand_row = find_row_by_content(ws, cols["rate"] if "rate" in cols else "AF", "Grand Total")
        if grand_row: ws.cell(grand_row, column_index_from_string(cols.get("pkg", "H"))).value = total_list + make_fee + int(round((total_list + make_fee) * 0.05))

    rem_row = find_row_by_content(ws, "B", "Remarks：")
    if rem_row:
        for i, r_text in enumerate(remarks):
            ws.cell(rem_row + 1 + i, 2).value = r_text

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks, total_list):
    font_b64 = None
    font_path = "NotoSansTC-Regular.ttf"
    if os.path.exists(font_path):
        with open(font_path, "rb") as f: font_b64 = base64.b64encode(f.read()).decode("utf-8")
    
    font_face = f"@font-face {{ font-family: 'NotoSansTC'; src: url(data:font/ttf;base64,{font_b64}) format('truetype'); }}" if font_b64 else ""
    
    html = f"""
    <html><head><style>
    {font_face}
    body {{ font-family: 'NotoSansTC', sans-serif; font-size: 10px; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ border: 1px solid #000; padding: 4px; text-align: center; }}
    .head {{ background-color: #4472C4; color: white; }}
    </style></head><body>
    <h3>Media Schedule</h3>
    <p>客戶: {c_name} | 產品: {p_display} | 期間: {start_dt} ~ {end_dt}</p>
    <table>
    <tr class="head"><th>媒體</th><th>區域</th><th>店數</th><th>秒數</th><th>單價(List)</th><th>總價(List)</th></tr>
    """
    for r in rows:
        html += f"<tr><td>{r['media']}</td><td>{r['region']}</td><td>{r['program_num']}</td><td>{r['seconds']}</td><td>{r['rate_display']:,}</td><td>{r['pkg_display']:,}</td></tr>"
    html += f"<tr><td colspan='5' style='text-align:right'><b>Total (List Price)</b></td><td><b>{total_list:,}</b></td></tr>"
    html += f"</table><br><div style='white-space: pre-wrap;'>{'<br>'.join(remarks)}</div></body></html>"
    return html

# =========================================================
# 6. UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v72.1)")

auto_tpl, source, msgs = load_default_template()
template_bytes = auto_tpl

if auto_tpl: st.success(f"✅ 已載入系統公版 ({source})")
else:
    tpl = st.file_uploader("上傳 Excel 模板", type=["xlsx"])
    if tpl: template_bytes = tpl.read()

format_type = st.radio("", ["Dongwu", "Shenghuo"], horizontal=True, label_visibility="collapsed")
c1, c2, c3 = st.columns(3)
with c1: client_name = st.text_input("客戶名稱", "萬國通路")
with c2: product_name = st.text_input("產品名稱", "統一布丁")
with c3: total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)
c4, c5 = st.columns(2)
with c4: start_date = st.date_input("開始日", datetime(2026, 1, 1))
with c5: end_date = st.date_input("結束日", datetime(2026, 1, 31))
days_count = (end_date - start_date).days + 1

with st.expander("📝 備註設定"):
    rc1, rc2, rc3 = st.columns(3)
    sign_deadline = rc1.date_input("回簽截止", datetime.now() + timedelta(days=3))
    billing_month = rc2.text_input("請款月", "2026年2月")
    payment_date = rc3.date_input("付款日", datetime(2026, 3, 31))

st.write("媒體設定：")
col_cb1, col_cb2, col_cb3 = st.columns(3)
with col_cb1: is_rad = st.checkbox("全家廣播", value=True)
with col_cb2: is_fv = st.checkbox("新鮮視", value=False)
with col_cb3: is_cf = st.checkbox("家樂福", value=False)

if "rad_share" not in st.session_state: st.session_state.rad_share = 100
if "fv_share" not in st.session_state: st.session_state.fv_share = 0
if "cf_share" not in st.session_state: st.session_state.cf_share = 0

config = {}
if is_rad:
    st.markdown("#### 📻 全家廣播")
    is_nat = st.checkbox("全省聯播", True, key="rad_nat")
    regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_reg")
    secs = st.multiselect("秒數", [5,10,15,20,30], [20], key="rad_sec")
    share = st.slider("預算 %", 0, 100, 100, key="rad_share")
    if secs: config["全家廣播"] = {"is_national": is_nat, "regions": regs, "sec_shares": {secs[0]: 100}, "share": share}

if is_fv:
    st.markdown("#### 📺 新鮮視")
    is_nat = st.checkbox("全省聯播", False, key="fv_nat")
    regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
    secs = st.multiselect("秒數", [5,10,15,20,30], [10], key="fv_sec")
    share = st.slider("預算 %", 0, 100, 0, key="fv_share")
    if secs: config["新鮮視"] = {"is_national": is_nat, "regions": regs, "sec_shares": {secs[0]: 100}, "share": share}

if is_cf:
    st.markdown("#### 🛒 家樂福")
    secs = st.multiselect("秒數", [5,10,15,20,30], [20], key="cf_sec")
    share = st.slider("預算 %", 0, 100, 0, key="cf_share")
    if secs: config["家樂福"] = {"regions": ["全省"], "sec_shares": {secs[0]: 100}, "share": share}

if config:
    rows, total_list, logs = calculate_plan_data(config, total_budget_input, days_count)
    rem_txt = get_remarks_text(sign_deadline, billing_month, payment_date)
    prod_cost = 10000
    vat = int((total_list + prod_cost) * 0.05)
    grand_total = total_list + prod_cost + vat
    
    st.dataframe(pd.DataFrame(rows)[['media', 'region', 'spots', 'rate_display', 'pkg_display']])
    
    html_preview = generate_html_preview(rows, days_count, start_date, end_date, client_name, product_name, format_type, rem_txt, total_list)
    
    if template_bytes and rows:
        try:
            xlsx = generate_excel(format_type, start_date, end_date, client_name, product_name, rows, rem_txt, template_bytes, total_list)
            if xlsx:
                st.download_button("下載 Excel", xlsx, f"Cue_{safe_filename(client_name)}.xlsx")
                
                pdf_bytes, _, err = xlsx_bytes_to_pdf_bytes(xlsx)
                if pdf_bytes:
                    st.download_button(f"下載 PDF (LibreOffice)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                else:
                    st.warning(f"本地轉檔失敗，切換至 WeasyPrint: {err}")
                    pdf_bytes, err = html_to_pdf_weasyprint(html_preview)
                    if pdf_bytes:
                        st.download_button("下載 PDF (HTML Render)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf")
                    else:
                        st.error(f"PDF Error: {err}")

        except Exception as e:
            st.error(f"Error: {e}")
