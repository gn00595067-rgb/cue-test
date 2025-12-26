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
from datetime import timedelta, datetime, date
import xlsxwriter

# ==============================================================================
# 🛠️ [通用工具模組] (Utilities)
# ==============================================================================
def parse_count_to_int(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return int(x)
    s = str(x)
    m = re.findall(r"[\d,]+", s)
    if not m: return 0
    return int(m[0].replace(",", ""))

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def find_soffice_path():
    """尋找 LibreOffice 執行檔路徑"""
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

# ==============================================================================
# 🅰️ [模組 A：資料庫與設定] (Database & Config)
# ==============================================================================
# 系統設定
GOOGLE_DRIVE_FILE_ID = "11R1SA_hpFD5O_MGmYeh4BdtcUhK2bPta"
DEFAULT_FILENAME = "1209-Cue表相關資料.xlsx"

# 載入公版檔案
@st.cache_resource(ttl=600)
def load_default_template():
    # 1. 優先嘗試 Google Drive
    if GOOGLE_DRIVE_FILE_ID:
        url = f"https://drive.google.com/uc?export=download&id={GOOGLE_DRIVE_FILE_ID}"
        try:
            r = requests.get(url, timeout=20, allow_redirects=True)
            if r.status_code == 200 and b"<!DOCTYPE html>" not in r.content[:500]:
                return io.BytesIO(r.content), "雲端硬碟 (Google Drive)"
        except: pass

    # 2. 其次嘗試本地檔案
    if os.path.exists(DEFAULT_FILENAME):
        try:
            with open(DEFAULT_FILENAME, "rb") as f:
                return io.BytesIO(f.read()), "系統主機 (Local)"
        except: pass
    
    return None, None

# 資料庫 (2026 新制)
STORE_COUNTS = {
    "全省": "4,437店", "北區": "1,649店", "桃竹苗": "779店", "中區": "839店", "雲嘉南": "499店", "高屏": "490店", "東區": "181店",
    "新鮮視_全省": "3,124面", "新鮮視_北區": "1,127面", "新鮮視_桃竹苗": "616面", "新鮮視_中區": "528面",
    "新鮮視_雲嘉南": "365面", "新鮮視_高屏": "405面", "新鮮視_東區": "83面",
    "家樂福_量販": "68店",  # [Fix]
    "家樂福_超市": "249店" # [Fix]
}
STORE_COUNTS_NUM = {k: parse_count_to_int(v) for k, v in STORE_COUNTS.items()}
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# [List Price (牌價/顯示用), Net Price (實收/計算用)]
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

# ==============================================================================
# 🅱️ [模組 B：核心計算引擎] (Calculation Logic)
# ==============================================================================
def get_sec_factor(media_type, seconds): return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    """偶數排程演算法"""
    if days <= 0: return []
    if total_spots % 2 != 0: total_spots += 1
    half_spots = total_spots // 2
    base, rem = divmod(half_spots, days)
    half_schedule = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in half_schedule]

def calculate_plan_data(config, total_budget, days_count):
    """計算所有媒體的投放數據"""
    rows = []
    total_list_price_accum = 0
    debug_logs = []

    for m, cfg in config.items():
        m_budget = total_budget * (cfg["share"] / 100.0)
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget * (sec_pct / 100.0)
            if s_budget <= 0: continue
            factor = get_sec_factor(m, sec)
            
            if m in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m]
                
                # 計算用 (Net) vs 顯示用 (List) 區域設定
                if cfg["is_national"]:
                    calc_regs = ["全省"]        # 計算用全省 Net
                    display_regs = REGIONS_ORDER # 顯示展開 6 區
                else:
                    calc_regs = cfg["regions"]
                    display_regs = cfg["regions"]
                
                # 1. 實收逆推 (Net)
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
                
                # 2. 準備顯示資料 (List - Value Anchor)
                for r in display_regs:
                    # 每列顯示分區定價 (db[r][0])
                    rate_list = int((db[r][0] / db["Std_Spots"]) * factor)
                    pkg_list = rate_list * spots_final
                    
                    # Total 累積邏輯：全省聯播時，Total 來自全省定價
                    if cfg["is_national"]:
                        if r == "北區": # 只加一次全省總價
                            nat_list = db["全省"][0]
                            nat_rate = int((nat_list / db["Std_Spots"]) * factor)
                            total_list_price_accum += nat_rate * spots_final
                    else:
                        total_list_price_accum += pkg_list

                    rows.append({
                        "media": m, "region": r, 
                        "location": f"{r.replace('區', '')}區-{r}",
                        "program": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec,
                        "spots": spots_final, "schedule": sch,
                        "rate_list": rate_list, "pkg_display_val": pkg_list,
                        "is_pkg_start": (cfg["is_national"] and r == "北區"),
                        "is_pkg_member": cfg["is_national"]
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
                
                # 家樂福顯示邏輯 (List)
                rate_h = int((db["量販_全省"]["List"] / base_std) * factor)
                pkg_h = rate_h * spots_final
                total_list_price_accum += pkg_h
                
                rows.append({"media": m, "region": "全省量販", "location": "全省量販", "program": STORE_COUNTS_NUM["家樂福_量販"], "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch_h, "rate_list": rate_h, "pkg_display_val": pkg_h, "is_pkg_start": False, "is_pkg_member": False})
                
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({"media": m, "region": "全省超市", "location": "全省超市", "program": STORE_COUNTS_NUM["家樂福_超市"], "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "spots": spots_s, "schedule": sch_s, "rate_list": "計量販", "pkg_display_val": "計量販", "is_pkg_start": False, "is_pkg_member": False})
    
    return rows, total_list_price_accum, debug_logs

# ==============================================================================
# ☪️ [模組 C：渲染引擎] (Renderer - Excel & PDF)
# ==============================================================================
def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    """將 Excel bytes 轉為 PDF bytes (優先使用 Excel COM, 其次 LibreOffice)"""
    # 1. Windows Excel COM
    if os.name == "nt":
        try:
            import win32com.client
            with tempfile.TemporaryDirectory() as tmp:
                xlsx_path = os.path.join(tmp, "cue.xlsx")
                pdf_path = os.path.join(tmp, "cue.pdf")
                with open(xlsx_path, "wb") as f: f.write(xlsx_bytes)
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False; excel.DisplayAlerts = False
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

    # 2. LibreOffice
    soffice = find_soffice_path()
    if soffice:
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
    return None, "Fail", "無可用的 Excel 轉檔引擎"

def generate_excel(rows, days_cnt, start_dt, end_dt, c_name, products, total_list, grand_total, budget, prod):
    media_order_map = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    rows.sort(key=lambda x: (media_order_map.get(x['media'], 99), x['seconds'], REGIONS_ORDER.index(x['region']) if x['region'] in REGIONS_ORDER else 99))
    
    used_media = sorted(list(set(r['media'] for r in rows)), key=lambda x: media_order_map.get(x, 99))
    mediums = "、".join(used_media)
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Media Schedule")
    
    # 格式定義
    fmt_title = workbook.add_format({'font_size': 18, 'bold': True, 'align': 'center', 'font_name': 'Arial'})
    fmt_header_left = workbook.add_format({'align': 'left', 'valign': 'top', 'bold': True, 'font_name': 'Arial', 'font_size': 10})
    fmt_col_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white', 'text_wrap': True, 'font_size': 10, 'font_name': 'Arial'})
    fmt_date_wk = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white', 'font_name': 'Arial'})
    fmt_date_we = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFD966', 'font_name': 'Arial'}) 
    fmt_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'font_name': 'Arial'})
    fmt_cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'text_wrap': True, 'font_name': 'Arial'})
    fmt_num = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0', 'font_size': 10, 'font_name': 'Arial'})
    fmt_spots = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFF2CC', 'font_size': 10, 'font_name': 'Arial'})
    fmt_total = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#E2EFDA', 'num_format': '#,##0', 'font_size': 10, 'font_name': 'Arial'})
    fmt_discount = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'font_color': 'red', 'num_format': '#,##0', 'font_size': 10, 'font_name': 'Arial'})
    fmt_grand_total = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFC107', 'num_format': '#,##0', 'font_size': 10, 'font_name': 'Arial'})

    # 繪製表頭
    worksheet.merge_range('A1:AJ1', "Media Schedule", fmt_title)
    info = [("客戶名稱：", c_name), ("Product：", products), ("Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"), ("Medium :", mediums)]
    for i, (label, val) in enumerate(info):
        worksheet.write(2+i, 0, label, fmt_header_left)
        worksheet.write(2+i, 1, val, fmt_header_left)

    worksheet.write(6, 6, f"{start_dt.month}月", fmt_cell)
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    curr = start_dt
    for i in range(days_cnt):
        col_idx = 7 + i
        wd = curr.weekday()
        fmt = fmt_date_we if wd >= 5 else fmt_date_wk
        worksheet.write(7, col_idx, curr.day, fmt)
        worksheet.write(8, col_idx, weekdays[wd], fmt)
        curr += timedelta(days=1)

    # 🌟 Header: 顯示 (Net) 但填 List 值
    headers = ["Station", "Location", "Program", "Day-part", "Size", "rate\n(Net)", "Package-cost\n(Net)"]
    for i, h in enumerate(headers): worksheet.write(8, i, h, fmt_col_header)
    
    last_col = 7 + days_cnt
    worksheet.write(8, last_col, "檔次", fmt_col_header)

    current_row = 9
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]['media'] == row['media'] and rows[j]['seconds'] == row['seconds']: j += 1
        group_size = j - i
        
        m_name = row['media']
        if "全家廣播" in m_name: m_name = "全家便利商店\n通路廣播廣告"
        if "新鮮視" in m_name: m_name = "全家便利商店\n新鮮視廣告"
        
        # Merge Station Name
        if group_size > 1:
            worksheet.merge_range(current_row, 0, current_row + group_size - 1, 0, m_name, fmt_cell_left)
        else:
            worksheet.write(current_row, 0, m_name, fmt_cell_left)
            
        for k in range(group_size):
            r_data = rows[i + k]
            r_idx = current_row + k
            
            loc_txt = region_display(r_data['region'])
            if "北北基" in loc_txt and "廣播" in r_data['media']: loc_txt = "北區-北北基+東"
            
            worksheet.write(r_idx, 1, loc_txt, fmt_cell)
            worksheet.write(r_idx, 2, r_data['program'], fmt_cell)
            worksheet.write(r_idx, 3, r_data['daypart'], fmt_cell)
            worksheet.write(r_idx, 4, f"{r_data['seconds']}秒", fmt_cell)
            
            # Rate & Package: 顯示分區 List Price
            worksheet.write(r_idx, 5, r_data['rate_list'], fmt_num)
            worksheet.write(r_idx, 6, r_data['pkg_display_val'], fmt_num)

            for d_idx, s_val in enumerate(r_data['schedule']):
                worksheet.write(r_idx, 7 + d_idx, s_val, fmt_cell)
            worksheet.write(r_idx, last_col, r_data['spots'], fmt_spots)

        current_row += group_size
        i = j

    # Total Row
    worksheet.write(current_row, 2, "Total (List Price)", fmt_total)
    worksheet.write(current_row, 5, "", fmt_total)
    # 🌟 Total: 顯示全省優惠打包價 (Value Anchor Override)
    worksheet.write(current_row, 6, total_list, fmt_total)
    worksheet.write(current_row, last_col, sum(r['spots'] for r in rows), fmt_spots)
    
    # Footer
    current_row += 1
    worksheet.write(current_row, 6, "製作", fmt_cell)
    worksheet.write(current_row, 7, prod, fmt_num)
    current_row += 1
    worksheet.write(current_row, 6, "專案優惠價 (Budget)", fmt_cell)
    worksheet.write(current_row, 7, budget, fmt_discount)
    current_row += 1
    vat_val = int(round((budget + prod) * 0.05))
    worksheet.write(current_row, 6, "5% VAT", fmt_cell)
    worksheet.write(current_row, 7, vat_val, fmt_num)
    current_row += 1
    final_total = budget + prod + vat_val
    worksheet.write(current_row, 6, "Grand Total", fmt_grand_total)
    worksheet.write(current_row, 7, final_total, fmt_grand_total)

    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:E', 12)
    worksheet.set_column('F:G', 12)
    worksheet.set_column(7, last_col, 4)
    workbook.close()
    return output

# ==============================================================================
# 🇩 [模組 D：前端介面] (Streamlit UI)
# ==============================================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v68.3")
st.title("📺 媒體 Cue 表生成器")

template_bytes, source_type = load_default_template()
has_template = (template_bytes is not None)

if has_template:
    st.success(f"✅ 已載入系統公版 ({source_type})")
else:
    st.warning("⚠️ 無法載入公版，請手動上傳")
    tpl = st.file_uploader("上傳 Excel 模板", type=["xlsx"])
    if tpl: 
        with open(DEFAULT_FILENAME, "wb") as f: f.write(tpl.read())
        st.experimental_rerun()

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

# 執行
if config:
    rows, total_list_accum, logs = calculate_plan_data(config, total_budget_input, days_count)
    
    # 簡易預覽
    st.write("### 預覽 (與 Excel 輸出一致)")
    st.dataframe(pd.DataFrame(rows)[['media', 'region', 'spots', 'rate_list', 'pkg_display_val']])
    
    # Debug Panel
    with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
        st.markdown("#### 1. 本次預算分配 (Waterfall)")
        for log in logs:
            color = "green" if log["status"] == "達標" else "red"
            st.markdown(f"**{log['media']} ({log['sec']}秒)**: 預算${log['budget']:,.0f} | 執行{log['spots']}檔 -> <span style='color:{color}'><b>{log['status']}</b></span>", unsafe_allow_html=True)

    # 產出
    if has_template and rows:
        try:
            prod_cost = 10000
            vat = int(round((total_budget_input + prod_cost) * 0.05))
            grand_total = total_budget_input + prod_cost + vat
            p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}"
            
            xlsx = generate_excel(rows, days_count, start_date, end_date, client_name, p_str, total_list_accum, grand_total, total_budget_input, prod_cost)
            st.download_button("下載 Excel", xlsx, f"Cue_{client_name}.xlsx")
            
            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx.getvalue())
            if pdf_bytes:
                st.download_button(f"下載 PDF ({method})", pdf_bytes, f"Cue_{client_name}.pdf")
            else:
                st.error(f"PDF 產出失敗: {err}")
        except Exception as e: st.error(f"Excel 產出錯誤: {e}")
