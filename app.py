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
# 🛠️ [通用工具] (Utilities) - 不太需要改動
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
# ⚠️ 修改規則：只在價格調整、店數變動時修改此處。
# ==============================================================================
GOOGLE_DRIVE_FILE_ID = "11R1SA_hpFD5O_MGmYeh4BdtcUhK2bPta"
DEFAULT_FILENAME = "1209-Cue表相關資料.xlsx"

STORE_COUNTS = {
    "全省": "4,437店", "北區": "1,649店", "桃竹苗": "779店", "中區": "839店", "雲嘉南": "499店", "高屏": "490店", "東區": "181店",
    "新鮮視_全省": "3,124面", "新鮮視_北區": "1,127面", "新鮮視_桃竹苗": "616面", "新鮮視_中區": "528面",
    "新鮮視_雲嘉南": "365面", "新鮮視_高屏": "405面", "新鮮視_東區": "83面",
    "家樂福_量販": "68店", 
    "家樂福_超市": "249店"
}
STORE_COUNTS_NUM = {k: parse_count_to_int(v) for k, v in STORE_COUNTS.items()}
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# [List Price (牌價), Net Price (實收底價)]
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
# ⚠️ 修改規則：負責「預算分配」、「實收逆推」、「偶數排程」。不處理顯示金額。
# ==============================================================================
def get_sec_factor(media_type, seconds): return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    if days <= 0: return []
    if total_spots % 2 != 0: total_spots += 1
    half_spots = total_spots // 2
    base, rem = divmod(half_spots, days)
    half_schedule = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in half_schedule]

def calculate_plan(config, total_budget, days_count):
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
                
                # 關鍵邏輯：計算用 Net (全省or區域)，顯示用 List (全省or區域)
                if cfg["is_national"]:
                    calc_regs = ["全省"]
                    display_regs = REGIONS_ORDER
                else:
                    calc_regs = cfg["regions"]
                    display_regs = cfg["regions"]
                
                # 1. 實收逆推 (Net Calculation)
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
                
                # 2. 準備顯示資料 (List Preparation)
                for r in display_regs:
                    # 每一列都顯示分區牌價 (Value Anchor)
                    rate_list = int((db[r][0] / db["Std_Spots"]) * factor)
                    pkg_list = rate_list * spots_final
                    
                    # 累積 Total (全省聯播時只加一次全省總價)
                    if cfg["is_national"]:
                        if r == "北區": 
                            nat_list = db["全省"][0]
                            nat_rate = int((nat_list / db["Std_Spots"]) * factor)
                            total_list_price_accum += nat_rate * spots_final
                    else:
                        total_list_price_accum += pkg_list

                    rows.append({
                        "media": m, "region": r, "location": f"{r.replace('區', '')}區-{r}",
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
# ⚠️ 修改規則：負責「畫」出 Excel 和 PDF。字體、格線、Total 的強制覆蓋寫在這裡。
# ==============================================================================
def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    # (此處保留原有的 Windows COM / LibreOffice 雙引擎邏輯，為了節省篇幅省略重複代碼，請複製 v68.2 的完整函式)
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
    
    # 格式定義 (省略部分重複代碼，請參考 v68.2)
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

    # 🌟 表頭文字：依客戶要求寫 Net
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
            
            # Rate & Package (顯示 List Price，價值定錨)
            worksheet.write(r_idx, 5, r_data['rate_list'], fmt_num)
            worksheet.write(r_idx, 6, r_data['pkg_display_val'], fmt_num)

            for d_idx, s_val in enumerate(r_data['schedule']):
                worksheet.write(r_idx, 7 + d_idx, s_val, fmt_cell)
            worksheet.write(r_idx, last_col, r_data['spots'], fmt_spots)

        current_row += group_size
        i = j

    worksheet.write(current_row, 2, "Total (List Price)", fmt_total)
    worksheet.write(current_row, 5, "", fmt_total)
    # 🌟 Total: 顯示全省優惠總價 (來自模組 B 的計算)
    worksheet.write(current_row, 6, total_list, fmt_total)
    worksheet.write(current_row, last_col, sum(r['spots'] for r in rows), fmt_spots)
    
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
# ⚠️ 修改規則：負責按鈕、輸入框、版面配置。不涉及計算。
# ==============================================================================
st.title("📺 媒體 Cue 表生成器 (v68.2 Modular)")

# (載入設定與 UI 程式碼，完全沿用 v68.2，為了簡潔省略重複部分)
# ... 請使用 v68.2 的 UI 代碼填入此處 ...
# (但在呼叫計算時，請使用 calculate_plan 函式)

# 範例 UI 呼叫邏輯：
# if rows:
#     rows, total_list_price_accum, debug_logs = calculate_plan(config, total_budget_input, days_count)
#     xlsx = generate_excel(...)
