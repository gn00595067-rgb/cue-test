import streamlit as st
import pandas as pd
import math
import io
import xlsxwriter
from datetime import timedelta, datetime

# ==========================================
# 1. 基礎資料與設定 (2026 新制)
# ==========================================

STORE_COUNTS = {
    "全省": "4,437店",
    "北區": "北北基 1,649店",
    "桃竹苗": "桃竹苗 779店",
    "中區": "中彰投 839店",
    "雲嘉南": "雲嘉南 499店",
    "高屏": "高高屏 490店",
    "東區": "宜花東 181店",
    "新鮮視_全省": "3,124面",
    "新鮮視_北區": "北北基 1,127面",
    "新鮮視_桃竹苗": "桃竹苗 616面",
    "新鮮視_中區": "中彰投 528面",
    "新鮮視_雲嘉南": "雲嘉南 365面",
    "新鮮視_高屏": "高高屏 405面",
    "新鮮視_東區": "宜花東 83面",
    "家樂福_量販": "68店",
    "家樂福_超市": "249店"
}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# --- 價格資料庫 (List=定價, Net=實收價) 2026版 ---
PRICING_DB = {
    "全家廣播": {
        "Std_Spots": 480, 
        "Base_Sec": 30,    
        "Day_Part": "07:00-23:00",
        "全省": [400000, 320000], 
        "北區": [250000, 200000], "桃竹苗": [150000, 120000],
        "中區": [150000, 120000], "雲嘉南": [100000, 80000], 
        "高屏": [100000, 80000], "東區": [62500, 50000]
    },
    "新鮮視": {
        "Std_Spots": 504, 
        "Base_Sec": 10,    
        "Day_Part": "07:00-23:00",
        "全省": [150000, 120000], 
        "北區": [150000, 120000], "桃竹苗": [120000, 96000],
        "中區": [90000, 72000], "雲嘉南": [75000, 60000], 
        "高屏": [75000, 60000], "東區": [45000, 36000]
    },
    "家樂福": {
        "Base_Sec": 20,
        "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420, "Day_Part": "09:00-23:00"},
        "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720, "Day_Part": "00:00-24:00"} 
    }
}

# --- 秒數折扣係數 ---
SEC_FACTORS = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5, 5: 0.25},
    "新鮮視":   {30: 3.0, 20: 2.0, 15: 1.5, 10: 1.0, 5: 0.5},
    "家樂福":   {30: 1.5, 20: 1.0, 15: 0.85, 10: 0.65, 5: 0.35}
}

def get_sec_factor(media_type, seconds):
    factors = SEC_FACTORS.get(media_type, {})
    return factors.get(seconds, 1.0) 

def calculate_schedule(total_spots, days):
    """ 
    偶數排程演算法 (Even Distribution Strategy)
    1. 將總檔次除以 2
    2. 分配到每天
    3. 結果乘以 2
    """
    if days == 0: return []
    
    if total_spots % 2 != 0: total_spots += 1
        
    half_spots = total_spots // 2
    schedule = [0] * days
    
    base = half_spots // days
    remaining = half_spots % days
    
    for i in range(days):
        schedule[i] = base
        if i < remaining:
            schedule[i] += 1
            
    final_schedule = [x * 2 for x in schedule]
    return final_schedule

# ==========================================
# 2. UI 設定
# ==========================================

st.set_page_config(layout="wide", page_title="Cue Sheet Generator 2026 (v60.5)")
st.title("📺 媒體 Cue 表生成器")

# --- 1. 基本資料 ---
with st.container():
    st.markdown("### 1. 基本資料設定")
    with st.expander("📝 點擊展開/收合基本資料", expanded=True):
        col_b1, col_b2 = st.columns([1, 1])
        with col_b1:
            client_name = st.text_input("客戶名稱", "萬國通路")
            start_date = st.date_input("開始日", datetime(2025, 1, 1))
        with col_b2:
            total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)
            end_date = st.date_input("結束日", datetime(2025, 1, 31))
        
        days_count = (end_date - start_date).days + 1
        st.info(f"📅 走期共 **{days_count}** 天")

# --- [新增] 備註設定 ---
with st.expander("📝 備註欄位設定 (Remarks)", expanded=False):
    rc1, rc2, rc3 = st.columns(3)
    sign_deadline = rc1.date_input("回簽截止日", datetime.now() + timedelta(days=3))
    billing_month = rc2.text_input("請款月份", "2026年2月")
    payment_date = rc3.date_input("付款兌現日", datetime(2026, 3, 31))

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

# --- 2. 媒體設定 ---
config_media = {}
st.markdown("### 2. 媒體投放設定 (連動總和 100%)")

col_m1, col_m2, col_m3 = st.columns(3)
remaining_global_share = 100 

# 全家廣播
with col_m1:
    st.markdown("#### 📻 全家廣播")
    fm_act = st.checkbox("開啟", value=True, key="fm_act")
    if fm_act:
        is_nat = st.checkbox("全省聯播", value=True, key="fm_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, key="fm_reg")
        _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="fm_sec")
        secs = sorted(_secs_input)
        share = st.slider("預算佔比%", 0, remaining_global_share, min(70, remaining_global_share), key="fm_share")
        remaining_global_share -= share
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"fm_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.write(f"🔹 {secs[-1]}秒: {ls}%")
        elif secs: sec_shares[secs[0]] = 100
        config_media["全家廣播"] = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# 新鮮視
with col_m2:
    st.markdown("#### 📺 新鮮視")
    fv_act = st.checkbox("開啟", value=True, key="fv_act")
    if fv_act:
        is_nat = st.checkbox("全省聯播 ", value=False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區", "桃竹苗"], key="fv_reg")
        _secs_input = st.multiselect("秒數", DURATIONS, default=[10], key="fv_sec")
        secs = sorted(_secs_input)
        limit = remaining_global_share
        default_val = min(20, limit)
        share = st.slider("預算佔比% ", 0, limit, default_val, key="fv_share")
        remaining_global_share -= share
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比 ", 0, ls, int(ls/2), key=f"fv_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.write(f"🔹 {secs[-1]}秒: {ls}%")
        elif secs: sec_shares[secs[0]] = 100
        config_media["新鮮視"] = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# 家樂福
with col_m3:
    st.markdown("#### 🛒 家樂福")
    cf_act = st.checkbox("開啟", value=True, key="cf_act")
    if cf_act:
        st.write("區域：全省")
        _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="cf_sec")
        secs = sorted(_secs_input)
        share = remaining_global_share
        st.info(f"預算佔比: **{share}%** (自動填滿)")
        st.progress(share / 100.0 if share <= 100 else 1.0)
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比  ", 0, ls, int(ls/2), key=f"cf_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.write(f"🔹 {secs[-1]}秒: {ls}%")
        elif secs: sec_shares[secs[0]] = 100
        config_media["家樂福"] = {"regions": ["全省"], "seconds": secs, "share": share, "sec_shares": sec_shares}

# ==========================================
# 3. 計算邏輯 (核心引擎 v60.5)
# ==========================================

final_rows = []
all_secs = set()
total_list_price_accum = 0
debug_logs = [] 

if sum(m["share"] for m in config_media.values()) > 0:
    for m_type, cfg in config_media.items():
        media_budget = total_budget_input * (cfg["share"] / 100.0)
        
        for sec, sec_share in cfg["sec_shares"].items():
            all_secs.add(f"{sec}秒")
            sec_budget = media_budget * (sec_share / 100.0)
            if sec_budget <= 0: continue
            
            factor = get_sec_factor(m_type, sec)
            
            log_item = {
                "media": m_type, "sec": sec, "budget": sec_budget, 
                "status": "OK", "reason": "", "spots": 0, 
                "std": 0, "factor": factor, "unit_cost": 0
            }

            if m_type in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m_type]
                std_spots = db["Std_Spots"]
                day_part = db["Day_Part"]
                
                calc_regions = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regions = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                # --- Step 1: 裡子 (Net) 算檔次 ---
                temp_net_unit_sum = 0
                for reg in calc_regions:
                    net_price_total = db[reg][1] # 取 Net Price
                    unit_net = (net_price_total / std_spots) * factor
                    temp_net_unit_sum += unit_net
                
                if temp_net_unit_sum == 0: continue
                
                # 逆推 + 懲罰 + 偶數修正
                initial_spots = math.ceil(sec_budget / temp_net_unit_sum)
                is_under_target = initial_spots < std_spots
                multiplier = 1.1 if is_under_target else 1.0
                
                final_unit_net = temp_net_unit_sum * multiplier
                target_spots = math.ceil(sec_budget / final_unit_net)
                if target_spots % 2 != 0: target_spots += 1 
                if target_spots == 0: target_spots = 2
                
                log_item["spots"] = target_spots
                log_item["std"] = std_spots
                log_item["unit_cost"] = final_unit_net
                log_item["status"] = "未達標" if is_under_target else "達標"
                log_item["reason"] = f"觸發 x1.1" if is_under_target else "費率正常"

                daily_sch = calculate_schedule(target_spots, days_count)
                
                # --- Step 2: 面子 (List) 填報表 ---
                for i, reg in enumerate(display_regions):
                    reg_list_total = db.get(reg, [0,0])[0] if cfg["is_national"] else db[reg][0]
                    
                    rate_list_display = int((reg_list_total / std_spots) * factor)
                    pkg_display_val = rate_list_display * target_spots
                    
                    if not cfg["is_national"]:
                        total_list_price_accum += pkg_display_val
                    elif cfg["is_national"] and reg == "北區":
                        nat_list_total = db["全省"][0]
                        total_list_price_accum += int((nat_list_total / std_spots) * factor * target_spots)

                    prog_name = STORE_COUNTS.get(reg, reg)
                    if m_type == "新鮮視": prog_name = STORE_COUNTS.get(f"新鮮視_{reg}", reg)
                    
                    final_rows.append({
                        "media": m_type, "region": reg, 
                        "location": f"{reg.replace('區', '')}區-{reg}" if m_type=="全家廣播" else f"{reg.replace('區', '')}區-{reg}",
                        "program": prog_name, 
                        "daypart": day_part,
                        "seconds": sec, "schedule": daily_sch, "spots": target_spots,
                        "rate_list": rate_list_display, 
                        "pkg_display_val": pkg_display_val,
                        "is_pkg_start": (cfg["is_national"] and reg == "北區"), 
                        "is_pkg_member": cfg["is_national"]
                    })

            elif m_type == "家樂福":
                db = PRICING_DB["家樂福"]
                
                # [FIX] 使用 DB 中的基準檔次，不再寫死 420/720
                base_std_spots = db["量販_全省"]["Std_Spots"]
                sup_std_spots = db["超市_全省"]["Std_Spots"]
                
                base_net = db["量販_全省"]["Net"] 
                unit_net = (base_net / base_std_spots) * factor
                
                initial_spots = math.ceil(sec_budget / unit_net)
                is_under_target = initial_spots < base_std_spots
                multiplier = 1.1 if is_under_target else 1.0
                
                final_unit_net = unit_net * multiplier
                target_spots = math.ceil(sec_budget / final_unit_net)
                if target_spots % 2 != 0: target_spots += 1
                if target_spots == 0: target_spots = 2

                log_item["spots"] = target_spots
                log_item["std"] = base_std_spots
                log_item["unit_cost"] = final_unit_net
                log_item["status"] = "未達標" if is_under_target else "達標"
                log_item["reason"] = f"觸發 x1.1" if is_under_target else "正常"

                sch = calculate_schedule(target_spots, days_count)
                
                # List 顯示
                base_list = db["量販_全省"]["List"]
                rate_list_display = int((base_list / base_std_spots) * factor)
                pkg_list_display = rate_list_display * target_spots
                
                total_list_price_accum += pkg_list_display
                
                final_rows.append({
                    "media": "家樂福", "region": "全省量販", "location": "全省量販", "program": STORE_COUNTS["家樂福_量販"],
                    "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "schedule": sch, "spots": target_spots,
                    "rate_list": rate_list_display, "pkg_display_val": pkg_list_display,
                    "is_pkg_start": False, "is_pkg_member": False
                })
                
                spots_s = int(target_spots * (sup_std_spots / base_std_spots))
                sch_s = calculate_schedule(spots_s, days_count)
                final_rows.append({
                    "media": "家樂福", "region": "全省超市", "location": "全省超市", "program": STORE_COUNTS["家樂福_超市"],
                    "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "schedule": sch_s, "spots": spots_s,
                    "rate_list": "計量販", "pkg_display_val": "計量販",
                    "is_pkg_start": False, "is_pkg_member": False
                })
            
            debug_logs.append(log_item)

media_order_map = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
final_rows.sort(key=lambda x: media_order_map.get(x['media'], 99))

def parse_sec_int(s):
    return int(s.replace("秒", ""))
sorted_secs_list = sorted(list(all_secs), key=parse_sec_int)
product_str = "、".join(sorted_secs_list)

prod_cost = 10000
vat = int(round((total_budget_input + prod_cost) * 0.05))
grand_total = total_budget_input + prod_cost + vat
discount_ratio_val = (total_budget_input / total_list_price_accum * 100) if total_list_price_accum > 0 else 0
discount_ratio_str = f"{discount_ratio_val:.1f}% (約 {discount_ratio_val/10:.1f} 折)"

# ==========================================
# 4. 生成 HTML 預覽
# ==========================================

def generate_html_preview(rows, days_cnt, start_dt, c_name, products, total_list, grand_total, budget, prod, remarks):
    used_media = sorted(list(set(r['media'] for r in rows)), key=lambda x: media_order_map.get(x, 99))
    mediums_str = "、".join(used_media)

    date_header_row1 = f"<th class='header-blue' colspan='{days_cnt}'>{start_dt.month}月</th>"
    date_header_row2 = ""
    date_header_row3 = ""
    curr = start_dt
    weekdays_map = ["一", "二", "三", "四", "五", "六", "日"]
    for i in range(days_cnt):
        wd = curr.weekday()
        cls = "header-yellow" if wd >= 5 else "header-blue"
        date_header_row2 += f"<th class='{cls}'>{curr.day}</th>"
        date_header_row3 += f"<th class='{cls}'>{weekdays_map[wd]}</th>"
        curr += timedelta(days=1)
        
    data_rows_html = ""
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]['media'] == row['media'] and rows[j]['seconds'] == row['seconds']:
            j += 1
        group_size = j - i
        m_name = row['media']
        if "全家廣播" in m_name: m_name = "全家便利商店<br>通路廣播廣告"
        if "新鮮視" in m_name: m_name = "全家便利商店<br>新鮮視廣告"
        
        for k in range(group_size):
            r_data = rows[i+k]
            tr = "<tr>"
            if k == 0:
                tr += f"<td rowspan='{group_size}' class='align-left'>{m_name}</td>"
            
            loc_txt = r_data['location']
            if "北北基" in loc_txt and "廣播" in r_data['media']: loc_txt = "北區-北北基+東"
            
            tr += f"<td>{loc_txt}</td>"
            tr += f"<td>{r_data['program']}</td>"
            tr += f"<td>{r_data['daypart']}</td>"
            tr += f"<td>{r_data['seconds']}秒</td>"
            
            rate_disp = f"{r_data['rate_list']:,}" if isinstance(r_data['rate_list'], int) else r_data['rate_list']
            tr += f"<td class='align-right'>{rate_disp}</td>"
            
            pkg_disp = f"{r_data['pkg_display_val']:,}" if isinstance(r_data['pkg_display_val'], int) else r_data['pkg_display_val']
            
            if row['is_pkg_start']:
                 if k == 0 and group_size > 1:
                     tr += f"<td rowspan='{group_size}' class='align-right'>{pkg_disp}</td>" 
            elif row['is_pkg_member']:
                pass
            else:
                tr += f"<td class='align-right'>{pkg_disp}</td>"
            
            for s_val in r_data['schedule']:
                tr += f"<td>{s_val}</td>"
            tr += f"<td class='cell-yellow'>{r_data['spots']}</td>"
            tr += "</tr>"
            data_rows_html += tr
        i = j
    
    vat_val = int(round((budget + prod) * 0.05))
    final_total = budget + prod + vat_val

    css_style = """
    <style>
        .preview-table { width: 100%; border-collapse: collapse; font-family: "Microsoft JhengHei", "Arial", sans-serif; font-size: 13px; color: #000; min-width: 1200px; background-color: #ffffff; }
        .preview-table th, .preview-table td { border: 1px solid #555; padding: 8px; text-align: center; vertical-align: middle; }
        .header-blue { background-color: #2c3e50; color: white !important; font-weight: bold; }
        .header-yellow { background-color: #f1c40f; color: #000 !important; font-weight: bold; }
        .cell-yellow { background-color: #fff3cd; color: #000 !important; font-weight: bold; }
        .row-total { background-color: #d4edda; color: #000 !important; font-weight: bold; }
        .row-grand-total { background-color: #ffc107; color: #000 !important; font-weight: bold; font-size: 15px; border-top: 2px solid #000; }
        .align-left { text-align: left; }
        .align-right { text-align: right; }
        .remarks { margin-top: 20px; font-size: 14px; text-align: left; line-height: 1.6; white-space: pre-wrap; color: #333; background-color: #f9f9f9; padding: 10px; border-radius: 5px; border: 1px solid #ddd; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #e6f7ff; }
    </style>
    """

    # 備註顯示邏輯
    remarks_html = ""
    if remarks:
        remarks_html = f"<div class='remarks'><b>Remarks：</b><br>{'<br>'.join(remarks)}</div>"

    html = f"""
    {css_style}
    <div style="overflow-x: auto; width: 100%;">
        <table class="preview-table">
            <tr>
                <td colspan="5" class="align-left" style="background-color:#fff; border:none;">
                    <b>客戶名稱：</b> {c_name}<br><b>Product：</b> {products}<br><b>Period：</b> {start_dt.strftime('%Y. %m. %d')} - {end_date.strftime('%Y. %m. %d')}<br><b>Medium：</b> {mediums_str}
                </td>
                <td colspan="{days_cnt + 3}" style="background-color:#fff; border:none;"></td>
            </tr>
            <tr><th colspan="7" style="border:none;"></th>{date_header_row1}<th style="border:none;"></th></tr>
            <tr>
                <th rowspan="2" class="header-blue">Station</th><th rowspan="2" class="header-blue">Location</th><th rowspan="2" class="header-blue">Program</th>
                <th rowspan="2" class="header-blue">Day-part</th><th rowspan="2" class="header-blue">Size</th><th rowspan="2" class="header-blue">rate (List)</th>
                <th rowspan="2" class="header-blue">Package-cost<br>(List)</th>{date_header_row2}<th rowspan="2" class="header-blue">檔次</th>
            </tr>
            <tr>{date_header_row3}</tr>
            {data_rows_html}
            <tr class="row-total">
                <td colspan="5" class="align-right">Total (List Price)</td>
                <td class="align-right">{sum(r['rate_list'] for r in rows if isinstance(r['rate_list'], int)):,}</td>
                <td class="align-right">{total_list:,}</td>
                <td colspan="{days_cnt}"></td>
                <td class="cell-yellow">{sum(r['spots'] for r in rows)}</td>
            </tr>
            <tr><td colspan="6" class="align-right">製作</td><td class="align-right">{prod:,}</td><td colspan="{days_cnt + 1}"></td></tr>
            <tr><td colspan="6" class="align-right">專案優惠價 (Budget)</td><td class="align-right" style="color:red; font-weight:bold;">{budget:,}</td><td colspan="{days_cnt + 1}"></td></tr>
            <tr><td colspan="6" class="align-right">5% VAT</td><td class="align-right">{vat_val:,}</td><td colspan="{days_cnt + 1}"></td></tr>
            <tr class="row-grand-total"><td colspan="6" class="align-right">Grand Total</td><td class="align-right">{final_total:,}</td><td colspan="{days_cnt + 1}"></td></tr>
        </table>
        {remarks_html}
    </div>
    """
    return html

def generate_excel(rows, days_cnt, start_dt, c_name, products, total_list, grand_total, budget, prod, remarks):
    used_media = sorted(list(set(r['media'] for r in rows)), key=lambda x: media_order_map.get(x, 99))
    mediums = "、".join(used_media)
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Media Schedule")
    
    fmt_title = workbook.add_format({'font_size': 18, 'bold': True, 'align': 'center'})
    fmt_header_left = workbook.add_format({'align': 'left', 'valign': 'top', 'bold': True})
    fmt_col_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white', 'text_wrap': True, 'font_size': 10})
    fmt_date_wk = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white'})
    fmt_date_we = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFD966'}) 
    fmt_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    fmt_cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'text_wrap': True})
    fmt_num = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0', 'font_size': 10})
    fmt_spots = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFF2CC', 'font_size': 10})
    fmt_total = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#E2EFDA', 'num_format': '#,##0', 'font_size': 10})
    fmt_discount = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'font_color': 'red', 'num_format': '#,##0', 'font_size': 10})
    fmt_grand_total = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFC107', 'num_format': '#,##0', 'font_size': 10})

    worksheet.merge_range('A1:AJ1', "Media Schedule", fmt_title)
    
    info = [("客戶名稱：", c_name), ("Product：", products), ("Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_date.strftime('%Y. %m. %d')}"), ("Medium :", mediums)]
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

    headers = ["Station", "Location", "Program", "Day-part", "Size", "rate (List)", "Package-cost\n(List)"]
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
            loc_txt = r_data['location']
            if "北北基" in loc_txt and "廣播" in r_data['media']: loc_txt = "北區-北北基+東"
            worksheet.write(r_idx, 1, loc_txt, fmt_cell)
            worksheet.write(r_idx, 2, r_data['program'], fmt_cell)
            worksheet.write(r_idx, 3, r_data['daypart'], fmt_cell)
            worksheet.write(r_idx, 4, f"{r_data['seconds']}秒", fmt_cell)
            
            # Rate (List)
            if isinstance(r_data['rate_list'], int):
                worksheet.write(r_idx, 5, r_data['rate_list'], fmt_num)
            else:
                worksheet.write(r_idx, 5, r_data['rate_list'], fmt_cell)
            
            # Package (List)
            if r_data['is_pkg_start']:
                 if k == 0 and group_size > 1:
                     worksheet.merge_range(current_row, 6, current_row + group_size - 1, 6, r_data['pkg_display_val'], fmt_num)
                 elif k == 0:
                     worksheet.write(r_idx, 6, r_data['pkg_display_val'], fmt_num)
            elif not r_data['is_pkg_member']:
                 if isinstance(r_data['pkg_display_val'], int):
                     worksheet.write(r_idx, 6, r_data['pkg_display_val'], fmt_num)
                 else:
                     worksheet.write(r_idx, 6, r_data['pkg_display_val'], fmt_cell)

            for d_idx, s_val in enumerate(r_data['schedule']):
                worksheet.write(r_idx, 7 + d_idx, s_val, fmt_cell)
            worksheet.write(r_idx, last_col, r_data['spots'], fmt_spots)

        current_row += group_size
        i = j

    worksheet.write(current_row, 2, "Total (List Price)", fmt_total)
    worksheet.write(current_row, 5, sum(r['rate_list'] for r in rows if isinstance(r['rate_list'], int)), fmt_total)
    worksheet.write(current_row, 6, total_list, fmt_total)
    
    total_spots_daily = [0] * days_cnt
    for r in rows:
        for idx, val in enumerate(r['schedule']):
            total_spots_daily[idx] += val
    for idx, val in enumerate(total_spots_daily):
        worksheet.write(current_row, 7+idx, val, fmt_cell)
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

    # 寫入備註
    if remarks:
        current_row += 2
        worksheet.write(current_row, 0, "Remarks：", fmt_header_left)
        for idx, remark in enumerate(remarks):
            worksheet.write(current_row + 1 + idx, 0, remark, fmt_header_left)

    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:E', 12)
    worksheet.set_column('F:G', 12)
    worksheet.set_column(7, last_col, 4)
    workbook.close()
    return output

# ==========================================
# 5. 結果顯示與下載
# ==========================================

st.markdown("### 3. 計算結果摘要")
m1, m2, m3 = st.columns(3)
m1.metric("客戶預算 (未稅)", f"{total_budget_input:,}")
m2.metric("折扣後總金額 (含稅)", f"{grand_total:,}", help="預算 + 製作 + 稅")
m3.metric("牌價折扣率", discount_ratio_str, delta_color="normal")

st.markdown("### 4. Cue 表網頁預覽")

if final_rows:
    remarks_text = get_remarks_text(sign_deadline, billing_month, payment_date)
    html_preview = generate_html_preview(final_rows, days_count, start_date, client_name, product_str, total_list_price_accum, grand_total, total_budget_input, prod_cost, remarks_text)
    st.components.v1.html(html_preview, height=600, scrolling=True)

    xlsx_data = generate_excel(final_rows, days_count, start_date, client_name, product_str, total_list_price_accum, grand_total, total_budget_input, prod_cost, remarks_text)
    
    st.download_button(
        label="📥 下載 Excel Cue表 (.xlsx)",
        data=xlsx_data.getvalue(),
        file_name=f"CueSheet_{client_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
