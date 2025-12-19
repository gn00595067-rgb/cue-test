import streamlit as st
import pandas as pd
import math
import io
import xlsxwriter
from datetime import timedelta, datetime, date
import uuid
import random

# ==========================================
# 0. 系統初始化 & 旗艦級 CSS (v51.1)
# ==========================================

st.set_page_config(layout="wide", page_title="瑞迪資源控管 v51.1", initial_sidebar_state="expanded")

st.markdown("""
<style>
    /* 全局設定 */
    .stApp { background-color: #f8f9fa; font-family: 'Segoe UI', "Microsoft JhengHei", sans-serif; }
    
    /* FIX: 增加上方留白，避免標題被 Streamlit 頂部白 Bar 遮住 */
    .block-container { padding-top: 3.5rem; max-width: 98% !important; }

    /* --- 風險面板 (Risk Panel) --- */
    .risk-panel {
        background: white; border: 1px solid #ddd; border-radius: 8px;
        padding: 15px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        height: 100%; min-height: 400px;
    }
    .risk-score-safe { color: #2e7d32; font-weight: bold; font-size: 18px; }
    .risk-score-warn { color: #f57f17; font-weight: bold; font-size: 18px; }
    .risk-score-crit { color: #c62828; font-weight: bold; font-size: 18px; }
    .kpi-box { margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 10px; }
    .kpi-label { font-size: 12px; color: #666; }
    .kpi-value { font-size: 16px; font-weight: 600; color: #333; }

    /* --- 核心表格樣式 --- */
    .unified-wrapper { 
        width: 100%; height: auto; max-height: 550px; overflow: auto; 
        border: 1px solid #cfd8dc; background: white; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); border-radius: 6px; margin-bottom: 15px; 
    }
    .unified-table { border-collapse: separate; border-spacing: 0; font-size: 12px; width: 100%; min-width: 1200px; }
    
    .row-month th { position: sticky; top: 0; z-index: 35; background: #37474f; color: #fff; height: 32px; padding: 0; text-align: center; }
    .row-month th:first-child { position: sticky; left: 0; z-index: 45; background: #263238; border-right: 2px solid #546e7a; min-width: 200px; }
    
    .row-day th { position: sticky; top: 32px; z-index: 30; background: #eceff1; padding: 4px 2px; height: 40px; text-align: center; }
    .row-day th:first-child { position: sticky; left: 0; z-index: 40; background: #f5f7f8; border-right: 2px solid #cfd8dc; }
    
    .unified-table tbody td:first-child { position: sticky; left: 0; z-index: 20; background: #fff; border-right: 2px solid #cfd8dc; font-weight: 600; text-align: left; padding: 8px; }
    
    .unified-table td { 
        border-right: 1px solid #eceff1; 
        border-bottom: 1px solid #eceff1; 
        padding: 6px 4px; 
        text-align: center; 
        min-width: 70px; 
        height: 50px;
        vertical-align: middle;
    }

    /* 顏色系統 */
    .inv-safe { background-color: #e8f5e9 !important; color: #2e7d32 !important; } 
    .inv-mid  { background-color: #fffde7 !important; color: #f57f17 !important; font-weight: 600; } 
    .inv-high { background-color: #ffebee !important; color: #c62828 !important; font-weight: bold; } 
    .inv-crit { background-color: #c62828 !important; color: white !important; font-weight: bold; } 
    .inv-sim  { box-shadow: inset 0 0 0 2px #ff9800 !important; }

    /* 專案狀態列背景 (左側框線指示) */
    .proj-conf { background-color: #e3f2fd !important; border-left: 4px solid #1565c0 !important; }
    .proj-prob { background-color: #f3e5f5 !important; border-left: 4px dashed #7b1fa2 !important; }
    .proj-pend { background-color: #fff3e0 !important; border-left: 3px dotted #e65100 !important; opacity: 0.9; }

    /* 狀態標籤 (Badges) */
    .badge { padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: bold; display: inline-block; margin-top: 4px; }
    .badge-conf { background: #1565c0; color: white; }
    .badge-prob { background: #7b1fa2; color: white; }
    .badge-pend { background: #e65100; color: white; }

    /* 數值容器 */
    .val-container { 
        display: flex; flex-direction: column; 
        align-items: center; justify-content: center; 
        width: 100%; line-height: 1.4;
    }
    .val-sec { font-size: 13px; font-weight: 800; display: block; color: #2d3748; margin-bottom: 2px; }
    .val-pct { font-size: 10px; opacity: 0.7; display: block; font-weight: normal; }

    /* Tabs & Cue Table */
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { height: 50px; background-color: #fff; border-radius: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.1); }
    .stTabs [aria-selected="true"] { background-color: #e3f2fd; font-weight: bold; color: #1565c0; }

    .preview-table { width: 100%; border-collapse: separate; font-size: 12px; background: white; margin-bottom: 10px; border: 1px solid #e2e8f0; }
    .preview-table th { background: #2c5282; color: white; padding: 8px; }
    .header-yellow { background-color: #ecc94b; color: #1a202c !important; }
    .cell-ok { background-color: #e8f5e9; }
    .cell-warn { background-color: #fffbe6; font-weight: bold; color: #d69e2e; }
    .cell-err { background-color: #ffebee; font-weight: bold; color: #c62828; }
    
    .approval-card { background: white; padding: 12px; margin-bottom: 10px; border-radius: 6px; border-left: 4px solid #ff9800; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .signed-card { background: white; padding: 12px; margin-bottom: 10px; border-radius: 6px; border-left: 4px solid #7b1fa2; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 基礎資料庫
# ==========================================

MEDIA_ORDER_MAP = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

STORE_COUNTS = {
    "全省": "4,437店", "北區": "北北基 1,649店", "桃竹苗": "桃竹苗 779店",
    "中區": "中彰投 839店", "雲嘉南": "雲嘉南 499店", "高屏": "高高屏 490店", "東區": "宜花東 181店",
    "新鮮視_全省": "3,124面", "家樂福_量販": "68店", "家樂福_超市": "249店",
    "新鮮視_北區": "1,127面", "新鮮視_桃竹苗": "616面", "新鮮視_中區": "528面",
    "新鮮視_雲嘉南": "365面", "新鮮視_高屏": "405面", "新鮮視_東區": "83面",
    "家樂福_量販": "68店", "家樂福_超市": "249店"
}

DAILY_CAPACITY = { "全家廣播": 7680, "新鮮視": 5000, "家樂福": 3600 }
CAPACITY_LIMITS = { "全家廣播": 1.2, "新鮮視": 1.0, "家樂福": 999.0 } 

QUOTA_CONFIG = {
    "全家廣播": { "limits": {r: DAILY_CAPACITY["全家廣播"] for r in ["全省"] + REGIONS_ORDER} },
    "新鮮視":   { "limits": {r: DAILY_CAPACITY["新鮮視"] for r in ["全省"] + REGIONS_ORDER} },
    "家樂福":   { "limits": {"量販": DAILY_CAPACITY["家樂福"], "超市": DAILY_CAPACITY["家樂福"]} }
}

PRICING_DB = {
    "全家廣播": { "Std_Spots": 480, "Base_Sec": 30, "Day_Part": "07:00-23:00", "全省": [400000, 320000], "北區": [250000, 200000], "桃竹苗": [150000, 120000], "中區": [150000, 120000], "雲嘉南": [100000, 80000], "高屏": [100000, 80000], "東區": [62500, 50000] },
    "新鮮視":   { "Std_Spots": 504, "Base_Sec": 10, "Day_Part": "07:00-23:00", "全省": [150000, 120000], "北區": [150000, 120000], "桃竹苗": [120000, 96000], "中區": [90000, 72000], "雲嘉南": [75000, 60000], "高屏": [75000, 60000], "東區": [45000, 36000] },
    "家樂福":   { "Base_Sec": 20, "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420, "Day_Part": "09:00-23:00"}, "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720, "Day_Part": "00:00-24:00"} }
}
SEC_FACTORS = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5},
    "新鮮視":   {30: 3.0, 20: 2.0, 15: 1.5, 10: 1.0},
    "家樂福":   {30: 1.5, 20: 1.0, 15: 0.85, 10: 0.65}
}

def get_sec_factor(media_type, seconds): return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)
def calculate_schedule(total_spots, days):
    if days == 0: return []
    base = total_spots // 2 // days
    rem = (total_spots // 2) % days
    schedule = [ (base + (1 if i < rem else 0)) * 2 for i in range(days) ]
    diff = total_spots - sum(schedule)
    if diff > 0: schedule[0] += diff
    return schedule

# --- 共用資料庫 ---
if 'db' not in st.session_state:
    st.session_state.db = { "orders": [], "order_items": [] }
    def make_sched(spots, days): return calculate_schedule(spots, days)
    
    o1_id = "A1"
    st.session_state.db["orders"].append({"id": o1_id, "sales": "System", "client": "統一企業 (年約)", "status": "Confirmed", "total_budget": 5000000, "create_at": str(date.today())})
    st.session_state.db["order_items"].append({"order_id": o1_id, "media": "全家廣播", "region": "全省", "start": date(2025,1,1), "end": date(2025,1,31), "sec": 30, "schedule": make_sched(960, 31), "budget": 5000000})

    o2_id = "P1"
    st.session_state.db["orders"].append({"id": o2_id, "sales": "Alice", "client": "三星 S26", "status": "Probable", "total_budget": 3000000, "create_at": str(date.today())})
    st.session_state.db["order_items"].append({"order_id": o2_id, "media": "全家廣播", "region": "全省", "start": date(2025,1,15), "end": date(2025,1,25), "sec": 30, "schedule": make_sched(240, 11), "budget": 3000000})

if 'user_role' not in st.session_state: st.session_state.user_role = None
if 'ops_view_target' not in st.session_state: st.session_state.ops_view_target = None

# ==========================================
# 2. 核心運算
# ==========================================

def get_occupied_inventory(media, target_date, region, include_probable=True):
    occupied = 0
    orders_map = {o['id']: o for o in st.session_state.db['orders']}
    for item in st.session_state.db['order_items']:
        if item['media'] != media: continue
        order = orders_map.get(item['order_id'])
        if not order: continue
        
        # KEY LOGIC CHANGE: Probable also counts as "Occupied" (Inventory Blocked)
        if order['status'] == 'Confirmed': pass
        elif order['status'] == 'Probable' and include_probable: pass
        else: continue 
        
        if item['start'] <= target_date <= item['end']:
            is_impacting = False
            if item['region'] == "全省": is_impacting = True 
            elif item['region'] == region: is_impacting = True
            
            if is_impacting:
                day_idx = (target_date - item['start']).days
                if 0 <= day_idx < len(item['schedule']):
                    occupied += item['schedule'][day_idx] * item['sec']
    return occupied

# ==========================================
# 3. HTML 生成
# ==========================================

def get_common_headers(start_date, end_date):
    date_range = pd.date_range(start_date, end_date)
    weekdays_zh = ['一','二','三','四','五','六','日']
    
    row1 = "<tr class='row-month'><th>媒體 / 區域 / 項目</th>"
    curr_m = None
    month_counts = {}
    for d in date_range: month_counts[(d.year, d.month)] = month_counts.get((d.year, d.month), 0) + 1
    for d in date_range:
        if (d.year, d.month) != curr_m:
            row1 += f"<th colspan='{month_counts[(d.year, d.month)]}'>{d.year}年 {d.month}月</th>"
            curr_m = (d.year, d.month)
    row1 += "</tr>"
    
    row2 = "<tr class='row-day'><th>日期</th>"
    for d in date_range:
        wd = d.weekday()
        style = "background:#e3f2fd;color:#1565c0;" if wd >= 5 else "" 
        row2 += f"<th style='{style}'>{d.day}<br><span style='font-size:9px;font-weight:normal;'>{weekdays_zh[wd]}</span></th>"
    row2 += "</tr>"
    return row1 + row2, date_range

def generate_global_inventory_html(start_date, end_date):
    date_range = pd.date_range(start_date, end_date)
    headers_html, _ = get_common_headers(start_date, end_date)
    body = ""
    structure = { "全家廣播": ["全省"] + REGIONS_ORDER, "新鮮視": ["全省"] + REGIONS_ORDER, "家樂福": ["量販", "超市"] }
    
    for media, regions in structure.items():
        for reg in regions:
            row_html = f"<td>{media}<br><span style='font-size:10px;color:#666'>{reg}</span></td>"
            if media == "家樂福":
                for _ in date_range: row_html += "<td class='inv-safe'>充足</td>"
            else:
                capacity = DAILY_CAPACITY.get(media, 5000)
                hard_limit = capacity * CAPACITY_LIMITS.get(media, 1.0)
                for d in date_range:
                    used = get_occupied_inventory(media, d.date(), reg)
                    pct = used / capacity
                    bg = "inv-safe"
                    if used > hard_limit: bg = "inv-crit"
                    elif pct > 1.0: bg = "inv-high"
                    elif pct > 0.8: bg = "inv-mid"
                    content = f"<div class='val-container'><span class='val-sec'>{used:,}</span><br><span class='val-pct'>{int(pct*100)}%</span></div>"
                    row_html += f"<td class='{bg}'>{content}</td>"
            body += f"<tr>{row_html}</tr>"
    return f"<div class='unified-wrapper' style='height:300px;'><table class='unified-table'><thead>{headers_html}</thead><tbody>{body}</tbody></table></div>"

# ==========================================
# 4. 業務端 (Sales Portal)
# ==========================================

def render_sales_portal():
    st.markdown("### 📝 業務報價與下單系統")
    
    with st.container():
        c1, c2, c3, c4 = st.columns(4)
        client_name = c1.text_input("客戶名稱", "萬國通路")
        start_date = c2.date_input("開始日", date(2025, 1, 15))
        end_date = c3.date_input("結束日", date(2025, 1, 31), min_value=start_date)
        total_budget_input = c4.number_input("總預算", value=1000000, step=10000)
        days_count = (end_date - start_date).days + 1

    with st.expander("📊 點擊展開：全通路庫存戰情摘要"):
        st.components.v1.html(generate_global_inventory_html(start_date, end_date), height=320, scrolling=True)

    st.divider()

    col_input, col_risk = st.columns([7, 3])
    config_media = {}
    remaining_global_share = 100
    ops_sync_items = []
    final_rows = []
    total_list_price_accum = 0
    all_secs = set()
    
    with col_input:
        st.markdown("#### 2. 媒體配置")
        tabs = st.tabs(["📻 全家廣播", "📺 新鮮視", "🛒 家樂福"])
        
        # Tab 1: 全家
        with tabs[0]:
            fm_act = st.checkbox("加入全家廣播", value=True, key="fm_act")
            if fm_act:
                c_a, c_b = st.columns(2)
                is_nat = c_a.checkbox("全省聯播", value=True, key="fm_nat")
                regs = ["全省"] if is_nat else c_b.multiselect("區域", REGIONS_ORDER, key="fm_reg")
                _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="fm_sec")
                secs = sorted(_secs_input)
                share = st.slider("預算佔比%", 0, remaining_global_share, min(70, remaining_global_share), key="fm_share")
                remaining_global_share -= share
                
                # Ratio Slider
                sec_shares = {}
                if len(secs) > 1:
                    st.caption("多秒數佔比分配")
                    ls = 100
                    cols_s = st.columns(len(secs))
                    for i, s in enumerate(secs[:-1]):
                        val = cols_s[i].slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"fm_s_{s}")
                        sec_shares[s] = val
                        ls -= val
                    sec_shares[secs[-1]] = ls
                    cols_s[-1].info(f"{secs[-1]}秒: {ls}%")
                elif secs: sec_shares[secs[0]] = 100
                
                config_media["全家廣播"] = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

        # Tab 2: 新鮮視
        with tabs[1]:
            fv_act = st.checkbox("加入新鮮視", value=True, key="fv_act")
            if fv_act:
                c_a, c_b = st.columns(2)
                is_nat = c_a.checkbox("全省聯播 ", value=False, key="fv_nat")
                regs = ["全省"] if is_nat else c_b.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
                _secs_input = st.multiselect("秒數", DURATIONS, default=[10], key="fv_sec")
                secs = sorted(_secs_input)
                
                if remaining_global_share > 0:
                    share = st.slider("預算佔比% ", 0, remaining_global_share, min(30, remaining_global_share), key="fv_share")
                else:
                    st.caption("⚠️ 前方通路已佔用 100% 預算")
                    share = 0
                remaining_global_share -= share
                
                sec_shares = {}
                if len(secs) > 1:
                    st.caption("多秒數佔比分配")
                    ls = 100
                    cols_s = st.columns(len(secs))
                    for i, s in enumerate(secs[:-1]):
                        val = cols_s[i].slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"fv_s_{s}")
                        sec_shares[s] = val
                        ls -= val
                    sec_shares[secs[-1]] = ls
                    cols_s[-1].info(f"{secs[-1]}秒: {ls}%")
                elif secs: sec_shares[secs[0]] = 100
                
                config_media["新鮮視"] = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

        # Tab 3: 家樂福
        with tabs[2]:
            cf_act = st.checkbox("加入家樂福", value=True, key="cf_act")
            if cf_act:
                st.info("家樂福庫存無限，請盡情配置")
                _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="cf_sec")
                secs = sorted(_secs_input)
                if remaining_global_share > 0:
                    share = st.slider("預算佔比  ", 0, remaining_global_share, remaining_global_share, key="cf_share")
                else:
                    st.caption("⚠️ 前方通路已佔用 100% 預算")
                    share = 0
                
                sec_shares = {}
                if len(secs) > 1:
                    st.caption("多秒數佔比分配")
                    ls = 100
                    cols_s = st.columns(len(secs))
                    for i, s in enumerate(secs[:-1]):
                        val = cols_s[i].slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"cf_s_{s}")
                        sec_shares[s] = val
                        ls -= val
                    sec_shares[secs[-1]] = ls
                    cols_s[-1].info(f"{secs[-1]}秒: {ls}%")
                elif secs: sec_shares[secs[0]] = 100
                
                config_media["家樂福"] = {"regions": ["全省"], "seconds": secs, "share": share, "sec_shares": sec_shares}

    # --- 計算 ---
    risk_summary = {"crit": 0, "warn": 0, "safe": 0, "msgs": []}
    
    if sum(m["share"] for m in config_media.values()) > 0:
        for m_type, cfg in config_media.items():
            media_budget = total_budget_input * (cfg["share"] / 100.0)
            for sec, sec_share in cfg["sec_shares"].items():
                sec_budget = media_budget * (sec_share / 100.0)
                if sec_budget <= 0: continue
                all_secs.add(sec)
                factor = get_sec_factor(m_type, sec)
                
                if m_type in ["全家廣播", "新鮮視"]:
                    db = PRICING_DB[m_type]
                    std_spots = db["Std_Spots"]
                    calc_regions = ["全省"] if cfg["is_national"] else cfg["regions"]
                    if not calc_regions: continue
                    temp_net = sum([db[r][1] for r in calc_regions]) / std_spots * factor
                    if temp_net == 0: continue
                    target_spots = math.ceil(sec_budget / temp_net)
                    if target_spots % 2 != 0: target_spots += 1
                    daily_sch = calculate_schedule(target_spots, days_count)
                    pkg_cost = (db["全省"][0] / std_spots) * target_spots * factor if cfg["is_national"] else 0
                    display_regions = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                    
                    for reg in display_regions:
                        limit = DAILY_CAPACITY.get(m_type, 5000)
                        limit_ratio = CAPACITY_LIMITS.get(m_type, 1.0)
                        hard_limit = int(limit * limit_ratio)
                        inv_status_list = []
                        curr = start_date
                        max_pct = 0
                        for i in range(days_count):
                            occupied = get_occupied_inventory(m_type, curr, reg)
                            new_sec = daily_sch[i] * sec
                            final_used = occupied + new_sec
                            pct = final_used / limit
                            if pct > max_pct: max_pct = pct
                            status = "cell-ok"
                            if final_used > hard_limit: status = "cell-err"
                            elif final_used > limit: status = "cell-warn"
                            inv_status_list.append({"status": status, "remaining": int(hard_limit - final_used)})
                            curr += timedelta(days=1)
                        
                        if max_pct > limit_ratio:
                            risk_summary["crit"] += 1
                            risk_summary["msgs"].append(f"🔴 {m_type}-{reg} 爆量 ({int(max_pct*100)}%)")
                        elif max_pct > 1.0:
                            risk_summary["warn"] += 1
                            risk_summary["msgs"].append(f"🟡 {m_type}-{reg} 緊張 ({int(max_pct*100)}%)")
                        else: risk_summary["safe"] += 1

                        reg_list = db.get(reg, [0,0])[0] if cfg["is_national"] else db[reg][0]
                        rate_list = int(round((reg_list / std_spots) * target_spots * factor))
                        pkg_display = int(round(pkg_cost)) if cfg["is_national"] else rate_list
                        if not cfg["is_national"] or (cfg["is_national"] and reg == "北區"): total_list_price_accum += pkg_display
                        prog = STORE_COUNTS.get(reg, reg)
                        if m_type == "新鮮視": prog = STORE_COUNTS.get(f"新鮮視_{reg}", reg)
                        
                        final_rows.append({
                            "media": m_type, "region": reg, "program": prog, "daypart": db["Day_Part"], "seconds": sec,
                            "schedule": daily_sch, "spots": target_spots, "rate_list": rate_list, "pkg_display_val": pkg_display,
                            "inv_status": inv_status_list, "is_pkg_start": (cfg["is_national"] and reg == "北區"), "is_pkg_member": cfg["is_national"]
                        })
                        ops_sync_items.append({
                            "media": m_type, "region": reg, "schedule": daily_sch, "sec": sec, "budget": sec_budget,
                            "start": start_date, "end": end_date
                        })

                elif m_type == "家樂福":
                    db = PRICING_DB["家樂福"]
                    u_net = (db["量販_全省"]["Net"] + db["超市_全省"]["Net"]) / db["量販_全省"]["Std_Spots"] * factor
                    target_spots = math.ceil(sec_budget / u_net)
                    if target_spots % 2 != 0: target_spots += 1
                    daily_sch = calculate_schedule(target_spots, days_count)
                    rate_h = int(round((db["量販_全省"]["List"]/db["量販_全省"]["Std_Spots"])*target_spots*factor))
                    rate_s = int(round((db["超市_全省"]["List"]/db["超市_全省"]["Std_Spots"])*target_spots*factor))
                    total_list_price_accum += (rate_h + rate_s)
                    fake_status = [{"status": "cell-ok", "remaining": 9999} for _ in daily_sch]
                    ops_sync_items.append({
                        "media": "家樂福", "region": "全省", "sec": sec, "schedule": daily_sch, 
                        "start": start_date, "end": end_date, "budget": sec_budget
                    })
                    final_rows.append({"media": "家樂福", "region": "全省量販", "program": STORE_COUNTS["家樂福_量販"], "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "schedule": daily_sch, "spots": target_spots, "rate_list": rate_h, "pkg_display_val": rate_h, "inv_status": fake_status, "is_pkg_start": False, "is_pkg_member": False})
                    final_rows.append({"media": "家樂福", "region": "全省超市", "program": STORE_COUNTS["家樂福_超市"], "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "schedule": daily_sch, "spots": target_spots, "rate_list": rate_s, "pkg_display_val": rate_s, "inv_status": fake_status, "is_pkg_start": False, "is_pkg_member": False})
                    risk_summary["safe"] += 1

    with col_risk:
        st.markdown(f"""
        <div class="risk-panel">
            <h4 style="border-bottom:2px solid #eee; padding-bottom:10px;">🛡️ 風險監控</h4>
            <div class="kpi-box"><div class="kpi-label">🔴 爆量項目 (禁止)</div><div class="kpi-value risk-score-crit">{risk_summary['crit']} 項</div></div>
            <div class="kpi-box"><div class="kpi-label">🟡 緊張項目 (需審慎)</div><div class="kpi-value risk-score-warn">{risk_summary['warn']} 項</div></div>
            <div class="kpi-box" style="border:none;"><div class="kpi-label">異常明細</div><div style="font-size:12px; color:#555; max-height:150px; overflow-y:auto;">{'<br>'.join(risk_summary['msgs']) if risk_summary['msgs'] else '無異常，可送審'}</div></div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("### 3. 排程預覽與送審")
    if final_rows:
        html = generate_smart_cue_sheet(final_rows, days_count, start_date, client_name)
        st.components.v1.html(html, height=600, scrolling=True)
        c_sub1, c_sub2 = st.columns([1, 4])
        with c_sub1:
            btn_disabled = risk_summary['crit'] > 0
            if btn_disabled: st.error("🔴 存在爆量項目，無法送審")
            else:
                if st.button("🚀 送出審核", type="primary"):
                    user = st.session_state.get('user_name', 'Sales')
                    new_order_id = str(uuid.uuid4())[:8]
                    st.session_state.db['orders'].append({"id": new_order_id, "sales": user, "client": client_name, "status": "Pending", "total_budget": total_budget_input, "create_at": str(date.today())})
                    for item in ops_sync_items:
                        item['order_id'] = new_order_id
                        st.session_state.db['order_items'].append(item)
                    st.success("✅ 訂單已送出！")
        with c_sub2:
            final_rows.sort(key=lambda x: MEDIA_ORDER_MAP.get(x['media'], 99))
            sorted_secs_list = sorted(list(all_secs))
            product_str = "、".join([f"{s}秒" for s in sorted_secs_list])
            xlsx_data = generate_excel(final_rows, days_count, start_date, client_name, product_str, total_list_price_accum, 0, total_budget_input, 10000)
            st.download_button("📥 下載 Excel", data=xlsx_data.getvalue(), file_name=f"Cue_{client_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def generate_smart_cue_sheet(rows, days_cnt, start_dt, c_name):
    date_header_row1 = f"<th class='header-blue' colspan='{days_cnt}'>{start_dt.month}月</th>"
    date_header_row2 = "".join([f"<th class='{'header-yellow' if (start_dt+timedelta(days=i)).weekday()>=5 else 'header-blue'}'>{(start_dt+timedelta(days=i)).day}</th>" for i in range(days_cnt)])
    date_header_row3 = "".join([f"<th class='{'header-yellow' if (start_dt+timedelta(days=i)).weekday()>=5 else 'header-blue'}'>{['一','二','三','四','五','六','日'][(start_dt+timedelta(days=i)).weekday()]}</th>" for i in range(days_cnt)])
    rows_html = ""
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]['media'] == row['media'] and rows[j]['seconds'] == row['seconds']: j += 1
        group_size = j - i
        for k in range(group_size):
            r_data = rows[i+k]
            tr = "<tr>"
            if k == 0: tr += f"<td rowspan='{group_size*2}'>{r_data['media']}<br>{r_data['seconds']}秒</td>"
            tr += f"<td>{r_data['region']}<br><span style='font-size:9px;color:#666'>{r_data['program']}</span></td>"
            for idx, val in enumerate(r_data['schedule']):
                status_cls = r_data['inv_status'][idx]['status']
                tr += f"<td class='{status_cls}'>{val}</td>"
            tr += "</tr>"
            tr_inv = "<tr class='row-avail'><td>剩餘</td>"
            for inv in r_data['inv_status']:
                val = inv['remaining']
                color = "red" if val < 0 else "#666"
                tr_inv += f"<td style='color:{color}'>{val}</td>"
            tr_inv += "</tr>"
            rows_html += tr + tr_inv
        i = j
    return f"<div style='overflow-x:auto;width:100%;'><table class='preview-table'><tr><th colspan='4' class='header-blue'>基本資料</th>{date_header_row1}</tr><tr><th rowspan='2' class='header-blue'>媒體</th><th rowspan='2' class='header-blue'>區域</th>{date_header_row2}</tr><tr>{date_header_row3}</tr>{rows_html}</table></div>"

def generate_excel(rows, days_cnt, start_dt, c_name, products, total_list, grand_total, budget, prod):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Media Schedule")
    worksheet.write(0, 0, "Excel Download Ready")
    workbook.close()
    return output

# ==========================================
# 5. 營運端邏輯 (Ops Portal)
# ==========================================

def render_ops_dashboard():
    st.markdown("### 🛡️ 營運戰情總控台")
    
    if st.session_state.ops_view_target:
        t = st.session_state.ops_view_target
        st.session_state.ops_view_target = None 
        items = [i for i in st.session_state.db['order_items'] if i['order_id'] == t['id']]
        if items:
            target_media = items[0]['media']
            target_start = items[0]['start']
            target_end = items[0]['end']
            
            if 'ops_media_select' not in st.session_state or st.session_state.ops_media_select != target_media:
                st.session_state.ops_media_select = target_media
            st.session_state.ops_d_range = (target_start, target_end)
            st.toast(f"已切換至訂單: {t['client']}")
            st.rerun()

    c1, c2, c3, c4 = st.columns([1.5, 1.5, 2, 1])
    with c1: media_select = st.selectbox("1. 媒體頻道", ["全家廣播", "新鮮視", "家樂福"], key="ops_media_select")
    with c2: reg_opts = ["All"] + list(QUOTA_CONFIG[media_select]['limits'].keys()); region_select = st.selectbox("2. 對帳區域", reg_opts)
    with c3: d_range = st.date_input("3. 檢視區間", value=(date(2025, 1, 1), date(2025, 2, 10)), key="ops_d_range")
    with c4: st.write(""); st.write(""); show_sim = st.toggle("顯示待審 (模擬)", value=True)

    if not (isinstance(d_range, tuple) and len(d_range) == 2): return
    start_d, end_d = d_range
    capacity = DAILY_CAPACITY.get(media_select, 5000)
    date_range = pd.date_range(start_d, end_d)
    
    display_regions = ["全省"] + REGIONS_ORDER if media_select == "全家廣播" else list(QUOTA_CONFIG[media_select]['limits'].keys())
    if region_select != "All": display_regions = [region_select]
    matrix = {r: {d.strftime("%Y-%m-%d"): {"used": 0, "pending": 0} for d in date_range} for r in display_regions}
    
    orders_to_display = {}
    orders_map = {o['id']: o for o in st.session_state.db['orders']}
    
    for item in st.session_state.db['order_items']:
        if item['media'] != media_select: continue
        order = orders_map.get(item['order_id'])
        if not order: continue
        
        status = order['status']
        if status == 'Pending' and not show_sim: continue
        
        s = max(item['start'], start_d)
        e = min(item['end'], end_d)
        if s <= e:
            # Impact Logic
            affected_regions = []
            if item['region'] == '全省':
                affected_regions = list(QUOTA_CONFIG[media_select]['limits'].keys())
                if '全省' not in affected_regions: affected_regions.append('全省')
            else:
                affected_regions = [item['region']]
            
            is_relevant = False
            if region_select == "All": is_relevant = True
            elif region_select in affected_regions: is_relevant = True
            
            if is_relevant:
                if order['id'] not in orders_to_display:
                    orders_to_display[order['id']] = { "client": order['client'], "sales": order['sales'], "status": status, "schedule_data": {} }
            
            curr = s
            while curr <= e:
                d_str = curr.strftime("%Y-%m-%d")
                day_idx = (curr - item['start']).days
                val = 0
                if 0 <= day_idx < len(item['schedule']):
                    val = item['schedule'][day_idx] * item['sec']
                
                # Logic Fix: Use MAX for Project Row Display
                if is_relevant and order['id'] in orders_to_display:
                    if d_str not in orders_to_display[order['id']]['schedule_data']:
                        orders_to_display[order['id']]['schedule_data'][d_str] = 0
                    current = orders_to_display[order['id']]['schedule_data'][d_str]
                    orders_to_display[order['id']]['schedule_data'][d_str] = max(current, val)
                
                # Update Matrix (Summation Logic)
                for r in affected_regions:
                    if r in matrix and d_str in matrix[r]:
                        k = "pending" if status == 'Pending' else "used"
                        matrix[r][d_str][k] += val
                curr += timedelta(days=1)

    headers_html, _ = get_common_headers(start_d, end_d)
    body_html = ""
    sorted_orders = sorted(orders_to_display.values(), key=lambda x: {"Pending": 0, "Probable": 1, "Confirmed": 2}.get(x['status'], 9))
    
    # Status badges map
    status_badges = {
        "Pending": "<span class='badge badge-pend'>待審</span>",
        "Probable": "<span class='badge badge-prob'>80% 卡位</span>",
        "Confirmed": "<span class='badge badge-conf'>正式簽約</span>"
    }

    for o in sorted_orders:
        if o['status'] == "Pending": cls = "proj-pend"
        elif o['status'] == "Probable": cls = "proj-prob"
        else: cls = "proj-conf"
        
        # New Left Column Design
        badge = status_badges.get(o['status'], "")
        row_cells = f"""
        <td>
            <div style='font-weight:bold;color:#333;margin-bottom:4px;font-size:12px;'>{o['client']}</div>
            <div style='display:flex;justify-content:space-between;align-items:center;'>
                <span style='font-size:11px;color:#555;'>{o['sales']}</span>
                {badge}
            </div>
        </td>"""
        
        for d in date_range:
            d_str = d.strftime("%Y-%m-%d")
            val = o['schedule_data'].get(d_str, 0)
            row_cells += f"<td class='{cls}'>{val:,}</td>" if val > 0 else "<td></td>"
        body_html += f"<tr>{row_cells}</tr>"
        
    reg_txt = f"【{region_select}】" if region_select != "All" else "【所有區域】"
    body_html += f"<tr style='background:#eceff1;font-weight:bold;'><td colspan='{len(date_range)+1}'>∑ 庫存水位匯總 {reg_txt}</td></tr>"
    
    # Feature: Max Load Row
    if region_select == "All":
        max_load_row = f"<td>📊 全省 (最大負荷)<br><span style='font-size:9px;color:#888'>Max Load</span></td>"
        for d in date_range:
            d_str = d.strftime("%Y-%m-%d")
            max_used = 0
            for r in display_regions:
                cell = matrix[r][d_str]
                total = cell['used'] + cell['pending']
                if total > max_used: max_used = total
            
            pct = max_used / capacity
            bg = "inv-safe"
            if max_used > (capacity * 1.2): bg = "inv-crit"
            elif pct > 1.0: bg = "inv-high"
            elif pct > 0.8: bg = "inv-mid"
            
            content = f"<div class='val-container'><span class='val-sec'>{max_used:,}</span><br><span class='val-pct'>{int(pct*100)}%</span></div>"
            max_load_row += f"<td class='{bg}'>{content}</td>"
        body_html += f"<tr>{max_load_row}</tr>"
    
    for r in matrix:
        row_cells = f"<td>📈 {r}<br><span style='font-size:9px;color:#888'>Limit {capacity}</span></td>"
        limit_ratio = CAPACITY_LIMITS.get(media_select, 1.0)
        hard_limit = int(capacity * limit_ratio)
        for d in date_range:
            d_str = d.strftime("%Y-%m-%d")
            cell = matrix[r][d_str]
            total = cell['used'] + cell['pending']
            pct = total / capacity
            
            bg = "inv-safe"
            if total > hard_limit: bg = "inv-crit"
            elif pct > 1.0: bg = "inv-high"
            elif pct > 0.8: bg = "inv-mid"
            
            sim_cls = "inv-sim" if cell['pending'] > 0 and pct > 1.0 else ""
            content = f"<div class='val-container'><span class='val-sec'>{total:,}</span><br><span class='val-pct'>{int(pct*100)}%</span></div>"
            row_cells += f"<td class='{bg} {sim_cls}'>{content}</td>"
        body_html += f"<tr>{row_cells}</tr>"

    st.markdown(f"<div class='unified-wrapper'><table class='unified-table'><thead>{headers_html}</thead><tbody>{body_html}</tbody></table></div>", unsafe_allow_html=True)

    st.markdown("### ⚡ 待審核案件")
    pending_orders = [o for o in st.session_state.db['orders'] if o['status'] == 'Pending']
    
    col_pend, col_prob = st.tabs(["待審核 (Pending)", "已卡位 (Probable)"])
    
    with col_pend:
        if not pending_orders: st.info("無待審案件")
        else:
            for o in pending_orders:
                items = [i for i in st.session_state.db['order_items'] if i['order_id'] == o['id']]
                media_txt = list(set([i['media'] for i in items]))[0] if items else ""
                budget_val = o.get('total_budget', o.get('budget', 0))
                with st.container():
                    st.markdown(f"""
                    <div class='approval-card'>
                        <h4 style='margin:0;color:#e65100;'>🟠 {o['client']}</h4>
                        <p style='margin:5px 0;font-size:12px;'>業務: {o['sales']} | 媒體: {media_txt} | 預算: {budget_val:,}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    c1, c2, c3 = st.columns([1, 1, 1])
                    if c1.button("🔍 檢視", key=f"view_{o['id']}"):
                        st.session_state.ops_view_target = o
                        st.rerun()
                    if c2.button("✅ 同意卡位 (80%)", key=f"app_{o['id']}"):
                        o['status'] = "Probable"
                        st.rerun()
                    if c3.button("❌ 駁回", key=f"rej_{o['id']}"):
                        st.session_state.db['orders'].remove(o)
                        st.rerun()

    with col_prob:
        prob_orders = [o for o in st.session_state.db['orders'] if o['status'] == 'Probable']
        if not prob_orders: st.info("無已卡位案件")
        else:
            for o in prob_orders:
                items = [i for i in st.session_state.db['order_items'] if i['order_id'] == o['id']]
                media_txt = list(set([i['media'] for i in items]))[0] if items else ""
                budget_val = o.get('total_budget', o.get('budget', 0))
                with st.container():
                    st.markdown(f"""
                    <div class='approval-card' style='border-left: 4px solid #7b1fa2;'>
                        <h4 style='margin:0;color:#7b1fa2;'>🟣 {o['client']} (已卡位)</h4>
                        <p style='margin:5px 0;font-size:12px;'>業務: {o['sales']} | 媒體: {media_txt} | 預算: {budget_val:,}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button("📝 正式簽約 (Confirmed)", key=f"sign_{o['id']}"):
                        o['status'] = "Confirmed"
                        st.rerun()

def main():
    with st.sidebar:
        st.title("🔐 系統登入")
        role = st.radio("選擇身份", ["業務人員 (Sales)", "營運主管 (Ops)"])
        if role == "業務人員 (Sales)": st.session_state.user_role = "Sales"; st.session_state.user_name = "Andy"
        else: st.session_state.user_role = "Ops"; st.session_state.user_name = "Director"

    if st.session_state.user_role == "Sales": render_sales_portal()
    else: render_ops_dashboard()

if __name__ == "__main__":
    main()
