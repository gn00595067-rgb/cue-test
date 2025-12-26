# -*- coding: utf-8 -*-
import streamlit as st
import math
import io
from datetime import timedelta, datetime, date

import xlsxwriter

# PDF (HTML -> PDF)
try:
    from weasyprint import HTML
    WEASY_OK = True
except Exception:
    WEASY_OK = False


# ==========================================
# 1) 基礎資料與設定 (2026 新制)
# ==========================================

STORE_COUNTS_INT = {
    "全省": 4437,
    "北區": 1649, "桃竹苗": 779, "中區": 839, "雲嘉南": 499, "高屏": 490, "東區": 181,

    "新鮮視_全省": 3124,
    "新鮮視_北區": 1127, "新鮮視_桃竹苗": 616, "新鮮視_中區": 528,
    "新鮮視_雲嘉南": 365, "新鮮視_高屏": 405, "新鮮視_東區": 83,

    "家樂福_量販": 68,
    "家樂福_超市": 249,
}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# List=牌價(顯示用), Net=實收(運算用)
PRICING_DB = {
    "全家廣播": {
        "Std_Spots": 480,
        "Day_Part": "00:00-24:00",
        "全省":   [400000, 320000],
        "北區":   [250000, 200000],
        "桃竹苗": [150000, 120000],
        "中區":   [150000, 120000],
        "雲嘉南": [100000,  80000],
        "高屏":   [100000,  80000],
        "東區":   [ 62500,  50000],
    },
    "新鮮視": {
        "Std_Spots": 504,
        "Day_Part": "07:00-22:00",
        "全省":   [150000, 120000],
        "北區":   [150000, 120000],
        "桃竹苗": [120000,  96000],
        "中區":   [ 90000,  72000],
        "雲嘉南": [ 75000,  60000],
        "高屏":   [ 75000,  60000],
        "東區":   [ 45000,  36000],
    },
    "家樂福": {
        "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420, "Day_Part": "09:00-23:00"},
        "超市_全省": {"List": 100000, "Net":  80000, "Std_Spots": 720, "Day_Part": "00:00-24:00"},
    }
}

SEC_FACTORS = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5, 5: 0.25},
    "新鮮視":   {30: 3.0, 20: 2.0,  15: 1.5,  10: 1.0, 5: 0.5},
    "家樂福":   {30: 1.5, 20: 1.0,  15: 0.85, 10: 0.65, 5: 0.35}
}

def get_sec_factor(media_type, seconds):
    # 跟你範例一致：沒有對應就 1.0
    return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    """
    Even Distribution Strategy (跟你範例一致)
    - 強制偶數
    - 先除2平均分，再乘2
    """
    if days <= 0:
        return []
    if total_spots % 2 != 0:
        total_spots += 1
    half = total_spots // 2
    base, rem = divmod(half, days)
    arr = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in arr]

def fmt_int(v):
    try:
        if isinstance(v, (int, float)):
            return f"{int(round(v)):,}"
        return str(v)
    except:
        return str(v)

def region_display(region):
    # 你示例的顯示
    mp = {
        "北區": "北區-北北基",
        "桃竹苗": "桃區-桃竹苗",
        "中區": "中區-中彰投",
        "雲嘉南": "雲嘉南區-雲嘉南",
        "高屏": "高屏區-高屏",
        "東區": "東區-宜花東",
        "全省量販": "全省量販",
        "全省超市": "全省超市",
    }
    return mp.get(region, region)

def station_name(media):
    if media == "全家廣播":
        return "全家便利商店<br>通路廣播廣告"
    if media == "新鮮視":
        return "全家便利商店<br>新鮮視廣告"
    return "家樂福"


# ==========================================
# 2) 核心計算（完全對齊你範例的 group + merge 概念）
# ==========================================

def calc_radio_or_fv_rows(media, cfg, sec, sec_budget, days_count):
    """
    回傳 rows (每分區一列) + group_bundle_pkg (全省時才有) + group_penalty
    你最新規則：
    - Rate(Net) 顯示「分區總價」(List 算) = unit_list * spots
    - Package-cost(Net):
        * 全省：表格中只顯示一次（合併格），數值用「全省 List」算打包價，未達標 *1.1
              分區列不要 *1.1（避免太怪）
        * 非全省：每列 Package = Rate * (1.1 if penalty)
    """
    db = PRICING_DB[media]
    std = db["Std_Spots"]
    factor = get_sec_factor(media, sec)
    day_part = db["Day_Part"]

    is_national = bool(cfg.get("is_national", False))
    selected_regions = cfg.get("regions", [])

    if is_national:
        calc_regions = ["全省"]         # 內部算檔次只用全省 Net（跟你範例一致）
        display_regions = REGIONS_ORDER[:]  # 但表格一定展開 6 區
    else:
        calc_regions = selected_regions[:]  # 內部算檔次用選到的區
        display_regions = selected_regions[:]  # 表格只顯示選到的區

    if not display_regions:
        return [], None, False, None  # 沒選分區就不產生

    # Step A: Net 算 spots
    unit_net_sum = 0.0
    for reg in calc_regions:
        net_price = db[reg][1]
        unit_net_sum += (net_price / std) * factor

    if unit_net_sum <= 0:
        return [], None, False, None

    spots_init = math.ceil(sec_budget / unit_net_sum)
    penalty = spots_init < std
    mult = 1.1 if penalty else 1.0

    spots = math.ceil(sec_budget / (unit_net_sum * mult))
    if spots % 2 != 0:
        spots += 1
    if spots <= 0:
        spots = 2

    sch = calculate_schedule(spots, days_count)

    # Step B: List 填表
    rows = []
    bundle_pkg = None

    # 全省打包價（只在全省時顯示一次，且可 *1.1）
    if is_national:
        nat_list = db["全省"][0]
        unit_list_nat = int((nat_list / std) * factor)
        bundle_pkg = unit_list_nat * spots
        if penalty:
            bundle_pkg = int(round(bundle_pkg * 1.1))

    for idx, reg in enumerate(display_regions):
        # 分區顯示用 List
        reg_list = db[reg][0]  # 全省情境下也要用分區 List（你規格書的欺敵邏輯）
        unit_list_reg = int((reg_list / std) * factor)

        # ✅ Rate(Net) = 分區總價
        rate_total = unit_list_reg * spots

        # ✅ Package-cost(Net)
        if is_national:
            # 全省時：分區列不要 *1.1（避免價差過度）
            pkg_display = None  # 讓 renderer 用 merge cell 顯示 bundle_pkg
        else:
            pkg_display = int(round(rate_total * (1.1 if penalty else 1.0)))

        # program
        if media == "新鮮視":
            prog = STORE_COUNTS_INT.get(f"新鮮視_{reg}", 0)
        else:
            prog = STORE_COUNTS_INT.get(reg, 0)

        rows.append({
            "media": media,
            "seconds": sec,
            "region": reg,
            "location": region_display(reg),
            "program": prog,
            "daypart": day_part,
            "schedule": sch,
            "spots": spots,

            "rate_total": rate_total,      # G 欄（總價）
            "pkg_display": pkg_display,    # H 欄（非全省時逐列顯示；全省時 None）
            "is_national": is_national,
            "penalty": penalty,

            # renderer 用：全省時在該 group 第1列放合併格
            "is_pkg_start": (is_national and idx == 0),
            "is_pkg_member": is_national,
            "bundle_pkg": bundle_pkg,
        })

    return rows, bundle_pkg, penalty, spots


def calc_carrefour_rows(cfg, sec, sec_budget, days_count):
    """
    家樂福：你範例做法
    - spots 用量販 Net/Std 算（未達標 *1.1）
    - 量販列顯示 List -> Rate & Package
    - 超市列顯示「計量販」
    """
    db = PRICING_DB["家樂福"]
    factor = get_sec_factor("家樂福", sec)

    base = db["量販_全省"]
    sup  = db["超市_全省"]

    std_base = base["Std_Spots"]
    std_sup  = sup["Std_Spots"]

    unit_net = (base["Net"] / std_base) * factor
    if unit_net <= 0:
        return [], None, False, None

    spots_init = math.ceil(sec_budget / unit_net)
    penalty = spots_init < std_base
    mult = 1.1 if penalty else 1.0

    spots = math.ceil(sec_budget / (unit_net * mult))
    if spots % 2 != 0:
        spots += 1
    if spots <= 0:
        spots = 2

    sch_base = calculate_schedule(spots, days_count)

    unit_list_base = int((base["List"] / std_base) * factor)
    rate_total_base = unit_list_base * spots
    pkg_total_base  = int(round(rate_total_base * (1.1 if penalty else 1.0)))  # 家樂福照你範例：逐列顯示

    rows = [{
        "media": "家樂福",
        "seconds": sec,
        "region": "全省量販",
        "location": "全省量販",
        "program": STORE_COUNTS_INT.get("家樂福_量販", 0),
        "daypart": base["Day_Part"],
        "schedule": sch_base,
        "spots": spots,
        "rate_total": rate_total_base,
        "pkg_display": pkg_total_base,
        "is_national": True,
        "penalty": penalty,
        "is_pkg_start": False,
        "is_pkg_member": False,
        "bundle_pkg": None,
    }]

    # 超市 spots 比例換算
    spots_sup = int(round(spots * (std_sup / std_base)))
    if spots_sup % 2 != 0:
        spots_sup += 1
    sch_sup = calculate_schedule(spots_sup, days_count)
    rows.append({
        "media": "家樂福",
        "seconds": sec,
        "region": "全省超市",
        "location": "全省超市",
        "program": STORE_COUNTS_INT.get("家樂福_超市", 0),
        "daypart": sup["Day_Part"],
        "schedule": sch_sup,
        "spots": spots_sup,
        "rate_total": "計量販",
        "pkg_display": "計量販",
        "is_national": True,
        "penalty": penalty,
        "is_pkg_start": False,
        "is_pkg_member": False,
        "bundle_pkg": None,
    })
    return rows, None, penalty, spots


def build_final_rows(config_media, total_budget, days_count):
    final_rows = []
    debug_logs = []

    # 用來算 Total Package：全省 group 的 bundle 只加一次
    bundle_seen = set()
    total_pkg = 0
    total_rate = 0

    for media, cfg in config_media.items():
        media_budget = total_budget * (cfg["share"] / 100.0)
        for sec, sec_share in cfg["sec_shares"].items():
            sec_budget = media_budget * (sec_share / 100.0)
            if sec_budget <= 0:
                continue

            if media in ["全家廣播", "新鮮視"]:
                rows, bundle_pkg, penalty, spots = calc_radio_or_fv_rows(media, cfg, sec, sec_budget, days_count)
                if not rows:
                    continue
                final_rows.extend(rows)

                # debug
                std = PRICING_DB[media]["Std_Spots"]
                debug_logs.append({
                    "media": media, "sec": sec, "budget": sec_budget,
                    "spots": spots, "std": std, "penalty": penalty,
                    "is_national": cfg.get("is_national", False),
                    "reason": "觸發 x1.1" if penalty else "費率正常"
                })

            elif media == "家樂福":
                rows, bundle_pkg, penalty, spots = calc_carrefour_rows(cfg, sec, sec_budget, days_count)
                if not rows:
                    continue
                final_rows.extend(rows)
                debug_logs.append({
                    "media": media, "sec": sec, "budget": sec_budget,
                    "spots": spots, "std": PRICING_DB["家樂福"]["量販_全省"]["Std_Spots"],
                    "penalty": penalty, "is_national": True,
                    "reason": "觸發 x1.1" if penalty else "費率正常"
                })

    # 排序：媒體->秒數->區域
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    region_rank = {r: i for i, r in enumerate(REGIONS_ORDER)}
    def rr(x):
        if x["region"] in region_rank:
            return region_rank[x["region"]]
        if x["region"] == "全省量販":
            return 90
        if x["region"] == "全省超市":
            return 91
        return 999

    final_rows.sort(key=lambda x: (media_order.get(x["media"], 99), x["seconds"], rr(x)))

    # Total Rate / Total Package
    # - Rate：所有 int 的 rate_total 加總
    total_rate = sum([r["rate_total"] for r in final_rows if isinstance(r["rate_total"], int)])

    # - Package：
    #   * 全省 group：bundle_pkg 加一次
    #   * 非全省：逐列 pkg_display 加總
    for r in final_rows:
        if r.get("is_national") and r.get("bundle_pkg") is not None and r.get("is_pkg_start"):
            key = (r["media"], r["seconds"])
            if key not in bundle_seen:
                bundle_seen.add(key)
                total_pkg += int(r["bundle_pkg"])
        elif (not r.get("is_pkg_member")) and isinstance(r.get("pkg_display"), int):
            total_pkg += int(r["pkg_display"])

    return final_rows, debug_logs, total_rate, total_pkg


# ==========================================
# 3) HTML 預覽（rowspan 擬真）
# ==========================================

def generate_html_preview(rows, days_cnt, start_dt, end_dt, client_name, product_str, total_rate, total_pkg, budget, prod_cost):
    used_media = sorted(list(set(r["media"] for r in rows)), key=lambda x: {"全家廣播":1,"新鮮視":2,"家樂福":3}.get(x,99))
    mediums_str = "、".join(used_media)

    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    curr = start_dt
    date_row1 = f"<th class='header-blue' colspan='{days_cnt}'>{start_dt.month}月</th>"
    date_row2 = ""
    date_row3 = ""
    for i in range(days_cnt):
        wd = curr.weekday()
        cls = "header-yellow" if wd >= 5 else "header-blue"
        date_row2 += f"<th class='{cls}'>{curr.day}</th>"
        date_row3 += f"<th class='{cls}'>{weekdays[wd]}</th>"
        curr += timedelta(days=1)

    css = """
    <style>
      .preview-wrap { width:100%; overflow-x:auto; }
      table.preview { border-collapse:collapse; font-family:"Microsoft JhengHei","Arial",sans-serif; font-size:13px; min-width:1200px; background:#fff; }
      th, td { border:1px solid #555; padding:6px 8px; text-align:center; vertical-align:middle; white-space:nowrap; }
      .header-blue { background:#2c3e50; color:#fff; font-weight:bold; }
      .header-yellow { background:#f1c40f; color:#000; font-weight:bold; }
      .cell-yellow { background:#fff3cd; font-weight:bold; }
      .row-total { background:#d4edda; font-weight:bold; }
      .row-grand { background:#ffc107; font-weight:bold; border-top:2px solid #000; }
      .left { text-align:left; white-space:normal; }
      .right { text-align:right; }
      tr:nth-child(even) { background:#f2f2f2; }
      tr:hover { background:#e6f7ff; }
    </style>
    """

    # group by (media, seconds) for rowspan station & (national) package-cost
    data_rows_html = ""
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]["media"] == row["media"] and rows[j]["seconds"] == row["seconds"]:
            j += 1
        group = rows[i:j]
        group_size = len(group)

        station_html = station_name(row["media"])

        for k, r in enumerate(group):
            tr = "<tr>"
            if k == 0:
                tr += f"<td rowspan='{group_size}' class='left'>{station_html}</td>"

            tr += f"<td>{r['location']}</td>"
            tr += f"<td class='right'>{fmt_int(r['program'])}</td>"
            tr += f"<td>{r['daypart']}</td>"
            tr += f"<td>{r['seconds']}秒</td>"

            # Rate(Net) = 總價
            rate_disp = fmt_int(r["rate_total"])
            tr += f"<td class='right'>{rate_disp}</td>"

            # Package-cost(Net)
            if r.get("is_pkg_start") and r.get("bundle_pkg") is not None:
                # 全省：整段合併
                tr += f"<td rowspan='{group_size}' class='right'>{fmt_int(r['bundle_pkg'])}</td>"
            elif r.get("is_pkg_member"):
                # 全省 group 的其他列不畫 package cell
                pass
            else:
                tr += f"<td class='right'>{fmt_int(r.get('pkg_display',''))}</td>"

            # schedule days
            sch = r.get("schedule", [])
            for d in sch[:days_cnt]:
                tr += f"<td>{fmt_int(d)}</td>"

            tr += f"<td class='cell-yellow'>{fmt_int(sum(sch[:days_cnt]))}</td>"
            tr += "</tr>"
            data_rows_html += tr

        i = j

    # total row schedule
    totals = [0]*days_cnt
    for r in rows:
        sch = r.get("schedule", [])
        for idx in range(min(days_cnt, len(sch))):
            totals[idx] += sch[idx]

    vat = int(round((budget + prod_cost) * 0.05))
    grand = budget + prod_cost + vat

    html = f"""
    {css}
    <div class="preview-wrap">
      <table class="preview">
        <tr>
          <td colspan="5" class="left" style="background:#fff; border:none;">
            <b>客戶名稱：</b>{client_name}<br>
            <b>Product：</b>{product_str}<br>
            <b>Period：</b>{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}<br>
            <b>Medium：</b>{mediums_str}
          </td>
          <td colspan="{days_cnt+3}" style="background:#fff; border:none;"></td>
        </tr>

        <tr><th colspan="7" style="border:none;"></th>{date_row1}<th style="border:none;"></th></tr>
        <tr>
          <th rowspan="2" class="header-blue">Station</th>
          <th rowspan="2" class="header-blue">Location</th>
          <th rowspan="2" class="header-blue">Program</th>
          <th rowspan="2" class="header-blue">Day-part</th>
          <th rowspan="2" class="header-blue">Size</th>
          <th rowspan="2" class="header-blue">Rate<br>(Net)</th>
          <th rowspan="2" class="header-blue">Package-cost<br>(Net)</th>
          {date_row2}
          <th rowspan="2" class="header-blue">檔次</th>
        </tr>
        <tr>{date_row3}</tr>

        {data_rows_html}

        <tr class="row-total">
          <td colspan="5" class="right">Total</td>
          <td class="right">{fmt_int(total_rate)}</td>
          <td class="right">{fmt_int(total_pkg)}</td>
          {"".join([f"<td>{fmt_int(x)}</td>" for x in totals])}
          <td class="cell-yellow">{fmt_int(sum(totals))}</td>
        </tr>

        <tr><td colspan="6" class="right">製作</td><td class="right">{fmt_int(prod_cost)}</td><td colspan="{days_cnt+1}"></td></tr>
        <tr><td colspan="6" class="right">專案優惠價 (Budget)</td><td class="right" style="color:red; font-weight:bold;">{fmt_int(budget)}</td><td colspan="{days_cnt+1}"></td></tr>
        <tr><td colspan="6" class="right">5% VAT</td><td class="right">{fmt_int(vat)}</td><td colspan="{days_cnt+1}"></td></tr>
        <tr class="row-grand"><td colspan="6" class="right">Grand Total</td><td class="right">{fmt_int(grand)}</td><td colspan="{days_cnt+1}"></td></tr>
      </table>
    </div>
    """
    return html


# ==========================================
# 4) Excel (xlsxwriter) 產出（用 merge_range 擬真）
# ==========================================

def generate_excel(rows, days_cnt, start_dt, end_dt, client_name, product_str, total_rate, total_pkg, budget, prod_cost):
    used_media = sorted(list(set(r["media"] for r in rows)), key=lambda x: {"全家廣播":1,"新鮮視":2,"家樂福":3}.get(x,99))
    mediums = "、".join(used_media)

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Media Schedule")

    # formats
    fmt_title = wb.add_format({'font_size': 18, 'bold': True, 'align': 'center'})
    fmt_info = wb.add_format({'align': 'left', 'valign': 'top', 'bold': True})
    fmt_head = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1,
                              'bg_color': '#4472C4', 'font_color': 'white', 'text_wrap': True, 'font_size': 10})
    fmt_date_wk = wb.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1,
                                 'bg_color': '#4472C4', 'font_color': 'white'})
    fmt_date_we = wb.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1,
                                 'bg_color': '#FFD966'})
    fmt_cell = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    fmt_left = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'text_wrap': True})
    fmt_num = wb.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0', 'font_size': 10})
    fmt_spots = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True,
                               'bg_color': '#FFF2CC', 'font_size': 10})
    fmt_total = wb.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True,
                               'bg_color': '#E2EFDA', 'num_format': '#,##0', 'font_size': 10})
    fmt_discount = wb.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True,
                                  'font_color': 'red', 'num_format': '#,##0', 'font_size': 10})
    fmt_grand = wb.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True,
                               'bg_color': '#FFC107', 'num_format': '#,##0', 'font_size': 10})

    # title
    # A1:?? 依 days_cnt 動態
    last_col = 7 + days_cnt  # 0-based: Station~Package=6, days start at 7, last is 7+days-1, spots at 7+days
    ws.merge_range(0, 0, 0, last_col, "Media Schedule", fmt_title)

    # info
    info = [
        ("客戶名稱：", client_name),
        ("Product：", product_str),
        ("Period：", f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"),
        ("Medium：", mediums),
    ]
    for idx, (k, v) in enumerate(info):
        ws.write(2 + idx, 0, k, fmt_info)
        ws.write(2 + idx, 1, v, fmt_info)

    # month header
    ws.write(6, 6, f"{start_dt.month}月", fmt_cell)

    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    curr = start_dt
    for i in range(days_cnt):
        col_idx = 7 + i
        wd = curr.weekday()
        fmt = fmt_date_we if wd >= 5 else fmt_date_wk
        ws.write(7, col_idx, curr.day, fmt)
        ws.write(8, col_idx, weekdays[wd], fmt)
        curr += timedelta(days=1)

    # headers row
    headers = ["Station", "Location", "Program", "Day-part", "Size", "Rate\n(Net)", "Package-cost\n(Net)"]
    for c, h in enumerate(headers):
        ws.write(8, c, h, fmt_head)
    ws.write(8, last_col, "檔次", fmt_head)

    # data rows with group merge
    current_row = 9
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]["media"] == row["media"] and rows[j]["seconds"] == row["seconds"]:
            j += 1
        group = rows[i:j]
        group_size = len(group)

        # Station merge per group
        station_txt = station_name(row["media"]).replace("<br>", "\n")
        if group_size > 1:
            ws.merge_range(current_row, 0, current_row + group_size - 1, 0, station_txt, fmt_left)
        else:
            ws.write(current_row, 0, station_txt, fmt_left)

        # Package merge if national (bundle)
        # 找 is_pkg_start 的那一列
        bundle_val = None
        for r in group:
            if r.get("is_pkg_start") and r.get("bundle_pkg") is not None:
                bundle_val = r["bundle_pkg"]
                break
        if bundle_val is not None:
            # merge H column (index 6)
            if group_size > 1:
                ws.merge_range(current_row, 6, current_row + group_size - 1, 6, bundle_val, fmt_num)
            else:
                ws.write(current_row, 6, bundle_val, fmt_num)

        # fill group rows
        for k, r in enumerate(group):
            rr = current_row + k
            ws.write(rr, 1, r["location"], fmt_cell)
            ws.write(rr, 2, r["program"], fmt_num)
            ws.write(rr, 3, r["daypart"], fmt_cell)
            ws.write(rr, 4, f"{r['seconds']}秒", fmt_cell)

            # rate
            if isinstance(r["rate_total"], int):
                ws.write(rr, 5, r["rate_total"], fmt_num)
            else:
                ws.write(rr, 5, r["rate_total"], fmt_cell)

            # package (non-national only)
            if bundle_val is None:
                v = r.get("pkg_display", "")
                if isinstance(v, int):
                    ws.write(rr, 6, v, fmt_num)
                else:
                    ws.write(rr, 6, v, fmt_cell)
            else:
                # national group other rows skip column 6
                pass

            # schedule
            sch = r.get("schedule", [])
            for d_idx, s_val in enumerate(sch[:days_cnt]):
                ws.write(rr, 7 + d_idx, s_val, fmt_cell)

            ws.write(rr, last_col, sum(sch[:days_cnt]), fmt_spots)

        current_row += group_size
        i = j

    # total row
    ws.write(current_row, 4, "Total", fmt_total)
    ws.write(current_row, 5, total_rate, fmt_total)
    ws.write(current_row, 6, total_pkg, fmt_total)

    # daily totals
    totals = [0]*days_cnt
    for r in rows:
        sch = r.get("schedule", [])
        for idx in range(min(days_cnt, len(sch))):
            totals[idx] += sch[idx]
    for idx, val in enumerate(totals):
        ws.write(current_row, 7 + idx, val, fmt_cell)
    ws.write(current_row, last_col, sum(totals), fmt_spots)

    # footer
    current_row += 1
    ws.write(current_row, 6, "製作", fmt_cell)
    ws.write(current_row, 7, prod_cost, fmt_num)

    current_row += 1
    ws.write(current_row, 6, "專案優惠價 (Budget)", fmt_cell)
    ws.write(current_row, 7, budget, fmt_discount)

    vat = int(round((budget + prod_cost) * 0.05))
    current_row += 1
    ws.write(current_row, 6, "5% VAT", fmt_cell)
    ws.write(current_row, 7, vat, fmt_num)

    grand = budget + prod_cost + vat
    current_row += 1
    ws.write(current_row, 6, "Grand Total", fmt_grand)
    ws.write(current_row, 7, grand, fmt_grand)

    # widths
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 16)
    ws.set_column('C:C', 10)
    ws.set_column('D:D', 14)
    ws.set_column('E:E', 8)
    ws.set_column('F:G', 14)
    ws.set_column(7, last_col, 4)

    wb.close()
    output.seek(0)
    return output.getvalue()


def html_to_pdf_bytes(html_str):
    if not WEASY_OK:
        return None
    try:
        return HTML(string=html_str).write_pdf()
    except Exception:
        return None


# ==========================================
# 5) Streamlit UI
# ==========================================

st.set_page_config(layout="wide", page_title="Cue Sheet Pro (擬真版)")
st.title("📺 媒體 Cue 表生成器（擬真版）")

# 基本資料
st.markdown("### 1) 基本資料")
col1, col2, col3 = st.columns(3)
with col1:
    client_name = st.text_input("客戶名稱", "萬國通路")
with col2:
    product_name = st.text_input("產品名稱", "統一布丁")
with col3:
    total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)

col4, col5 = st.columns(2)
with col4:
    start_date = st.date_input("開始日", datetime(2026, 1, 1))
with col5:
    end_date = st.date_input("結束日", datetime(2026, 1, 31))

days_count = (end_date - start_date).days + 1
if days_count <= 0:
    st.error("結束日必須晚於開始日")
    st.stop()
if days_count > 31:
    st.warning("此版本擬真以 31 天為上限，會只顯示前 31 天。")
days_count = min(days_count, 31)
st.info(f"📅 走期共 **{days_count}** 天（擬真顯示）")

prod_cost = 10000

st.markdown("### 2) 媒體投放設定（總和 100%）")

# 你範例那種 share flow：這裡簡化成直接填三個 slider，總和自動 normalize（避免 UI 卡死）
colm1, colm2, colm3 = st.columns(3)
with colm1:
    st.markdown("#### 📻 全家廣播")
    fm_act = st.checkbox("開啟", True, key="fm_act")
    fm_nat = st.checkbox("全省聯播", True, key="fm_nat") if fm_act else False
    fm_regs = []
    if fm_act:
        if fm_nat:
            fm_regs = ["全省"]
        else:
            fm_regs = st.multiselect("分區", REGIONS_ORDER, default=["北區", "桃竹苗", "中區"], key="fm_regs")
    fm_secs = st.multiselect("秒數", DURATIONS, default=[20], key="fm_secs") if fm_act else []
    fm_share = st.slider("預算佔比 %", 0, 100, 70, key="fm_share") if fm_act else 0

with colm2:
    st.markdown("#### 📺 新鮮視")
    fv_act = st.checkbox("開啟", True, key="fv_act")
    fv_nat = st.checkbox("全省聯播 ", False, key="fv_nat") if fv_act else False
    fv_regs = []
    if fv_act:
        if fv_nat:
            fv_regs = ["全省"]
        else:
            fv_regs = st.multiselect("分區 ", REGIONS_ORDER, default=["北區", "桃竹苗"], key="fv_regs")
    fv_secs = st.multiselect("秒數 ", DURATIONS, default=[10], key="fv_secs") if fv_act else []
    fv_share = st.slider("預算佔比 % ", 0, 100, 20, key="fv_share") if fv_act else 0

with colm3:
    st.markdown("#### 🛒 家樂福")
    cf_act = st.checkbox("開啟 ", True, key="cf_act")
    cf_secs = st.multiselect("秒數  ", DURATIONS, default=[20], key="cf_secs") if cf_act else []
    cf_share = st.slider("預算佔比 %  ", 0, 100, 10, key="cf_share") if cf_act else 0

# normalize shares to sum=100 for active medias
raw = []
if fm_act: raw.append(("全家廣播", fm_share))
if fv_act: raw.append(("新鮮視", fv_share))
if cf_act: raw.append(("家樂福", cf_share))
sum_share = sum(v for _, v in raw)
if sum_share <= 0:
    st.warning("請至少開啟一個媒體並給預算佔比")
    st.stop()

norm = [(k, v * 100.0 / sum_share) for k, v in raw]

def build_sec_shares(secs):
    secs = sorted(secs)
    if not secs:
        return {}
    # 先簡化：一律平均（你若要跟 v60.5 一樣逐秒 slider，我也可以再加回去）
    per = 100 // len(secs)
    mp = {s: per for s in secs}
    mp[secs[0]] += (100 - per * len(secs))
    return mp

config_media = {}
for k, share in norm:
    if k == "全家廣播":
        if fm_nat:
            regs = ["全省"]
        else:
            regs = fm_regs
        config_media[k] = {
            "is_national": fm_nat,
            "regions": regs if not fm_nat else ["全省"],
            "seconds": sorted(fm_secs),
            "share": share,
            "sec_shares": build_sec_shares(fm_secs)
        }
    elif k == "新鮮視":
        if fv_nat:
            regs = ["全省"]
        else:
            regs = fv_regs
        config_media[k] = {
            "is_national": fv_nat,
            "regions": regs if not fv_nat else ["全省"],
            "seconds": sorted(fv_secs),
            "share": share,
            "sec_shares": build_sec_shares(fv_secs)
        }
    elif k == "家樂福":
        config_media[k] = {
            "is_national": True,
            "regions": ["全省"],
            "seconds": sorted(cf_secs),
            "share": share,
            "sec_shares": build_sec_shares(cf_secs)
        }

# 防呆：非全省時如果 regions 空，就不產生 rows
for m, cfg in list(config_media.items()):
    if m in ["全家廣播", "新鮮視"] and (not cfg.get("is_national")) and (not cfg.get("regions")):
        st.error(f"{m} 沒勾全省但分區為空，請至少選 1 個分區。")
        st.stop()

# 計算
final_rows, debug_logs, total_rate, total_pkg = build_final_rows(config_media, total_budget_input, days_count)

# product string（顯示用）
all_secs = sorted(set([r["seconds"] for r in final_rows]))
product_str = "、".join([f"{s}秒" for s in all_secs]) + f" {product_name}" if all_secs else product_name

# 結果摘要
st.markdown("### 3) 結果摘要")
m1, m2, m3 = st.columns(3)
vat = int(round((total_budget_input + prod_cost) * 0.05))
grand_total = total_budget_input + prod_cost + vat
m1.metric("客戶預算 (未稅)", f"{total_budget_input:,}")
m2.metric("折扣後總金額 (含稅)", f"{grand_total:,}")
m3.metric("表格顯示 Total Package-cost", f"{total_pkg:,}")

with st.expander("💡 Debug（檔次/是否未達標）", expanded=False):
    for log in debug_logs:
        nat_txt = "全省" if log["is_national"] else "分區"
        st.write(f"- {log['media']} / {log['sec']}秒 / {nat_txt}：預算 {log['budget']:,.0f} → 檔次 {log['spots']}（Std {log['std']}）{(' → 未達標 x1.1' if log['penalty'] else '')}")

# HTML preview
st.markdown("### 4) 擬真預覽（HTML rowSpan）")
html = generate_html_preview(
    final_rows, days_count, start_date, end_date,
    client_name, product_str, total_rate, total_pkg, total_budget_input, prod_cost
)
st.components.v1.html(html, height=650, scrolling=True)

# downloads
st.markdown("### 5) 下載")
xlsx_bytes = generate_excel(
    final_rows, days_count, start_date, end_date,
    client_name, product_str, total_rate, total_pkg,
    total_budget_input, prod_cost
)
st.download_button(
    "📥 下載 Excel (.xlsx)",
    data=xlsx_bytes,
    file_name=f"Cue_{client_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

if WEASY_OK:
    pdf_bytes = html_to_pdf_bytes(html)
    if pdf_bytes:
        st.download_button(
            "📥 下載 PDF（HTML 擬真）",
            data=pdf_bytes,
            file_name=f"Cue_{client_name}.pdf",
            mime="application/pdf"
        )
else:
    st.info("PDF 需要 weasyprint：pip install weasyprint")
