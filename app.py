import streamlit as st
import math
import io
import os
import re
import shutil
import tempfile
import subprocess
from datetime import timedelta, datetime, date
from copy import copy

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font

# =========================================================
# 0) 基礎工具
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

# =========================================================
# 1) 核心資料庫（依你的規格）
# =========================================================
STORE_COUNTS_CONFIG = {
    "全省": 4437,
    "北區": 1649, "桃竹苗": 779, "中區": 839, "雲嘉南": 499, "高屏": 490, "東區": 181,
    "新鮮視_全省": 3124,
    "新鮮視_北區": 1127, "新鮮視_桃竹苗": 616, "新鮮視_中區": 528,
    "新鮮視_雲嘉南": 365, "新鮮視_高屏": 405, "新鮮視_東區": 83,
    "家樂福_量販": 68, "家樂福_超市": 249
}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

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

REGION_DISPLAY_6 = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市",
}
def region_display(region: str) -> str:
    return REGION_DISPLAY_6.get(region, region)

def get_sec_factor(media_type, seconds):
    return SEC_FACTORS_CONFIG.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    """偶數分配：先除2分配每天，再*2"""
    if days <= 0:
        return []
    if total_spots % 2 != 0:
        total_spots += 1
    half = total_spots // 2
    base, rem = divmod(half, days)
    half_schedule = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in half_schedule]

# =========================================================
# 2) Excel 輔助：safe write / style copy / merge處理
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

def apply_center(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, indent=al.indent)

def apply_left(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True, indent=al.indent)

def apply_right(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True, indent=al.indent)

def copy_row_with_style_fix(ws, src_row, dst_row, max_col):
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    shift = dst_row - src_row
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
                dc.value = Translator(v, origin=sc.coordinate).translate_formula(row_shift=shift, col_shift=0)
            except:
                dc.value = v
        else:
            dc.value = v

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

def find_cell_exact_in_ws(ws, text):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == text:
                return cell.row, cell.column
    return None

def find_row_contains_any(ws, col_letters, keyword):
    for col in col_letters:
        col_idx = column_index_from_string(col)
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, col_idx).value
            if isinstance(v, str) and keyword in v:
                return r
    return None

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

def set_schedule(ws, row, start_col_letter, max_days, schedule_list):
    start_col = column_index_from_string(start_col_letter)
    for i in range(max_days):
        v = schedule_list[i] if (schedule_list and i < len(schedule_list)) else None
        safe_write_rc(ws, row, start_col + i, v)

# =========================================================
# 3) 站台名稱（依你的模板習慣）
# =========================================================
def station_title(format_type, media_key):
    if media_key == "全家廣播":
        return "全家便利商店\n通路廣播廣告" if format_type == "Dongwu" else "全家便利商店\n廣播通路廣告"
    if media_key == "新鮮視":
        return "全家便利商店\n新鮮視廣告"
    return "家樂福"

# =========================================================
# 4) 規則核心：計算 rows（你要求的複雜規則）
# =========================================================
def compute_rows(config, total_budget_net, days_count):
    rows = []
    debug = []

    for media, cfg in config.items():
        media_budget = total_budget_net * (cfg["share"] / 100.0)

        for sec, sec_pct in cfg["sec_shares"].items():
            sec_budget = media_budget * (sec_pct / 100.0)
            if sec_budget <= 0:
                continue

            factor = get_sec_factor(media, sec)

            # -------------------------
            # A) 全家廣播 / 新鮮視
            # -------------------------
            if media in ["全家廣播", "新鮮視"]:
                mdb = PRICING_CONFIG[media]
                std = mdb["Std_Spots"]
                daypart = mdb["Day_Part"]
                prices = mdb["Prices"]

                is_national = bool(cfg.get("is_national", False))
                selected_regions = cfg.get("regions", [])

                # 沒勾全省又沒選區域 -> 不產生
                if (not is_national) and (not selected_regions):
                    continue

                # 計算用 region：全省用全省 Net；分區用所選 regions Net 加總
                if is_national:
                    calc_regions = ["全省"]
                    display_regions = REGIONS_ORDER[:]   # 6 區展開
                else:
                    calc_regions = selected_regions[:]   # 只算選的
                    display_regions = selected_regions[:] # 只顯示選的

                # ---- 裡子(Net)算 spots ----
                net_unit_sum = 0.0
                for reg in calc_regions:
                    net_price = prices[reg]["Net"]
                    net_unit_sum += (net_price / std) * factor

                if net_unit_sum <= 0:
                    continue

                spots_init = math.ceil(sec_budget / net_unit_sum)
                under_target = spots_init < std
                penalty = 1.1 if under_target else 1.0

                final_unit = net_unit_sum * penalty
                spots_final = math.ceil(sec_budget / final_unit)
                if spots_final % 2 != 0:
                    spots_final += 1
                if spots_final <= 0:
                    spots_final = 2

                sch = calculate_schedule(spots_final, days_count)

                # ---- 面子(List)寫入顯示欄位 ----
                # Rate(Net) = 分區總價（不是單檔）
                # Package(Net)：
                #   - 有全省：做「打包價合併格」，值 = 全省List總價（必要時*1.1），分區Rate不再*1.1
                #   - 無全省：若未達標 -> Rate/Package 都要*1.1；達標 -> Rate=Package
                nat_pkg_total = None
                if is_national:
                    nat_list = prices["全省"]["List"]
                    nat_pkg_total = int(round((nat_list / std) * factor * spots_final * (1.1 if under_target else 1.0)))

                for idx, reg in enumerate(display_regions):
                    reg_list = prices[reg]["List"]
                    base_total = (reg_list / std) * factor * spots_final

                    if is_national:
                        rate_total = int(round(base_total))  # 分區價不乘1.1
                        pkg_total = None  # 用 merged cell 顯示 nat_pkg_total
                        is_pkg_start = (idx == 0)  # 合併起點（第一列）
                        is_pkg_member = True
                    else:
                        # 分區購買：未達標 -> 全都*1.1；達標 -> 不乘
                        mult = 1.1 if under_target else 1.0
                        rate_total = int(round(base_total * mult))
                        pkg_total = rate_total
                        is_pkg_start = False
                        is_pkg_member = False

                    # Program 店/面數
                    if media == "新鮮視":
                        program_num = STORE_COUNTS_CONFIG.get(f"新鮮視_{reg}", 0)
                    else:
                        program_num = STORE_COUNTS_CONFIG.get(reg, 0)

                    rows.append({
                        "media_type": media,
                        "is_national": is_national,
                        "under_target": under_target,
                        "region": reg,
                        "program_num": program_num,
                        "daypart": daypart,
                        "seconds": sec,
                        "schedule": sch,
                        "spots": spots_final,
                        "rate_total": rate_total,          # Rate(Net) 欄：分區總價
                        "pkg_total": pkg_total,            # 非全省才逐列顯示
                        "nat_pkg_total": nat_pkg_total,    # 全省打包價
                        "is_pkg_start": is_pkg_start,
                        "is_pkg_member": is_pkg_member
                    })

                debug.append({
                    "media": media, "sec": sec, "budget": sec_budget,
                    "spots": spots_final, "std": std,
                    "status": "未達標" if under_target else "達標",
                    "penalty": penalty, "factor": factor
                })

            # -------------------------
            # B) 家樂福（固定全省）
            # -------------------------
            elif media == "家樂福":
                cdb = PRICING_CONFIG["家樂福"]["Configs"]
                base = cdb["量販_全省"]
                sup  = cdb["超市_全省"]

                base_std = base["Std"]
                factor = get_sec_factor("家樂福", sec)

                unit_net = (base["Net"] / base_std) * factor
                spots_init = math.ceil(sec_budget / unit_net)
                under_target = spots_init < base_std
                penalty = 1.1 if under_target else 1.0

                spots_final = math.ceil(sec_budget / (unit_net * penalty))
                if spots_final % 2 != 0:
                    spots_final += 1
                if spots_final <= 0:
                    spots_final = 2

                sch_h = calculate_schedule(spots_final, days_count)

                # 家樂福：Rate = 不含1.1的總價；Package = 反映1.1（如你示例）
                rate_total_h = int(round((base["List"] / base_std) * factor * spots_final))
                pkg_total_h  = int(round(rate_total_h * (1.1 if under_target else 1.0)))

                rows.append({
                    "media_type": "家樂福",
                    "is_national": True,
                    "under_target": under_target,
                    "region": "全省量販",
                    "program_num": STORE_COUNTS_CONFIG["家樂福_量販"],
                    "daypart": base["Time"],
                    "seconds": sec,
                    "schedule": sch_h,
                    "spots": spots_final,
                    "rate_total": rate_total_h,
                    "pkg_total": pkg_total_h,
                    "nat_pkg_total": None,
                    "is_pkg_start": False,
                    "is_pkg_member": False
                })

                # 超市檔次依比例
                spots_s = int(spots_final * (sup["Std"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({
                    "media_type": "家樂福",
                    "is_national": True,
                    "under_target": under_target,
                    "region": "全省超市",
                    "program_num": STORE_COUNTS_CONFIG["家樂福_超市"],
                    "daypart": sup["Time"],
                    "seconds": sec,
                    "schedule": sch_s,
                    "spots": spots_s,
                    "rate_total": "計量販",
                    "pkg_total": "計量販",
                    "nat_pkg_total": None,
                    "is_pkg_start": False,
                    "is_pkg_member": False
                })

                debug.append({
                    "media": "家樂福", "sec": sec, "budget": sec_budget,
                    "spots": spots_final, "std": base_std,
                    "status": "未達標" if under_target else "達標",
                    "penalty": penalty, "factor": factor
                })

    # 排序：媒體 -> 秒數 -> 區域順序
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    region_rank = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}
    rows.sort(key=lambda x: (
        media_order.get(x["media_type"], 99),
        int(x["seconds"]),
        region_rank.get(x["region"], 999),
    ))
    return rows, debug

# =========================================================
# 5) 依樣板定位 blocks（以 Total 為終點）
# =========================================================
def detect_blocks(ws, meta):
    total_pos = find_cell_exact_in_ws(ws, meta["total_label"])
    if not total_pos:
        raise ValueError("找不到 Total（樣板中必須有一格文字等於 'Total'）")
    total_row = total_pos[0]

    search_cols = meta.get("anchor_search_cols", ["B","C","D"])
    anchors = meta["anchors"]

    starts = []
    for m, kw in anchors.items():
        srow = find_row_contains_any(ws, search_cols, kw)
        if srow:
            starts.append((m, srow))
    starts.sort(key=lambda x: x[1])

    blocks = []
    for i, (m, srow) in enumerate(starts):
        erow = (starts[i+1][1] - 1) if i+1 < len(starts) else (total_row - 1)
        if erow >= srow:
            blocks.append({"media": m, "start": srow, "end": erow})

    if not blocks:
        raise ValueError("找不到任何媒體區段 blocks（anchors 可能與樣板不一致）")

    return blocks, total_row

# =========================================================
# 6) Block Renderer（Dongwu：Station整塊合併；Daypart/Seconds同秒數合併；Package 依規則合併）
# =========================================================
def build_groups_for_render(rows):
    grouped = {}
    for r in rows:
        grouped.setdefault(r["media_type"], {}).setdefault(r["seconds"], []).append(r)

    result = {}
    for media, sec_map in grouped.items():
        sec_groups = []
        for sec, items in sorted(sec_map.items(), key=lambda x: int(x[0])):
            # 同秒數群組內保持區域順序
            items_sorted = sorted(items, key=lambda x: REGIONS_ORDER.index(x["region"]) if x["region"] in REGIONS_ORDER else 999)

            # 是否為全省群組（radio/fv 才會出現 nat_pkg_total）
            is_national = any(i.get("is_national") for i in items_sorted)
            nat_pkg_total = None
            for it in items_sorted:
                if it.get("nat_pkg_total") is not None:
                    nat_pkg_total = it["nat_pkg_total"]
                    break

            sec_groups.append({
                "seconds": sec,
                "daypart": items_sorted[0].get("daypart", ""),
                "rows": items_sorted,
                "is_national": is_national,
                "nat_pkg_total": nat_pkg_total
            })
        result[media] = sec_groups
    return result

def render_block_dongwu(ws, meta, blk, sec_groups):
    cols = meta["cols"]
    s, e = blk["start"], blk["end"]
    max_col = ws.max_column
    style_row = s

    # 展平 lines
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

    # 解除會影響合併的舊 merge
    unmerge_overlaps_by_cols(ws, s, new_end, ["B", "E", "F", "H"])

    # Station 欄整塊合併
    ws.merge_cells(f"{cols['station']}{s}:{cols['station']}{new_end}")
    top_station = ws[f"{cols['station']}{s}"]
    top_station.value = station_title("Dongwu", blk["media"])
    apply_center(top_station)

    # 寫每列
    for i, line in enumerate(lines):
        r = s + i
        item = line["item"]

        # Location
        safe_write(ws, f"{cols['location']}{r}", region_display(item["region"]))
        apply_center(ws[f"{cols['location']}{r}"])

        # Program
        safe_write(ws, f"{cols['program']}{r}", int(item["program_num"]) if isinstance(item["program_num"], int) else item["program_num"])
        ws[f"{cols['program']}{r}"].number_format = "#,##0"
        apply_center(ws[f"{cols['program']}{r}"])

        # Rate(Net) -> 分區總價
        rate_cell = ws[f"{cols['rate']}{r}"]
        rv = item["rate_total"]
        safe_write(ws, f"{cols['rate']}{r}", rv)
        if isinstance(rv, int):
            rate_cell.number_format = "#,##0"
            apply_right(rate_cell)
        else:
            apply_center(rate_cell)

        # Schedule + Spots
        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)
        ws[f"{meta['total_col']}{r}"].number_format = "#,##0"
        apply_center(ws[f"{meta['total_col']}{r}"])

        # Package-cost：非全省才逐列寫
        if item.get("pkg_total") is not None:
            pv = item["pkg_total"]
            pkg_cell = ws[f"{cols['pkg']}{r}"]
            safe_write(ws, f"{cols['pkg']}{r}", pv)
            if isinstance(pv, int):
                pkg_cell.number_format = "#,##0"
                apply_right(pkg_cell)
            else:
                apply_center(pkg_cell)

    # 以秒數群組合併 Daypart / Seconds；Package 僅在全省時合併顯示打包價
    for g_start, g_end, g in group_ranges:
        r1 = s + g_start
        r2 = s + g_end

        # Daypart 合併
        ws.merge_cells(f"{cols['daypart']}{r1}:{cols['daypart']}{r2}")
        cE = ws[f"{cols['daypart']}{r1}"]
        cE.value = g["daypart"]
        apply_center(cE)

        # Seconds 合併（廣播/新鮮視顯示數字；家樂福顯示 xx秒）
        ws.merge_cells(f"{cols['seconds']}{r1}:{cols['seconds']}{r2}")
        cF = ws[f"{cols['seconds']}{r1}"]
        if blk["media"] == "家樂福":
            cF.value = f"{g['seconds']}秒"
        else:
            cF.value = int(g["seconds"])
        apply_center(cF)

        # Package 欄位：只有「全省」才合併顯示打包價
        if g.get("is_national") and g.get("nat_pkg_total") is not None and blk["media"] in ["全家廣播", "新鮮視"]:
            ws.merge_cells(f"{cols['pkg']}{r1}:{cols['pkg']}{r2}")
            cH = ws[f"{cols['pkg']}{r1}"]
            cH.value = int(g["nat_pkg_total"])
            cH.number_format = "#,##0"
            apply_right(cH)

def render_block_shenghuo(ws, meta, blk, sec_groups):
    # 生活樣板：通常不做大量合併，照列寫即可（保留樣板樣式）
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
        item = line["item"]
        g = line["group"]

        # Station
        safe_write(ws, f"{cols['station']}{r}", station_title("Shenghuo", blk["media"]))
        apply_center(ws[f"{cols['station']}{r}"])

        # Location / Program / Daypart / Seconds
        safe_write(ws, f"{cols['location']}{r}", region_display(item["region"]))
        apply_center(ws[f"{cols['location']}{r}"])

        safe_write(ws, f"{cols['program']}{r}", int(item["program_num"]) if isinstance(item["program_num"], int) else item["program_num"])
        ws[f"{cols['program']}{r}"].number_format = "#,##0"
        apply_center(ws[f"{cols['program']}{r}"])

        safe_write(ws, f"{cols['daypart']}{r}", g["daypart"])
        apply_center(ws[f"{cols['daypart']}{r}"])

        safe_write(ws, f"{cols['seconds']}{r}", f"{item['seconds']}秒廣告")
        apply_center(ws[f"{cols['seconds']}{r}"])

        # 專案價欄（若樣板有）
        proj_col = cols.get("proj_price")
        if proj_col:
            # 生活樣板若欄位叫專案價，就用 pkg_total 或 rate_total（視你的樣板用途）
            v = item.get("pkg_total")
            if v is None:
                v = item.get("rate_total")
            safe_write(ws, f"{proj_col}{r}", v)
            if isinstance(v, int):
                ws[f"{proj_col}{r}"].number_format = "#,##0"
                apply_right(ws[f"{proj_col}{r}"])
            else:
                apply_center(ws[f"{proj_col}{r}"])

        # Schedule + Spots
        set_schedule(ws, r, meta["schedule_start_col"], meta["max_days"], item.get("schedule", []))
        spot_sum = sum(item.get("schedule", [])[:meta["max_days"]])
        safe_write(ws, f"{meta['total_col']}{r}", spot_sum)
        ws[f"{meta['total_col']}{r}"].number_format = "#,##0"
        apply_center(ws[f"{meta['total_col']}{r}"])

# =========================================================
# 7) Footer 美化（只做必要對齊，不破壞樣板視覺）
# =========================================================
def polish_footer_minimal(ws, meta):
    lbl = meta["footer_labels"]
    for key in ["make", "vat", "grand"]:
        pos = find_cell_exact_in_ws(ws, lbl.get(key, ""))
        if not pos:
            continue
        r, c = pos
        label_cell = ws.cell(r, c)
        val_cell = ws.cell(r, c + 1)
        apply_left(label_cell)
        apply_right(val_cell)
        if isinstance(val_cell.value, (int, float)):
            val_cell.number_format = "#,##0"
        if key == "grand":
            label_cell.font = copy(label_cell.font) if label_cell.font else Font(bold=True)
            val_cell.font = copy(val_cell.font) if val_cell.font else Font(bold=True)

# =========================================================
# 8) Sheet Meta（如你的樣板不同，可以只改這裡）
# =========================================================
SHEET_META = {
    "Dongwu": {
        "sheet_name": "東吳-格式",
        "date_start_cell": "I7",
        "schedule_start_col": "I",
        "max_days": 31,
        "total_col": "AN",  # 檔次/每日合計欄
        "anchors": {"全家廣播": "通路廣播廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "anchor_search_cols": ["B", "C", "D"],
        "header_cells": {"client": "C3", "product": "C4", "period": "C5", "medium": "C6", "month": "I6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "rate": "G", "pkg": "H"},
        # 改成你要的表頭
        "header_override": {"G7": "Rate\n(Net)", "H7": "Package-cost\n(Net)"},
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
    },
    "Shenghuo": {
        "sheet_name": "聲活-格式",
        "date_start_cell": "G7",
        "schedule_start_col": "G",
        "max_days": 23,
        "total_col": "AD",
        "anchors": {"全家廣播": "廣播通路廣告", "新鮮視": "新鮮視廣告", "家樂福": "家樂福"},
        "anchor_search_cols": ["B", "C", "D"],
        "header_cells": {"client": "C5", "product": "C6", "month": "G6"},
        "cols": {"station": "B", "location": "C", "program": "D", "daypart": "E", "seconds": "F", "proj_price": "AF"},
        "header_override": {"AF7": "專案價\n(Net)"},
        "total_label": "Total",
        "footer_labels": {"make": "製作", "vat": "5% VAT", "grand": "Grand Total"},
    }
}

# =========================================================
# 9) 產生 Excel（用樣板填值，保留字型/顏色/logo）
# =========================================================
def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, template_bytes):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if meta["sheet_name"] not in wb.sheetnames:
        raise ValueError(f"缺少分頁：{meta['sheet_name']}（請確認樣板分頁名）")

    # 只保留目標 sheet
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
        safe_write(ws, hc["medium"], "、".join(sorted(set([r["media_type"] for r in rows]))))
    if "month" in hc:
        safe_write(ws, hc["month"], f" {start_dt.month}月")

    # 日期起始
    safe_write(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))

    # 表頭覆寫（只覆文字，不改樣式）
    for addr, text in meta.get("header_override", {}).items():
        safe_write(ws, addr, text)

    # 找 blocks & total
    blocks, total_row_orig = detect_blocks(ws, meta)

    # 先解除資料區可能影響渲染的 merge（避免 Total/Footer 被吞）
    data_start = min(b["start"] for b in blocks)
    data_end = total_row_orig - 1
    unmerge_overlaps_by_cols(ws, data_start, data_end, ["B", "E", "F", "H"])

    # 分群
    media_groups = build_groups_for_render(rows)

    # 由下往上渲染（避免 insert/delete 影響上方定位）
    for blk in sorted(blocks, key=lambda b: b["start"], reverse=True):
        m = blk["media"]
        sec_groups = media_groups.get(m, [])

        # 沒資料 -> 刪整段
        if not sec_groups:
            ws.delete_rows(blk["start"], amount=(blk["end"] - blk["start"] + 1))
            continue

        if format_type == "Dongwu":
            render_block_dongwu(ws, meta, blk, sec_groups)
        else:
            render_block_shenghuo(ws, meta, blk, sec_groups)

    # 重新找 Total row
    total_pos = find_cell_exact_in_ws(ws, meta["total_label"])
    if not total_pos:
        raise ValueError("渲染後找不到 Total（可能樣板 Total 文字不同）")
    total_row = total_pos[0]

    # Total daily spots
    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [0] * eff_days
    for x in rows:
        sch = x.get("schedule", [])
        for d in range(min(eff_days, len(sch))):
            daily_sums[d] += sch[d]
    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))
    ws[f"{meta['total_col']}{total_row}"].number_format = "#,##0"
    apply_center(ws[f"{meta['total_col']}{total_row}"])

    # Total Rate / Package（Dongwu: G/H；Shenghuo: AF 之類）
    if format_type == "Dongwu":
        cols = meta["cols"]
        rate_sum = 0
        pkg_sum = 0

        # Rate：加總所有 int 的 rate_total
        for r in rows:
            if isinstance(r.get("rate_total"), int):
                rate_sum += r["rate_total"]

        # Package：加總「非全省逐列pkg」 + 「全省打包價」每組只算一次
        seen_national_group = set()
        for r in rows:
            if r["media_type"] in ["全家廣播", "新鮮視"] and r.get("is_national") and r.get("nat_pkg_total") is not None:
                key = (r["media_type"], r["seconds"])
                if key not in seen_national_group:
                    pkg_sum += int(r["nat_pkg_total"])
                    seen_national_group.add(key)
            else:
                pv = r.get("pkg_total")
                if isinstance(pv, int):
                    pkg_sum += pv

        safe_write(ws, f"{cols['rate']}{total_row}", rate_sum)
        ws[f"{cols['rate']}{total_row}"].number_format = "#,##0"
        apply_right(ws[f"{cols['rate']}{total_row}"])

        safe_write(ws, f"{cols['pkg']}{total_row}", pkg_sum)
        ws[f"{cols['pkg']}{total_row}"].number_format = "#,##0"
        apply_right(ws[f"{cols['pkg']}{total_row}"])

    else:
        # Shenghuo：用 proj_price 放總價（若你要改成別欄可調 meta）
        cols = meta["cols"]
        proj_col = cols.get("proj_price")
        if proj_col:
            total_proj = 0
            for r in rows:
                pv = r.get("pkg_total")
                if pv is None:
                    pv = r.get("rate_total")
                if isinstance(pv, int):
                    total_proj += pv
            safe_write(ws, f"{proj_col}{total_row}", total_proj)
            ws[f"{proj_col}{total_row}"].number_format = "#,##0"
            apply_right(ws[f"{proj_col}{total_row}"])

    # Footer：製作/VAT/Grand Total（沿用樣板數值，這裡只確保格式）
    polish_footer_minimal(ws, meta)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 10) Windows Excel COM：匯出高擬真 PDF（最像公司範例）
# =========================================================
def xlsx_to_pdf_excel_com(xlsx_bytes: bytes, sheet_name: str):
    if os.name != "nt":
        return None, "Fail", "非 Windows 無法使用 Excel COM（要 1:1 請用 Windows + Excel）"

    try:
        import win32com.client  # pywin32
    except Exception:
        return None, "Fail", "缺少 pywin32：請 pip install pywin32"

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
            ws = wb.Worksheets(sheet_name)

            # --- 防裁切必要設定（不破壞樣板視覺，只做「不要切」） ---
            try:
                # PrintArea = UsedRange
                used = ws.UsedRange
                addr = used.Address
                ws.PageSetup.PrintArea = addr

                # Fit 1 page wide
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesWide = 1
                ws.PageSetup.FitToPagesTall = 1

                # 橫向（如果樣板已經是橫向，其實不會變）
                # xlLandscape = 2
                ws.PageSetup.Orientation = 2
            except:
                pass

            # 匯出
            # 0 = xlTypePDF
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
                return f.read(), "Excel COM", ""
        return None, "Fail", "Excel COM 匯出失敗（未產生 pdf）"

# =========================================================
# 11) UI
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro (Template 1:1)")

st.title("📺 Cue Sheet Pro（樣板 1:1 擬真版）")
st.caption("目標：Excel/PDF 盡量與公司樣板完全一致（字型/顏色/Logo/框線皆以樣板為準）")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TPL_DW = os.path.join(BASE_DIR, "東吳樣板.xlsx")
DEFAULT_TPL_SH = os.path.join(BASE_DIR, "生活樣板.xlsx")

format_type = st.radio("格式", ["Dongwu", "Shenghuo"], horizontal=True)
tpl_path = DEFAULT_TPL_DW if format_type == "Dongwu" else DEFAULT_TPL_SH

with st.expander("📌 樣板來源（務必放同資料夾以達到 1:1）", expanded=True):
    st.write(f"目前預設樣板路徑：`{tpl_path}`")
    st.write("若你要用另一份樣板，可以在這裡上傳覆蓋。")
    upl = st.file_uploader("上傳樣板 (xlsx)（可選）", type=["xlsx"])
    template_bytes = None
    if upl:
        template_bytes = upl.read()
        st.success("✅ 已使用上傳的樣板")
    else:
        if os.path.exists(tpl_path):
            with open(tpl_path, "rb") as f:
                template_bytes = f.read()
            st.success("✅ 已載入本機樣板檔")
        else:
            st.error("❌ 找不到樣板檔。請把『東吳樣板.xlsx / 生活樣板.xlsx』放在程式同資料夾，或直接上傳。")

st.markdown("## 1) 基本資料")
c1, c2, c3 = st.columns(3)
with c1:
    client_name = st.text_input("客戶名稱", "萬國通路")
with c2:
    product_name = st.text_input("產品名稱", "統一布丁")
with c3:
    total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)

d1, d2 = st.columns(2)
with d1:
    start_date = st.date_input("開始日", date(2026, 1, 1))
with d2:
    end_date = st.date_input("結束日", date(2026, 1, 31))

days_count = (end_date - start_date).days + 1
st.info(f"📅 走期共 **{days_count}** 天")

st.markdown("## 2) 媒體投放設定（總和 100%）")

# 初始化 share
if "rad_share" not in st.session_state:
    st.session_state.rad_share = 60
if "fv_share" not in st.session_state:
    st.session_state.fv_share = 40
if "cf_share" not in st.session_state:
    st.session_state.cf_share = 0

def normalize_shares(active_keys):
    if not active_keys:
        return
    share = 100 // len(active_keys)
    for k in active_keys:
        st.session_state[k] = share
    rem = 100 - sum(st.session_state[k] for k in active_keys)
    st.session_state[active_keys[0]] += rem

def on_media_toggle():
    active = []
    if st.session_state.get("cb_rad"): active.append("rad_share")
    if st.session_state.get("cb_fv"):  active.append("fv_share")
    if st.session_state.get("cb_cf"):  active.append("cf_share")
    normalize_shares(active)

col_cb1, col_cb2, col_cb3 = st.columns(3)
with col_cb1:
    is_rad = st.checkbox("全家廣播", value=True, key="cb_rad", on_change=on_media_toggle)
with col_cb2:
    is_fv = st.checkbox("新鮮視", value=True, key="cb_fv", on_change=on_media_toggle)
with col_cb3:
    is_cf = st.checkbox("家樂福", value=False, key="cb_cf", on_change=on_media_toggle)

config = {}
m1, m2, m3 = st.columns(3)

if is_rad:
    with m1:
        st.subheader("📻 全家廣播")
        is_nat = st.checkbox("全省聯播", value=True, key="rad_nat")
        regs = ["全省"] if is_nat else st.multiselect("分區", REGIONS_ORDER, default=["北區","中區","高屏"], key="rad_regs")
        secs = sorted(st.multiselect("秒數", DURATIONS, default=[20], key="rad_secs"))
        st.slider("預算 %", 0, 100, key="rad_share")
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in secs[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"rad_s_{s}")
                sec_shares[s] = v
                ls -= v
            sec_shares[secs[-1]] = ls
        elif secs:
            sec_shares[secs[0]] = 100

        config["全家廣播"] = {
            "is_national": is_nat,
            "regions": regs if not is_nat else ["全省"],
            "seconds": secs,
            "share": st.session_state.rad_share,
            "sec_shares": sec_shares
        }

if is_fv:
    with m2:
        st.subheader("📺 新鮮視")
        is_nat = st.checkbox("全省聯播 ", value=False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("分區", REGIONS_ORDER, default=["北區","中區","高屏"], key="fv_regs")
        secs = sorted(st.multiselect("秒數 ", DURATIONS, default=[10, 35], key="fv_secs"))
        st.slider("預算 % ", 0, 100, key="fv_share")
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in secs[:-1]:
                v = st.slider(f"{s}秒 % ", 0, ls, int(ls/2), key=f"fv_s_{s}")
                sec_shares[s] = v
                ls -= v
            sec_shares[secs[-1]] = ls
        elif secs:
            sec_shares[secs[0]] = 100

        config["新鮮視"] = {
            "is_national": is_nat,
            "regions": regs if not is_nat else ["全省"],
            "seconds": secs,
            "share": st.session_state.fv_share,
            "sec_shares": sec_shares
        }

if is_cf:
    with m3:
        st.subheader("🛒 家樂福")
        secs = sorted(st.multiselect("秒數", DURATIONS, default=[20, 30, 45], key="cf_secs"))
        st.slider("預算 %", 0, 100, key="cf_share")
        sec_shares = {}
        if len(secs) > 1:
            ls = 100
            for s in secs[:-1]:
                v = st.slider(f"{s}秒 %", 0, ls, int(ls/2), key=f"cf_s_{s}")
                sec_shares[s] = v
                ls -= v
            sec_shares[secs[-1]] = ls
        elif secs:
            sec_shares[secs[0]] = 100

        config["家樂福"] = {
            "is_national": True,
            "regions": ["全省"],
            "seconds": secs,
            "share": st.session_state.cf_share,
            "sec_shares": sec_shares
        }

# share總和提示（不強制，但提醒）
share_sum = sum(v["share"] for v in config.values()) if config else 0
if share_sum != 100 and config:
    st.warning(f"⚠️ 目前媒體預算總和為 {share_sum}%，建議調整成 100%（不然會影響計算）")

# =========================================================
# 12) 生成
# =========================================================
st.markdown("## 3) 生成與下載")

btn = st.button("🚀 生成 Excel / PDF", type="primary", disabled=(template_bytes is None or not config))
if btn:
    rows, debug_logs = compute_rows(config, total_budget_input, days_count)

    if not rows:
        st.error("沒有產生任何資料列：請確認你有勾媒體、秒數、以及（非全省時）有選分區。")
        st.stop()

    # 產品顯示字串：列出所有秒數
    all_secs = sorted(set(int(r["seconds"]) for r in rows))
    p_str = "、".join([f"{s}秒" for s in all_secs]) + f" {product_name}"

    try:
        xlsx_bytes = generate_excel_from_template(
            format_type=format_type,
            start_dt=start_date,
            end_dt=end_date,
            client_name=client_name,
            product_display_str=p_str,
            rows=rows,
            template_bytes=template_bytes
        )

        st.download_button(
            "📥 下載 Excel",
            data=xlsx_bytes,
            file_name=f"Cue_{safe_filename(client_name)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 高擬真 PDF（Excel COM）
        pdf_bytes, method, err = xlsx_to_pdf_excel_com(xlsx_bytes, SHEET_META[format_type]["sheet_name"])
        if pdf_bytes:
            st.download_button(
                f"📥 下載 PDF（{method} 高擬真）",
                data=pdf_bytes,
                file_name=f"Cue_{safe_filename(client_name)}.pdf",
                mime="application/pdf"
            )
            st.success("✅ PDF 已用 Excel 高擬真匯出（最接近公司範例）")
        else:
            st.warning(f"PDF 無法高擬真產出：{err}\n\n若你要 1:1 PDF，請用 Windows + 已安裝 Excel 的環境跑 Streamlit。")

        with st.expander("🧪 Debug（檔次/達標/1.1）", expanded=False):
            for d in debug_logs:
                st.write(d)

    except Exception as e:
        st.error(f"生成失敗：{e}")
