import streamlit as st
import math
import io
import os
import shutil
import tempfile
import subprocess
import re
from datetime import timedelta, datetime
from copy import copy

import requests
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment

# =========================================================
# 0) 基礎工具函式（必須放最前面）
# =========================================================
def parse_count_to_int(x):
    """將含有逗號的字串或數字轉為整數"""
    if x is None:
        return 0
    if isinstance(x, (int, float)):
        return int(x)
    s = str(x)
    m = re.findall(r"[\d,]+", s)
    if not m:
        return 0
    return int(m[0].replace(",", ""))

def sanitize_text(s):
    """清理不可見控制字元/替代字元，避免 PDF 亂碼、怪符號"""
    if s is None:
        return s
    s = str(s)
    s = s.replace("\uFFFD", "").replace("\uFFFE", "").replace("\ufeff", "")
    s = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", "", s)  # 保留 \t \n
    s = re.sub(r"[ ]{2,}", " ", s).strip()
    return s

def find_soffice_path():
    """找 LibreOffice/soffice（雲端通常在 /usr/bin/soffice）"""
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
    """
    PDF 最高擬真策略：
    1) Windows 本機：Excel COM ExportAsFixedFormat（最像）
    2) Linux / Streamlit Cloud：LibreOffice headless convert（很像）
    3) 都沒有 → 回傳 None（讓 UI fallback）
    """
    # 1) Windows：Excel COM（最像）
    if os.name == "nt":
        try:
            import win32com.client  # pip install pywin32
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
                    try:
                        if wb is not None:
                            wb.Close(False)
                    except:
                        pass
                    excel.Quit()

                if os.path.exists(pdf_path):
                    return open(pdf_path, "rb").read(), "excel_com", ""
                return None, "excel_com_fail", "excel export no output"
        except Exception as e:
            return None, "excel_com_fail", str(e)

    # 2) LibreOffice headless
    soffice = find_soffice_path()
    if soffice:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                xlsx_path = os.path.join(tmp, "cue.xlsx")
                with open(xlsx_path, "wb") as f:
                    f.write(xlsx_bytes)

                # ✅雲端穩定性關鍵：固定 UserInstallation profile（避免 lock / first-run wizard）
                lo_profile = os.path.join(tmp, "lo_profile")
                os.makedirs(lo_profile, exist_ok=True)

                # file URI（Windows/Linux 都安全）
                lo_uri = "file:///" + lo_profile.replace("\\", "/")

                cmd = [
                    soffice,
                    "--headless", "--invisible", "--nologo", "--nofirststartwizard",
                    "--norestore", "--nolockcheck",
                    f"-env:UserInstallation={lo_uri}",
                    "--convert-to", "pdf:calc_pdf_Export",
                    "--outdir", tmp,
                    xlsx_path
                ]
                p = subprocess.run(cmd, capture_output=True, text=True)

                if p.returncode != 0:
                    return None, "libreoffice_fail", (p.stderr or p.stdout or "LibreOffice convert failed")

                # 期待輸出 cue.pdf，但 LO 有時會輸出不同檔名
                pdf_path = os.path.join(tmp, "cue.pdf")
                if not os.path.exists(pdf_path):
                    for fn in os.listdir(tmp):
                        if fn.lower().endswith(".pdf"):
                            pdf_path = os.path.join(tmp, fn)
                            break

                if os.path.exists(pdf_path):
                    return open(pdf_path, "rb").read(), "libreoffice", ""
                return None, "libreoffice_fail", "libreoffice export no output"
        except Exception as e:
            return None, "libreoffice_fail", str(e)

    return None, "none", "no_excel_com_and_no_soffice"

def ensure_noto_tc_ttf():
    """WeasyPrint fallback 用：確保中文字體在本機可用（避免亂碼）"""
    font_path = "NotoSansTC-Regular.ttf"
    if os.path.exists(font_path) and os.path.getsize(font_path) > 100000:
        return font_path
    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/TTF/TraditionalChinese/NotoSansTC-Regular.ttf"
    try:
        r = requests.get(url, timeout=20)
        if r.status_code == 200 and len(r.content) > 100000:
            with open(font_path, "wb") as f:
                f.write(r.content)
            return font_path
    except:
        pass
    return font_path if os.path.exists(font_path) else None

def html_to_pdf_weasyprint(html_str: str):
    """
    ⚠️擬真度一定低於 xlsx→pdf（這是最後 fallback）
    """
    try:
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
    except Exception as e:
        return None, f"WeasyPrint not installed: {e}"

    font_path = ensure_noto_tc_ttf()
    font_abs = os.path.abspath(font_path) if font_path else ""

    font_config = FontConfiguration()
    css_str = f"""
    @page {{
        size: A4 landscape;
        margin: 0.5cm;
    }}
    @font-face {{
        font-family: 'NotoSansTC';
        src: url('file://{font_abs}');
    }}
    body {{
        font-family: 'NotoSansTC', sans-serif !important;
        font-size: 8pt;
    }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{
        border: 0.5pt solid #555;
        padding: 2px 3px;
        text-align: center;
        vertical-align: middle;
        white-space: nowrap;
    }}
    tr {{ page-break-inside: avoid; }}
    """
    try:
        pdf_bytes = HTML(string=html_str).write_pdf(
            stylesheets=[CSS(string=css_str)],
            font_config=font_config
        )
        return pdf_bytes, ""
    except Exception as e:
        return None, f"PDF Render Error: {str(e)}"

# =========================================================
# 1) Streamlit 頁面設定 & 自動載入模板
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v63.1 (Cloud PDF 擬真提升)")

GOOGLE_DRIVE_FILE_ID = "11R1SA_hpFD5O_MGmYeh4BdtcUhK2bPta"
DEFAULT_FILENAME = "1209-Cue表相關資料.xlsx"

@st.cache_resource(ttl=600)
def load_default_template():
    status_msg = []
    # Google Drive（若權限可直連）
    if GOOGLE_DRIVE_FILE_ID:
        url = f"https://drive.google.com/uc?export=download&id={GOOGLE_DRIVE_FILE_ID}"
        try:
            r = requests.get(url, timeout=20, allow_redirects=True)
            if r.status_code == 200:
                if b"<!DOCTYPE html>" in r.content[:500]:
                    status_msg.append("⚠️ Drive 下載權限受限（可能需登入 / 不是直連下載）")
                else:
                    return r.content, "雲端硬碟 (Google Drive)", status_msg
        except Exception as e:
            status_msg.append(f"❌ Drive 連線錯誤: {e}")

    # 本機檔（你本機測試用；雲端通常沒有）
    if os.path.exists(DEFAULT_FILENAME):
        try:
            with open(DEFAULT_FILENAME, "rb") as f:
                return f.read(), "系統主機 (Local)", status_msg
        except Exception as e:
            status_msg.append(f"❌ 本地讀取失敗: {e}")

    return None, None, status_msg

# =========================================================
# 2) CSS（預覽更像 Excel）
# =========================================================
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700&display=swap');
  .stApp { background-color: #f4f4f4; font-family: 'Noto Sans TC', sans-serif; }

  .preview-wrapper {
      background: #fff;
      padding: 12px;
      border: 1px solid #cfcfcf;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
      margin-bottom: 16px;
      overflow: auto;
      white-space: nowrap;
  }

  table.excel-table {
      border-collapse: collapse;
      font-size: 11px;
      min-width: 1200px;
      color: #000;
      font-family: Arial, "Microsoft JhengHei", sans-serif;
  }

  .excel-table th, .excel-table td {
      border: 1px solid #000;
      padding: 3px 6px;
      text-align: center;
      vertical-align: middle;
      height: 20px;
      line-height: 1.2;
      white-space: nowrap;
  }

  .excel-table thead th {
      position: sticky; top: 0; z-index: 20;
      border-bottom: 2px solid #000;
  }
  .excel-table thead tr:nth-child(2) th { top: 24px; z-index: 19; }

  .bg-dw-head { background-color: #4472C4; color: white; font-weight: 700; }
  .bg-sh-head { background-color: #BDD7EE; color: black; font-weight: 700; }
  .bg-weekend { background-color: #FFD966; color: black; }
  .bg-total   { background-color: #FFF2CC; font-weight: 700; }

  .align-left  { text-align: left !important; padding-left: 6px !important; }
  .align-right { text-align: right !important; padding-right: 6px !important; font-family: Consolas, monospace; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 3) 資料庫
# =========================================================
STORE_COUNTS_RAW = {
    "全省": "4,437店",
    "北區": "1,649店", "桃竹苗": "779店", "中區": "839店", "雲嘉南": "499店", "高屏": "490店", "東區": "181店",
    "新鮮視_全省": "3,124面",
    "新鮮視_北區": "1,127面", "新鮮視_桃竹苗": "616面", "新鮮視_中區": "528面",
    "新鮮視_雲嘉南": "365面", "新鮮視_高屏": "405面", "新鮮視_東區": "83面",
    "家樂福_量販": "67店", "家樂福_超市": "250店"
}
STORE_COUNTS_NUM = {k: parse_count_to_int(v) for k, v in STORE_COUNTS_RAW.items()}

# 2026：全家廣播以 6 區（你確認不是 bug）
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

PRICING_DB = {
    "全家廣播": {
        "Std_Spots": 480, "Day_Part": "00:00-24:00",
        "全省": [400000, 320000],
        "北區": [250000, 200000], "桃竹苗": [150000, 120000], "中區": [150000, 120000],
        "雲嘉南": [100000, 80000], "高屏": [100000, 80000], "東區": [62500, 50000]
    },
    "新鮮視": {
        "Std_Spots": 504, "Day_Part": "07:00-22:00",
        "全省": [150000, 120000],
        "北區": [150000, 120000], "桃竹苗": [120000, 96000], "中區": [90000, 72000],
        "雲嘉南": [75000, 60000], "高屏": [75000, 60000], "東區": [45000, 36000]
    },
    "家樂福": {
        "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420, "Day_Part": "09:00-23:00"},
        "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720, "Day_Part": "00:00-24:00"},
    }
}

SEC_FACTORS = {
    "全家廣播": {30: 1.0, 20: 0.85, 15: 0.65, 10: 0.5, 5: 0.25},
    "新鮮視": {30: 3.0, 20: 2.0, 15: 1.5, 10: 1.0, 5: 0.5},
    "家樂福": {30: 1.5, 20: 1.0, 15: 0.85, 10: 0.65, 5: 0.35}
}

def get_sec_factor(media_type, seconds):
    return SEC_FACTORS.get(media_type, {}).get(seconds, 1.0)

def calculate_schedule(total_spots, days):
    """固定偶數、對半分配後乘回 2 => 每天都偶數"""
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
        sanitize_text(f"1.請於 {d_str}前 回簽及進單，方可順利上檔。"),
        sanitize_text("2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇時段滿檔，上檔時間挪後或更換至同級時段。"),
        sanitize_text("3.通路店鋪數與開機率開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。"),
        sanitize_text("4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。"),
        sanitize_text(f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。"),
        sanitize_text(f"6.付款兌現日期：{p_str}")
    ]

REGION_DISPLAY_6 = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市",
}
def region_display(region: str) -> str:
    return sanitize_text(REGION_DISPLAY_6.get(region, region))

# =========================================================
# 4) Excel 生成（openpyxl：模板填空）
# =========================================================
def _get_master_cell(ws, cell):
    if not isinstance(cell, MergedCell):
        return cell
    r, c = cell.row, cell.column
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
            return ws.cell(row=mr.min_row, column=mr.min_col)
    return None

def safe_write(ws, addr: str, value):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master:
            master.value = value
        return
    cell.value = value

def safe_write_rc(ws, row: int, col: int, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        master = _get_master_cell(ws, cell)
        if master:
            master.value = value
        return
    cell.value = value

def apply_center_style(cell):
    al = cell.alignment or Alignment()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=al.wrap_text, indent=al.indent)

def copy_row_with_style_fix(ws, src_row, dst_row, max_col):
    """複製樣式 + 公式平移（避免插列後格式跑掉）"""
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    row_shift = dst_row - src_row
    for c in range(1, max_col + 1):
        sc = ws.cell(src_row, c)
        dc = ws.cell(dst_row, c)

        if isinstance(sc, MergedCell) or isinstance(dc, MergedCell):
            continue

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
    """最後補強：確保某些欄位全部置中（含 merged master cell）"""
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
        except:
            pass

def set_schedule(ws, row, start_col_letter, max_days, schedule_list):
    start_col = column_index_from_string(start_col_letter)
    for i in range(max_days):
        v = schedule_list[i] if (schedule_list and i < len(schedule_list)) else None
        safe_write_rc(ws, row, start_col + i, v)

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

def hide_unused_sheets(wb, keep_sheet_names, mode="veryHidden"):
    mode = mode if mode in ("hidden", "veryHidden") else "veryHidden"
    for sh in wb.worksheets:
        if sh.title not in keep_sheet_names:
            sh.sheet_state = mode

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
        "force_center_cols": [],
    }
}

def generate_excel_from_template(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, template_bytes):
    meta = SHEET_META[format_type]
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if meta["sheet_name"] not in wb.sheetnames:
        raise ValueError(f"缺少分頁：{meta['sheet_name']}")

    ws = wb[meta["sheet_name"]]
    hide_unused_sheets(wb, [meta["sheet_name"]], mode="veryHidden")

    # Header
    hc = meta["header_cells"]
    if "client" in hc:
        safe_write(ws, hc["client"], sanitize_text(client_name))
    if "product" in hc:
        safe_write(ws, hc["product"], sanitize_text(product_display_str))
    if "period" in hc:
        safe_write(ws, hc["period"], f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y.%m. %d')}")
    if "medium" in hc:
        safe_write(ws, hc["medium"], sanitize_text(" ".join(sorted(set([r["media_type"] for r in rows])))))
    if "month" in hc:
        safe_write(ws, hc["month"], f" {start_dt.month}月")

    safe_write(ws, meta["date_start_cell"], datetime(start_dt.year, start_dt.month, start_dt.day))

    for addr, text in meta.get("header_override", {}).items():
        safe_write(ws, addr, sanitize_text(text))

    # 找 Total row
    total_cell = find_cell_exact(ws, meta["total_label"])
    if not total_cell:
        raise ValueError("找不到 Total（模板可能變更）")
    total_row = total_cell[0]

    cols = meta["cols"]

    # 找各媒體區塊起點（以 Station 欄關鍵字）
    sec_start = {}
    for m_key, kw in meta["anchors"].items():
        r0 = find_first_row_contains(ws, cols["station"], kw)
        if r0:
            sec_start[m_key] = r0
    if not sec_start:
        raise ValueError("模板找不到任何區塊 anchor（請確認 Station 欄是否有關鍵字）")

    # 產生區塊範圍
    sec_order = sorted(sec_start.items(), key=lambda x: x[1])
    sec_ranges = []
    for i, (k, sr) in enumerate(sec_order):
        next_start = sec_order[i + 1][1] if i + 1 < len(sec_order) else total_row
        sec_ranges.append((k, sr, next_start - 1))

    # 排序（媒體內：秒數→區域）
    reg_map = {r: i for i, r in enumerate(REGIONS_ORDER + ["全省量販", "全省超市"])}
    def sort_key(x):
        return (x.get("seconds", 0), reg_map.get(x.get("region", ""), 999))

    grouped = {
        "全家廣播": sorted([r for r in rows if r["media_type"] == "全家廣播"], key=sort_key),
        "新鮮視":   sorted([r for r in rows if r["media_type"] == "新鮮視"], key=sort_key),
        "家樂福":   sorted([r for r in rows if r["media_type"] == "家樂福"], key=sort_key),
    }

    # 插列：用區塊最後一列當 pattern row（更接近原樣式）
    for k, sr, er in sorted(sec_ranges, key=lambda x: x[1], reverse=True):
        data = grouped.get(k, [])
        needed = len(data)
        if needed <= 0:
            continue
        existing = er - sr + 1
        if needed > existing:
            insert_n = needed - existing
            insert_at = er + 1
            ws.insert_rows(insert_at, amount=insert_n)
            pattern_row = er
            max_col = ws.max_column
            for rr in range(insert_at, insert_at + insert_n):
                copy_row_with_style_fix(ws, pattern_row, rr, max_col)

    # 插列後重新抓 Total / ranges（row 會變）
    total_cell = find_cell_exact(ws, meta["total_label"])
    if not total_cell:
        raise ValueError("插列後找不到 Total（模板可能異常）")
    total_row = total_cell[0]

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
        name = "通路廣播廣告" if m == "全家廣播" else "新鮮視廣告" if m == "新鮮視" else "家樂福"
        if format_type == "Shenghuo" and m == "全家廣播":
            name = "廣播通路廣告"
        return sanitize_text(prefix + name)

    # Dongwu：每個秒數區塊只在第一列寫 E/F/G/H（避免 merged-cell 子格覆寫）
    def group_by_seconds(data_list):
        by_sec = {}
        for r in data_list:
            by_sec.setdefault(r["seconds"], []).append(r)
        out = []
        for s in sorted(by_sec.keys()):
            sec_rows = sorted(by_sec[s], key=lambda x: reg_map.get(x.get("region", ""), 999))
            out.append((s, sec_rows))
        return out

    written_rows = []

    for m, sr, er in sec_ranges:
        data = grouped.get(m, [])
        if not data:
            continue

        if meta["station_merge"]:
            unmerge_col_overlap(ws, cols["station"], sr, er)
            merge_rng = f"{cols['station']}{sr}:{cols['station']}{sr + len(data) - 1}"
            ws.merge_cells(merge_rng)
            top_cell = ws[f"{cols['station']}{sr}"]
            top_cell.value = station_title(m)
            apply_center_style(top_cell)

        row_ptr = sr

        if format_type == "Dongwu":
            for sec_val, sec_rows in group_by_seconds(data):
                for idx, r in enumerate(sec_rows):
                    if not meta["station_merge"]:
                        c_station = ws[f"{cols['station']}{row_ptr}"]
                        c_station.value = station_title(m)
                        apply_center_style(c_station)

                    safe_write(ws, f"{cols['location']}{row_ptr}", region_display(r["region"]))
                    safe_write(ws, f"{cols['program']}{row_ptr}", int(r.get("program_num", 0)))

                    if idx == 0:
                        safe_write(ws, f"{cols['daypart']}{row_ptr}", sanitize_text(r["daypart"]))
                        if m == "家樂福":
                            safe_write(ws, f"{cols['seconds']}{row_ptr}", sanitize_text(f"{r['seconds']}秒"))
                        else:
                            safe_write(ws, f"{cols['seconds']}{row_ptr}", int(r["seconds"]))

                        safe_write(ws, f"{cols['rate']}{row_ptr}", r.get("rate_list", ""))
                        safe_write(ws, f"{cols['pkg']}{row_ptr}", r.get("pkg_display_val", ""))

                        # 置中保險（master cell）
                        for kcol in (cols["daypart"], cols["seconds"], cols["rate"], cols["pkg"]):
                            cc = ws[f"{kcol}{row_ptr}"]
                            if isinstance(cc, MergedCell):
                                mc = _get_master_cell(ws, cc)
                                if mc:
                                    cc = mc
                            apply_center_style(cc)

                    set_schedule(ws, row_ptr, meta["schedule_start_col"], meta["max_days"], r["schedule"])
                    spot_sum = sum(r["schedule"][:meta["max_days"]])
                    safe_write(ws, f"{meta['total_col']}{row_ptr}", spot_sum)

                    written_rows.append(row_ptr)
                    row_ptr += 1

        else:
            for r in data:
                if not meta["station_merge"]:
                    c_station = ws[f"{cols['station']}{row_ptr}"]
                    c_station.value = station_title(m)
                    apply_center_style(c_station)

                safe_write(ws, f"{cols['location']}{row_ptr}", region_display(r["region"]))
                safe_write(ws, f"{cols['program']}{row_ptr}", int(r.get("program_num", 0)))
                safe_write(ws, f"{cols['daypart']}{row_ptr}", sanitize_text(r["daypart"]))
                safe_write(ws, f"{cols['seconds']}{row_ptr}", sanitize_text(f"{r['seconds']}秒廣告"))
                safe_write(ws, f"{cols['proj_price']}{row_ptr}", r.get("pkg_display_val", 0) if isinstance(r.get("pkg_display_val", 0), int) else 0)

                set_schedule(ws, row_ptr, meta["schedule_start_col"], meta["max_days"], r["schedule"])
                spot_sum = sum(r["schedule"][:meta["max_days"]])
                safe_write(ws, f"{meta['total_col']}{row_ptr}", spot_sum)

                written_rows.append(row_ptr)
                row_ptr += 1

    # Total row
    eff_days = min((end_dt - start_dt).days + 1, meta["max_days"])
    daily_sums = [sum([x["schedule"][d] for x in rows if d < len(x["schedule"])]) for d in range(eff_days)]
    set_schedule(ws, total_row, meta["schedule_start_col"], meta["max_days"], daily_sums)
    safe_write(ws, f"{meta['total_col']}{total_row}", sum(daily_sums))

    total_pkg = sum([x.get("pkg_display_val", 0) for x in rows if isinstance(x.get("pkg_display_val", 0), int)])
    pkg_col = cols.get("pkg") or cols.get("proj_price")
    safe_write(ws, f"{pkg_col}{total_row}", total_pkg)

    # Footer
    lbl = meta["footer_labels"]
    def write_footer(key, val):
        pos = find_cell_exact(ws, lbl.get(key, ""))
        if pos:
            safe_write_rc(ws, pos[0], pos[1] + 1, int(val))

    make_fee = 10000
    pos_make = find_cell_exact(ws, lbl["make"])
    if pos_make:
        v = ws.cell(pos_make[0], pos_make[1] + 1).value
        if isinstance(v, (int, float)) and v > 0:
            make_fee = int(v)
        else:
            safe_write_rc(ws, pos_make[0], pos_make[1] + 1, make_fee)

    vat = int(round((total_pkg + make_fee) * 0.05))
    write_footer("vat", vat)
    write_footer("grand", total_pkg + make_fee + vat)

    # Remarks
    rem_pos = find_cell_exact(ws, "Remarks：")
    if rem_pos:
        for i, rm in enumerate(remarks_list):
            safe_write_rc(ws, rem_pos[0] + 1 + i, rem_pos[1], sanitize_text(rm))

    # ✅補強置中（避免 E/F/G/H 不置中）
    if format_type == "Dongwu" and written_rows:
        min_r, max_r = min(written_rows), total_row
        force_center_columns_range(ws, meta["force_center_cols"], min_r, max_r)

    # ✅PDF 擬真度關鍵：強制列印/縮放/print area（東吳 + 聲活 共用）
    try:
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2
        ws.page_margins.header = 0.1
        ws.page_margins.footer = 0.1

        last_col = meta["total_col"]
        last_row = total_row + 20
        if rem_pos:
            last_row = max(last_row, rem_pos[0] + 10)
        ws.print_area = f"A1:{last_col}{last_row}"
    except:
        pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 5) HTML Preview（僅供網頁顯示；PDF 以 xlsx→pdf 為主）
# =========================================================
def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks):
    header_cls = "bg-dw-head" if format_type == "Dongwu" else "bg-sh-head"
    media_order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    eff_days = min(days_cnt, 31 if format_type == "Dongwu" else 23)

    date_headers_1 = ""
    date_headers_2 = ""
    curr = start_dt
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for _ in range(eff_days):
        wd = curr.weekday()
        bg = "bg-weekend" if (format_type == "Dongwu" and wd >= 5) else header_cls
        if format_type == "Shenghuo":
            bg = header_cls
        date_headers_1 += f"<th class='{bg}'>{curr.day}</th>"
        date_headers_2 += f"<th class='{bg}'>{weekdays[wd]}</th>"
        curr += timedelta(days=1)

    if format_type == "Dongwu":
        cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(List)", "Package<br>(List)"]
    else:
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "專案價"]
    th_fixed = "".join([f"<th class='{header_cls}' rowspan='2'>{c}</th>" for c in cols_def])

    def region_rank(r):
        if r in REGIONS_ORDER:
            return REGIONS_ORDER.index(r)
        return 99

    rows_sorted = sorted(
        rows,
        key=lambda x: (media_order.get(x["media_type"], 99), x.get("seconds", 0), region_rank(x.get("region", "")))
    )

    tbody = ""
    media_group_counts = {}
    for r in rows_sorted:
        media_group_counts[r["media_type"]] = media_group_counts.get(r["media_type"], 0) + 1
    media_printed = {m: False for m in media_group_counts}

    # Dongwu：在預覽也模擬「秒數區塊」只顯示一次 daypart/size/rate/pkg
    last_key = None

    for r in rows_sorted:
        m = r["media_type"]
        sec = r.get("seconds", "")
        key = (m, sec)

        tbody += "<tr>"

        if format_type == "Dongwu":
            if not media_printed[m]:
                rowspan = media_group_counts[m]
                display_name = "全家便利商店<br>通路廣播廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
                tbody += f"<td rowspan='{rowspan}' class='align-left'>{display_name}</td>"
                media_printed[m] = True

            tbody += f"<td>{region_display(r['region'])}</td>"
            tbody += f"<td class='align-right'>{int(r.get('program_num', 0)):,}</td>"

            if key != last_key:
                tbody += f"<td>{sanitize_text(r['daypart'])}</td>"
                sec_txt = f"{r['seconds']}秒" if m == "家樂福" else f"{int(r['seconds'])}"
                tbody += f"<td>{sanitize_text(sec_txt)}</td>"

                rate = r.get("rate_list", "")
                pkg = r.get("pkg_display_val", "")
                rate_disp = f"{rate:,}" if isinstance(rate, int) else sanitize_text(str(rate))
                pkg_disp = f"{pkg:,}" if isinstance(pkg, int) else sanitize_text(str(pkg))
                tbody += f"<td class='align-right'>{rate_disp}</td>"
                tbody += f"<td class='align-right'>{pkg_disp}</td>"
            else:
                tbody += "<td></td><td></td><td></td><td></td>"

            for d in r["schedule"][:eff_days]:
                tbody += f"<td>{d}</td>"
            tbody += f"<td class='bg-total'>{sum(r['schedule'][:eff_days])}</td></tr>"
            last_key = key

        else:
            display_name = "全家便利商店<br>廣播通路廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
            tbody += f"<td class='align-left'>{display_name}</td>"
            tbody += f"<td>{region_display(r['region'])}</td>"
            tbody += f"<td class='align-right'>{int(r.get('program_num', 0)):,}</td>"
            tbody += f"<td>{sanitize_text(r['daypart'])}</td>"
            tbody += f"<td>{sanitize_text(str(r['seconds']) + '秒廣告')}</td>"
            pkg = r.get("pkg_display_val", 0)
            pkg_disp = f"{pkg:,}" if isinstance(pkg, int) else sanitize_text(str(pkg))
            tbody += f"<td class='align-right'>{pkg_disp}</td>"
            for d in r["schedule"][:eff_days]:
                tbody += f"<td>{d}</td>"
            tbody += f"<td class='bg-total'>{sum(r['schedule'][:eff_days])}</td></tr>"

    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]
    total_spots = sum(totals)
    total_pkg = sum([r.get("pkg_display_val", 0) for r in rows if isinstance(r.get("pkg_display_val", 0), int)])

    if format_type == "Dongwu":
        tfoot = f"<tr class='bg-total'><td colspan='5' class='align-left'>Total</td><td></td><td class='align-right'>{total_pkg:,}</td>"
    else:
        tfoot = f"<tr class='bg-total'><td colspan='5' class='align-left'>Total</td><td class='align-right'>{total_pkg:,}</td>"
    for t in totals:
        tfoot += f"<td>{t}</td>"
    tfoot += f"<td>{total_spots}</td></tr>"

    remarks_html = "<div style='margin-top:10px; font-size:12px; line-height:1.6; border-top:2px solid #000; padding-top:8px; text-align:left;'>" \
                   "<b style='text-decoration:underline;'>Remarks：</b><br>" + "<br>".join([sanitize_text(x) for x in remarks]) + "</div>"

    return f"""
    <div class="preview-wrapper">
      <div style="margin-bottom:8px;">
        <b>客戶：</b>{sanitize_text(c_name)} &nbsp; <b>產品：</b>{sanitize_text(p_display)}<br>
        <span style="color:#666;">走期：{start_dt} ~ {end_dt}</span>
      </div>
      <table class="excel-table">
        <thead>
          <tr>
            {th_fixed}{date_headers_1}
            <th class="{header_cls}" rowspan="2">檔次</th>
          </tr>
          <tr>{date_headers_2}</tr>
        </thead>
        <tbody>
          {tbody}
          {tfoot}
        </tbody>
      </table>
      {remarks_html}
    </div>
    """

# =========================================================
# 6) UI Main
# =========================================================
st.title("📺 媒體 Cue 表生成器 (v63.1：雲端 PDF 擬真提升)")

auto_tpl, source, msgs = load_default_template()
template_bytes = auto_tpl

if msgs:
    for m in msgs:
        st.caption(m)

if auto_tpl:
    st.success(f"✅ 已載入系統公版 ({source})")
else:
    st.warning("⚠️ 無法載入公版，請手動上傳模板")
    tpl = st.file_uploader("上傳 Excel 模板（1209-Cue表相關資料.xlsx）", type=["xlsx"])
    if tpl:
        template_bytes = tpl.read()

st.markdown("### 1. 選擇格式")
format_type = st.radio("", ["Dongwu", "Shenghuo"], horizontal=True, label_visibility="collapsed")

with st.container():
    st.markdown("### 2. 基本資料設定")
    with st.expander("📝 點擊展開/收合基本資料", expanded=True):
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
m1, m2, m3 = st.columns(3)
config = {}
rem_budget = 100

with m1:
    if st.checkbox("全家廣播", True):
        is_nat = st.checkbox("全省聯播（以6區出表）", True, key="rad_nat")
        regs = REGIONS_ORDER if is_nat else st.multiselect("區域", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_reg")
        secs = st.multiselect("秒數", DURATIONS, [20], key="rad_sec")
        share = st.slider("預算 %", 0, 100, 60, key="rad_share")
        rem_budget -= share

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

        config["全家廣播"] = {"is_national": is_nat, "regions": regs, "seconds": sorted(secs), "share": share, "sec_shares": sec_shares}

with m2:
    if st.checkbox("新鮮視", True):
        is_nat = st.checkbox("全省聯播（以6區出表）", False, key="fv_nat")
        regs = REGIONS_ORDER if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
        secs = st.multiselect("秒數", DURATIONS, [10], key="fv_sec")
        share = st.slider("預算 %", 0, rem_budget, min(20, rem_budget), key="fv_share")
        rem_budget -= share

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

        config["新鮮視"] = {"is_national": is_nat, "regions": regs, "seconds": sorted(secs), "share": share, "sec_shares": sec_shares}

with m3:
    if st.checkbox("家樂福", True):
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.info(f"剩餘預算: {rem_budget}%")

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

        config["家樂福"] = {"regions": ["全省"], "seconds": sorted(secs), "share": rem_budget, "sec_shares": sec_shares}

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
                calc_regs = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                display_regs = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]

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

                debug_logs.append({
                    "media": m, "sec": sec, "budget": s_budget,
                    "unit_cost": unit_net_sum * penalty, "spots": spots_final,
                    "std": db["Std_Spots"],
                    "status": "未達標" if penalty > 1 else "達標",
                    "reason": "懲罰 x1.1" if penalty > 1 else "費率正常"
                })

                for r in display_regs:
                    rate_list = int((db[r][0] / db["Std_Spots"]) * factor)
                    pkg_list = rate_list * spots_final
                    rows.append({
                        "media_type": m,
                        "region": r,
                        "program_num": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m == "新鮮視" else r, 0),
                        "daypart": db["Day_Part"],
                        "seconds": sec,
                        "spots": spots_final,
                        "schedule": sch,
                        "rate_list": rate_list,
                        "pkg_display_val": pkg_list
                    })

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
                debug_logs.append({
                    "media": m, "sec": sec, "budget": s_budget,
                    "unit_cost": unit_net * penalty, "spots": spots_final,
                    "std": base_std,
                    "status": "未達標" if penalty > 1 else "達標",
                    "reason": "懲罰 x1.1" if penalty > 1 else "費率正常"
                })

                rate_h = int((db["量販_全省"]["List"] / base_std) * factor)
                rows.append({
                    "media_type": m,
                    "region": "全省量販",
                    "program_num": STORE_COUNTS_NUM["家樂福_量販"],
                    "daypart": db["量販_全省"]["Day_Part"],
                    "seconds": sec,
                    "spots": spots_final,
                    "schedule": sch_h,
                    "rate_list": rate_h,
                    "pkg_display_val": rate_h * spots_final
                })

                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({
                    "media_type": m,
                    "region": "全省超市",
                    "program_num": STORE_COUNTS_NUM["家樂福_超市"],
                    "daypart": db["超市_全省"]["Day_Part"],
                    "seconds": sec,
                    "spots": spots_s,
                    "schedule": sch_s,
                    "rate_list": "計量販",
                    "pkg_display_val": "計量販"
                })

p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}" if rows else ""
rem = get_remarks_text(sign_deadline, billing_month, payment_date)

with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
    for log in debug_logs:
        color = "green" if log["status"] == "達標" else "red"
        st.markdown(
            f"**{log['media']} ({log['sec']}秒)**：預算 ${log['budget']:,.0f} | 檔次 {log['spots']} "
            f"→ <span style='color:{color}'><b>{log['status']}</b></span>（{log['reason']}）",
            unsafe_allow_html=True
        )

if rows:
    html = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem)
    st.components.v1.html(html, height=720, scrolling=True)

    if template_bytes:
        try:
            xlsx_bytes = generate_excel_from_template(format_type, start_date, end_date, client_name, p_str, rows, rem, template_bytes)

            st.download_button(
                "📥 下載 Excel（模板擬真版）",
                xlsx_bytes,
                file_name=f"Cue_{sanitize_text(client_name)}_{format_type}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ✅PDF：優先 xlsx→pdf（最像），失敗才 fallback WeasyPrint
            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_bytes)
            if pdf_bytes:
                st.download_button(
                    f"📄 下載 PDF（{method}：擬真度高）",
                    pdf_bytes,
                    file_name=f"Cue_{sanitize_text(client_name)}_{format_type}.pdf",
                    mime="application/pdf"
                )
                st.caption(f"PDF 產生方式：{method}")
            else:
                st.warning(f"⚠️ 目前環境無法用 Excel/LibreOffice 產 PDF（{method}）。改用 WeasyPrint fallback（擬真度會下降）。")
                pdf2, err2 = html_to_pdf_weasyprint(html)
                if pdf2:
                    st.download_button(
                        "📄 下載 PDF（WeasyPrint fallback）",
                        pdf2,
                        file_name=f"Cue_{sanitize_text(client_name)}_{format_type}.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.error(f"PDF 產出失敗：{err} | fallback：{err2}")

            # 小提示：雲端是否有 soffice
            soffice = find_soffice_path()
            if soffice:
                st.caption(f"✅ 雲端偵測到 LibreOffice: {soffice}")
            else:
                st.caption("⚠️ 雲端未偵測到 LibreOffice/soffice（若你有 packages.txt 安裝，重新部署後會出現）")

        except Exception as e:
            st.error(f"Excel 產出錯誤: {e}")
    else:
        st.warning("請上傳模板以啟用下載。")
else:
    st.info("請先設定媒體/秒數/預算，系統才會產生 Cue 表。")
