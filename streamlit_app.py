"""
스마트공장 공급기업 수집기 — SmartFind
"""
from __future__ import annotations

import io, json, time, random
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "output"
DATA_DIR.mkdir(exist_ok=True)

BASE_URL = "https://www.smart-factory.kr"
API_URL = f"{BASE_URL}/usr/bg/fs/ma/FixesSplySrch/selectFixesSplyContainer.do"

st.set_page_config(page_title="SmartFind", page_icon="🏭", layout="centered")

# ── CSS ──
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
.stApp{font-family:'Inter',system-ui,sans-serif}
#MainMenu,footer{visibility:hidden}
header[data-testid="stHeader"]{background:transparent!important}
.block-container{max-width:720px!important;padding-top:2rem!important}

.hero{text-align:center;padding:48px 0 32px}
.hero .brand{font-size:2.2rem;font-weight:900;letter-spacing:-2px;margin-bottom:4px}
.hero .brand span{color:#111}.hero .brand b{color:#0066FF}
.hero .sub{color:#9CA3AF;font-size:.82rem;font-weight:500;letter-spacing:2px}

.stat-row{display:flex;gap:12px;margin:20px 0}
.stat-box{flex:1;background:#F8FAFF;border:1px solid #E5E9F2;border-radius:12px;padding:16px;text-align:center}
.stat-box .num{font-size:1.5rem;font-weight:800;color:#111;letter-spacing:-1px}
.stat-box .lbl{font-size:.68rem;color:#9CA3AF;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-top:4px}

.done-box{background:linear-gradient(135deg,#F0FDF4,#ECFDF5);border:1px solid #BBF7D0;border-radius:14px;padding:28px;text-align:center;margin:24px 0}
.done-box .icon{font-size:2.5rem;margin-bottom:8px}
.done-box .msg{font-size:1.1rem;font-weight:700;color:#059669}
.done-box .detail{font-size:.82rem;color:#6B7280;margin-top:6px}
</style>""", unsafe_allow_html=True)


# ── Hero ──
st.markdown("""<div class="hero">
<div class="brand"><span>SMART</span><b>FIND</b></div>
<div class="sub">스마트공장 공급기업 수집기</div>
</div>""", unsafe_allow_html=True)


def _fetch_all(progress_bar, status_text) -> list[dict]:
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json", "Content-Type": "application/json",
        "Referer": f"{BASE_URL}/usr/bg/fs/ma/fixesSplySrch", "Origin": BASE_URL,
    })
    s.get(BASE_URL, timeout=15)

    first = s.post(API_URL, json={"key": "splyList"}, timeout=30).json()
    total_pages = int(first["paginationInfo"]["totalPageCount"])
    total_count = int(first["paginationInfo"]["totalCount"])
    all_items = list(first.get("splyInstList", []))
    status_text.text(f"전체 {total_count:,}건 발견 · 수집 중...")

    for page in range(2, total_pages + 1):
        time.sleep(random.uniform(0.15, 0.35))
        payload = {"key": "splyList", "currentPage": str(page), "recordCountPerPage": "50"}
        for attempt in range(3):
            try:
                items = s.post(API_URL, json=payload, timeout=30).json().get("splyInstList", [])
                all_items.extend(items)
                break
            except Exception:
                time.sleep(2 ** attempt)
        progress_bar.progress(page / total_pages, text=f"{page}/{total_pages} 페이지")

    progress_bar.progress(1.0, text="완료!")
    return all_items


def _build_excel(items: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "공급기업 목록"

    # ── 컬럼 정의: API 필드 → 엑셀 헤더 ──
    columns = [
        ("instNm",              "업체명",       22),
        ("splyInstSe",          "공급기업 구분", 20),
        ("splyCnstcNocs",       "구축건수",     10),
        ("splyLctnCdNm",        "지역",         8),
        ("splyWholEmpCnt",      "종사자규모",   12),
        ("splySlsAmt",          "매출규모(억)", 14),
        ("slsYr",               "매출연도",     10),
        ("splyDgstfnScr",       "만족도(5점)",  12),
        ("splyMfrcSpcltyFldNm", "주력분야",     28),
        ("splySpcltyFldNm",     "전문분야",     40),
        ("splyTpbizNm",         "특화업종",     45),
        ("rprsvNm",             "대표자",       10),
        ("fndnYmd",             "설립일",       12),
        ("instAddr",            "주소",         35),
        ("rprsTelno",           "대표전화",     16),
        ("rprsFxno",            "팩스",         16),
        ("hmpgAddr",            "홈페이지",     25),
        ("brno",                "사업자번호",   14),
    ]

    keys = [c[0] for c in columns]
    headers = [c[1] for c in columns]
    widths = [c[2] for c in columns]

    # ── 스타일 정의 ──
    hdr_font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1B3A5C", end_color="1B3A5C", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cell_font = Font(name="맑은 고딕", size=9)
    cell_font_bold = Font(name="맑은 고딕", size=9, bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    thin = Side(style="thin", color="D9DEE4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    stripe_fill = PatternFill(start_color="F5F7FA", end_color="F5F7FA", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    green_font = Font(name="맑은 고딕", size=9, color="059669", bold=True)
    blue_font = Font(name="맑은 고딕", size=9, color="2563EB")
    orange_font = Font(name="맑은 고딕", size=9, color="D97706", bold=True)

    # ── 헤더 행 ──
    ws.row_dimensions[1].height = 32
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = border

    # ── 데이터 행 ──
    for row_idx, item in enumerate(items, 2):
        is_stripe = row_idx % 2 == 0
        fill = stripe_fill if is_stripe else white_fill
        ws.row_dimensions[row_idx].height = 22

        for col_idx, key in enumerate(keys, 1):
            val = item.get(key, "")
            if val is None or str(val).strip() in ("None", "null", "."):
                val = ""

            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = border
            cell.fill = fill

            header = headers[col_idx - 1]

            # 숫자 포맷
            if key in ("splyCnstcNocs", "splyWholEmpCnt"):
                try:
                    val = int(val) if val else 0
                except (ValueError, TypeError):
                    val = 0
                cell.value = val
                cell.alignment = center
                cell.font = green_font if val > 0 else cell_font

            elif key == "splySlsAmt":
                try:
                    val = int(val) if val else 0
                except (ValueError, TypeError):
                    val = 0
                cell.value = f"{val}억" if val else "-"
                cell.alignment = center
                cell.font = cell_font_bold if val else cell_font

            elif key == "splyDgstfnScr":
                try:
                    val = float(val) if val else 0
                except (ValueError, TypeError):
                    val = 0
                if val > 0:
                    stars = "★" * int(val) + ("½" if val % 1 >= 0.5 else "")
                    cell.value = f"{stars} ({val})"
                    cell.font = orange_font
                else:
                    cell.value = "-"
                    cell.font = cell_font
                cell.alignment = center

            elif key == "slsAmt":
                try:
                    val = int(val) if val else 0
                except (ValueError, TypeError):
                    val = 0
                cell.value = val
                cell.number_format = '#,##0'
                cell.alignment = right
                cell.font = cell_font

            elif key == "instNm":
                cell.value = str(val)
                cell.font = cell_font_bold
                cell.alignment = left

            elif key == "splyInstSe":
                cell.value = str(val)
                cell.alignment = center
                if "제조" in str(val) and "서비스" in str(val):
                    cell.font = Font(name="맑은 고딕", size=9, color="7C3AED", bold=True)
                elif "서비스" in str(val):
                    cell.font = Font(name="맑은 고딕", size=9, color="0891B2", bold=True)
                elif "공방" in str(val):
                    cell.font = Font(name="맑은 고딕", size=9, color="D97706", bold=True)
                else:
                    cell.font = blue_font

            elif key in ("splyMfrcSpcltyFldNm", "splySpcltyFldNm", "splyTpbizNm", "instAddr"):
                cell.value = str(val) if val else "-"
                cell.font = cell_font
                cell.alignment = left

            elif key == "hmpgAddr":
                cell.value = str(val) if val else ""
                cell.font = Font(name="맑은 고딕", size=9, color="2563EB", underline="single") if val else cell_font
                cell.alignment = left

            else:
                cell.value = str(val) if val else "-"
                cell.font = cell_font
                cell.alignment = center

    # ── 컬럼 너비 ──
    for col_idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    # ── 고정 & 필터 ──
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    # ── 요약 시트 ──
    ws2 = wb.create_sheet("요약")
    ws2.sheet_properties.tabColor = "0066FF"

    summary_hdr_fill = PatternFill(start_color="0066FF", end_color="0066FF", fill_type="solid")
    summary_hdr_font = Font(name="맑은 고딕", bold=True, size=11, color="FFFFFF")
    summary_font = Font(name="맑은 고딕", size=10)
    summary_val_font = Font(name="맑은 고딕", size=12, bold=True, color="1B3A5C")

    ws2.column_dimensions["A"].width = 20
    ws2.column_dimensions["B"].width = 18

    summary_data = [
        ("전체 기업수", f"{len(items):,}건"),
        ("스마트제조", f"{sum(1 for i in items if '제조' in str(i.get('splyInstSe', '')))}건"),
        ("스마트서비스", f"{sum(1 for i in items if '서비스' in str(i.get('splyInstSe', '')))}건"),
        ("스마트공방", f"{sum(1 for i in items if '공방' in str(i.get('splyInstSe', '')))}건"),
        ("수집일시", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("데이터 출처", "smart-factory.kr"),
    ]

    ws2.cell(row=1, column=1, value="항목").font = summary_hdr_font
    ws2.cell(row=1, column=1).fill = summary_hdr_fill
    ws2.cell(row=1, column=1).alignment = center
    ws2.cell(row=1, column=1).border = border
    ws2.cell(row=1, column=2, value="값").font = summary_hdr_font
    ws2.cell(row=1, column=2).fill = summary_hdr_fill
    ws2.cell(row=1, column=2).alignment = center
    ws2.cell(row=1, column=2).border = border

    for r, (lbl, val) in enumerate(summary_data, 2):
        c1 = ws2.cell(row=r, column=1, value=lbl)
        c1.font = summary_font
        c1.border = border
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2 = ws2.cell(row=r, column=2, value=val)
        c2.font = summary_val_font
        c2.border = border
        c2.alignment = center

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── 최근 수집 결과 표시 ──
latest_files = sorted(DATA_DIR.glob("smart_factory_suppliers_*.json"), reverse=True)
if latest_files:
    with open(latest_files[0], "r", encoding="utf-8") as f:
        cached = json.load(f)
    ts = latest_files[0].stem.replace("smart_factory_suppliers_", "")

    mfg = sum(1 for i in cached if "제조" in str(i.get("splyInstSe", "")))
    svc = sum(1 for i in cached if "서비스" in str(i.get("splyInstSe", "")))
    ws_cnt = sum(1 for i in cached if "공방" in str(i.get("splyInstSe", "")))

    st.markdown(f"""<div class="stat-row">
    <div class="stat-box"><div class="num">{len(cached):,}</div><div class="lbl">전체 기업</div></div>
    <div class="stat-box"><div class="num">{mfg:,}</div><div class="lbl">스마트제조</div></div>
    <div class="stat-box"><div class="num">{svc:,}</div><div class="lbl">스마트서비스</div></div>
    <div class="stat-box"><div class="num">{ws_cnt:,}</div><div class="lbl">스마트공방</div></div>
    </div>""", unsafe_allow_html=True)

    st.caption(f"📅 마지막 수집: {ts}")

    excel_bytes = _build_excel(cached)
    st.download_button(
        "📥 엑셀 다운로드",
        excel_bytes,
        f"스마트공장_공급기업_{ts}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")

# ── 수집 버튼 ──
if st.button("🚀 수집 시작", type="primary", use_container_width=True):
    progress = st.progress(0, text="연결 중...")
    status = st.empty()

    items = _fetch_all(progress, status)

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    with open(DATA_DIR / f"smart_factory_suppliers_{ts}.json", "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

    st.markdown(f"""<div class="done-box">
    <div class="icon">✅</div>
    <div class="msg">수집 완료! {len(items):,}건</div>
    <div class="detail">{ts} · smart-factory.kr</div>
    </div>""", unsafe_allow_html=True)

    excel_bytes = _build_excel(items)
    st.download_button(
        "📥 엑셀 다운로드",
        excel_bytes,
        f"스마트공장_공급기업_{ts}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.rerun()
