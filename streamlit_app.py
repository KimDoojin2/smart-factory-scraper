"""
스마트공장 공급기업 탐색기 — Dashboard
InterX Design System 기반
"""
from __future__ import annotations

import json, time, random, logging, re
from datetime import datetime
from pathlib import Path
from collections import Counter

import streamlit as st
import pandas as pd
import requests

# ── Config ──
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "output"
DATA_DIR.mkdir(exist_ok=True)

BASE_URL = "https://www.smart-factory.kr"
API_URL = f"{BASE_URL}/usr/bg/fs/ma/FixesSplySrch/selectFixesSplyContainer.do"
PAGE_SIZE = 50

st.set_page_config(
    page_title="스마트공장 공급기업 탐색기",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
#  Design System — InterX-inspired Dark/Light
# =============================================================================
P = "#0066FF"
P_L = "#3388FF"
P_D = "#0052CC"
ACC = "#00C9A7"
GA = "#059669"; GB = "#2563EB"; GC = "#D97706"; GD = "#DC2626"

def _get_theme():
    return st.session_state.get("theme_mode", "light")

def _is_dark():
    return _get_theme() == "dark"

def _t():
    if _is_dark():
        return dict(
            bg="#0A0A0F", bg2="#111118", bg3="#1A1A24",
            card="#14141E", card_hover="#1C1C28",
            border="rgba(255,255,255,.08)", border2="rgba(255,255,255,.12)",
            text="#E8E8ED", text2="#A0A0B0", text3="#6B6B7B",
            shadow="rgba(0,0,0,.4)",
            nav_bg="linear-gradient(180deg,#06061A 0%,#0D1B3E 100%)",
            input_bg="#1A1A24", table_stripe="rgba(255,255,255,.02)",
        )
    return dict(
        bg="#F4F6FA", bg2="#FFFFFF", bg3="#EDF0F7",
        card="#FFFFFF", card_hover="#F8FAFF",
        border="rgba(0,0,0,.06)", border2="rgba(0,0,0,.10)",
        text="#111827", text2="#6B7280", text3="#9CA3AF",
        shadow="rgba(0,0,0,.04)",
        nav_bg="linear-gradient(180deg,#0D1B3E 0%,#162D5A 100%)",
        input_bg="#F9FAFB", table_stripe="rgba(0,0,0,.015)",
    )

# ── Intro animation ──
if "intro_shown" not in st.session_state:
    st.session_state.intro_shown = True
    st.markdown(
        '<style>@keyframes sf-fade{0%{opacity:0;transform:scale(.85) translateY(14px)}'
        '15%{opacity:1;transform:scale(1.02)}40%{opacity:1;transform:scale(1)}'
        '100%{opacity:0;transform:scale(.97) translateY(-8px)}}'
        '@keyframes sf-bg{0%,70%{opacity:1}100%{opacity:0;pointer-events:none;visibility:hidden}}'
        '.sf-intro{position:fixed;inset:0;z-index:99999;background:#0D1B3E;'
        'display:flex;align-items:center;justify-content:center;animation:sf-bg 2s ease forwards}'
        '.sf-intro .logo{animation:sf-fade 2s ease forwards;text-align:center}'
        '.sf-intro .mark{font-size:2.4rem;font-weight:900;letter-spacing:-1px;'
        'font-family:Inter,system-ui,sans-serif;color:#fff}'
        '.sf-intro .mark b{color:#00C9A7}'
        '.sf-intro .sub{color:rgba(255,255,255,.35);font-size:.72rem;letter-spacing:4px;'
        'margin-top:8px;font-weight:500}</style>'
        '<div class="sf-intro"><div class="logo">'
        '<div class="mark">SMART<b>FIND</b></div>'
        '<div class="sub">SUPPLIER INTELLIGENCE</div>'
        '</div></div>',
        unsafe_allow_html=True,
    )

t = _t()
st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

.stApp{{background:{t['bg']};font-family:'Inter',system-ui,-apple-system,sans-serif;color:{t['text']}}}
#MainMenu,footer{{visibility:hidden}}
header[data-testid="stHeader"]{{background:transparent !important;backdrop-filter:none !important}}
button[kind="headerNoPadding"],
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"]{{visibility:visible !important;z-index:999 !important}}

section[data-testid="stSidebar"]{{
    background:{t['card']};min-width:260px;max-width:260px;
    border-right:1px solid {t['border']};
}}
section[data-testid="stSidebar"] .stRadio label{{color:{t['text']} !important;font-weight:600;font-size:.82rem}}
section[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"]{{color:{t['text']} !important}}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label{{
    padding:10px 16px !important;border-radius:10px !important;margin:2px 0;transition:all .2s;
}}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover{{
    background:rgba(0,102,255,.06) !important;
}}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[data-checked="true"],
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:has(input:checked){{
    background:rgba(0,102,255,.10) !important;
    border-left:3px solid {P} !important;
}}
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p{{color:{t['text3']} !important}}

.block-container,[data-testid="stMainBlockContainer"]{{
    padding-top:0 !important;padding-bottom:0 !important;
    padding-left:1.5rem !important;padding-right:1.5rem !important;max-width:100% !important;
}}

/* Nav */
.nav-bar{{
    background:{t['nav_bg']};padding:12px 32px;margin:0 -1.5rem 0 -1.5rem;
    display:flex;align-items:center;justify-content:space-between;
    border-bottom:1px solid rgba(0,201,167,.15);
}}
.nav-bar .brand{{font-size:1.25rem;font-weight:900;letter-spacing:-1px}}
.nav-bar .brand span{{color:#fff}}.nav-bar .brand b{{color:{ACC}}}
.nav-bar .meta{{display:flex;align-items:center;gap:18px}}
.nav-bar .meta-item{{color:rgba(255,255,255,.4);font-size:.68rem;font-weight:500;letter-spacing:.5px;display:flex;align-items:center;gap:5px}}
.nav-bar .meta-dot{{width:5px;height:5px;border-radius:50%;background:#22C55E;display:inline-block}}

/* KPI */
.kpi-card{{
    background:{t['card']};border:1px solid {t['border']};border-radius:16px;padding:22px 20px;
    position:relative;overflow:hidden;transition:all .3s cubic-bezier(.25,.8,.25,1);
}}
.kpi-card::before{{
    content:'';position:absolute;top:0;left:0;right:0;height:3px;
    background:linear-gradient(90deg,{P},{ACC});transform:scaleX(0);transform-origin:left;
    transition:transform .3s;
}}
.kpi-card:hover{{border-color:rgba(0,102,255,.2);box-shadow:0 8px 30px {t['shadow']};transform:translateY(-2px)}}
.kpi-card:hover::before{{transform:scaleX(1)}}
.kpi-val{{font-size:1.9rem;font-weight:800;color:{t['text']};line-height:1;letter-spacing:-1px}}
.kpi-label{{font-size:.63rem;color:{t['text3']};margin-top:8px;font-weight:600;text-transform:uppercase;letter-spacing:1.2px}}

/* Section */
.sec-h{{display:flex;align-items:center;gap:10px;margin:28px 0 14px}}
.sec-h .dot{{width:3px;height:22px;border-radius:2px;background:linear-gradient(180deg,{P},{ACC})}}
.sec-h .txt{{font-size:.9rem;font-weight:700;color:{t['text']};letter-spacing:-.3px}}

/* Supplier Row */
.s-row{{
    background:{t['card']};border:1px solid {t['border']};border-radius:12px;
    padding:14px 18px;margin-bottom:6px;border-left:3px solid transparent;
    transition:all .2s;display:flex;align-items:center;gap:14px;cursor:pointer;
}}
.s-row:hover{{border-left-color:{P};background:{t['card_hover']};transform:translateX(2px)}}
.s-badge{{
    min-width:34px;height:26px;display:inline-flex;align-items:center;justify-content:center;
    border-radius:8px;font-size:.68rem;font-weight:800;color:#fff;flex-shrink:0;
}}
.s-title{{font-size:.84rem;font-weight:600;color:{t['text']};flex:1;line-height:1.4}}
.s-meta{{font-size:.7rem;color:{t['text3']};font-weight:500}}

/* Pill */
.pill{{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;font-size:.66rem;font-weight:600}}
.pill-mfg{{background:rgba(0,102,255,.08);color:{P}}}
.pill-svc{{background:rgba(0,201,167,.08);color:{ACC}}}
.pill-ws{{background:rgba(217,119,6,.08);color:{GC}}}
.pill-both{{background:rgba(139,92,246,.08);color:#8B5CF6}}

/* Detail card */
.detail-card{{
    background:{t['card']};border:1px solid {t['border']};border-radius:14px;
    padding:24px;margin:8px 0;
}}
.detail-card .d-label{{font-size:.68rem;color:{t['text3']};font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:4px}}
.detail-card .d-val{{font-size:.88rem;color:{t['text']};font-weight:500;line-height:1.5}}

/* Progress bar */
.progress-wrap{{background:{t['bg3']};border-radius:8px;height:6px;overflow:hidden;margin:6px 0}}
.progress-bar{{height:100%;border-radius:8px;transition:width .5s ease}}

/* Table */
.stDataFrame{{border-radius:12px;overflow:hidden}}
</style>""", unsafe_allow_html=True)


# =============================================================================
#  Data Layer
# =============================================================================
FIELD_MAP = {
    "instCd": "기관코드", "instNm": "업체명", "splyInstSe": "구분",
    "splySpcltyFldNm": "전문분야", "splyMfrcSpcltyFldNm": "제조전문분야",
    "splyTpbizNm": "업종", "rprsvNm": "대표자", "fndnYmd": "설립일",
    "instAddr": "주소", "instDaddr": "상세주소", "splyLctnCdNm": "지역",
    "rprsTelno": "대표전화", "rprsFxno": "팩스", "hmpgAddr": "홈페이지",
    "brno": "사업자번호", "splyWholEmpCnt": "직원수", "slsAmt": "매출액",
    "slsYr": "매출연도", "splySlsAmt": "매출액(억)", "splyCnstcNocs": "구축건수",
    "splyDgstfnScr": "만족도점수", "splyAprvYmd": "승인일",
}


def _load_latest_json() -> list[dict]:
    jsons = sorted(DATA_DIR.glob("smart_factory_suppliers_*.json"), reverse=True)
    if not jsons:
        return []
    with open(jsons[0], "r", encoding="utf-8") as f:
        return json.load(f)


def _fetch_from_api(progress_callback=None) -> list[dict]:
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Referer": f"{BASE_URL}/usr/bg/fs/ma/fixesSplySrch",
        "Origin": BASE_URL,
    })
    s.get(BASE_URL, timeout=15)

    first = s.post(API_URL, json={"key": "splyList"}, timeout=30).json()
    if "paginationInfo" not in first:
        return []

    total_pages = int(first["paginationInfo"]["totalPageCount"])
    total_count = int(first["paginationInfo"]["totalCount"])
    all_items = list(first.get("splyInstList", []))

    for page in range(2, total_pages + 1):
        time.sleep(random.uniform(0.2, 0.5))
        payload = {"key": "splyList", "currentPage": str(page), "recordCountPerPage": str(PAGE_SIZE)}
        for attempt in range(3):
            try:
                r = s.post(API_URL, json=payload, timeout=30)
                items = r.json().get("splyInstList", [])
                all_items.extend(items)
                break
            except Exception:
                time.sleep(2 ** attempt)

        if progress_callback:
            progress_callback(page / total_pages, f"{page}/{total_pages} 페이지 ({len(all_items)}건)")

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    with open(DATA_DIR / f"smart_factory_suppliers_{ts}.json", "w", encoding="utf-8") as f:
        json.dump(all_items, f, ensure_ascii=False, indent=2)

    return all_items


def _to_df(items: list[dict]) -> pd.DataFrame:
    if not items:
        return pd.DataFrame()
    df = pd.DataFrame(items)
    rename = {k: v for k, v in FIELD_MAP.items() if k in df.columns}
    df = df.rename(columns=rename)
    if "매출액" in df.columns:
        df["매출액"] = pd.to_numeric(df["매출액"], errors="coerce").fillna(0).astype(int)
    if "직원수" in df.columns:
        df["직원수"] = pd.to_numeric(df["직원수"], errors="coerce").fillna(0).astype(int)
    if "구축건수" in df.columns:
        df["구축건수"] = pd.to_numeric(df["구축건수"], errors="coerce").fillna(0).astype(int)
    if "만족도점수" in df.columns:
        df["만족도점수"] = pd.to_numeric(df["만족도점수"], errors="coerce").fillna(0)
    return df


# =============================================================================
#  Nav Bar
# =============================================================================
def _nav():
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    cnt = len(st.session_state.get("data", []))
    st.markdown(f"""<div class="nav-bar">
    <div class="brand"><span>SMART</span><b>FIND</b></div>
    <div class="meta">
        <div class="meta-item"><span class="meta-dot"></span> {cnt:,}건 수집</div>
        <div class="meta-item">{ts}</div>
    </div>
    </div>""", unsafe_allow_html=True)


# =============================================================================
#  Sidebar
# =============================================================================
def _sidebar():
    with st.sidebar:
        st.markdown(f"### 🏭 SmartFind")
        st.caption("스마트공장 공급기업 탐색기")

        page = st.radio(
            "메뉴",
            ["📊 대시보드", "🔍 기업 탐색", "📋 전체 목록", "⚡ 데이터 수집"],
            label_visibility="collapsed",
        )

        st.markdown("---")

        theme = st.radio("테마", ["☀️ 라이트", "🌙 다크"], horizontal=True, label_visibility="collapsed")
        st.session_state.theme_mode = "dark" if "다크" in theme else "light"

        st.markdown("---")
        st.caption(f"v1.0 · {datetime.now().strftime('%Y-%m-%d')}")

    return page


def _pill(se: str) -> str:
    if "제조" in str(se) and "서비스" in str(se):
        return '<span class="pill pill-both">제조+서비스</span>'
    if "제조" in str(se):
        return '<span class="pill pill-mfg">스마트제조</span>'
    if "서비스" in str(se):
        return '<span class="pill pill-svc">스마트서비스</span>'
    if "공방" in str(se):
        return '<span class="pill pill-ws">스마트공방</span>'
    return f'<span class="pill pill-mfg">{se}</span>'


def _se_color(se: str) -> str:
    if "서비스" in str(se):
        return ACC
    if "공방" in str(se):
        return GC
    return P


def _fmt_amt(v) -> str:
    try:
        v = int(v)
        if v >= 100_000_000:
            return f"{v / 100_000_000:.1f}억"
        if v >= 10_000:
            return f"{v / 10_000:.0f}만"
        return f"{v:,}"
    except (ValueError, TypeError):
        return "-"


# =============================================================================
#  Pages
# =============================================================================

def _page_dashboard(df: pd.DataFrame):
    if df.empty:
        st.info("수집 데이터가 없습니다. '데이터 수집' 메뉴에서 먼저 수집하세요.")
        return

    st.markdown('<div class="sec-h"><div class="dot"></div><div class="txt">핵심 지표</div></div>', unsafe_allow_html=True)

    total = len(df)
    mfg = len(df[df.get("구분", pd.Series(dtype=str)).str.contains("제조", na=False)])
    svc = len(df[df.get("구분", pd.Series(dtype=str)).str.contains("서비스", na=False)])
    ws = len(df[df.get("구분", pd.Series(dtype=str)).str.contains("공방", na=False)])
    regions = df["지역"].nunique() if "지역" in df.columns else 0
    avg_emp = int(df["직원수"].mean()) if "직원수" in df.columns and df["직원수"].sum() > 0 else 0
    avg_score = df["만족도점수"].mean() if "만족도점수" in df.columns else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    for col, val, label in [
        (c1, f"{total:,}", "전체 기업"),
        (c2, f"{mfg:,}", "스마트제조"),
        (c3, f"{svc:,}", "스마트서비스"),
        (c4, f"{ws:,}", "스마트공방"),
        (c5, f"{regions}", "활동 지역"),
        (c6, f"{avg_score:.1f}", "평균 만족도"),
    ]:
        col.markdown(f'<div class="kpi-card"><div class="kpi-val">{val}</div><div class="kpi-label">{label}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="sec-h"><div class="dot"></div><div class="txt">지역별 분포</div></div>', unsafe_allow_html=True)

    col_l, col_r = st.columns([3, 2])

    with col_l:
        if "지역" in df.columns:
            region_counts = df["지역"].value_counts().head(15)
            max_count = region_counts.max()
            html = ""
            for region, count in region_counts.items():
                pct = count / max_count * 100
                html += f"""<div style="display:flex;align-items:center;gap:10px;margin:4px 0">
                    <div style="width:40px;font-size:.78rem;font-weight:600;color:{t['text2']};text-align:right">{region}</div>
                    <div class="progress-wrap" style="flex:1">
                        <div class="progress-bar" style="width:{pct}%;background:linear-gradient(90deg,{P},{ACC})"></div>
                    </div>
                    <div style="width:50px;font-size:.75rem;font-weight:700;color:{t['text']}">{count:,}건</div>
                </div>"""
            st.markdown(html, unsafe_allow_html=True)

    with col_r:
        if "전문분야" in df.columns:
            st.markdown('<div class="sec-h"><div class="dot"></div><div class="txt">전문분야 TOP 10</div></div>', unsafe_allow_html=True)
            fields = []
            for val in df["전문분야"].dropna():
                fields.extend([x.strip() for x in str(val).split(",") if x.strip()])
            field_counts = Counter(fields).most_common(10)
            for fname, cnt in field_counts:
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;padding:5px 0;'
                    f'border-bottom:1px solid {t["border"]};font-size:.78rem">'
                    f'<span style="color:{t["text"]};font-weight:500">{fname}</span>'
                    f'<span style="color:{P};font-weight:700">{cnt:,}</span></div>',
                    unsafe_allow_html=True,
                )

    st.markdown('<div class="sec-h"><div class="dot"></div><div class="txt">구축건수 TOP 10</div></div>', unsafe_allow_html=True)
    if "구축건수" in df.columns and "업체명" in df.columns:
        top10 = df.nlargest(10, "구축건수")[["업체명", "구분", "전문분야", "지역", "구축건수", "만족도점수"]].reset_index(drop=True)
        top10.index = top10.index + 1
        st.dataframe(top10, use_container_width=True, height=400)


def _page_search(df: pd.DataFrame):
    if df.empty:
        st.info("수집 데이터가 없습니다.")
        return

    st.markdown('<div class="sec-h"><div class="dot"></div><div class="txt">기업 탐색</div></div>', unsafe_allow_html=True)

    fc1, fc2, fc3, fc4 = st.columns(4)

    with fc1:
        q = st.text_input("🔍 업체명 / 분야 검색", placeholder="예: MES, 로봇, ERP...")
    with fc2:
        se_options = ["전체"] + sorted(df["구분"].dropna().unique().tolist()) if "구분" in df.columns else ["전체"]
        se = st.selectbox("구분", se_options)
    with fc3:
        region_options = ["전체"] + sorted(df["지역"].dropna().unique().tolist()) if "지역" in df.columns else ["전체"]
        region = st.selectbox("지역", region_options)
    with fc4:
        sort_opt = st.selectbox("정렬", ["구축건수 ↓", "만족도 ↓", "직원수 ↓", "매출액 ↓", "업체명"])

    filtered = df.copy()
    if q:
        mask = filtered.apply(lambda r: q.lower() in " ".join(str(v).lower() for v in r.values), axis=1)
        filtered = filtered[mask]
    if se != "전체":
        filtered = filtered[filtered["구분"].str.contains(se, na=False)]
    if region != "전체":
        filtered = filtered[filtered["지역"] == region]

    sort_map = {
        "구축건수 ↓": ("구축건수", False),
        "만족도 ↓": ("만족도점수", False),
        "직원수 ↓": ("직원수", False),
        "매출액 ↓": ("매출액", False),
        "업체명": ("업체명", True),
    }
    scol, sasc = sort_map.get(sort_opt, ("업체명", True))
    if scol in filtered.columns:
        filtered = filtered.sort_values(scol, ascending=sasc)

    st.caption(f"검색 결과: **{len(filtered):,}건** / 전체 {len(df):,}건")

    per_page = 20
    total_pages = max(1, (len(filtered) + per_page - 1) // per_page)
    page_num = st.number_input("페이지", min_value=1, max_value=total_pages, value=1, label_visibility="collapsed")
    page_data = filtered.iloc[(page_num - 1) * per_page: page_num * per_page]

    for _, row in page_data.iterrows():
        name = row.get("업체명", "")
        se_val = row.get("구분", "")
        field = row.get("전문분야", "-")
        region_val = row.get("지역", "")
        cnstc = row.get("구축건수", 0)
        score = row.get("만족도점수", 0)
        emp = row.get("직원수", 0)

        pill_html = _pill(se_val)
        color = _se_color(se_val)

        with st.expander(f"**{name}**  ·  {region_val}  ·  {field[:30]}  ·  구축 {cnstc}건"):
            ic1, ic2 = st.columns(2)
            with ic1:
                info_items = [
                    ("업체명", name),
                    ("구분", se_val),
                    ("전문분야", row.get("전문분야", "-")),
                    ("제조전문분야", row.get("제조전문분야", "-")),
                    ("업종", row.get("업종", "-")),
                    ("대표자", row.get("대표자", "-")),
                    ("설립일", row.get("설립일", "-")),
                ]
                for lbl, val in info_items:
                    if val and str(val) not in ("", "-", "None", "nan"):
                        st.markdown(
                            f'<div style="display:flex;gap:8px;padding:3px 0;font-size:.8rem">'
                            f'<span style="color:{t["text3"]};min-width:80px;font-weight:600">{lbl}</span>'
                            f'<span style="color:{t["text"]}">{val}</span></div>',
                            unsafe_allow_html=True,
                        )

            with ic2:
                info_items2 = [
                    ("주소", row.get("주소", "-")),
                    ("대표전화", row.get("대표전화", "-")),
                    ("팩스", row.get("팩스", "-")),
                    ("홈페이지", row.get("홈페이지", "-")),
                    ("직원수", f"{emp}명" if emp else "-"),
                    ("매출액", _fmt_amt(row.get("매출액", 0))),
                    ("구축건수", f"{cnstc}건"),
                    ("만족도", f"{'⭐' * int(score)}{' (' + str(score) + ')' if score else '-'}"),
                ]
                for lbl, val in info_items2:
                    if val and str(val) not in ("", "-", "None", "nan", "0명", "0건"):
                        st.markdown(
                            f'<div style="display:flex;gap:8px;padding:3px 0;font-size:.8rem">'
                            f'<span style="color:{t["text3"]};min-width:80px;font-weight:600">{lbl}</span>'
                            f'<span style="color:{t["text"]}">{val}</span></div>',
                            unsafe_allow_html=True,
                        )

    st.caption(f"페이지 {page_num} / {total_pages}")


def _page_table(df: pd.DataFrame):
    if df.empty:
        st.info("수집 데이터가 없습니다.")
        return

    st.markdown('<div class="sec-h"><div class="dot"></div><div class="txt">전체 목록</div></div>', unsafe_allow_html=True)

    show_cols = ["업체명", "구분", "전문분야", "지역", "대표자", "직원수", "매출액(억)", "구축건수", "만족도점수", "대표전화", "홈페이지"]
    available = [c for c in show_cols if c in df.columns]

    st.dataframe(
        df[available],
        use_container_width=True,
        height=700,
        column_config={
            "매출액(억)": st.column_config.NumberColumn("매출액(억)", format="%d억"),
            "만족도점수": st.column_config.NumberColumn("만족도", format="%.1f"),
            "홈페이지": st.column_config.LinkColumn("홈페이지"),
        },
    )

    csv = df[available].to_csv(index=False).encode("utf-8-sig")
    st.download_button("📥 CSV 다운로드", csv, f"smart_factory_{datetime.now():%Y%m%d}.csv", "text/csv")


def _page_collect():
    st.markdown('<div class="sec-h"><div class="dot"></div><div class="txt">데이터 수집</div></div>', unsafe_allow_html=True)

    existing = sorted(DATA_DIR.glob("smart_factory_suppliers_*.json"), reverse=True)

    col1, col2 = st.columns([2, 1])
    with col1:
        if existing:
            latest = existing[0]
            ts_str = latest.stem.replace("smart_factory_suppliers_", "")
            with open(latest, "r", encoding="utf-8") as f:
                cnt = len(json.load(f))
            st.markdown(
                f'<div class="detail-card">'
                f'<div class="d-label">최근 수집</div>'
                f'<div class="d-val">{ts_str} · {cnt:,}건</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.warning("아직 수집된 데이터가 없습니다.")

    with col2:
        if existing:
            st.markdown(
                f'<div class="detail-card">'
                f'<div class="d-label">저장 파일</div>'
                f'<div class="d-val">{len(existing)}개</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown("---")

    if st.button("🚀 새로 수집 시작", type="primary", use_container_width=True):
        bar = st.progress(0, text="수집 준비 중...")
        status = st.empty()

        def on_progress(pct, msg):
            bar.progress(pct, text=msg)

        try:
            items = _fetch_from_api(progress_callback=on_progress)
            bar.progress(1.0, text="완료!")
            status.success(f"✅ {len(items):,}건 수집 완료!")
            st.session_state["data"] = items
            st.session_state["df"] = _to_df(items)
            time.sleep(1)
            st.rerun()
        except Exception as e:
            status.error(f"❌ 수집 실패: {e}")

    if existing:
        st.markdown("---")
        st.caption("📂 수집 이력")
        for f in existing[:5]:
            name = f.stem.replace("smart_factory_suppliers_", "")
            size_kb = f.stat().st_size / 1024
            st.markdown(
                f'<div style="display:flex;justify-content:space-between;padding:6px 0;'
                f'border-bottom:1px solid {t["border"]};font-size:.78rem">'
                f'<span style="color:{t["text"]}">{name}</span>'
                f'<span style="color:{t["text3"]}">{size_kb:.0f} KB</span></div>',
                unsafe_allow_html=True,
            )


# =============================================================================
#  Main
# =============================================================================
def main():
    if "data" not in st.session_state:
        items = _load_latest_json()
        st.session_state["data"] = items
        st.session_state["df"] = _to_df(items)

    _nav()
    page = _sidebar()
    df = st.session_state.get("df", pd.DataFrame())

    if "대시보드" in page:
        _page_dashboard(df)
    elif "탐색" in page:
        _page_search(df)
    elif "목록" in page:
        _page_table(df)
    elif "수집" in page:
        _page_collect()


main()
