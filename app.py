"""
테라로사 외부몰 주문 취합 — Streamlit 앱
"""
import re
import io
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════
#  상수
# ══════════════════════════════════════════════════════════════════
DRIP_GREEN   = "E2EFDA"
HEADER_GRAY  = "D9D9D9"
BORDER_COLOR = "AAAAAA"
HEADERS      = ["품목명", "중량", "옵션", "수량", "상품코드"]
COL_WIDTHS   = [42, 10, 35, 8, 18]
GROUP_ORDER  = {"세트": 0, "기타": 1, "드립백": 2, "원두": 3}
WEIGHT_ORDER = {"1kg": 0, "500g": 1, "250g": 2, "150g": 3}
WEIGHT_PAT   = re.compile(r"^(\d+(?:\.\d+)?(?:kg|g))$", re.IGNORECASE)

OPTION_REMOVE = [
    "/구매 안함", "/플러스", "[플러스] ", "/불필요", "/필요",
    "불필요", "필요", "/상자 없음", ":상자 없음", "상자 없음",
    "/포장 없음", "테라로사 시그니처 ", "[Online Exclusive] ",
    "[Online Exclusive/플러스] ",
]
OPTION_REPLACE = {
    "중간 분쇄(드립용)": "드립용",
    "드립&커피메이커":   "드립용",
    "가는 분쇄(에스프레소용)": "에스프레소용",
}


# ══════════════════════════════════════════════════════════════════
#  엑셀 서식 헬퍼
# ══════════════════════════════════════════════════════════════════
def _border():
    s = Side(style="thin", color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)

def _header(cell):
    cell.font      = Font(name="Arial", size=10, bold=True)
    cell.fill      = PatternFill("solid", start_color=HEADER_GRAY)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _border()

def _data(cell, bg=None, center=False):
    cell.font      = Font(name="Arial", size=10)
    cell.fill      = PatternFill("solid", start_color=bg) if bg else PatternFill()
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
    cell.border    = _border()

def _blank(ws, row, ncols=5):
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c, value="").border = _border()


# ══════════════════════════════════════════════════════════════════
#  처리 로직
# ══════════════════════════════════════════════════════════════════
def _clean_name(name: str) -> str:
    name = re.sub(r"\[테라로사\]\s*", "", str(name).strip())
    name = re.sub(r"\s+\d+(?:\.\d+)?(?:kg|g)$", "", name.strip())
    return name.strip()

def _clean_option(opt: str) -> str:
    opt = str(opt).strip()
    for r in OPTION_REMOVE:
        opt = opt.replace(r, "")
    for old, new in OPTION_REPLACE.items():
        opt = opt.replace(old, new)
    opt = opt.strip()
    return "" if opt == "단품" else opt

def _split_weight(raw_name: str, option: str):
    weight, opt_out = "", option
    if ":" in option:
        parts = [p.strip() for p in option.split(":")]
        wp, np_ = [], []
        for p in parts:
            (wp if WEIGHT_PAT.match(p) else np_).append(p)
        if wp:
            weight  = wp[0]
            opt_out = ":".join(np_)
    else:
        m = re.search(r"(\d+(?:\.\d+)?(?:kg|g))", option, re.IGNORECASE)
        if m:
            weight  = m.group(1)
            opt_out = option.replace(weight, "").strip()
    if not weight:
        m2 = re.search(r"\s+(\d+(?:\.\d+)?(?:kg|g))$", str(raw_name).strip(), re.IGNORECASE)
        if m2:
            weight = m2.group(1)
    return weight, opt_out.strip().strip(":").strip()

def _classify(name: str, weight: str) -> str:
    if "드립백" in name: return "드립백"
    if "세트"  in name: return "세트"
    if weight and re.search(r"\d+(?:kg|g)", str(weight), re.IGNORECASE): return "원두"
    return "기타"

def _sort(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_g"] = df["분류"].map(GROUP_ORDER)
    df["_w"] = df["중량"].apply(lambda w: WEIGHT_ORDER.get(str(w), 99))
    df = df.sort_values(["_g", "품목명", "_w", "옵션"])
    return df.drop(columns=["_g", "_w"]).reset_index(drop=True)

def _build_matcher(codes_df: pd.DataFrame):
    rows = []
    for _, r in codes_df.iterrows():
        cp = re.sub(r"\s+", " ", re.sub(r"\[테라로사\]\s*", "", str(r["상품명"]))).strip()
        rows.append((cp, str(r["옵션"]).strip(), str(r["코드"])))

    def match(product, weight, option):
        p = re.sub(r"\s+", " ", re.sub(r"\[테라로사\]\s*", "", str(product))).strip()
        w, o = str(weight).strip(), str(option).strip()
        results = []
        for cp, cw, code in rows:
            exact   = cp == p
            partial = (p in cp) or (cp in p)
            wm, om  = (cw == w if w else False), (cw == o if o else False)
            if exact:
                if w and wm:   results.append((1, code))
                elif o and om: results.append((2, code))
                else:          results.append((3, code))
            elif partial and (wm or om):
                results.append((4, code))
        return min(results)[1] if results else ""

    return match

def _agg(df: pd.DataFrame, include_channel: bool) -> pd.DataFrame:
    def key(r):
        no_weight = r["분류"] in ("세트", "드립백")
        parts = ([r["채널"]] if include_channel else []) + \
                [r["분류"], r["품목명"], "" if no_weight else r["중량"], r["옵션"]]
        return tuple(parts)
    df = df.copy()
    df["_k"] = df.apply(key, axis=1)
    return (df.groupby("_k", sort=False)
              .agg(채널=("채널","first"), 품목명=("품목명","first"),
                   중량=("중량","first"), 옵션=("옵션","first"),
                   수량=("수량","sum"),   분류=("분류","first"))
              .reset_index(drop=True))


# ══════════════════════════════════════════════════════════════════
#  시트 작성
# ══════════════════════════════════════════════════════════════════
def _sheet1(ws, df):
    ws.title = "전체 주문취합"
    for c, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        _header(ws.cell(row=1, column=c, value=h))
        ws.column_dimensions[get_column_letter(c)].width = w
    row, prev = 2, None
    for _, r in df.iterrows():
        if prev is not None and r["품목명"] != prev:
            _blank(ws, row); row += 1
        bg   = DRIP_GREEN if r["분류"] == "드립백" else None
        vals = [r["품목명"], r["중량"], r["옵션"], int(r["수량"]), r["상품코드"]]
        for c, v in enumerate(vals, 1):
            _data(ws.cell(row=row, column=c, value=v or ""), bg=bg, center=(c in [2,4,5]))
        row += 1; prev = r["품목명"]

def _sheet2(ws, df):
    ws.title = "원두 중량 합산"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 15
    for c, h in enumerate(["품목명", "중량(kg)"], 1):
        _header(ws.cell(row=1, column=c, value=h))
    def to_kg(w, q):
        m = re.match(r"(\d+(?:\.\d+)?)(kg|g)", str(w).lower())
        if not m: return 0
        v = float(m.group(1))
        return (v if m.group(2)=="kg" else v/1000) * q
    beans = df[df["분류"]=="원두"].copy()
    beans["_kg"] = beans.apply(lambda r: to_kg(r["중량"], r["수량"]), axis=1)
    for i, (_, r) in enumerate(beans.groupby("품목명")["_kg"].sum().reset_index().sort_values("품목명").iterrows(), 2):
        _data(ws.cell(row=i, column=1, value=r["품목명"]))
        _data(ws.cell(row=i, column=2, value=round(r["_kg"],3)), center=True)

def _sheet3(ws, df, ch_order):
    ws.title = "채널별 주문취합"
    for c, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        _header(ws.cell(row=1, column=c, value=h))
        ws.column_dimensions[get_column_letter(c)].width = w
    row, first = 2, True
    for ch in ch_order:
        sub = _sort(df[df["채널"]==ch].copy())
        if sub.empty: continue
        if not first: _blank(ws, row); row += 1
        for c in range(1, 6):
            cell = ws.cell(row=row, column=c, value=ch if c==1 else "")
            cell.font      = Font(name="Arial", size=10, bold=True)
            cell.fill      = PatternFill("solid", start_color=HEADER_GRAY)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _border()
        row += 1; prev = None
        for _, r in sub.iterrows():
            if prev is not None and r["품목명"] != prev:
                _blank(ws, row); row += 1
            bg   = DRIP_GREEN if r["분류"]=="드립백" else None
            vals = [r["품목명"], r["중량"], r["옵션"], int(r["수량"]), r["상품코드"]]
            for c, v in enumerate(vals, 1):
                _data(ws.cell(row=row, column=c, value=v or ""), bg=bg, center=(c in [2,4,5]))
            row += 1; prev = r["품목명"]
        first = False


# ══════════════════════════════════════════════════════════════════
#  메인 처리
# ══════════════════════════════════════════════════════════════════
def process(orders_bytes: bytes, codes_bytes: bytes) -> bytes:
    orders_df = pd.read_excel(io.BytesIO(orders_bytes), dtype=str)
    orders_df.columns = ["채널", "상품명", "옵션", "수량"]
    orders_df["수량"] = pd.to_numeric(orders_df["수량"], errors="coerce").fillna(0)

    codes_df = pd.read_excel(io.BytesIO(codes_bytes), dtype=str)
    codes_df.columns = ["코드", "상품명", "옵션"]
    codes_df = codes_df.fillna("")

    matcher = _build_matcher(codes_df)
    ch_order = list(dict.fromkeys(orders_df["채널"].tolist()))

    rows = []
    for _, r in orders_df.iterrows():
        opt    = _clean_option(str(r["옵션"]))
        weight, opt_clean = _split_weight(str(r["상품명"]), opt)
        rows.append({"채널": r["채널"], "품목명": _clean_name(str(r["상품명"])),
                     "중량": weight, "옵션": opt_clean, "수량": r["수량"]})
    df = pd.DataFrame(rows)
    df["분류"] = df.apply(lambda r: _classify(r["품목명"], r["중량"]), axis=1)

    total = _agg(df, include_channel=False)
    by_ch = _agg(df, include_channel=True)

    for frame in (total, by_ch):
        frame["상품코드"] = frame.apply(
            lambda r: matcher(r["품목명"], r["중량"], r["옵션"]), axis=1)

    total_s = _sort(total)
    by_ch_s = _sort(by_ch)

    wb  = Workbook()
    _sheet1(wb.active, total_s)
    _sheet2(wb.create_sheet(), total_s)
    _sheet3(wb.create_sheet(), by_ch_s, ch_order)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def validate(orders_bytes: bytes, result_bytes: bytes) -> dict:
    orders_df = pd.read_excel(io.BytesIO(orders_bytes), dtype=str)
    orders_df.columns = ["채널", "상품명", "옵션", "수량"]
    orders_df["수량"] = pd.to_numeric(orders_df["수량"], errors="coerce").fillna(0)

    orig_total = orders_df["수량"].sum()
    orig_by_ch = orders_df.groupby("채널")["수량"].sum().to_dict()

    res_total = pd.read_excel(io.BytesIO(result_bytes), sheet_name="전체 주문취합")
    res_qty   = res_total.dropna(subset=["품목명"])["수량"].sum()

    ch_sheet = pd.read_excel(io.BytesIO(result_bytes), sheet_name="채널별 주문취합")
    cur, recs = None, []
    for _, row in ch_sheet.iterrows():
        if pd.notna(row["품목명"]) and pd.isna(row["수량"]):
            cur = str(row["품목명"])
        elif pd.notna(row.get("수량")) and pd.notna(row.get("품목명")):
            recs.append({"채널": cur, "수량": float(row["수량"])})
    res_by_ch = pd.DataFrame(recs).groupby("채널")["수량"].sum().to_dict() if recs else {}

    return {
        "orig_total":   int(orig_total),
        "result_total": int(res_qty),
        "total_match":  abs(orig_total - res_qty) < 0.01,
        "channels": {
            ch: {"orig": int(orig_by_ch.get(ch,0)),
                 "result": int(res_by_ch.get(ch,0)),
                 "match": abs(orig_by_ch.get(ch,0) - res_by_ch.get(ch,0)) < 0.01}
            for ch in sorted(set(list(orig_by_ch)+list(res_by_ch)))
        },
    }


# ══════════════════════════════════════════════════════════════════
#  Streamlit UI
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="테라로사 주문 취합", page_icon="☕", layout="centered")

st.title("☕ 테라로사 외부몰 주문 취합")
st.caption("주문 파일과 상품코드 파일을 업로드하면 취합 엑셀을 자동으로 생성합니다.")
st.divider()

col1, col2 = st.columns(2)
with col1:
    orders_file = st.file_uploader(
        "📋 주문 파일 (.xlsx)",
        type=["xlsx"],
        help="A열:쇼핑몰명 / B열:상품명 / C열:옵션 / D열:수량",
    )
with col2:
    codes_file = st.file_uploader(
        "🗂️ 상품코드 파일 (.xlsx)",
        type=["xlsx"],
        help="외부몰상품코드.xlsx",
    )

st.divider()

if orders_file and codes_file:
    if st.button("🚀 취합 파일 생성", type="primary", use_container_width=True):
        with st.spinner("처리 중..."):
            try:
                orders_bytes = orders_file.read()
                codes_bytes  = codes_file.read()
                result_bytes = process(orders_bytes, codes_bytes)

                today    = date.today().strftime("%Y%m%d")
                filename = f"외부몰_주문취합_{today}.xlsx"

                st.success("✅ 취합 완료! 아래 버튼으로 다운로드하세요.")

                st.download_button(
                    label="⬇️ 취합 파일 다운로드",
                    data=result_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                # ── 수량 검증 ──
                st.divider()
                st.subheader("📊 수량 검증 결과")
                report = validate(orders_bytes, result_bytes)

                total_ok = report["total_match"]
                total_icon = "✅" if total_ok else "❌"
                st.metric(
                    label=f"{total_icon} 전체 수량",
                    value=f"{report['result_total']}개",
                    delta=f"원본 {report['orig_total']}개 {'일치' if total_ok else '불일치'}",
                    delta_color="normal" if total_ok else "inverse",
                )

                st.write("**채널별 수량**")
                rows_val = []
                for ch, info in report["channels"].items():
                    rows_val.append({
                        "채널":   ch,
                        "원본":   info["orig"],
                        "취합":   info["result"],
                        "일치 여부": "✅" if info["match"] else "❌",
                    })
                st.dataframe(
                    pd.DataFrame(rows_val).set_index("채널"),
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"❌ 오류가 발생했습니다: {e}")
else:
    st.info("👆 위에서 파일 두 개를 모두 업로드하면 버튼이 활성화됩니다.")

with st.expander("처리 규칙 보기"):
    st.markdown("""
| 구분 | 내용 |
|------|------|
| **브랜드명 제거** | `[테라로사]` 제거 (`[테라로사 X 미르]` 유지) |
| **중량 분리** | 상품명/옵션에서 `숫자g/kg` 추출 → B열 |
| **옵션 정리** | `단품`, `필요/불필요`, `상자 없음` 등 제거 |
| **분류** | 드립백 → 세트 → 원두(중량 있음) → 기타 |
| **합산** | 전체(채널 무관) / 채널별 각각 수량 합산 |
| **정렬** | 세트→기타→드립백→원두, 품목명 가나다, 중량 내림차순 |
| **상품코드** | 품목명+중량 → 품목명+옵션 → 품목명 순서로 매칭 |

**출력 시트:** 전체 주문취합 / 원두 중량 합산 / 채널별 주문취합
    """)
