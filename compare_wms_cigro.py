import pandas as pd
import os
import re
import sys
import glob
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# 설정
# ==========================================
WMS_ORDER_COL    = "주문번호"
WMS_SUBORDER_COL = "부주문코드"
WMS_SKU_COL      = "상품코드"
WMS_QTY_COL      = "수량"
WMS_CHANNEL_COL  = "매출처"

CIGRO_ORDER_COL   = "order_id"
CIGRO_SKU_COL     = "match_sku"
CIGRO_QTY_COL     = "sku_적용_후_수량"
CIGRO_CHANNEL_COL = "channel_name"
CIGRO_STATUS_COL  = "status"

# Cigro와 주문번호 직접 비교 가능한 WMS 채널
# ON008(스마트스토어): WMS 주문번호 끝자리 1→0 변환하면 Cigro order_id와 일치
# ON054: Cigro와 주문번호 직접 비교 가능
MATCHABLE_WMS_CHANNELS = {"ON003", "ON017", "ON040", "ON046", "ON008", "ON054"}

# 파손 재발송: Cigro에 주문 없는 것이 정상 → 전체 검사 제외
EXEMPT_WMS_CHANNELS = {"ON032", "ON033"}

WMS_DATE_COL = "출고일자"
WMS_LOT_COL  = "LOT"


# ==========================================
# 유틸
# ==========================================

def resolve_input_file(base_dir):
    if len(sys.argv) > 1:
        path = sys.argv[1]
        return path if os.path.isabs(path) else os.path.join(base_dir, path)
    xlsx_files = glob.glob(os.path.join(base_dir, "*.xlsx"))
    xlsx_files = [f for f in xlsx_files if "분석결과" not in os.path.basename(f)]
    return max(xlsx_files, key=os.path.getmtime) if xlsx_files else None


def norm(v):
    """하이픈·공백·언더스코어 제거 후 대문자 (주문번호 정규화)"""
    if pd.isna(v):
        return ""
    return re.sub(r"[\s\-_]", "", str(v)).strip().upper()


def sep(char="=", n=60):
    print(char * n)


def to_cigro_key(order_norm, channel):
    """WMS 주문번호 → Cigro 비교용 키"""
    if channel == "ON008" and len(order_norm) == 16 and order_norm.endswith("1"):
        return order_norm[:-1] + "0"
    return order_norm


def make_dup_key(row):
    """중복 탐지용 복합키 (부주문코드+SKU 또는 주문번호+SKU)"""
    sub = norm(row[WMS_SUBORDER_COL])
    sku = norm(row[WMS_SKU_COL])
    if sub:
        return f"SUB||{sub}||{sku}"
    return f"ORD||{norm(row[WMS_ORDER_COL])}||{sku}"


def style_ws(ws, 판정_col=None):
    """헤더 스타일, 행 색상, 열 너비, 틀 고정"""
    HDR_FILL  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    HDR_FONT  = Font(color="FFFFFF", bold=True, size=10)
    OK_FILL   = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")   # 연초록
    BAD_FILL  = PatternFill(start_color="FFDCE0", end_color="FFDCE0", fill_type="solid")   # 연빨강
    WARN_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")   # 연노랑

    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    if 판정_col and ws.max_row > 1:
        headers = [c.value for c in ws[1]]
        ci = next((i + 1 for i, h in enumerate(headers) if h == 판정_col), None)
        if ci:
            for row in ws.iter_rows(min_row=2):
                val = str(row[ci - 1].value or "")
                if "정상" in val:
                    fill = OK_FILL
                elif "확인필요" in val:
                    fill = BAD_FILL
                elif "미지정" in val:
                    fill = WARN_FILL
                else:
                    fill = None
                if fill:
                    for cell in row:
                        cell.fill = fill

    for col in ws.columns:
        sample = list(col)[:101]
        max_len = max((len(str(c.value or "")) for c in sample), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 50)

    ws.freeze_panes = "A2"


# ==========================================
# 메인
# ==========================================
def main():
    base_dir   = os.path.dirname(os.path.abspath(__file__))
    input_path = resolve_input_file(base_dir)

    if not input_path or not os.path.exists(input_path):
        print("[오류] 비교할 엑셀 파일을 찾을 수 없습니다.")
        print("사용법: python compare_wms_cigro.py [파일명.xlsx]")
        return

    sep()
    print(f"  파일    : {os.path.basename(input_path)}")
    print(f"  실행시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    sep()

    # ──────────────────────────────────────────────
    # 1. 데이터 로드 및 필터
    # ──────────────────────────────────────────────
    wms   = pd.read_excel(input_path, sheet_name=0, dtype=str)
    cigro = pd.read_excel(input_path, sheet_name=1, dtype=str)

    wms_valid   = wms[wms[WMS_ORDER_COL].notna()].copy()
    wms_qty_num = pd.to_numeric(wms_valid[WMS_QTY_COL], errors="coerce").fillna(0)
    n_zero_qty  = (wms_qty_num <= 0).sum()
    wms_valid   = wms_valid[wms_qty_num > 0].copy()
    n_exempt    = wms_valid[WMS_CHANNEL_COL].isin(EXEMPT_WMS_CHANNELS).sum()
    wms_valid   = wms_valid[~wms_valid[WMS_CHANNEL_COL].isin(EXEMPT_WMS_CHANNELS)].copy()
    is_사은품   = wms_valid[WMS_ORDER_COL].astype(str).str.contains("사은품", na=False)
    n_사은품    = is_사은품.sum()
    wms_valid   = wms_valid[~is_사은품].copy()

    # 정규화 컬럼 추가
    wms_valid["_주문_n"]    = wms_valid[WMS_ORDER_COL].apply(norm)
    wms_valid["_부주문_n"]  = wms_valid[WMS_SUBORDER_COL].apply(norm)
    wms_valid["_sku"]       = wms_valid[WMS_SKU_COL].str.strip().str.upper()
    wms_valid["_dup_key"]   = wms_valid.apply(make_dup_key, axis=1)
    wms_valid["_wms_qty"]   = pd.to_numeric(wms_valid[WMS_QTY_COL], errors="coerce")
    wms_valid["_cigro_key"] = wms_valid.apply(
        lambda r: to_cigro_key(r["_주문_n"], r[WMS_CHANNEL_COL]), axis=1
    )

    # 채널 분류
    wms_matchable    = wms_valid[wms_valid[WMS_CHANNEL_COL].isin(MATCHABLE_WMS_CHANNELS)].copy()
    wms_unknown      = wms_valid[~wms_valid[WMS_CHANNEL_COL].isin(MATCHABLE_WMS_CHANNELS)].copy()
    n_unknown        = len(wms_unknown)
    unknown_channels = sorted(wms_unknown[WMS_CHANNEL_COL].unique().tolist()) if n_unknown else []

    # Cigro 정규화
    cigro["_order_n"]   = cigro[CIGRO_ORDER_COL].apply(norm)
    cigro["_sku"]       = cigro[CIGRO_SKU_COL].str.strip().str.upper()
    cigro_valid         = cigro[cigro["_order_n"] != ""].copy()
    cigro_valid["_qty"] = pd.to_numeric(cigro_valid[CIGRO_QTY_COL], errors="coerce")
    cigro_qty_map       = cigro_valid.groupby(["_order_n", "_sku"])["_qty"].sum()
    cigro_order_set     = set(cigro_valid["_order_n"])
    cigro_교환_set      = (
        set(cigro_valid[cigro_valid[CIGRO_STATUS_COL] == "교환"]["_order_n"])
        if CIGRO_STATUS_COL in cigro_valid.columns else set()
    )

    # ──────────────────────────────────────────────
    # 2. 현황 출력
    # ──────────────────────────────────────────────
    print(f"\n[ 데이터 현황 ]")
    print(f"  WMS 출고   : {len(wms_valid):,}건")
    print(f"               (원본 {len(wms):,}행  수량=0 {n_zero_qty}건 / 파손재발송 {n_exempt}건 / 사은품 {n_사은품}건 제외)")
    ch_counts    = cigro_valid[CIGRO_CHANNEL_COL].value_counts()
    cigro_ch_str = "  /  ".join(f"{ch} {cnt}건" for ch, cnt in ch_counts.items())
    print(f"  Cigro 주문 : {cigro_valid['_order_n'].nunique():,}건  ({cigro_ch_str})")
    if n_unknown > 0:
        print(f"  미지정 채널: {n_unknown}행  ({', '.join(unknown_channels)})  ← Cigro 비교 제외 중")

    # ──────────────────────────────────────────────
    # 검사 ① 중복출고
    # ──────────────────────────────────────────────
    sep()
    print("[ 검사 ① ] 중복 출고")
    print("  같은 주문번호+SKU가 WMS에 2번 이상 출고된 것을 탐지합니다.")
    sep("-")

    dup_count_map = wms_valid.groupby("_dup_key").size()
    dup_keys      = set(dup_count_map[dup_count_map >= 2].index)

    # 초기값 (중복 없는 경우 대비)
    agg_df        = pd.DataFrame()
    dup_wms       = pd.DataFrame()
    n_qty_ok      = 0
    n_lot_ok      = 0
    n_교환_ok     = 0
    n_ok_dup      = 0
    n_problem_dup = 0
    problem_keys  = set()

    if not dup_keys:
        print("  결과: 중복 출고 없음")
    else:
        dup_wms = wms_valid[wms_valid["_dup_key"].isin(dup_keys)].copy()

        # 키별 집계
        agg_df = dup_wms.groupby("_dup_key").agg(
            출고횟수=("_주문_n", "count"),
            WMS_총수량=("_wms_qty", "sum"),
        ).reset_index()

        # 주문 메타 조인
        key_meta = (
            wms_valid[wms_valid["_dup_key"].isin(dup_keys)]
            .groupby("_dup_key")[[WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL]]
            .first()
        )
        agg_df = agg_df.join(key_meta, on="_dup_key")

        # Cigro 수량 조회
        def get_cigro_qty(row):
            key = to_cigro_key(norm(row[WMS_ORDER_COL]), row[WMS_CHANNEL_COL])
            sku = norm(row[WMS_SKU_COL])
            return cigro_qty_map.get((key, sku), None)

        agg_df["Cigro_수량"]     = agg_df.apply(get_cigro_qty, axis=1)
        agg_df["Cigro_주문있음"] = agg_df["Cigro_수량"].notna()
        agg_df["수량차이"]       = agg_df["WMS_총수량"] - agg_df["Cigro_수량"]

        # 정상 판정
        qty_ok_set = set(
            agg_df[agg_df["Cigro_주문있음"] & (agg_df["수량차이"] == 0)]["_dup_key"]
        )

        lot_ok_set = set()
        if WMS_LOT_COL in wms.columns:
            for dk in dup_keys:
                rows  = dup_wms[dup_wms["_dup_key"] == dk]
                dates = rows[WMS_DATE_COL].dropna().unique() if WMS_DATE_COL in rows.columns else []
                lots  = rows[WMS_LOT_COL].dropna().unique()
                if len(dates) == 1 and len(lots) >= 2:
                    lot_ok_set.add(dk)

        교환_ok_set = set()
        for dk in dup_keys:
            rows = dup_wms[dup_wms["_dup_key"] == dk]
            if set(rows["_주문_n"]) & cigro_교환_set:
                교환_ok_set.add(dk)

        def judge(dk):
            if dk in qty_ok_set:  return "정상 - 수량합산 분할출고"
            if dk in lot_ok_set:  return "정상 - LOT 분할출고"
            if dk in 교환_ok_set: return "정상 - 교환 재출고"
            return "확인필요 ★"

        agg_df["판정"] = agg_df["_dup_key"].apply(judge)

        problem_keys  = set(agg_df[agg_df["판정"] == "확인필요 ★"]["_dup_key"])
        n_problem_dup = len(problem_keys)
        n_qty_ok      = len(qty_ok_set)
        n_lot_ok      = len(lot_ok_set - qty_ok_set)
        n_교환_ok     = len(교환_ok_set - qty_ok_set - lot_ok_set)
        n_ok_dup      = n_qty_ok + n_lot_ok + n_교환_ok

        print(f"  중복 조합: {len(dup_keys)}건  /  해당 WMS 행: {len(dup_wms)}행")
        print()
        print(f"  ┌ 정상 - 수량합산 분할출고 : {n_qty_ok}건")
        print(f"  ├ 정상 - LOT 분할출고      : {n_lot_ok}건")
        print(f"  ├ 정상 - 교환 재출고       : {n_교환_ok}건")
        print(f"  └ 확인 필요 ★              : {n_problem_dup}건  → 엑셀 \"중복출고\" 시트")
        if n_problem_dup > 0:
            prob_by_ch = (
                dup_wms[dup_wms["_dup_key"].isin(problem_keys)]
                .groupby(WMS_CHANNEL_COL).size().sort_values(ascending=False)
            )
            print(f"    └ 채널별: " + "  /  ".join(f"{ch} {cnt}행" for ch, cnt in prob_by_ch.items()))

    # ──────────────────────────────────────────────
    # 검사 ② 미등록출고
    # ──────────────────────────────────────────────
    sep()
    print("[ 검사 ② ] 미등록 출고")
    print(f"  대상 채널: {sorted(MATCHABLE_WMS_CHANNELS)}")
    sep("-")

    wms_unmatched  = wms_matchable[~wms_matchable["_cigro_key"].isin(cigro_order_set)].copy()
    n_matchable    = len(wms_matchable)
    n_confirmed    = n_matchable - len(wms_unmatched)
    n_unregistered = len(wms_unmatched)

    print(f"  대상 채널 WMS 출고   : {n_matchable:,}행")
    print(f"  Cigro 주문 확인됨    : {n_confirmed:,}행  ✓")
    if wms_unmatched.empty:
        print("  미등록 출고          : 없음  ✓")
    else:
        by_ch = wms_unmatched.groupby(WMS_CHANNEL_COL).size().sort_values(ascending=False)
        print(f"  미등록 출고          : {n_unregistered:,}행  ★  → 엑셀 \"미등록출고\" 시트")
        print(f"    └ 채널별: " + "  /  ".join(f"{ch} {cnt}행" for ch, cnt in by_ch.items()))

    # 미지정 채널 안내
    if n_unknown > 0:
        sep()
        print("[ 확인 필요 ] 미지정 채널 (MATCHABLE·EXEMPT 어디에도 속하지 않음)")
        sep("-")
        by_unk = wms_unknown.groupby(WMS_CHANNEL_COL).size().sort_values(ascending=False)
        for ch, cnt in by_unk.items():
            print(f"  {ch}: {cnt}행")
        print(f"  → 총 {n_unknown}행  →  엑셀 \"채널확인필요\" 시트")
        print(f"    ※ Cigro 비교 가능 채널이면 MATCHABLE_WMS_CHANNELS에 추가 필요")

    # Cigro 미출고 참고
    wms_cigro_key_set = set(wms_valid["_cigro_key"])
    cigro_missed_set  = cigro_order_set - wms_cigro_key_set

    sep()
    print("[ 참고 ] Cigro 주문 있으나 WMS 출고 없는 것  (날짜 범위 차이 가능)")
    sep("-")
    if cigro_missed_set:
        lines = []
        for ch_name, sub_c in cigro_valid.groupby(CIGRO_CHANNEL_COL):
            miss = len(set(sub_c["_order_n"]) - wms_cigro_key_set)
            if miss > 0:
                lines.append(f"{ch_name}: {miss}건")
        print(f"  미출고 주문: {len(cigro_missed_set):,}건  ({' / '.join(lines)})")
    else:
        print("  미출고 주문 없음  ✓")

    # ──────────────────────────────────────────────
    # 최종 요약
    # ──────────────────────────────────────────────
    sep()
    print("[ 최종 결과 요약 ]")
    sep("-")
    if n_problem_dup == 0 and n_unregistered == 0 and n_unknown == 0:
        print("  이상 없음  ✓  중복출고 · 미등록출고 · 미지정채널 모두 없음")
    else:
        icon = "★" if n_problem_dup > 0 else "✓"
        print(f"  {icon} 중복출고    : {n_problem_dup}건" + ("  → \"중복출고\" 시트" if n_problem_dup > 0 else ""))
        icon = "★" if n_unregistered > 0 else "✓"
        print(f"  {icon} 미등록출고  : {n_unregistered}행" + ("  → \"미등록출고\" 시트" if n_unregistered > 0 else ""))
        icon = "⚠" if n_unknown > 0 else "✓"
        print(f"  {icon} 미지정채널  : {n_unknown}행" + (f"  ({', '.join(unknown_channels)})  → \"채널확인필요\" 시트" if n_unknown > 0 else ""))

    if n_ok_dup > 0:
        parts = []
        if n_qty_ok  > 0: parts.append(f"수량합산 {n_qty_ok}건")
        if n_lot_ok  > 0: parts.append(f"LOT분할 {n_lot_ok}건")
        if n_교환_ok > 0: parts.append(f"교환재출고 {n_교환_ok}건")
        print(f"  ✓ 정상처리    : {n_ok_dup}건  ({', '.join(parts)})")
    print(f"  ✓ 제외됨      : 파손재발송 {n_exempt}건 / 사은품 {n_사은품}건")

    # ──────────────────────────────────────────────
    # 엑셀 저장
    # ──────────────────────────────────────────────
    key_cols    = [WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL, WMS_QTY_COL]
    extra_cols  = ["상품명", "고객명", "출고일자", "출고상태", "운송장번호"]
    front       = key_cols + [c for c in extra_cols if c in wms.columns]
    rest        = [c for c in wms.columns if c not in front and not c.startswith("_")]
    detail_cols = front + rest

    run_time    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    output_path = os.path.join(base_dir, "분석결과.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # ── 시트 0: 분析요약 ──────────────────────────
        cigro_unique_cnt = cigro_valid["_order_n"].nunique()
        rows = [
            ["항목", "값", "비고"],
            ["실행시각", run_time, ""],
            ["입력파일", os.path.basename(input_path), ""],
            ["", "", ""],
            ["── WMS 출고 현황 ──", "", ""],
            ["WMS 원본 행수", len(wms), "행"],
            ["  수량=0 제외", n_zero_qty, "행"],
            ["  파손재발송 제외", n_exempt, "행 (ON032, ON033)"],
            ["  사은품 제외", n_사은품, "행"],
            ["WMS 유효 행수", len(wms_valid), "행"],
            ["  Cigro 비교 가능 채널", len(wms_matchable), "행 / " + ", ".join(sorted(MATCHABLE_WMS_CHANNELS))],
            ["  미지정 채널", n_unknown, ("행 / " + ", ".join(unknown_channels)) if unknown_channels else "없음"],
            ["", "", ""],
            ["── Cigro 주문 현황 ──", "", ""],
            ["Cigro 총 주문", cigro_unique_cnt, "건"],
        ]
        for ch, cnt in cigro_valid[CIGRO_CHANNEL_COL].value_counts().items():
            rows.append([f"  {ch}", cnt, "건"])
        rows += [
            ["", "", ""],
            ["── 검사 결과 ──", "", ""],
            ["① 중복출고  확인필요", n_problem_dup, "→ '중복출고' 시트" if n_problem_dup > 0 else "없음"],
            ["    정상 - 수량합산 분할출고", n_qty_ok, "건"],
            ["    정상 - LOT 분할출고", n_lot_ok, "건"],
            ["    정상 - 교환 재출고", n_교환_ok, "건"],
            ["② 미등록출고", n_unregistered, "→ '미등록출고' 시트" if n_unregistered > 0 else "없음"],
            ["⚠ 채널확인필요 (미지정채널)", n_unknown, "→ '채널확인필요' 시트" if n_unknown > 0 else "없음"],
        ]
        pd.DataFrame(rows[1:], columns=rows[0]).to_excel(writer, sheet_name="분析요약", index=False)
        style_ws(writer.sheets["분析요약"])

        # ── 시트 1: 중복출고 ─────────────────────────
        if not agg_df.empty:
            JUDGE_ORDER = {
                "확인필요 ★":            0,
                "정상 - 수량합산 분할출고": 1,
                "정상 - LOT 분할출고":     2,
                "정상 - 교환 재출고":      3,
            }
            agg_df["_sort"] = agg_df["판정"].map(JUDGE_ORDER).fillna(9)
            agg_disp = (
                agg_df.sort_values(["_sort", WMS_CHANNEL_COL, WMS_ORDER_COL])
                .drop(columns=["_sort", "_dup_key"])
            )
            out_cols   = ["판정", WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL,
                          "출고횟수", "WMS_총수량", "Cigro_주문있음", "Cigro_수량", "수량차이"]
            out_cols   = [c for c in out_cols if c in agg_disp.columns]
            rename_map = {
                "출고횟수":       "WMS 출고횟수",
                "WMS_총수량":     "WMS 총수량",
                "Cigro_주문있음": "Cigro 주문확인",
                "Cigro_수량":     "Cigro 수량",
                "수량차이":       "수량차이 (WMS-Cigro)",
            }
            agg_disp[out_cols].rename(columns=rename_map).to_excel(
                writer, sheet_name="중복출고", index=False
            )
            style_ws(writer.sheets["중복출고"], 판정_col="판정")

            # 중복출고 상세 (확인필요만)
            if problem_keys and not dup_wms.empty:
                prob_detail = (
                    dup_wms[dup_wms["_dup_key"].isin(problem_keys)][detail_cols]
                    .sort_values([WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SKU_COL])
                    .copy()
                )
                prob_detail.insert(0, "총출고횟수",
                    prob_detail.apply(lambda r: dup_count_map.get(make_dup_key(r), ""), axis=1)
                )
                prob_detail.to_excel(writer, sheet_name="중복출고_상세", index=False)
                style_ws(writer.sheets["중복출고_상세"])

        # ── 시트 2: 미등록출고 ───────────────────────
        if n_unregistered > 0:
            (
                wms_unmatched[detail_cols]
                .sort_values([WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SKU_COL])
                .to_excel(writer, sheet_name="미등록출고", index=False)
            )
            style_ws(writer.sheets["미등록출고"])

        # ── 시트 3: 채널확인필요 ─────────────────────
        if n_unknown > 0:
            unk_disp = (
                wms_unknown[detail_cols]
                .sort_values([WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SKU_COL])
                .copy()
            )
            unk_disp.insert(0, "판정", "미지정채널 - 확인필요")
            unk_disp.to_excel(writer, sheet_name="채널확인필요", index=False)
            style_ws(writer.sheets["채널확인필요"], 판정_col="판정")

    sep()
    print(f"\n저장 완료 → 분析결과.xlsx")
    sheets_info = [("분析요약", "전체 통계 요약")]
    if not agg_df.empty:
        sheets_info.append(("중복출고", f"{len(agg_df)}건 전체 (확인필요 {n_problem_dup}건 / 정상 {n_ok_dup}건)"))
        if problem_keys:
            sheets_info.append(("중복출고_상세", "확인필요 조합의 WMS 원본 행"))
    if n_unregistered > 0:
        sheets_info.append(("미등록출고", f"{n_unregistered}행"))
    if n_unknown > 0:
        sheets_info.append(("채널확인필요", f"{n_unknown}행  ({', '.join(unknown_channels)})"))
    for sname, label in sheets_info:
        print(f"  [{sname}]  {label}")
    print()


if __name__ == "__main__":
    main()
