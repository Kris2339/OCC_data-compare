import pandas as pd
import os
import re
import sys
import glob
from datetime import datetime

# ==========================================
# 설정 영역
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

# Cigro 주문번호와 WMS 주문번호를 비교할 수 있는 채널
# ON008(스마트스토어): WMS 주문번호 끝자리 1→0 변환하면 Cigro order_id와 일치
MATCHABLE_WMS_CHANNELS = {"ON003", "ON017", "ON040", "ON046", "ON008"}

# 파손 재발송 채널: Cigro에 주문이 없는 것이 정상 → 모든 검사에서 제외
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
    """하이픈·공백·언더스코어 제거 후 대문자 통일 (주문번호 정규화용)"""
    if pd.isna(v):
        return ""
    return re.sub(r"[\s\-_]", "", str(v)).strip().upper()


def sep(char="=", n=60):
    print(char * n)


def to_cigro_key(order_norm, channel):
    """
    WMS 주문번호 → Cigro order_id 변환
    - ON008(스마트스토어): WMS 주문번호 끝자리가 1, Cigro는 0
      예) WMS 2026040122335441 → Cigro 2026040122335440
    - 그 외: 그대로 사용
    """
    if channel == "ON008" and len(order_norm) == 16 and order_norm.endswith("1"):
        return order_norm[:-1] + "0"
    return order_norm


def make_dup_key(row):
    """
    중복 탐지용 키 생성
    - 부주문코드가 있으면 → 부주문코드 + SKU  (네이버 등: 부주문코드가 실제 주문 단위)
    - 부주문코드가 없으면 → 주문번호  + SKU  (카카오 등: 주문번호가 실제 주문 단위)
    """
    sub = norm(row[WMS_SUBORDER_COL])
    sku = norm(row[WMS_SKU_COL])
    if sub:
        return f"SUB||{sub}||{sku}"
    return f"ORD||{norm(row[WMS_ORDER_COL])}||{sku}"


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

    # ──────────────────────────────────────────────
    # 1. 데이터 로드 및 정규화
    # ──────────────────────────────────────────────
    sep()
    print(f"  파일    : {os.path.basename(input_path)}")
    print(f"  실행시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    sep()

    wms   = pd.read_excel(input_path, sheet_name=0, dtype=str)
    cigro = pd.read_excel(input_path, sheet_name=1, dtype=str)

    # WMS: 주문번호 없는 행, 수량 0 이하, 파손재발송(EXEMPT) 채널 제거
    wms_valid   = wms[wms[WMS_ORDER_COL].notna()].copy()
    wms_qty_num = pd.to_numeric(wms_valid[WMS_QTY_COL], errors="coerce").fillna(0)
    n_zero_qty  = (wms_qty_num <= 0).sum()
    wms_valid   = wms_valid[wms_qty_num > 0].copy()
    n_exempt    = wms_valid[WMS_CHANNEL_COL].isin(EXEMPT_WMS_CHANNELS).sum()
    wms_valid   = wms_valid[~wms_valid[WMS_CHANNEL_COL].isin(EXEMPT_WMS_CHANNELS)].copy()

    wms_valid["_주문_n"]   = wms_valid[WMS_ORDER_COL].apply(norm)
    wms_valid["_부주문_n"] = wms_valid[WMS_SUBORDER_COL].apply(norm)
    wms_valid["_sku"]      = wms_valid[WMS_SKU_COL].str.strip().str.upper()
    wms_valid["_dup_key"]  = wms_valid.apply(make_dup_key, axis=1)
    wms_valid["_wms_qty"]  = pd.to_numeric(wms_valid[WMS_QTY_COL], errors="coerce")
    # Cigro 매칭용 키: ON008은 끝자리 1→0 변환
    wms_valid["_cigro_key"] = wms_valid.apply(
        lambda r: to_cigro_key(r["_주문_n"], r[WMS_CHANNEL_COL]), axis=1
    )

    # Cigro: 정규화
    cigro["_order_n"] = cigro[CIGRO_ORDER_COL].apply(norm)
    cigro["_sku"]     = cigro[CIGRO_SKU_COL].str.strip().str.upper()
    cigro_valid       = cigro[cigro["_order_n"] != ""].copy()
    cigro_valid["_qty"] = pd.to_numeric(cigro_valid[CIGRO_QTY_COL], errors="coerce")
    cigro_qty_map       = cigro_valid.groupby(["_order_n", "_sku"])["_qty"].sum()
    cigro_order_set     = set(cigro_valid["_order_n"])
    # 교환 상태 주문: Cigro order_id가 WMS 주문번호와 동일 형식 (변환 불필요)
    cigro_교환_set = set(
        cigro_valid[cigro_valid[CIGRO_STATUS_COL] == "교환"]["_order_n"]
    ) if CIGRO_STATUS_COL in cigro_valid.columns else set()

    # ──────────────────────────────────────────────
    # 2. 데이터 현황 출력
    # ──────────────────────────────────────────────
    print(f"\n[ 데이터 현황 ]")
    print(f"  WMS 출고   : {len(wms_valid):,}건")
    print(f"               (원본 {len(wms):,}행에서 수량=0인 {n_zero_qty}건, 파손재발송(ON032/ON033) {n_exempt}건 제외)")
    ch_counts = cigro_valid[CIGRO_CHANNEL_COL].value_counts()
    cigro_ch_str = "  /  ".join(f"{ch} {cnt}건" for ch, cnt in ch_counts.items())
    print(f"  Cigro 주문 : {cigro_valid['_order_n'].nunique():,}건  ({cigro_ch_str})")

    # ──────────────────────────────────────────────
    # 검사 ① 중복 출고 탐지
    # ──────────────────────────────────────────────
    sep()
    print("[ 검사 ① ] 중복 출고")
    print("  같은 주문번호+SKU가 WMS에 2번 이상 출고된 것을 찾습니다.")
    print("  예) 3월 1일에 나간 주문이 3월 5일에 또 나간 경우")
    sep("-")

    dup_count_map = wms_valid.groupby("_dup_key").size()
    dup_keys      = set(dup_count_map[dup_count_map >= 2].index)
    n_normal      = (dup_count_map == 1).sum()

    if not dup_keys:
        print(f"  결과: 중복 출고 없음 — 전체 {n_normal:,}개 (주문+SKU) 조합이 모두 정상")
        dup_wms          = pd.DataFrame()
        agg              = pd.DataFrame()
        n_ok_dup         = 0
        n_problem_dup    = 0
        problem_dup_keys = set()
    else:
        # 각 중복 키의 WMS 집계
        key_to_주문  = wms_valid.groupby("_dup_key")[WMS_ORDER_COL].first()
        key_to_부주문 = wms_valid.groupby("_dup_key")[WMS_SUBORDER_COL].first()
        key_to_sku   = wms_valid.groupby("_dup_key")[WMS_SKU_COL].first()
        key_to_ch    = wms_valid.groupby("_dup_key")[WMS_CHANNEL_COL].first()

        dup_wms = wms_valid[wms_valid["_dup_key"].isin(dup_keys)].copy()
        agg = dup_wms.groupby("_dup_key").agg(
            WMS_출고횟수=("_주문_n", "count"),
            WMS_총수량=("_wms_qty", "sum")
        )
        agg["주문번호"]   = agg.index.map(key_to_주문)
        agg["부주문코드"] = agg.index.map(key_to_부주문)
        agg["상품코드"]   = agg.index.map(key_to_sku)
        agg["채널"]       = agg.index.map(key_to_ch)
        agg["Cigro_수량"] = agg.apply(
            lambda r: cigro_qty_map.get(
                (to_cigro_key(norm(r["주문번호"]), r["채널"]), norm(r["상품코드"])), None
            ), axis=1
        )
        agg["Cigro_주문확인"] = agg["Cigro_수량"].notna()
        agg["수량차이(WMS-Cigro)"] = agg["WMS_총수량"] - agg["Cigro_수량"]

        # 정상 분류 ①: WMS 총수량 = Cigro 수량 → 분할출고(정상)
        ok_qty_mask = agg["Cigro_주문확인"] & (agg["수량차이(WMS-Cigro)"] == 0)
        qty_ok_keys = set(agg[ok_qty_mask].index)

        # 정상 분류 ②: 같은 날 LOT만 달라 분할 출고된 경우
        lot_ok_keys = set()
        if WMS_LOT_COL in wms.columns:
            for dk in dup_keys:
                rows  = dup_wms[dup_wms["_dup_key"] == dk]
                dates = rows[WMS_DATE_COL].dropna().unique() if WMS_DATE_COL in rows.columns else []
                lots  = rows[WMS_LOT_COL].dropna().unique()
                if len(dates) == 1 and len(lots) >= 2:
                    lot_ok_keys.add(dk)

        # 정상 분류 ③: Cigro 교환 상태 주문 → 재출고가 맞음
        교환_ok_keys = set()
        for dk in dup_keys:
            rows = dup_wms[dup_wms["_dup_key"] == dk]
            if set(rows["_주문_n"]) & cigro_교환_set:
                교환_ok_keys.add(dk)

        ok_keys          = qty_ok_keys | lot_ok_keys | 교환_ok_keys
        problem_keys     = dup_keys - ok_keys
        n_ok_dup         = len(ok_keys)
        n_lot_ok         = len(lot_ok_keys - qty_ok_keys)
        n_교환_ok        = len(교환_ok_keys - qty_ok_keys - lot_ok_keys)
        n_problem_dup    = len(problem_keys)
        problem_dup_keys = problem_keys

        agg = agg.reset_index(drop=True)
        agg["_dup_key"] = dup_wms.groupby("_dup_key")["_dup_key"].first().values

        # 결과 출력
        total_dup_rows = len(dup_wms)
        print(f"  발견: {len(dup_keys)}개 (주문+SKU) 조합  /  해당 WMS 행 합계 {total_dup_rows}행")
        print()
        n_qty_ok = len(qty_ok_keys)
        print(f"  ┌ 정상 (수량합산=Cigro, 분할출고)     : {n_qty_ok}개 조합  → 엑셀 미포함")
        print(f"  ├ 정상 (같은날 LOT 분할출고)          : {n_lot_ok}개 조합  → 엑셀 미포함")
        print(f"  ├ 정상 (Cigro 교환 재출고)            : {n_교환_ok}개 조합  → 엑셀 미포함")
        print(f"  └ 확인 필요                           : {n_problem_dup}개 조합  → 엑셀 \"중복출고\" 시트")

        if n_problem_dup > 0:
            prob_rows  = dup_wms[dup_wms["_dup_key"].isin(problem_dup_keys)]
            prob_by_ch = prob_rows.groupby(WMS_CHANNEL_COL).size().sort_values(ascending=False)
            print(f"    └ 채널별: " + "  /  ".join(f"{ch} {cnt}행" for ch, cnt in prob_by_ch.items()))

    # ──────────────────────────────────────────────
    # 검사 ② 미등록 출고 탐지
    # ──────────────────────────────────────────────
    sep()
    print("[ 검사 ② ] 미등록 출고 (Cigro 주문 없이 WMS에서 출고된 것)")
    print("  Cigro에 주문 기록이 없는데 WMS에서 출고가 나간 경우를 찾습니다.")
    print(f"  분석 대상 채널: {sorted(MATCHABLE_WMS_CHANNELS)}")
    print(f"    ※ ON008(스마트스토어): WMS 주문번호 끝자리 1→0 변환 후 Cigro와 비교")
    sep("-")

    wms_matchable   = wms_valid[wms_valid[WMS_CHANNEL_COL].isin(MATCHABLE_WMS_CHANNELS)].copy()
    wms_unmatched   = wms_matchable[~wms_matchable["_cigro_key"].isin(cigro_order_set)].copy()
    n_matchable_wms = len(wms_matchable)
    n_confirmed     = n_matchable_wms - len(wms_unmatched)

    print(f"  대상 채널 WMS 출고    : {n_matchable_wms:,}행")
    print(f"  Cigro 주문 확인됨     : {n_confirmed:,}행  ✓")
    if wms_unmatched.empty:
        print(f"  Cigro 주문 없는 출고  : 없음  ✓")
    else:
        by_ch = wms_unmatched.groupby(WMS_CHANNEL_COL).size().sort_values(ascending=False)
        print(f"  Cigro 주문 없는 출고  : {len(wms_unmatched):,}행  ★  → 엑셀 \"미등록출고\" 시트")
        print(f"    └ 채널별: " + "  /  ".join(f"{ch} {cnt}행" for ch, cnt in by_ch.items()))

    # ──────────────────────────────────────────────
    # 참고: Cigro 미출고 현황
    # ──────────────────────────────────────────────
    # Cigro order_id 집합 vs WMS _cigro_key 집합으로 비교 (전 채널 통합)
    wms_cigro_key_set = set(wms_valid["_cigro_key"])
    cigro_missed_set  = set(cigro_valid["_order_n"]) - wms_cigro_key_set

    sep()
    print("[ 참고 ] Cigro에 주문은 있는데 WMS 출고 기록이 없는 것")
    print("  (날짜 범위 차이로 발생 가능 — 주문일과 출고일이 다를 수 있음)")
    sep("-")
    if cigro_missed_set:
        lines = []
        for ch_name, sub_c in cigro_valid.groupby(CIGRO_CHANNEL_COL):
            ids  = set(sub_c["_order_n"])
            miss = len(ids - wms_cigro_key_set)
            if miss > 0:
                lines.append(f"{ch_name}: {miss}건")
        print(f"  미출고 주문: {len(cigro_missed_set):,}건  ({' / '.join(lines)})")
    else:
        print(f"  미출고 주문 없음  ✓")

    # ──────────────────────────────────────────────
    # 최종 요약
    # ──────────────────────────────────────────────
    n_unregistered = len(wms_unmatched)
    sep()
    print("[ 최종 결과 요약 ]")
    sep("-")
    if n_problem_dup == 0 and n_unregistered == 0:
        print("  문제 없음  ✓  중복 출고 및 미등록 출고 모두 발견되지 않았습니다.")
    else:
        if n_problem_dup > 0:
            print(f"  ★ 중복 출고   : {n_problem_dup}개 조합  → 엑셀 \"중복출고\" 시트 확인")
        else:
            print(f"  ✓ 중복 출고   : 없음")
        if n_unregistered > 0:
            print(f"  ★ 미등록 출고 : {n_unregistered}행  → 엑셀 \"미등록출고\" 시트 확인")
        else:
            print(f"  ✓ 미등록 출고 : 없음")
    if n_ok_dup > 0:
        parts = []
        if n_qty_ok  > 0: parts.append(f"수량합산 분할출고 {n_qty_ok}건")
        if n_lot_ok  > 0: parts.append(f"LOT 분할출고 {n_lot_ok}건")
        if n_교환_ok > 0: parts.append(f"교환 재출고 {n_교환_ok}건")
        print(f"  ✓ 정상 처리   : {n_ok_dup}개 조합  ({', '.join(parts)})")
    print(f"  ※ 제외됨     : ON032/ON033 파손재발송 {n_exempt}건 (검사 대상 아님)")

    # ──────────────────────────────────────────────
    # 엑셀 저장
    # ──────────────────────────────────────────────
    key_cols   = [WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL, WMS_QTY_COL]
    extra_cols = ["상품명", "고객명", "출고일자", "출고상태", "운송장번호"]
    front = key_cols + [c for c in extra_cols if c in wms.columns]
    rest  = [c for c in wms.columns if c not in front and not c.startswith("_")]
    detail_cols = front + rest

    sheets = []

    # 시트 ①: 중복출고
    if n_problem_dup > 0:
        mask_problem = agg["_dup_key"].isin(problem_dup_keys)

        # 요약 (한 행 = 하나의 중복 조합)
        summary_cols = ["채널", "주문번호", "부주문코드", "상품코드",
                        "WMS_출고횟수", "WMS_총수량", "Cigro_주문확인",
                        "Cigro_수량", "수량차이(WMS-Cigro)"]
        summary_dup = agg.loc[mask_problem, summary_cols].sort_values(["채널", "주문번호"])

        # 상세 (중복에 해당하는 WMS 원본 행 전체)
        prob_keys    = set(agg.loc[mask_problem, "_dup_key"].dropna())
        detail_dup   = (
            dup_wms[dup_wms["_dup_key"].isin(prob_keys)][detail_cols]
            .sort_values([WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SKU_COL])
        )
        # 출고횟수 정보 추가 (몇 번째 출고인지 쉽게 보이도록)
        detail_dup.insert(0, "총출고횟수",
            detail_dup.apply(lambda r: dup_count_map.get(
                make_dup_key(r), ""), axis=1)
        )

        sheets += [
            ("중복출고_요약", summary_dup,
             "★ 같은 주문+SKU가 2번 이상 출고된 조합 목록 (WMS_출고횟수 확인)"),
            ("중복출고_상세", detail_dup,
             "★ 중복 출고에 해당하는 WMS 원본 행 (총출고횟수가 2 이상인 것)"),
        ]

    # 시트 ②: 미등록출고
    if n_unregistered > 0:
        unreg_display = wms_unmatched[detail_cols].sort_values(
            [WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SKU_COL]
        )
        sheets += [
            ("미등록출고", unreg_display,
             "★ Cigro에 주문 기록 없는데 WMS에서 출고된 행"),
        ]

    if sheets:
        output_path = os.path.join(base_dir, "분석결과.xlsx")
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sname, df, _ in sheets:
                df.to_excel(writer, sheet_name=sname, index=False)
        sep()
        print(f"\n저장 완료 → 분석결과.xlsx")
        for sname, df, label in sheets:
            print(f"  [{sname}]  {label}")
    else:
        sep()
        print(f"\n저장할 문제 항목 없음")
    print()


if __name__ == "__main__":
    main()
