import pandas as pd
import os
import re
import sys
import glob

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

# [주의] 네이버 스마트스토어(SMART_STORE / ON008)는
#  WMS가 '상품주문번호(발주번호)'를, Cigro가 '주문번호'를 저장하는
#  별개의 네이버 ID 체계를 사용하므로 주문번호로 직접 매칭 불가.
SMART_STORE_WMS_CHANNELS      = {"ON008"}
ID_UNMATCHABLE_CIGRO_CHANNELS = {"SMART_STORE"}


# ==========================================
# 유틸
# ==========================================

def resolve_input_file(base_dir):
    if len(sys.argv) > 1:
        path = sys.argv[1]
        return path if os.path.isabs(path) else os.path.join(base_dir, path)
    xlsx_files = glob.glob(os.path.join(base_dir, "*.xlsx"))
    xlsx_files = [f for f in xlsx_files if "중복출고" not in os.path.basename(f)]
    return max(xlsx_files, key=os.path.getmtime) if xlsx_files else None


def norm(v):
    """하이픈/공백/언더스코어 제거 후 대문자"""
    if pd.isna(v):
        return ""
    return re.sub(r"[\s\-_]", "", str(v)).strip().upper()


def sep(char="=", n=65):
    print(char * n)


def make_dup_key(row):
    """중복 탐지 키: 부주문코드가 있으면 부주문코드+SKU, 없으면 주문번호+SKU"""
    sub = norm(row[WMS_SUBORDER_COL])
    sku = norm(row[WMS_SKU_COL])
    if sub:
        return f"SUB||{sub}||{sku}"
    else:
        return f"ORD||{norm(row[WMS_ORDER_COL])}||{sku}"


# ==========================================
# 메인
# ==========================================
def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = resolve_input_file(base_dir)

    if not input_path or not os.path.exists(input_path):
        print("[오류] 비교할 엑셀 파일을 찾을 수 없습니다.")
        print("사용법: python compare_wms_cigro.py [파일명.xlsx]")
        return

    fname = os.path.basename(input_path)

    # ------------------------------------------
    # 1. 로드 및 정규화
    # ------------------------------------------
    sep()
    print(f"[ 입력 파일 ] {fname}")
    sep()

    wms   = pd.read_excel(input_path, sheet_name=0, dtype=str)
    cigro = pd.read_excel(input_path, sheet_name=1, dtype=str)

    # WMS: 주문번호 없는 행 제거, 수량 0 이하 제거
    wms_valid = wms[wms[WMS_ORDER_COL].notna()].copy()
    wms_qty_num = pd.to_numeric(wms_valid[WMS_QTY_COL], errors="coerce").fillna(0)
    n_zero_qty  = (wms_qty_num <= 0).sum()
    wms_valid   = wms_valid[wms_qty_num > 0].copy()

    wms_valid["_주문_n"]   = wms_valid[WMS_ORDER_COL].apply(norm)
    wms_valid["_부주문_n"] = wms_valid[WMS_SUBORDER_COL].apply(norm)
    wms_valid["_sku"]      = wms_valid[WMS_SKU_COL].str.strip().str.upper()
    wms_valid["_dup_key"]  = wms_valid.apply(make_dup_key, axis=1)
    wms_valid["_key_type"] = wms_valid["_부주문_n"].apply(
        lambda x: "부주문코드" if x else "주문번호"
    )
    wms_valid["_wms_qty"]  = pd.to_numeric(wms_valid[WMS_QTY_COL], errors="coerce")

    # Cigro: 정규화 (WMS 요약 출력 전에 수행 — 수량 일치 판단에 필요)
    cigro["_order_n"] = cigro[CIGRO_ORDER_COL].apply(norm)
    cigro["_sku"]     = cigro[CIGRO_SKU_COL].str.strip().str.upper()
    cigro_valid       = cigro[cigro["_order_n"] != ""].copy()
    cigro_valid["_qty"] = pd.to_numeric(cigro_valid[CIGRO_QTY_COL], errors="coerce")

    # Cigro 수량 집계 (order_id + SKU 기준 합산)
    cigro_qty_map = cigro_valid.groupby(["_order_n", "_sku"])["_qty"].sum()

    # ------------------------------------------
    # WMS 중복 탐지 + 사전 분류 (Cigro와 비교)
    #
    # 분류 기준:
    #   수량일치(OK) : WMS 총수량 == Cigro 수량  → 정상 (분할출고 등)
    #   과잉출고     : WMS 총수량 >  Cigro 수량  → 실제 문제
    #   부족출고     : WMS 총수량 <  Cigro 수량  → 확인 필요
    #   Cigro없음    : Cigro에 해당 주문 없음     → 확인 필요
    # ------------------------------------------
    dup_count_map = wms_valid.groupby("_dup_key").size()
    dup_keys      = set(dup_count_map[dup_count_map >= 2].index)

    key_to_주문  = wms_valid.groupby("_dup_key")[WMS_ORDER_COL].first()
    key_to_부주문 = wms_valid.groupby("_dup_key")[WMS_SUBORDER_COL].first()
    key_to_sku   = wms_valid.groupby("_dup_key")[WMS_SKU_COL].first()
    key_to_ch    = wms_valid.groupby("_dup_key")[WMS_CHANNEL_COL].first()
    key_to_type  = wms_valid.groupby("_dup_key")["_key_type"].first()

    dup_wms = wms_valid[wms_valid["_dup_key"].isin(dup_keys)].copy()

    agg = (
        dup_wms.groupby("_dup_key")
        .agg(WMS_출고횟수=("_주문_n", "count"), WMS_총수량=("_wms_qty", "sum"))
    )
    agg["주문번호"]   = agg.index.map(key_to_주문)
    agg["부주문코드"] = agg.index.map(key_to_부주문)
    agg["상품코드"]   = agg.index.map(key_to_sku)
    agg["채널"]       = agg.index.map(key_to_ch)
    agg["키유형"]     = agg.index.map(key_to_type)

    def get_cigro_qty(row):
        return cigro_qty_map.get((norm(row["주문번호"]), norm(row["상품코드"])), None)

    agg["Cigro_수량"] = agg.apply(get_cigro_qty, axis=1)
    agg["Cigro_존재"] = agg["Cigro_수량"].notna()
    agg["수량차이_WMS합산빼기Cigro"] = agg["WMS_총수량"] - agg["Cigro_수량"]

    mask_cigro = agg["Cigro_존재"]
    mask_over  = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] > 0)
    mask_under = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] < 0)
    mask_ok    = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] == 0)  # ← 정상 처리
    mask_nocig = ~agg["Cigro_존재"]

    n_over, n_under, n_ok, n_nocig = (
        mask_over.sum(), mask_under.sum(), mask_ok.sum(), mask_nocig.sum()
    )

    # 수량일치 키는 정상 → 문제 있는 중복에서 제외
    ok_dup_key_vals  = set(agg[mask_ok].index)          # agg index = _dup_key
    problem_dup_keys = dup_keys - ok_dup_key_vals

    agg = agg.reset_index(drop=True)
    agg["_dup_key"] = dup_wms.groupby("_dup_key")["_dup_key"].first().values

    # reset_index 후 mask 재계산 (index 정렬 필요)
    mask_cigro = agg["Cigro_존재"]
    mask_over  = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] > 0)
    mask_under = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] < 0)
    mask_ok    = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] == 0)
    mask_nocig = ~agg["Cigro_존재"]
    n_over, n_under, n_ok, n_nocig = mask_over.sum(), mask_under.sum(), mask_ok.sum(), mask_nocig.sum()

    # ------------------------------------------
    # 2. WMS 현황 출력
    # ------------------------------------------
    wms_unique_주문  = wms_valid["_주문_n"].nunique()
    wms_unique_부주문 = wms_valid.loc[wms_valid["_부주문_n"] != "", "_부주문_n"].nunique()
    wms_has_부주문   = (wms_valid["_부주문_n"] != "").sum()
    wms_unique_dup   = wms_valid["_dup_key"].nunique()

    print(f"\n[ WMS 출고 데이터 (Sheet1) ]")
    print(f"  원본 전체 행수                      : {len(wms):,}행")
    print(f"  수량 0 제외                         : {n_zero_qty:,}행")
    print(f"  유효 행수 (주문번호 있음+수량>0)     : {len(wms_valid):,}행")
    print(f"  고유 주문번호                        : {wms_unique_주문:,}건")
    print(f"  부주문코드 있는 행                   : {wms_has_부주문:,}행 / 고유 {wms_unique_부주문:,}건")
    print(f"  중복탐지 고유 키 (부주문/주문+SKU)   : {wms_unique_dup:,}건")
    print(f"  ── 정상 (1회 출고)                   : {(dup_count_map == 1).sum():,}건")
    print(f"  ── 정상 (분할출고 · Cigro수량일치)   : {n_ok:,}건")
    print(f"  ── 중복 의심 (확인 필요)             : {len(problem_dup_keys):,}건  ← 확인 필요")

    if problem_dup_keys:
        prob_rows   = wms_valid[wms_valid["_dup_key"].isin(problem_dup_keys)]
        prob_by_type = prob_rows.groupby("_key_type").size()
        prob_by_ch   = prob_rows.groupby(WMS_CHANNEL_COL).size().sort_values(ascending=False)
        dc = dup_count_map[dup_count_map >= 2]
        print(f"     중복 해당 WMS 행수              : {len(prob_rows):,}행")
        print(f"     평균 출고 횟수                  : {dc.mean():.1f}회 / 최대 {dc.max()}회")
        print(f"     ── 부주문코드 기준 중복          : {prob_by_type.get('부주문코드', 0):,}건")
        print(f"     ── 주문번호 기준 중복 (부주문없음): {prob_by_type.get('주문번호', 0):,}건")
        print(f"     채널별 분포: {dict(prob_by_ch)}")

    # ------------------------------------------
    # 3. WMS 채널별 현황
    # ------------------------------------------
    print(f"\n[ WMS 채널별 현황 ]")
    ch_stats = []
    for ch, sub in wms_valid.groupby(WMS_CHANNEL_COL):
        # 문제 있는 중복만 ★ 표시
        sub_prob_keys = problem_dup_keys & set(sub["_dup_key"])
        sub_prob_rows = sub[sub["_dup_key"].isin(sub_prob_keys)]
        has_부주문 = (sub["_부주문_n"] != "").sum()
        ch_stats.append({
            "채널": ch,
            "총행수": len(sub),
            "고유주문": sub["_주문_n"].nunique(),
            "부주문코드有": has_부주문,
            "중복키수": len(sub_prob_keys),
            "중복행수": len(sub_prob_rows),
        })
    for s in sorted(ch_stats, key=lambda x: -x["중복행수"]):
        flag = "  ★" if s["중복행수"] > 0 else ""
        print(f"  [{s['채널']:8}] 총{s['총행수']:4}행 | 고유주문:{s['고유주문']:4} | "
              f"부주문有:{s['부주문코드有']:4} | 중복:{s['중복키수']}건/{s['중복행수']}행{flag}")

    # ------------------------------------------
    # 4. Cigro 현황
    # ------------------------------------------
    print(f"\n[ Cigro 주문 데이터 (Sheet2) ]")
    print(f"  원본 전체 행수                      : {len(cigro):,}행")
    print(f"  유효 행수                           : {len(cigro_valid):,}행")
    print(f"  고유 주문번호                        : {cigro_valid['_order_n'].nunique():,}건")

    for ch_name, cnt in cigro_valid[CIGRO_CHANNEL_COL].value_counts().items():
        print(f"    [{ch_name}]: {cnt:,}행")

    # ------------------------------------------
    # 5. Cigro ↔ WMS 매칭 현황
    # ------------------------------------------
    cigro_주문_set = set(cigro_valid["_order_n"])
    wms_부주문_set = set(wms_valid.loc[wms_valid["_부주문_n"] != "", "_부주문_n"])

    match_via_주문   = len(cigro_주문_set & set(wms_valid["_주문_n"]))
    match_via_부주문 = len(cigro_주문_set & wms_부주문_set)
    unmatchable      = cigro_valid[
        cigro_valid[CIGRO_CHANNEL_COL].isin(ID_UNMATCHABLE_CIGRO_CHANNELS)
    ]["_order_n"].nunique()
    truly_unmatched  = len(cigro_주문_set) - match_via_주문 - match_via_부주문 - unmatchable

    sep("-")
    print("[ Cigro ↔ WMS 매칭 현황 ]")
    print(f"  Cigro order_id → WMS 주문번호 매칭   : {match_via_주문:,}건")
    print(f"  Cigro order_id → WMS 부주문코드 매칭 : {match_via_부주문:,}건")
    print(f"  ID 매칭 불가 (다른 ID체계 사용)      : {unmatchable:,}건  ← SMART_STORE: WMS=발주번호, Cigro=주문번호")
    print(f"  실제 미매칭 (확인 필요)               : {truly_unmatched:,}건")

    # SMART_STORE 전체 건수 비교 (날짜 기준 아님 — 주문일과 출고일이 다를 수 있음)
    on008    = wms_valid[wms_valid[WMS_CHANNEL_COL].isin(SMART_STORE_WMS_CHANNELS)]
    ss_cigro = cigro_valid[cigro_valid[CIGRO_CHANNEL_COL] == "SMART_STORE"]
    if not on008.empty and not ss_cigro.empty:
        w_total = on008["_주문_n"].nunique()
        c_total = ss_cigro["_order_n"].nunique()
        diff    = w_total - c_total
        bar     = "✓" if abs(diff) <= 20 else "△"
        print(f"\n  [SMART_STORE ↔ ON008] 전체 건수 비교 (ID 매칭 불가 → 건수로 검증, 날짜 무관)")
        print(f"    WMS ON008 고유주문: {w_total:,}건  |  Cigro SMART_STORE: {c_total:,}건  |  차이: {diff:+d}  {bar}")

    print(f"\n  채널별 매칭 상세:")
    for ch_name, sub_c in cigro_valid.groupby(CIGRO_CHANNEL_COL):
        ids    = set(sub_c["_order_n"])
        m주문  = len(ids & set(wms_valid["_주문_n"]))
        m부주문 = len(ids & wms_부주문_set)
        total  = sub_c["_order_n"].nunique()
        note   = "← ID체계 달라 매칭불가 (실제누락아님)" if ch_name in ID_UNMATCHABLE_CIGRO_CHANNELS else ""
        print(f"    [{ch_name:15}] 고유주문:{total:5} | "
              f"주문번호매칭:{m주문:4} | 부주문코드매칭:{m부주문:4} | "
              f"미매칭:{total-m주문-m부주문:4}  {note}")

    # ------------------------------------------
    # 6. 중복 분석 결과 출력
    # ------------------------------------------
    sep()
    print("[ 중복 출고 분석 결과 ]")
    sep()

    n_problem = len(problem_dup_keys)
    print(f"  WMS 중복 키 총계                    : {len(dup_keys):,}건")
    print(f"  ├ 정상 (분할출고 · Cigro수량일치)   : {n_ok:,}건  ← 문제 없음")
    print(f"  └ 확인 필요                         : {n_problem:,}건")

    if n_problem == 0 and n_ok == 0 and not dup_keys:
        print("  중복 출고된 키가 없습니다. 정상입니다.")

    if n_problem > 0 or n_ok > 0:
        print(f"     ├ WMS합산 > Cigro (과잉 출고)   : {n_over:,}건  ★ 주의!")
        print(f"     ├ WMS합산 < Cigro (부족 출고)   : {n_under:,}건")
        print(f"     └ Cigro에 없음 (WMS 단독)        : {n_nocig:,}건")

    # ------------------------------------------
    # 7. 상세 WMS 행 준비
    # ------------------------------------------
    key_cols   = [WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL, WMS_QTY_COL]
    extra_cols = ["상품명", "고객명", "출고일자", "출고상태", "운송장번호"]
    front = key_cols + [c for c in extra_cols if c in wms.columns]
    rest  = [c for c in wms.columns if c not in front and not c.startswith("_")]

    def get_detail(mask_summary):
        keys = set(agg.loc[mask_summary, "_dup_key"].dropna())
        return dup_wms[dup_wms["_dup_key"].isin(keys)][front + rest].sort_values(
            [WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL]
        )

    summary_cols = ["채널", "키유형", "주문번호", "부주문코드", "상품코드",
                    "WMS_출고횟수", "WMS_총수량", "Cigro_존재", "Cigro_수량",
                    "수량차이_WMS합산빼기Cigro"]

    summary_all   = agg[summary_cols].sort_values(["채널", "주문번호"])
    summary_over  = agg.loc[mask_over,  summary_cols]
    summary_under = agg.loc[mask_under, summary_cols]
    summary_ok    = agg.loc[mask_ok,    summary_cols]
    summary_nocig = agg.loc[mask_nocig, summary_cols]

    detail_over  = get_detail(mask_over)
    detail_under = get_detail(mask_under)
    detail_ok    = get_detail(mask_ok)
    detail_nocig = get_detail(mask_nocig)

    # ------------------------------------------
    # 8. 저장
    # ------------------------------------------
    output_path = os.path.join(base_dir, "중복출고_분석결과.xlsx")

    sheets = [
        ("전체_요약",              summary_all,   f"중복 {len(agg)}건 전체"),
        ("★과잉출고_요약",         summary_over,  f"WMS합산 > Cigro ({n_over}건)"),
        ("★과잉출고_상세행",       detail_over,   f"과잉출고 WMS 행 ({len(detail_over)}행)"),
        ("Cigro없음_요약",         summary_nocig, f"Cigro 미존재 ({n_nocig}건)"),
        ("Cigro없음_상세행",       detail_nocig,  f"Cigro없음 WMS 행 ({len(detail_nocig)}행)"),
        ("정상_분할출고_요약",     summary_ok,    f"분할출고·수량일치 (정상 {n_ok}건)"),
        ("정상_분할출고_상세행",   detail_ok,     f"분할출고 WMS 행 ({len(detail_ok)}행)"),
    ]
    if n_under > 0:
        sheets.insert(4, ("부족출고_요약",   summary_under, f"WMS합산 < Cigro ({n_under}건)"))
        sheets.insert(5, ("부족출고_상세행", detail_under,  f"부족출고 WMS 행 ({len(detail_under)}행)"))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sname, df, _ in sheets:
            df.to_excel(writer, sheet_name=sname, index=False)

    sep()
    print(f"\n저장 완료 → 중복출고_분석결과.xlsx")
    print(f"  {'시트명':<25} 내용")
    print(f"  {'-'*55}")
    for sname, df, label in sheets:
        print(f"  {sname:<25} {label}")
    print()


if __name__ == "__main__":
    main()
