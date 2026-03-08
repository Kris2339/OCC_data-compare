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

# 네이버 스마트스토어: WMS=발주번호, Cigro=주문번호 → ID 체계가 달라 주문번호 매칭 불가
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
    if pd.isna(v):
        return ""
    return re.sub(r"[\s\-_]", "", str(v)).strip().upper()


def sep(char="=", n=60):
    print(char * n)


def make_dup_key(row):
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

    sep()
    print(f"  파일: {os.path.basename(input_path)}")
    sep()

    # ------------------------------------------
    # 1. 로드 및 정규화
    # ------------------------------------------
    wms   = pd.read_excel(input_path, sheet_name=0, dtype=str)
    cigro = pd.read_excel(input_path, sheet_name=1, dtype=str)

    # WMS 정규화
    wms_valid   = wms[wms[WMS_ORDER_COL].notna()].copy()
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

    # Cigro 정규화
    cigro["_order_n"]   = cigro[CIGRO_ORDER_COL].apply(norm)
    cigro["_sku"]       = cigro[CIGRO_SKU_COL].str.strip().str.upper()
    cigro_valid         = cigro[cigro["_order_n"] != ""].copy()
    cigro_valid["_qty"] = pd.to_numeric(cigro_valid[CIGRO_QTY_COL], errors="coerce")
    cigro_qty_map       = cigro_valid.groupby(["_order_n", "_sku"])["_qty"].sum()

    # ------------------------------------------
    # WMS 중복 탐지 + Cigro 비교 사전 분류
    #   수량일치 : WMS 총수량 == Cigro 수량 → 정상 (분할출고 등), 제외
    #   과잉출고 : WMS 총수량 >  Cigro 수량 → ★ 문제
    #   Cigro없음: Cigro에 해당 주문 없음   → 확인 필요
    # ------------------------------------------
    dup_count_map = wms_valid.groupby("_dup_key").size()
    dup_keys      = set(dup_count_map[dup_count_map >= 2].index)

    if dup_keys:
        key_to_주문  = wms_valid.groupby("_dup_key")[WMS_ORDER_COL].first()
        key_to_부주문 = wms_valid.groupby("_dup_key")[WMS_SUBORDER_COL].first()
        key_to_sku   = wms_valid.groupby("_dup_key")[WMS_SKU_COL].first()
        key_to_ch    = wms_valid.groupby("_dup_key")[WMS_CHANNEL_COL].first()
        key_to_type  = wms_valid.groupby("_dup_key")["_key_type"].first()

        dup_wms = wms_valid[wms_valid["_dup_key"].isin(dup_keys)].copy()
        agg = dup_wms.groupby("_dup_key").agg(
            WMS_출고횟수=("_주문_n", "count"), WMS_총수량=("_wms_qty", "sum")
        )
        agg["주문번호"]   = agg.index.map(key_to_주문)
        agg["부주문코드"] = agg.index.map(key_to_부주문)
        agg["상품코드"]   = agg.index.map(key_to_sku)
        agg["채널"]       = agg.index.map(key_to_ch)
        agg["키유형"]     = agg.index.map(key_to_type)

        def get_cigro_qty(row):
            return cigro_qty_map.get((norm(row["주문번호"]), norm(row["상품코드"])), None)

        agg["Cigro_수량"]              = agg.apply(get_cigro_qty, axis=1)
        agg["Cigro_존재"]              = agg["Cigro_수량"].notna()
        agg["수량차이(WMS-Cigro)"]     = agg["WMS_총수량"] - agg["Cigro_수량"]

        # 수량일치는 정상 → 문제에서 제외
        ok_keys = set(agg[agg["Cigro_존재"] & (agg["수량차이(WMS-Cigro)"] == 0)].index)
        problem_dup_keys = dup_keys - ok_keys

        agg = agg.reset_index(drop=True)
        agg["_dup_key"] = dup_wms.groupby("_dup_key")["_dup_key"].first().values

        mask_over  = agg["Cigro_존재"] & (agg["수량차이(WMS-Cigro)"] > 0)
        mask_nocig = ~agg["Cigro_존재"]
        n_ok       = len(ok_keys)
    else:
        dup_wms          = pd.DataFrame()
        agg              = pd.DataFrame()
        problem_dup_keys = set()
        mask_over        = pd.Series(dtype=bool)
        mask_nocig       = pd.Series(dtype=bool)
        n_ok             = 0

    # ------------------------------------------
    # 2. WMS 현황
    # ------------------------------------------
    n_single  = (dup_count_map == 1).sum()
    n_problem = len(problem_dup_keys)

    print(f"\n[ WMS 출고 현황 ]")
    print(f"  출고 행수   : {len(wms_valid):,}행  (수량 0 제외: {n_zero_qty}행)")
    print(f"  고유 주문   : {wms_valid['_주문_n'].nunique():,}건")
    if n_ok > 0:
        print(f"  정상 (분할출고·수량일치 포함) : {n_single + n_ok:,}건")
    else:
        print(f"  정상        : {n_single:,}건")

    if n_problem > 0:
        prob_rows  = wms_valid[wms_valid["_dup_key"].isin(problem_dup_keys)]
        prob_by_ch = prob_rows.groupby(WMS_CHANNEL_COL).size().sort_values(ascending=False)
        ch_str     = "  /  ".join(f"{ch}: {cnt}건" for ch, cnt in prob_by_ch.items())
        print(f"  중복 의심   : {n_problem:,}건  ← 확인 필요")
        print(f"    └ {ch_str}")
    else:
        print(f"  중복 의심   : 0건")

    # ------------------------------------------
    # 3. Cigro 현황
    # ------------------------------------------
    print(f"\n[ Cigro 주문 현황 ]")
    print(f"  총 주문 건수 : {cigro_valid['_order_n'].nunique():,}건")
    ch_counts = cigro_valid[CIGRO_CHANNEL_COL].value_counts()
    print(f"  채널별       : " + "  |  ".join(f"{ch} {cnt}건" for ch, cnt in ch_counts.items()))

    # ------------------------------------------
    # 4. Cigro → WMS 출고 대조
    # ------------------------------------------
    cigro_matchable = cigro_valid[~cigro_valid[CIGRO_CHANNEL_COL].isin(ID_UNMATCHABLE_CIGRO_CHANNELS)]
    matchable_set   = set(cigro_matchable["_order_n"])
    wms_부주문_set  = set(wms_valid.loc[wms_valid["_부주문_n"] != "", "_부주문_n"])

    matched        = len(matchable_set & (set(wms_valid["_주문_n"]) | wms_부주문_set))
    truly_unmatched = len(matchable_set) - matched
    unmatchable    = cigro_valid[cigro_valid[CIGRO_CHANNEL_COL].isin(ID_UNMATCHABLE_CIGRO_CHANNELS)]["_order_n"].nunique()

    sep("-")
    print(f"[ Cigro → WMS 출고 대조 ]")
    print(f"  매칭 가능 채널 : {len(matchable_set):,}건 중 {matched:,}건 매칭됨")
    if truly_unmatched > 0:
        # 채널별 미매칭
        lines = []
        for ch_name, sub_c in cigro_matchable.groupby(CIGRO_CHANNEL_COL):
            ids    = set(sub_c["_order_n"])
            m      = len(ids & (set(wms_valid["_주문_n"]) | wms_부주문_set))
            miss   = len(ids) - m
            if miss > 0:
                lines.append(f"{ch_name}: {miss}건")
        print(f"  미출고 의심    : {truly_unmatched:,}건  ← 날짜 범위 차이일 수 있음")
        if lines:
            print(f"    └ " + "  /  ".join(lines))

    # SMART_STORE 건수 비교
    on008    = wms_valid[wms_valid[WMS_CHANNEL_COL].isin(SMART_STORE_WMS_CHANNELS)]
    ss_cigro = cigro_valid[cigro_valid[CIGRO_CHANNEL_COL] == "SMART_STORE"]
    if not on008.empty and not ss_cigro.empty:
        w_total = on008["_주문_n"].nunique()
        c_total = ss_cigro["_order_n"].nunique()
        diff    = w_total - c_total
        bar     = "✓" if abs(diff) <= 20 else "△"
        print(f"  SMART_STORE    : WMS {w_total:,}건 / Cigro {c_total:,}건 (차이 {diff:+d} {bar})")
        print(f"    └ 주문번호 체계가 달라 건수로만 비교 (ID 매칭 불가)")

    # ------------------------------------------
    # 5. 최종 결과
    # ------------------------------------------
    sep()
    n_over  = int(mask_over.sum())  if len(mask_over)  > 0 else 0
    n_nocig = int(mask_nocig.sum()) if len(mask_nocig) > 0 else 0

    print(f"[ 최종 결과 ]")
    if n_problem == 0:
        print(f"  중복 출고 문제 없음")
    else:
        print(f"  과잉 출고 (WMS > Cigro)  : {n_over:,}건  {'★ 주의!' if n_over > 0 else ''}")
        print(f"  WMS 단독 (Cigro 없음)    : {n_nocig:,}건  {'← 확인 필요' if n_nocig > 0 else ''}")

    # ------------------------------------------
    # 6. 상세 행 준비
    # ------------------------------------------
    if agg.empty:
        sep()
        print()
        return

    key_cols   = [WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL, WMS_QTY_COL]
    extra_cols = ["상품명", "고객명", "출고일자", "출고상태", "운송장번호"]
    front = key_cols + [c for c in extra_cols if c in wms.columns]
    rest  = [c for c in wms.columns if c not in front and not c.startswith("_")]

    def get_detail(mask):
        keys = set(agg.loc[mask, "_dup_key"].dropna())
        return dup_wms[dup_wms["_dup_key"].isin(keys)][front + rest].sort_values(
            [WMS_CHANNEL_COL, WMS_ORDER_COL, WMS_SUBORDER_COL, WMS_SKU_COL]
        )

    summary_cols = ["채널", "키유형", "주문번호", "부주문코드", "상품코드",
                    "WMS_출고횟수", "WMS_총수량", "Cigro_수량", "수량차이(WMS-Cigro)"]

    summary_over  = agg.loc[mask_over,  summary_cols]
    summary_nocig = agg.loc[mask_nocig, summary_cols]
    detail_over   = get_detail(mask_over)
    detail_nocig  = get_detail(mask_nocig)

    # ------------------------------------------
    # 7. 저장
    # ------------------------------------------
    output_path = os.path.join(base_dir, "중복출고_분석결과.xlsx")
    sheets = []
    if n_over > 0:
        sheets += [
            ("과잉출고",    summary_over,  f"WMS > Cigro ({n_over}건)"),
            ("과잉출고_상세", detail_over, f"WMS 행 ({len(detail_over)}행)"),
        ]
    if n_nocig > 0:
        sheets += [
            ("미확인",      summary_nocig,  f"Cigro 없음 ({n_nocig}건)"),
            ("미확인_상세", detail_nocig,   f"WMS 행 ({len(detail_nocig)}행)"),
        ]

    if sheets:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sname, df, _ in sheets:
                df.to_excel(writer, sheet_name=sname, index=False)
        sep()
        print(f"\n저장 완료 → 중복출고_분석결과.xlsx")
        for sname, df, label in sheets:
            print(f"  {sname:<12} {label}")
    else:
        sep()
        print(f"\n저장할 문제 항목 없음 (중복출고_분석결과.xlsx 생성 안 함)")
    print()


if __name__ == "__main__":
    main()
