import pandas as pd
import os
import re
import sys
import glob

# ==========================================
# 설정 영역
# ==========================================

# Sheet1: WMS 출고 데이터 컬럼명
WMS_ORDER_COL = "주문번호"
WMS_SKU_COL   = "상품코드"
WMS_QTY_COL   = "수량"

# Sheet2: Cigro 주문 데이터 컬럼명
CIGRO_ORDER_COL = "order_id"
CIGRO_SKU_COL   = "match_sku"
CIGRO_QTY_COL   = "sku_적용_후_수량"


# ==========================================
# 입력 파일 자동 탐색
# 우선순위: 1) 커맨드라인 인자  2) 폴더 내 최신 .xlsx
# ==========================================
def resolve_input_file(base_dir):
    if len(sys.argv) > 1:
        path = sys.argv[1]
        if not os.path.isabs(path):
            path = os.path.join(base_dir, path)
        return path

    xlsx_files = glob.glob(os.path.join(base_dir, "*.xlsx"))
    xlsx_files = [f for f in xlsx_files if "중복출고" not in os.path.basename(f)]
    if not xlsx_files:
        return None
    return max(xlsx_files, key=os.path.getmtime)


def normalize_order_no(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[\s\-_]", "", str(val)).strip()


def sep(char="=", n=62):
    print(char * n)


# ==========================================
# 메인 실행
# ==========================================
def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = resolve_input_file(base_dir)

    if not input_path or not os.path.exists(input_path):
        print("[오류] 비교할 엑셀 파일을 찾을 수 없습니다.")
        print("사용법: python compare_wms_cigro.py [파일명.xlsx]")
        return

    # ------------------------------------------
    # 1. 데이터 로드
    # ------------------------------------------
    fname = os.path.basename(input_path)
    sep()
    print(f"[ 입력 파일 ]")
    print(f"  {fname}")
    sep()

    wms   = pd.read_excel(input_path, sheet_name=0, dtype=str)
    cigro = pd.read_excel(input_path, sheet_name=1, dtype=str)

    # ------------------------------------------
    # 2. 정규화 키 생성
    # ------------------------------------------
    wms["_order_key"] = wms[WMS_ORDER_COL].apply(normalize_order_no)
    wms["_sku_key"]   = wms[WMS_SKU_COL].str.strip().str.upper()
    wms["_comp_key"]  = wms["_order_key"] + "||" + wms["_sku_key"]

    cigro["_order_key"] = cigro[CIGRO_ORDER_COL].apply(normalize_order_no)
    cigro["_sku_key"]   = cigro[CIGRO_SKU_COL].str.strip().str.upper()
    cigro["_comp_key"]  = cigro["_order_key"] + "||" + cigro["_sku_key"]

    # ------------------------------------------
    # 3. WMS 유효 행 / 현황 출력
    # ------------------------------------------
    wms_valid = wms[wms["_order_key"] != ""].copy()

    wms_combo_counts  = wms_valid.groupby("_comp_key").size()
    dup_keys          = set(wms_combo_counts[wms_combo_counts >= 2].index)
    normal_keys       = set(wms_combo_counts[wms_combo_counts == 1].index)

    print(f"\n[ WMS 출고 데이터 (Sheet1) ]")
    print(f"  원본 전체 행수                    : {len(wms):,}행")
    print(f"  주문번호 있는 유효 행수           : {len(wms_valid):,}행")
    print(f"  고유 주문번호                     : {wms_valid['_order_key'].nunique():,}건")
    print(f"  고유 (주문번호+SKU) 조합          : {wms_valid['_comp_key'].nunique():,}건")
    print(f"  ── 정상 (1회 출고)                : {len(normal_keys):,}건")
    print(f"  ── 중복 (2회 이상 출고)           : {len(dup_keys):,}건  ← 확인 필요")
    if dup_keys:
        dup_counts = wms_combo_counts[wms_combo_counts >= 2]
        print(f"     중복 해당 WMS 행수            : {wms_valid['_comp_key'].isin(dup_keys).sum():,}행")
        print(f"     평균 출고 횟수                : {dup_counts.mean():.1f}회")
        print(f"     최대 출고 횟수                : {dup_counts.max()}회")

    # ------------------------------------------
    # 4. Cigro 현황 출력
    # ------------------------------------------
    cigro_valid = cigro[cigro["_order_key"] != ""].copy()

    print(f"\n[ Cigro 주문 데이터 (Sheet2) ]")
    print(f"  원본 전체 행수                    : {len(cigro):,}행")
    print(f"  주문번호 있는 유효 행수           : {len(cigro_valid):,}행")
    print(f"  고유 주문번호                     : {cigro_valid['_order_key'].nunique():,}건")
    print(f"  고유 (주문번호+SKU) 조합          : {cigro_valid['_comp_key'].nunique():,}건")

    # ------------------------------------------
    # 5. 중복 분석
    # ------------------------------------------
    sep()
    print(f"[ 중복 출고 분석 결과 ]")
    sep()

    if not dup_keys:
        print("  중복 출고된 (주문번호+SKU) 조합이 없습니다. 정상입니다.")
        return

    # _comp_key → 원본 주문번호/상품코드 매핑 (WMS 원본값 기준)
    key_to_order = wms_valid.groupby("_comp_key")[WMS_ORDER_COL].first()
    key_to_sku   = wms_valid.groupby("_comp_key")[WMS_SKU_COL].first()

    # Cigro: _comp_key별 수량 합산
    cigro_valid["_cigro_qty"] = pd.to_numeric(cigro_valid[CIGRO_QTY_COL], errors="coerce")
    cigro_qty_map = cigro_valid.groupby("_comp_key")["_cigro_qty"].sum()

    # WMS 중복 그룹 집계
    dup_wms = wms_valid[wms_valid["_comp_key"].isin(dup_keys)].copy()
    dup_wms["_wms_qty"] = pd.to_numeric(dup_wms[WMS_QTY_COL], errors="coerce")

    agg = dup_wms.groupby("_comp_key").agg(WMS_출고횟수=("_order_key", "count"),
                                            WMS_총수량=("_wms_qty", "sum"))
    agg["Cigro_수량"]  = agg.index.map(cigro_qty_map)
    agg["Cigro_존재"]  = agg["Cigro_수량"].notna()
    agg["수량차이_WMS합산빼기Cigro"] = agg["WMS_총수량"] - agg["Cigro_수량"]

    # 주문번호/상품코드 원본값 복원 (인덱스 조작 없이 map으로)
    agg.insert(0, "주문번호", agg.index.map(key_to_order))
    agg.insert(1, "상품코드", agg.index.map(key_to_sku))
    agg = agg.reset_index(drop=True)

    # 카테고리 분류
    mask_cigro   = agg["Cigro_존재"]
    mask_over    = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] > 0)
    mask_under   = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] < 0)
    mask_ok      = mask_cigro & (agg["수량차이_WMS합산빼기Cigro"] == 0)
    mask_no_cigro = ~agg["Cigro_존재"]

    summary_over     = agg[mask_over].copy()
    summary_under    = agg[mask_under].copy()
    summary_ok       = agg[mask_ok].copy()
    summary_no_cigro = agg[mask_no_cigro].copy()

    n_over     = len(summary_over)
    n_under    = len(summary_under)
    n_ok       = len(summary_ok)
    n_no_cigro = len(summary_no_cigro)

    print(f"  WMS 중복 출고 (주문번호+SKU)      : {len(agg):,}건")
    print(f"  ├ Cigro에도 존재 (비교 가능)      : {mask_cigro.sum():,}건")
    print(f"  │  ├ WMS합산 > Cigro (과잉 출고)  : {n_over:,}건  ★ 주의!")
    print(f"  │  ├ WMS합산 < Cigro (부족 출고)  : {n_under:,}건")
    print(f"  │  └ WMS합산 = Cigro (수량 일치)  : {n_ok:,}건")
    print(f"  └ Cigro에 없음  (WMS 단독)        : {n_no_cigro:,}건")

    # ------------------------------------------
    # 6. 상세 WMS 행 준비
    # ------------------------------------------
    key_cols  = [WMS_ORDER_COL, WMS_SKU_COL, WMS_QTY_COL]
    extra_cols = ["상품명", "고객명", "출고일자", "출고상태", "운송장번호", "매출처"]
    front = key_cols + [c for c in extra_cols if c in wms.columns]
    rest  = [c for c in wms.columns if c not in front and not c.startswith("_")]
    display_cols = front + rest

    def detail_rows(keys):
        """_comp_key 집합에 해당하는 WMS 상세 행 반환 (주요 컬럼 앞에 배치)"""
        rows = dup_wms[dup_wms["_comp_key"].isin(keys)].copy()
        return rows[display_cols].sort_values([WMS_ORDER_COL, WMS_SKU_COL])

    # _comp_key set 추출 (agg에서 원본 주문번호/상품코드로 매핑된 것과 대응)
    comp_over     = set(dup_wms.loc[dup_wms[WMS_ORDER_COL].isin(summary_over["주문번호"])    & dup_wms[WMS_SKU_COL].isin(summary_over["상품코드"]),    "_comp_key"])
    comp_under    = set(dup_wms.loc[dup_wms[WMS_ORDER_COL].isin(summary_under["주문번호"])   & dup_wms[WMS_SKU_COL].isin(summary_under["상품코드"]),   "_comp_key"])
    comp_ok       = set(dup_wms.loc[dup_wms[WMS_ORDER_COL].isin(summary_ok["주문번호"])      & dup_wms[WMS_SKU_COL].isin(summary_ok["상품코드"]),      "_comp_key"])
    comp_no_cigro = set(dup_wms.loc[dup_wms[WMS_ORDER_COL].isin(summary_no_cigro["주문번호"]) & dup_wms[WMS_SKU_COL].isin(summary_no_cigro["상품코드"]), "_comp_key"])

    detail_over     = detail_rows(comp_over)
    detail_under    = detail_rows(comp_under)
    detail_ok       = detail_rows(comp_ok)
    detail_no_cigro = detail_rows(comp_no_cigro)

    # ------------------------------------------
    # 7. 결과 저장 (카테고리별 시트)
    # ------------------------------------------
    output_path = os.path.join(base_dir, "중복출고_분석결과.xlsx")

    sheets = [
        ("전체_요약",            agg,             "전체 중복 87건 요약"),
        ("★과잉출고_요약",       summary_over,    f"WMS합산 > Cigro 수량 ({n_over}건)"),
        ("★과잉출고_상세행",     detail_over,     f"과잉출고 WMS 원본 행 ({len(detail_over)}행)"),
        ("Cigro없음_요약",       summary_no_cigro,f"Cigro 미존재 WMS 중복 ({n_no_cigro}건)"),
        ("Cigro없음_상세행",     detail_no_cigro, f"Cigro 없음 WMS 원본 행 ({len(detail_no_cigro)}행)"),
        ("수량일치_중복_요약",   summary_ok,      f"수량 일치지만 중복 출고 ({n_ok}건)"),
        ("수량일치_중복_상세행", detail_ok,       f"수량일치 중복 WMS 원본 행 ({len(detail_ok)}행)"),
    ]
    if n_under > 0:
        sheets.insert(4, ("부족출고_요약",   summary_under, f"WMS합산 < Cigro 수량 ({n_under}건)"))
        sheets.insert(5, ("부족출고_상세행", detail_under,  f"부족출고 WMS 원본 행 ({len(detail_under)}행)"))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df, label in sheets:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    sep()
    print(f"\n저장 완료 → 중복출고_분석결과.xlsx")
    print(f"  {'시트명':<25}  {'내용'}")
    print(f"  {'-'*55}")
    for sheet_name, df, label in sheets:
        print(f"  {sheet_name:<25}  {label}")
    print()


if __name__ == "__main__":
    main()
