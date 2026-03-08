import pandas as pd
import os
import re

# ==========================================
# 설정 영역 - 파일 및 컬럼명 설정
# ==========================================

# 비교할 엑셀 파일 경로 (파일명만 입력하면 스크립트 폴더 기준으로 찾음)
INPUT_FILE = "data.xlsx"

# Sheet1: WMS 출고 데이터 컬럼명
WMS_ORDER_COL = "주문번호"   # 주문번호
WMS_SKU_COL   = "상품코드"   # SKU 코드
WMS_QTY_COL   = "출고수량"   # 출고 수량

# Sheet2: Cigro 주문 데이터 컬럼명
CIGRO_ORDER_COL = "주문번호"  # 주문번호
CIGRO_SKU_COL   = "상품코드"  # SKU 코드
CIGRO_QTY_COL   = "주문수량"  # 주문 수량


# ==========================================
# 주문번호 정규화 (채널마다 형식이 달라도 매칭되도록)
# ==========================================
def normalize_order_no(val):
    """
    주문번호에서 공백, 하이픈, 언더스코어 등을 제거하고 소문자로 통일.
    ex) 'ORD-2025-001' → 'ord2025001'
    """
    if pd.isna(val):
        return ""
    return re.sub(r"[\s\-_]", "", str(val)).lower().strip()


# ==========================================
# 메인 실행
# ==========================================
def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(base_dir, INPUT_FILE)

    if not os.path.exists(input_path):
        print(f"[오류] 파일을 찾을 수 없습니다: {input_path}")
        print("INPUT_FILE 변수에 올바른 파일명을 설정하세요.")
        return

    # ------------------------------------------
    # 1. 데이터 로드
    # ------------------------------------------
    print(f"파일 로드 중: {INPUT_FILE}\n")
    wms   = pd.read_excel(input_path, sheet_name=0, dtype=str)  # Sheet1: WMS
    cigro = pd.read_excel(input_path, sheet_name=1, dtype=str)  # Sheet2: Cigro

    print("=== Sheet1 (WMS 출고) 컬럼 목록 ===")
    print(list(wms.columns))
    print(f"총 {len(wms)}행\n")

    print("=== Sheet2 (Cigro 주문) 컬럼 목록 ===")
    print(list(cigro.columns))
    print(f"총 {len(cigro)}행\n")

    # ------------------------------------------
    # 2. 컬럼 존재 여부 검증
    # ------------------------------------------
    missing = []
    for col in [WMS_ORDER_COL, WMS_SKU_COL, WMS_QTY_COL]:
        if col not in wms.columns:
            missing.append(f"Sheet1 '{col}'")
    for col in [CIGRO_ORDER_COL, CIGRO_SKU_COL, CIGRO_QTY_COL]:
        if col not in cigro.columns:
            missing.append(f"Sheet2 '{col}'")
    if missing:
        print("[오류] 아래 컬럼을 찾지 못했습니다. 스크립트 상단 설정을 확인하세요:")
        for m in missing:
            print(f"  - {m}")
        return

    # ------------------------------------------
    # 3. 주문번호 정규화 키 추가
    # ------------------------------------------
    wms["_order_key"]   = wms[WMS_ORDER_COL].apply(normalize_order_no)
    cigro["_order_key"] = cigro[CIGRO_ORDER_COL].apply(normalize_order_no)
    wms["_sku_key"]     = wms[WMS_SKU_COL].str.strip().str.lower()
    cigro["_sku_key"]   = cigro[CIGRO_SKU_COL].str.strip().str.lower()

    # ------------------------------------------
    # 4. 주문번호 매칭률 먼저 확인
    # ------------------------------------------
    wms_orders   = set(wms["_order_key"]) - {""}
    cigro_orders = set(cigro["_order_key"]) - {""}
    matched_orders = wms_orders & cigro_orders
    only_wms    = wms_orders - cigro_orders
    only_cigro  = cigro_orders - wms_orders

    print("=" * 50)
    print("[ 주문번호 매칭 현황 ]")
    print(f"  WMS   고유 주문번호: {len(wms_orders):,}건")
    print(f"  Cigro 고유 주문번호: {len(cigro_orders):,}건")
    print(f"  양쪽 모두 존재    : {len(matched_orders):,}건")
    print(f"  WMS에만 존재      : {len(only_wms):,}건")
    print(f"  Cigro에만 존재    : {len(only_cigro):,}건")
    if len(wms_orders) > 0:
        rate = len(matched_orders) / len(wms_orders | cigro_orders) * 100
        print(f"  전체 매칭률       : {rate:.1f}%")
    print("=" * 50 + "\n")

    # ------------------------------------------
    # 5. WMS / Cigro 컬럼명을 구분 접두사로 통일
    # ------------------------------------------
    wms_renamed = wms.rename(columns={
        WMS_ORDER_COL: "주문번호",
        WMS_SKU_COL:   "SKU코드",
        WMS_QTY_COL:   "WMS_수량",
    })
    cigro_renamed = cigro.rename(columns={
        CIGRO_ORDER_COL: "주문번호",
        CIGRO_SKU_COL:   "SKU코드",
        CIGRO_QTY_COL:   "Cigro_수량",
    })

    # ------------------------------------------
    # 6. 주문번호 + SKU 기준 outer merge
    # ------------------------------------------
    merged = pd.merge(
        wms_renamed,
        cigro_renamed,
        left_on=["_order_key", "_sku_key"],
        right_on=["_order_key", "_sku_key"],
        how="outer",
        suffixes=("_wms", "_cigro"),
        indicator=True,
    )

    # ------------------------------------------
    # 7. 3개 그룹으로 분류
    # ------------------------------------------
    df_matched     = merged[merged["_merge"] == "both"].copy()
    df_wms_only    = merged[merged["_merge"] == "left_only"].copy()
    df_cigro_only  = merged[merged["_merge"] == "right_only"].copy()

    # 수량 비교 컬럼 추가 (매칭 데이터)
    if not df_matched.empty and "WMS_수량" in df_matched.columns and "Cigro_수량" in df_matched.columns:
        df_matched["WMS_수량_n"]   = pd.to_numeric(df_matched["WMS_수량"],   errors="coerce")
        df_matched["Cigro_수량_n"] = pd.to_numeric(df_matched["Cigro_수량"], errors="coerce")
        df_matched["수량_일치"] = df_matched["WMS_수량_n"] == df_matched["Cigro_수량_n"]
        df_matched["수량_차이"] = df_matched["WMS_수량_n"] - df_matched["Cigro_수량_n"]

    # 내부 키 컬럼 제거
    drop_cols = ["_order_key", "_sku_key", "_merge",
                 "WMS_수량_n", "Cigro_수량_n"]
    for df in [df_matched, df_wms_only, df_cigro_only]:
        df.drop(columns=[c for c in drop_cols if c in df.columns], inplace=True)

    # ------------------------------------------
    # 8. 결과 저장
    # ------------------------------------------
    output_filename = f"비교결과_WMS_vs_Cigro.xlsx"
    output_path = os.path.join(base_dir, output_filename)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_matched.to_excel(writer,    sheet_name="매칭 데이터",  index=False)
        df_wms_only.to_excel(writer,   sheet_name="WMS만 있음",   index=False)
        df_cigro_only.to_excel(writer, sheet_name="Cigro만 있음", index=False)

    print("[ SKU 단위 분류 결과 ]")
    print(f"  매칭 데이터  : {len(df_matched):,}행")
    print(f"  WMS만 있음   : {len(df_wms_only):,}행")
    print(f"  Cigro만 있음 : {len(df_cigro_only):,}행")
    print(f"\n저장 완료 → {output_filename}")


if __name__ == "__main__":
    main()
