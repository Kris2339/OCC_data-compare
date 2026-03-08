import pandas as pd
import os
import re

# ==========================================
# 설정 영역
# ==========================================

# 비교할 엑셀 파일명 (스크립트와 같은 폴더에 있어야 함)
INPUT_FILE = "출고진행_합산_20260226_20260308.xlsx"

# Sheet1: WMS 출고 데이터 컬럼명
WMS_ORDER_COL = "주문번호"   # 주문번호
WMS_SKU_COL   = "상품코드"   # SKU 코드
WMS_QTY_COL   = "수량"       # 출고 수량

# Sheet2: Cigro 주문 데이터 컬럼명
CIGRO_ORDER_COL = "order_id"          # 주문번호
CIGRO_SKU_COL   = "match_sku"         # SKU 코드
CIGRO_QTY_COL   = "sku_적용_후_수량"  # 수량 (SKU 환산 후)


# ==========================================
# 주문번호 정규화
# - 공백/하이픈/언더스코어 제거 후 비교
# - ex) '20260226-0000133' → '202602260000133'
# ==========================================
def normalize_order_no(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[\s\-_]", "", str(val)).strip()


# ==========================================
# 메인 실행
# ==========================================
def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(base_dir, INPUT_FILE)

    if not os.path.exists(input_path):
        print(f"[오류] 파일을 찾을 수 없습니다: {input_path}")
        return

    # ------------------------------------------
    # 1. 데이터 로드
    # ------------------------------------------
    print(f"파일 로드 중: {INPUT_FILE}")
    wms   = pd.read_excel(input_path, sheet_name=0, dtype=str)  # Sheet1: WMS
    cigro = pd.read_excel(input_path, sheet_name=1, dtype=str)  # Sheet2: Cigro
    print(f"  WMS   {len(wms):,}행 / Cigro {len(cigro):,}행 로드 완료\n")

    # ------------------------------------------
    # 2. 정규화 키 생성
    # ------------------------------------------
    wms["_order_key"]   = wms[WMS_ORDER_COL].apply(normalize_order_no)
    cigro["_order_key"] = cigro[CIGRO_ORDER_COL].apply(normalize_order_no)
    wms["_sku_key"]     = wms[WMS_SKU_COL].str.strip().str.upper()
    cigro["_sku_key"]   = cigro[CIGRO_SKU_COL].str.strip().str.upper()

    # ------------------------------------------
    # 3. 주문번호 매칭률 현황 출력
    # ------------------------------------------
    wms_orders     = set(wms["_order_key"]) - {""}
    cigro_orders   = set(cigro["_order_key"]) - {""}
    matched_orders = wms_orders & cigro_orders
    only_wms       = wms_orders - cigro_orders
    only_cigro     = cigro_orders - wms_orders
    total          = len(wms_orders | cigro_orders)

    print("=" * 55)
    print("[ 주문번호 매칭 현황 ]")
    print(f"  WMS   고유 주문번호 : {len(wms_orders):,}건")
    print(f"  Cigro 고유 주문번호 : {len(cigro_orders):,}건")
    print(f"  양쪽 모두 존재      : {len(matched_orders):,}건")
    print(f"  WMS에만 존재        : {len(only_wms):,}건")
    print(f"  Cigro에만 존재      : {len(only_cigro):,}건")
    if total > 0:
        print(f"  전체 매칭률         : {len(matched_orders)/total*100:.1f}%")
    print("=" * 55 + "\n")

    # ------------------------------------------
    # 4. WMS 컬럼 정리 (비교에 필요한 것만 앞에 배치)
    # ------------------------------------------
    wms_cols_front = ["주문번호", "상품코드", "수량", "상품명", "고객명",
                      "출고일자", "출고상태", "운송장번호", "매출처", "UserId", "Name"]
    wms_front = [c for c in wms_cols_front if c in wms.columns]
    wms_rest  = [c for c in wms.columns if c not in wms_front and not c.startswith("_")]
    wms_ordered = wms[wms_front + wms_rest + ["_order_key", "_sku_key"]].copy()

    cigro_cols_front = ["order_id", "match_sku", "sku_적용_후_수량", "quantity",
                        "product_name", "channel_name", "status", "payment_date"]
    cigro_front = [c for c in cigro_cols_front if c in cigro.columns]
    cigro_rest  = [c for c in cigro.columns if c not in cigro_front and not c.startswith("_")]
    cigro_ordered = cigro[cigro_front + cigro_rest + ["_order_key", "_sku_key"]].copy()

    # 컬럼 접두사 추가 (중복 방지)
    cigro_rename = {c: f"cigro_{c}" for c in cigro_ordered.columns
                    if c not in ["_order_key", "_sku_key"]}
    cigro_ordered = cigro_ordered.rename(columns=cigro_rename)

    # ------------------------------------------
    # 5. 주문번호 + SKU 기준 outer merge
    # ------------------------------------------
    merged = pd.merge(
        wms_ordered,
        cigro_ordered,
        on=["_order_key", "_sku_key"],
        how="outer",
        indicator=True,
    )

    # ------------------------------------------
    # 6. 3개 그룹 분류
    # ------------------------------------------
    df_matched    = merged[merged["_merge"] == "both"].copy()
    df_wms_only   = merged[merged["_merge"] == "left_only"].copy()
    df_cigro_only = merged[merged["_merge"] == "right_only"].copy()

    # 매칭 데이터: 수량 비교 컬럼 추가
    if not df_matched.empty:
        wms_qty   = pd.to_numeric(df_matched[WMS_QTY_COL],           errors="coerce")
        cigro_qty = pd.to_numeric(df_matched[f"cigro_{CIGRO_QTY_COL}"], errors="coerce")
        df_matched.insert(2, "수량_WMS",   wms_qty)
        df_matched.insert(3, "수량_Cigro", cigro_qty)
        df_matched.insert(4, "수량_일치",  wms_qty == cigro_qty)
        df_matched.insert(5, "수량_차이",  wms_qty - cigro_qty)

    # 내부 키 컬럼 / indicator 제거
    drop_cols = ["_order_key", "_sku_key", "_merge"]
    for df in [df_matched, df_wms_only, df_cigro_only]:
        df.drop(columns=[c for c in drop_cols if c in df.columns], inplace=True)

    # ------------------------------------------
    # 7. 결과 저장
    # ------------------------------------------
    output_path = os.path.join(base_dir, "비교결과_WMS_vs_Cigro.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_matched.to_excel(writer,    sheet_name="매칭 데이터",  index=False)
        df_wms_only.to_excel(writer,   sheet_name="WMS만 있음",   index=False)
        df_cigro_only.to_excel(writer, sheet_name="Cigro만 있음", index=False)

    print("[ SKU 단위 분류 결과 ]")
    print(f"  매칭 데이터  : {len(df_matched):,}행")
    print(f"  WMS만 있음   : {len(df_wms_only):,}행")
    print(f"  Cigro만 있음 : {len(df_cigro_only):,}행")

    if not df_matched.empty and "수량_일치" in df_matched.columns:
        qty_ok  = df_matched["수량_일치"].sum()
        qty_all = df_matched["수량_일치"].count()
        print(f"\n  수량 일치    : {qty_ok:,}/{qty_all:,}건 ({qty_ok/qty_all*100:.1f}%)")
        print(f"  수량 불일치  : {qty_all-qty_ok:,}건")

    print(f"\n저장 완료 → 비교결과_WMS_vs_Cigro.xlsx")


if __name__ == "__main__":
    main()
