import pandas as pd
import os
import re
import sys
import glob

# ==========================================
# 설정 영역
# ==========================================

# Sheet1: WMS 출고 데이터 컬럼명
WMS_ORDER_COL = "주문번호"   # 주문번호
WMS_SKU_COL   = "상품코드"   # SKU 코드
WMS_QTY_COL   = "수량"       # 출고 수량

# Sheet2: Cigro 주문 데이터 컬럼명
CIGRO_ORDER_COL = "order_id"          # 주문번호
CIGRO_SKU_COL   = "match_sku"         # SKU 코드
CIGRO_QTY_COL   = "sku_적용_후_수량"  # 수량 (SKU 환산 후)


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


# ==========================================
# 주문번호 정규화
# - 공백/하이픈/언더스코어 제거 후 비교
# - ex) '20260226-0000133' → '202602260000133'
# ==========================================
def normalize_order_no(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[\s\-_]", "", str(val)).strip()


def sep(char="=", n=60):
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
    # 3. WMS 현황 출력
    # ------------------------------------------
    wms_valid = wms[wms["_order_key"] != ""]
    wms_unique_orders = wms_valid["_order_key"].nunique()
    wms_unique_combos = wms_valid["_comp_key"].nunique()
    wms_total_rows    = len(wms_valid)

    # 중복 그룹: (주문번호+SKU)가 2번 이상 나온 것
    wms_combo_counts = wms_valid.groupby("_comp_key").size().rename("wms_출고횟수")
    dup_combos = wms_combo_counts[wms_combo_counts >= 2]
    dup_combo_keys = set(dup_combos.index)

    dup_rows = wms_valid[wms_valid["_comp_key"].isin(dup_combo_keys)]

    print(f"\n[ WMS 출고 데이터 (Sheet1) ]")
    print(f"  원본 전체 행수                    : {len(wms):,}행")
    print(f"  주문번호 있는 유효 행수           : {wms_total_rows:,}행")
    print(f"  고유 주문번호                     : {wms_unique_orders:,}건")
    print(f"  고유 (주문번호+SKU) 조합          : {wms_unique_combos:,}건")
    print(f"  ── 1회 출고 (정상)                : {(wms_combo_counts == 1).sum():,}건")
    print(f"  ── 2회 이상 출고 (중복 의심)      : {len(dup_combos):,}건  ← 확인 필요")
    if len(dup_combos) > 0:
        avg_dup = dup_combos.mean()
        max_dup = dup_combos.max()
        print(f"     중복 해당 WMS 행수            : {len(dup_rows):,}행")
        print(f"     평균 출고 횟수                : {avg_dup:.1f}회")
        print(f"     최대 출고 횟수                : {max_dup}회")

    # ------------------------------------------
    # 4. Cigro 현황 출력
    # ------------------------------------------
    cigro_valid = cigro[cigro["_order_key"] != ""]
    cigro_unique_orders = cigro_valid["_order_key"].nunique()
    cigro_unique_combos = cigro_valid["_comp_key"].nunique()

    print(f"\n[ Cigro 주문 데이터 (Sheet2) ]")
    print(f"  원본 전체 행수                    : {len(cigro):,}행")
    print(f"  주문번호 있는 유효 행수           : {len(cigro_valid):,}행")
    print(f"  고유 주문번호                     : {cigro_unique_orders:,}건")
    print(f"  고유 (주문번호+SKU) 조합          : {cigro_unique_combos:,}건")

    # ------------------------------------------
    # 5. 중복 출고 분석
    # ------------------------------------------
    sep()
    print(f"[ 중복 출고 분석 결과 ]")
    sep()

    if len(dup_combos) == 0:
        print("  중복 출고된 (주문번호+SKU) 조합이 없습니다. 정상입니다.")
        return

    # Cigro에서 (주문번호+SKU)별 수량 집계
    cigro_qty_map = (
        cigro_valid
        .assign(_cigro_qty=pd.to_numeric(cigro_valid[CIGRO_QTY_COL], errors="coerce"))
        .groupby("_comp_key")["_cigro_qty"]
        .sum()
    )

    # 중복 요약 테이블 구성
    wms_dup_agg = (
        wms_valid[wms_valid["_comp_key"].isin(dup_combo_keys)]
        .assign(_wms_qty=pd.to_numeric(wms_valid.loc[wms_valid["_comp_key"].isin(dup_combo_keys), WMS_QTY_COL], errors="coerce"))
        .groupby("_comp_key")
        .agg(
            WMS_출고횟수=("_order_key", "count"),
            WMS_총수량=("_wms_qty", "sum"),
        )
        .join(dup_combos)
        .join(cigro_qty_map.rename("Cigro_수량"))
    )

    # 주문번호 / SKU 복원
    wms_dup_agg[["주문번호", "상품코드"]] = (
        wms_dup_agg.index.str.split("||", expand=True).to_frame(index=False)[[0, 1]].values
    )
    wms_dup_agg["Cigro_존재"] = wms_dup_agg["Cigro_수량"].notna()
    wms_dup_agg["수량_차이(WMS합산-Cigro)"] = wms_dup_agg["WMS_총수량"] - wms_dup_agg["Cigro_수량"]

    # 컬럼 순서 정리
    summary = wms_dup_agg[["주문번호", "상품코드",
                            "WMS_출고횟수", "WMS_총수량",
                            "Cigro_존재", "Cigro_수량",
                            "수량_차이(WMS합산-Cigro)"]].copy()
    summary = summary.sort_values(["Cigro_존재", "수량_차이(WMS합산-Cigro)"],
                                  ascending=[False, False])

    # 그룹별 통계 출력
    in_cigro     = summary["Cigro_존재"].sum()
    not_in_cigro = (~summary["Cigro_존재"]).sum()

    print(f"  WMS 중복 출고 (주문번호+SKU)      : {len(summary):,}건")
    print(f"  ├ Cigro에도 존재 (비교 가능)      : {in_cigro:,}건")
    print(f"  └ Cigro에 없음  (WMS 단독)        : {not_in_cigro:,}건")

    if in_cigro > 0:
        cigro_subset = summary[summary["Cigro_존재"]]
        qty_ok  = (cigro_subset["수량_차이(WMS합산-Cigro)"] == 0).sum()
        qty_over = (cigro_subset["수량_차이(WMS합산-Cigro)"] > 0).sum()
        qty_under = (cigro_subset["수량_차이(WMS합산-Cigro)"] < 0).sum()
        print(f"\n  Cigro 존재 {in_cigro}건 중 수량 비교:")
        print(f"  ├ WMS합산 == Cigro (수량 일치) : {qty_ok:,}건")
        print(f"  ├ WMS합산  > Cigro (과잉 출고) : {qty_over:,}건  ← 주의!")
        print(f"  └ WMS합산  < Cigro (부족 출고) : {qty_under:,}건  ← 주의!")

    # ------------------------------------------
    # 6. 중복 상세 행 (WMS 원본 행 그대로)
    # ------------------------------------------
    # 앞쪽에 주요 컬럼 배치
    key_cols = [WMS_ORDER_COL, WMS_SKU_COL, WMS_QTY_COL]
    extra_cols = ["상품명", "고객명", "출고일자", "출고상태", "운송장번호", "매출처"]
    front = key_cols + [c for c in extra_cols if c in wms.columns]
    rest  = [c for c in wms.columns if c not in front and not c.startswith("_")]
    detail = (
        dup_rows[front + rest]
        .sort_values([WMS_ORDER_COL, WMS_SKU_COL])
        .copy()
    )

    # ------------------------------------------
    # 7. 결과 저장
    # ------------------------------------------
    output_path = os.path.join(base_dir, "중복출고_분석결과.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="중복출고_요약",   index=False)
        detail.to_excel(writer,  sheet_name="중복출고_상세행", index=False)

    sep()
    print(f"\n저장 완료 → 중복출고_분석결과.xlsx")
    print(f"  Sheet1 [중복출고_요약]   : {len(summary):,}건  (주문번호+SKU 조합 단위)")
    print(f"  Sheet2 [중복출고_상세행] : {len(detail):,}행  (WMS 원본 행 전체)\n")


if __name__ == "__main__":
    main()
