import requests
import pandas as pd
import os
import datetime
# ==========================================
# 1. [설정 영역] 날짜 및 사용자 목록 설정
# ==========================================
SEARCH_FROM_DT = "20251101"  # 조회 시작일
SEARCH_TO_DT   = "20260110"  # 조회 종료일
# 사용자 계정 목록 (이곳에 계정을 계속 추가하면 됩니다)
USERS = [
    { "UserId": "ONR01", "Name": "아이디1", "VendorCode": "2008" },
    { "UserId": "TLG01", "Name": "아이디2", "VendorCode": "2009" },
    { "UserId": "INST",  "Name": "아이디3", "VendorCode": "0108" }
]
# 기본 API 설정
BASE_URL = "http://svcbw3net.ebizway.co.kr/api/Data/Post"
COM_CD = "20246"
# ==========================================
# 2. 공통 함수 정의
# ==========================================
def call_api_and_save_excel(filename_prefix, payload, headers, user_info):
    """
    API 호출 후 엑셀로 저장하는 함수
    파일명 형식: [UserId]_파일명_시작일_종료일.xlsx
    """
    user_id = user_info["UserId"]

    try:
        # API 호출
        response = requests.post(BASE_URL, headers=headers, json=payload)
        response.raise_for_status()

        data = response.json()
        # 데이터 존재 여부 확인
        if not data:
            print(f"    Pass: 데이터가 없습니다. ({filename_prefix})")
            return
        # DataFrame 변환
        if isinstance(data, dict):
            df = pd.DataFrame([data])
        else:
            df = pd.DataFrame(data)
        # 저장 경로 및 파일명 설정
        current_folder = os.path.dirname(os.path.abspath(__file__))

        # 파일명에 [UserID]를 추가하여 구분
        filename = f"[{user_id}]_{filename_prefix}_{SEARCH_FROM_DT}_{SEARCH_TO_DT}.xlsx"
        file_path = os.path.join(current_folder, filename)
        # 엑셀 저장
        df.to_excel(file_path, index=False, engine='openpyxl')

        print(f"    Save: {filename} 저장 완료")
    except Exception as e:
        print(f"    Error: {e}")
# ==========================================
# 3. 메인 실행 (반복문 처리)
# ==========================================
def main():
    print(f"=== API 일괄 다운로드 시작 ({SEARCH_FROM_DT} ~ {SEARCH_TO_DT}) ===\n")
    # 사용자 목록을 하나씩 돌면서 작업 수행
    for user in USERS:
        u_id = user["UserId"]
        u_name = user["Name"]
        v_code = user["VendorCode"]
        print(f"▶ 사용자 처리 시작: {u_name} (ID: {u_id}, Vendor: {v_code})")
        # 1. 해당 사용자에 맞는 헤더 생성
        headers = {
            "Accept": "application/json",
            "Accept-Encoding": "gzip,deflate",
            "comcd": COM_CD,
            "usrid": u_id,  # <--- 사용자에 따라 ID 변경
            "Content-Type": "application/json; charset=utf-8",
            "Host": "svcbw3net.ebizway.co.kr",
            "User-Agent": "Python-API-Client"
        }
        # 2. 각 시나리오별 Payload 생성 (VendorCode 변경 적용)

        # (1) 재고현황
        payload_inventory = {
            "comcd": COM_CD,
            "mapinfo": "V4/Pages/MapW2025501.GetData2301_LOCATION",
            "sqlparam": {
                "vendorcd": v_code, # <--- 벤더코드 변경
                "브랜드": None, "상품코드": None, "상품명": None,
                "searchtype": "2", "warehouse": "", "업체바코드": None,
                "IsRefreshCache": True
            }
        }
        call_api_and_save_excel("1_재고현황", payload_inventory, headers, user)
        # (2) 입고조회
        payload_inbound = {
            "comcd": COM_CD,
            "mapinfo": "V4/Pages/MapP020020_v4.GetData16",
            "sqlparam": {
                "vendorcd": v_code, # <--- 벤더코드 변경
                "fromdt": SEARCH_FROM_DT, "todt": SEARCH_TO_DT,
                "브랜드": None, "상품명": None, "상품코드": None, "창고": None,
                "존": None, "행": None, "열": None, "단": None,
                "searchtype": "1", "returnyn": None, "LOT": None,
                "SELLBYDATE": None, "SERIALNUMBER": None, "업체바코드": None
            }
        }
        call_api_and_save_excel("2_입고조회", payload_inbound, headers, user)
        # (3) 반품내역
        payload_return = {
            "comcd": COM_CD,
            "mapinfo": "V4/Pages/MapWR030.GetData01",
            "sqlparam": {
                "vendorcd": v_code, # <--- 벤더코드 변경
                "fromdt": SEARCH_FROM_DT, "todt": SEARCH_TO_DT,
                "주문번호": None, "고객명": None, "운송장번호": None, "상품코드": None,
                "반품상태": "",
                "recfromdt": SEARCH_FROM_DT, "rectodt": SEARCH_TO_DT
            }
        }
        call_api_and_save_excel("3_반품내역", payload_return, headers, user)
        # (4) 출고 진행조회
        payload_outbound = {
            "comcd": COM_CD,
            "mapinfo": "V4/Pages/MapW246000.GetDataW246080",
            "sqlparam": {
                "vendorcd": v_code, # <--- 벤더코드 변경
                "fromdt": SEARCH_FROM_DT, "todt": SEARCH_TO_DT,
                "브랜드": None, "고객명": None, "운송장번호": None, "주문번호": None,
                "상품코드": None, "출고상태": None, "warehouse": None, "매출처": None
            }
        }
        call_api_and_save_excel("4_출고진행", payload_outbound, headers, user)
        print("-" * 50) # 구분선
    print("\n=== 모든 계정의 작업이 완료되었습니다 ===")
if __name__ == "__main__":
    main()
