import requests
import pandas as pd
import os
import datetime

# ==========================================
# 1. [설정 영역] 날짜 및 사용자 목록 설정
# ==========================================
_today = datetime.date.today()
SEARCH_FROM_DT = (_today - datetime.timedelta(days=10)).strftime("%Y%m%d")  # 오늘 기준 10일 전
SEARCH_TO_DT   = _today.strftime("%Y%m%d")                                  # 오늘

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
# 2. 출고 데이터 조회 함수
# ==========================================
def fetch_outbound(user_info):
    """
    단일 사용자의 출고 진행 데이터를 API로 가져와 DataFrame으로 반환.
    사용자 구분을 위해 'UserId', 'Name' 컬럼을 앞에 추가한다.
    """
    u_id   = user_info["UserId"]
    u_name = user_info["Name"]
    v_code = user_info["VendorCode"]

    headers = {
        "Accept": "application/json",
        "Accept-Encoding": "gzip,deflate",
        "comcd": COM_CD,
        "usrid": u_id,
        "Content-Type": "application/json; charset=utf-8",
        "Host": "svcbw3net.ebizway.co.kr",
        "User-Agent": "Python-API-Client"
    }

    payload = {
        "comcd": COM_CD,
        "mapinfo": "V4/Pages/MapW246000.GetDataW246080",
        "sqlparam": {
            "vendorcd": v_code,
            "fromdt": SEARCH_FROM_DT,
            "todt": SEARCH_TO_DT,
            "브랜드": None, "고객명": None, "운송장번호": None, "주문번호": None,
            "상품코드": None, "출고상태": None, "warehouse": None, "매출처": None
        }
    }

    try:
        response = requests.post(BASE_URL, headers=headers, json=payload)
        response.raise_for_status()
        data = response.json()

        if not data:
            print(f"  Pass: 데이터 없음 ({u_name} / {u_id})")
            return None

        df = pd.DataFrame([data] if isinstance(data, dict) else data)

        # 어떤 사용자 데이터인지 식별할 수 있도록 앞에 컬럼 추가
        df.insert(0, "Name",   u_name)
        df.insert(0, "UserId", u_id)

        print(f"  OK : {u_name} ({u_id}) - {len(df)}건")
        return df

    except Exception as e:
        print(f"  Error ({u_name} / {u_id}): {e}")
        return None


# ==========================================
# 3. 메인 실행
# ==========================================
def main():
    print(f"=== 출고 데이터 합산 다운로드 시작 ({SEARCH_FROM_DT} ~ {SEARCH_TO_DT}) ===\n")

    frames = []
    for user in USERS:
        print(f"▶ {user['Name']} (ID: {user['UserId']}, Vendor: {user['VendorCode']})")
        df = fetch_outbound(user)
        if df is not None:
            frames.append(df)

    if not frames:
        print("\n저장할 데이터가 없습니다.")
        return

    # 모든 사용자 데이터를 하나의 DataFrame으로 합치기
    combined = pd.concat(frames, ignore_index=True)

    # 저장 경로 및 파일명
    current_folder = os.path.dirname(os.path.abspath(__file__))
    filename  = f"출고진행_합산_{SEARCH_FROM_DT}_{SEARCH_TO_DT}.xlsx"
    file_path = os.path.join(current_folder, filename)

    combined.to_excel(file_path, index=False, engine="openpyxl")

    print(f"\n=== 저장 완료 ===")
    print(f"파일명 : {filename}")
    print(f"총 건수: {len(combined)}건")


if __name__ == "__main__":
    main()
