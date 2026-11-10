"""공공데이터포털 API 키가 잘 읽히는지 확인하는 스크립트.

읽는 순서
1) Windows 환경변수 DATA_GO_KR_API_KEY
2) 없으면 이 폴더의 .env 파일

실행: .venv\\Scripts\\python.exe check_api_key.py
      .venv\\Scripts\\python.exe check_api_key.py --call   (실제 API 1건 호출)
"""
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

KEY_NAME = "DATA_GO_KR_API_KEY"
ENV_FILE = Path(__file__).with_name(".env")

# override=False → Windows 환경변수가 이미 있으면 그 값을 우선 사용
source = "Windows 환경변수" if os.getenv(KEY_NAME) else ".env 파일"
load_dotenv(ENV_FILE, override=False)
key = os.getenv(KEY_NAME, "").strip()

if not key or key.startswith("여기에"):
    print(f"[미설정] {KEY_NAME} 값이 없어요. .env 파일이나 환경변수에 키를 넣어주세요.")
    sys.exit(1)

print(f"[확인] {KEY_NAME} 읽기 성공 (출처: {source}, 길이 {len(key)}자, 앞 4자리 {key[:4]}****)")

if "--call" in sys.argv:
    import requests

    url = "https://apis.data.go.kr/1160100/service/GetFinaStatInfoService_V2/getSummFinaStat"
    params = {"serviceKey": key, "resultType": "json", "numOfRows": 1, "pageNo": 1, "bizYear": "2024"}
    r = requests.get(url, params=params, timeout=15)
    print(f"[호출] HTTP {r.status_code}")
    print(r.text[:500])
