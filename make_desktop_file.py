import os
import pandas as pd
import random

# 아파트 실무 맞춤형 1년 치 과거 전표 샘플 데이터
items = [
    ("5101", "일반관리비", "3월 단지 공용 소모품 및 사무용품 구입"),
    ("5102", "청소비", "지하주차장 청소용역비 지급"),
    ("5102", "청소비", "재활용품 분리수거용 마대 및 쓰레기봉투 구입"),
    ("5103", "경비비", "경비원 피복비 및 제니토 일괄 구입"),
    ("5105", "승강기유지비", "101동 승강기 정기 안전점검 및 부품 교체"),
    ("5106", "수선유지비", "단지 외곽 보안등 LED 램프 교체 공사"),
    ("5107", "광고선전비/도서인쇄비", "입주자대표회의 선거 현수막 및 안내문 인쇄"),
    ("5101", "일반관리비", "사무실 복사기 토너 및 용지 구입"),
    ("5102", "청소비", "단지 정화조 주변 대청소 용역 대금"),
    ("5105", "승강기유지비", "102동 승강기 도어 인버터 수리비"),
    ("5102", "소독비", "단지 지하 및 승강기 정기 방역소독"),
    ("5106", "수선유지비", "공용현관 자동문 수리 및 힌지 교체"),
]

rows = []
months = [f"{m:02d}" for m in range(1, 13)]

# 1년 치 약 60건의 전표 데이터 시뮬레이션 생성
for i in range(1, 61):
    base = random.choice(items)
    month = random.choice(months)
    day = f"{random.randint(1, 28):02d}"
    date_str = f"2025-{month}-{day}"
    amount = random.randint(5, 200) * 10000
    
    rows.append({
        "전표일자": date_str,
        "계정코드": base[0],
        "계정명": base[1],
        "적요": f"{base[2]} ({month}월)",
        "차변금액": amount,
        "대변금액": amount
    })

df = pd.DataFrame(rows)

# 윈도우 바탕화면 경로 자동 탐지 후 저장
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
file_path = os.path.join(desktop_path, "sample_historical_entries.xlsx")

df.to_excel(file_path, index=False)
print(f"🎉 성공! 바탕화면에 파일이 저장되었습니다: {file_path}")
