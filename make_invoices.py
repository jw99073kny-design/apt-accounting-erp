import os
from PIL import Image, ImageDraw

# 저장할 폴더 생성
output_dir = "sample_invoices"
os.makedirs(output_dir, exist_ok=True)

# 생성할 영수증 샘플 목록 (파일명, 품목, 금액, 담당자/전화번호)
invoices_data = [
    ("receipt_01_led.png", "지하주차장 LED 등기구 교체 공사비", "350,000원", "홍길동 소장 (010-1234-5678)"),
    ("receipt_02_cleaning.png", "단지 정화조 주변 대청소 용역 대금", "1,200,000원", "김철수 과장 (010-9876-5432)"),
    ("receipt_03_elevator.png", "102동 승강기 도어 인버터 수리비", "450,000원", "(주)한국엘리베이터 (02-555-1234)"),
    ("receipt_04_disinfect.png", "단지 지하 및 승강기 정기 방역소독", "280,000원", "클린방역 이영수 (010-5555-7777)"),
    ("receipt_05_supplies.png", "사무실 복사기 토너 및 A4용지 구입", "110,000원", "오피스디포 (02-123-4567)")
]

for filename, item, amount, contact in invoices_data:
    # 이미지 생성 (가로 600, 세로 250 흰색 배경)
    img = Image.new('RGB', (600, 250), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    
    # 텍스트 그리기
    d.text((30, 30), f"거래 명세서 / 영수증", fill=(0, 0, 0))
    d.text((30, 70), f"품목: {item}", fill=(0, 0, 0))
    d.text((30, 110), f"금액: {amount}", fill=(0, 0, 0))
    d.text((30, 150), f"공급자/연락처: {contact}", fill=(100, 100, 100))
    
    file_path = os.path.join(output_dir, filename)
    img.save(file_path)

print(f"🎉 성공! '{output_dir}' 폴더 안에 테스트용 영수증 샘플 5장이 생성되었습니다!")
