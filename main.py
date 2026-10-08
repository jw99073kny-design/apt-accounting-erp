import csv
import json
import os

DATA_FILE = "invoices.json"
CSV_FILE = "invoices_report.csv"

# --- [3단계] 데이터 저장 및 불러오기 ---
def load_invoices():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_invoices(invoices):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(invoices, f, ensure_ascii=False, indent=4)
    print("💾 데이터가 'invoices.json' 파일에 안전하게 저장되었습니다.")

# --- [4단계] AI / OCR 명세서 인식 시뮬레이션 ---
def ai_ocr_process_invoice(file_path):
    print(f"\n🤖 [AI OCR 분석 중] '{file_path}' 파일 검토...")
    
    sample_ocr_db = {
        "invoice_elevator.png": {
            "vendor": "현대엘리베이터(주)",
            "date": "2026-10-07",
            "supply_value": 750000,
            "vat": 75000,
            "category": "수선유지비"
        },
        "invoice_fire.pdf": {
            "vendor": "(주)한국소방방재",
            "date": "2026-10-08",
            "supply_value": 450000,
            "vat": 45000,
            "category": "위탁관리비"
        },
        "invoice_suspicious.png": {
            "vendor": "대박건설(주)",
            "date": "2026-10-09",
            "supply_value": 15000000,
            "vat": 1500000,
            "category": "수선유지비"
        }
    }

    filename = os.path.basename(file_path)
    if filename in sample_ocr_db:
        return sample_ocr_db[filename]
    return {
        "vendor": "미등록 거래처",
        "date": "2026-10-07",
        "supply_value": 100000,
        "vat": 10000,
        "category": "잡지출"
    }

# --- [5단계] 회계 자동 검증 및 이상 지출 탐지 ---
def audit_invoice(invoices, category, supply_value, vat):
    expected_vat = int(supply_value * 0.1)
    if vat != expected_vat:
        return f"⚠️ 경고: 부가세 불일치 (추정: {expected_vat:,}원)"

    same_category_totals = [inv["total"] for inv in invoices if inv["category"] == category]
    
    if same_category_totals:
        avg_total = sum(same_category_totals) / len(same_category_totals)
        current_total = supply_value + vat
        if current_total > avg_total * 3:
            return f"🚨 이상 지출 주의 (평균 {int(avg_total):,}원 대비 {current_total / avg_total:.1f}배)"

    return "정상"

def add_invoice_from_ocr(invoices, file_path):
    data = ai_ocr_process_invoice(file_path)
    status = audit_invoice(invoices, data["category"], data["supply_value"], data["vat"])

    total = data["supply_value"] + data["vat"]
    new_id = len(invoices) + 1 if not invoices else max(inv["id"] for inv in invoices) + 1
    new_item = {
        "id": new_id,
        "vendor": data["vendor"],
        "date": data["date"],
        "supply_value": data["supply_value"],
        "vat": data["vat"],
        "total": total,
        "category": data["category"],
        "status": status,
        "source_file": file_path
    }
    invoices.append(new_item)
    save_invoices(invoices)

# --- [6단계] 대시보드 리포트 및 CSV 내보내기 ---
def print_dashboard(invoices):
    print("\n==========================================================================================")
    print("📊 [푸른솔아파트] 관리비 지출 현황 요약 대시보드")
    print("==========================================================================================")
    if not invoices:
        print("등록된 데이터가 없습니다.")
        return

    total_expense = sum(inv["total"] for inv in invoices)
    print(f"💰 총 전체 지출액: {total_expense:,}원 (총 {len(invoices)}건)")

    # 계정과목별 집계
    category_summary = {}
    for inv in invoices:
        cat = inv["category"]
        category_summary[cat] = category_summary.get(cat, 0) + inv["total"]

    print("\n[계정과목별 지출 합계]")
    for cat, amt in category_summary.items():
        print(f"  - {cat:<10}: {amt:>12,}원")

    # 이상 지출 항목 필터링
    anomalies = [inv for inv in invoices if "정상" not in inv.get("status", "정상")]
    if anomalies:
        print("\n🚨 [검토 필요 이상 지출 내역]")
        for a in anomalies:
            print(f"  - [ID {a['id']}] {a['vendor']} ({a['category']}) : {a['total']:,}원 | 사유: {a['status']}")
    print("==========================================================================================\n")

def export_to_csv(invoices):
    if not invoices:
        print("내보낼 데이터가 없습니다.")
        return

    fieldnames = ["id", "date", "vendor", "category", "supply_value", "vat", "total", "status", "source_file"]
    
    with open(CSV_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for inv in invoices:
            writer.writerow(inv)

    print(f"📑 엑셀용 CSV 파일이 '{CSV_FILE}' 경로로 성공적으로 내보내졌습니다!")

# --- 실행부 ---
if __name__ == "__main__":
    invoices = load_invoices()

    # 1. 요약 대시보드 출력
    print_dashboard(invoices)

    # 2. 엑셀 CSV 내보내기 실행
    export_to_csv(invoices)
    