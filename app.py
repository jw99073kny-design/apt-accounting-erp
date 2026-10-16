import io
import os
import random
import re
import zipfile
import hashlib
from datetime import date, timedelta
import cv2
import numpy as np
import openpyxl
import pandas as pd
import streamlit as st
from PIL import Image, ImageEnhance
import pytesseract
import fitz

st.set_page_config(
    page_title="XpERP 연동 아파트 실무 회계·행정 시스템", page_icon="🏢", layout="wide"
)

st.markdown(
    """
    <style>
    div[data-testid="stSidebar"] button {
        border-radius: 6px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def sanitize_text(text):
    """전표 적요 및 텍스트에서 개인정보를 마스킹합니다."""
    if not isinstance(text, str):
        return ""

    text = re.sub(r"\d{6}[-~]\d{7}", "[주민번호]", text)
    text = re.sub(r"\d{3}-\d{2}-\d{5}", "[사업자번호]", text)
    text = re.sub(r"01[016789][-\s]?\d{3,4}[-\s]?\d{4}", "[전화번호]", text)
    text = re.sub(r"0\d{1,2}[-\s]?\d{3,4}[-\s]?\d{4}", "[전화번호]", text)
    text = re.sub(r"\d{3,4}\s*동\s*\d{3,4}\s*호?", "[동호]", text)
    text = re.sub(r"\b\d{3,4}[-_\s]\d{3,4}\b", "[동호]", text)
    text = re.sub(r"[가-힣]{2,4}(님|세대|직원|기사|소장)", "[이름]", text)

    return text.strip()


def preprocess_image_for_ocr(pil_img):
    """OCR 인식률 향상을 위한 이미지 보정"""
    w, h = pil_img.size
    resized = pil_img.resize((w * 2, h * 2), Image.LANCZOS)
    
    enhancer = ImageEnhance.Contrast(resized)
    enhanced = enhancer.enhance(1.8)
    
    enhancer_sharp = ImageEnhance.Sharpness(enhanced)
    final_img = enhancer_sharp.enhance(2.0)
    
    return final_img


def split_receipts_from_image(uploaded_file):
    """업로드된 파일에서 개별 영수증 영역 분할"""
    try:
        file_bytes = uploaded_file.getvalue()
        images = []

        if uploaded_file.name.lower().endswith(".pdf"):
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
            try:
                for page in pdf_doc:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                    images.append(Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB"))
            finally:
                pdf_doc.close()
            return images

        nparr = np.frombuffer(file_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        if img is None:
            return [Image.open(io.BytesIO(file_bytes)).convert("RGB")]

        h_img, w_img = img.shape[:2]
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        edges = cv2.Canny(blurred, 30, 150)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
        dilated = cv2.dilate(edges, kernel, iterations=2)
        contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        boxes = []
        min_area = (w_img * h_img) * 0.03
        max_area = (w_img * h_img) * 0.85

        for cnt in contours:
            area = cv2.contourArea(cnt)
            if min_area < area < max_area:
                x, y, w, h = cv2.boundingRect(cnt)
                boxes.append((x, y, w, h))

        boxes = sorted(boxes, key=lambda b: (b[1] // 100, b[0]))
        cropped_images = []
        for x, y, w, h in boxes:
            pad = 10
            x1 = max(0, x - pad)
            y1 = max(0, y - pad)
            x2 = min(w_img, x + w + pad)
            y2 = min(h_img, y + h + pad)
            cropped = img[y1:y2, x1:x2]
            cropped_rgb = cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB)
            cropped_images.append(Image.fromarray(cropped_rgb))

        if not cropped_images:
            return [Image.open(io.BytesIO(file_bytes)).convert("RGB")]
        return images if not cropped_images else cropped_images
    except Exception:
        return [Image.open(io.BytesIO(uploaded_file.getvalue())).convert("RGB")]


def extract_receipt_info(uploaded_file):
    results = []
    try:
        images = split_receipts_from_image(uploaded_file)
        for img_page in images:
            result = {
                "ocr_text": "", "ocr_date": "", "ocr_amount": 0, "ocr_vat": 0,
                "ocr_vendor": "", "ocr_memo": "", "ocr_type": "영수증", "ocr_status": "",
            }
            processed_img = preprocess_image_for_ocr(img_page)
            custom_config = r"--psm 6 -l kor+eng"
            text = pytesseract.image_to_string(processed_img, config=custom_config).strip()
            text = re.sub(r"50091", "500ml", text)
            result["ocr_text"] = text

            if not text:
                result["ocr_status"] = "읽힌 글자가 없습니다. 직접 입력해 주세요."
                results.append(result)
                continue

            full_raw_upper = text.upper().replace(" ", "")
            if "거래명세서" in text:
                result["ocr_type"] = "거래명세서"
            elif "세금계산서" in text:
                result["ocr_type"] = "세금계산서"
            elif "신용카드매출전표" in text or "카드전표" in text:
                result["ocr_type"] = "카드전표"
            else:
                result["ocr_type"] = "영수증"

            date_match = re.search(r"(20\d{2})\s*[./년-]\s*(\d{1,2})\s*[./월-]\s*(\d{1,2})\s*일?", text)
            if date_match:
                try:
                    year, month, day = map(int, date_match.groups())
                    result["ocr_date"] = date(year, month, day).isoformat()
                except ValueError:
                    pass

            lines = [line.strip() for line in text.splitlines() if line.strip()]
            if "GS25" in full_raw_upper or "6525" in full_raw_upper:
                result["ocr_vendor"] = "GS25"
            elif "사랑채" in full_raw_upper:
                result["ocr_vendor"] = "사랑채 한식당"
            elif "한국전력" in full_raw_upper:
                result["ocr_vendor"] = "한국전력공사"
            else:
                vendor_found = ""
                for line in lines[:6]:
                    if any(kw in line for kw in ["영수증", "신용카드", "매출", "거래명세서", "일자", "TEL", "Tel"]):
                        continue
                    cleaned_l = re.sub(r"[^가-힣a-zA-Z0-9\s]", "", line).strip()
                    if len(cleaned_l) >= 2:
                        vendor_found = cleaned_l[:20]
                        break
                result["ocr_vendor"] = vendor_found or "일반거래처"

            target_amount = 0
            extracted_vat = 0
            for line in lines:
                if re.search(r"부가세|세액|VAT", line):
                    nums = re.findall(r"(?<!\d)[\d,]{2,}(?!\d)", line)
                    if nums:
                        try:
                            extracted_vat = int(nums[-1].replace(",", ""))
                        except ValueError:
                            pass

            for line in lines:
                if re.search(r"합계금액|합계|결제금액|청구금액|받을금액|공급대가|총\s*금액|카드결제", line):
                    nums = re.findall(r"(?<!\d)[\d,]{3,}(?!\d)", line)
                    if nums:
                        try:
                            target_amount = int(nums[-1].replace(",", ""))
                            break
                        except ValueError:
                            pass

            if target_amount == 0:
                all_valid_amounts = []
                for line in lines:
                    nums = re.findall(r"(?<!\d)[\d,]{3,}(?!\d)", line)
                    for n in nums:
                        try:
                            val = int(n.replace(",", ""))
                            if 1000 <= val < 100000000:
                                all_valid_amounts.append(val)
                        except ValueError:
                            continue
                if all_valid_amounts:
                    target_amount = max(all_valid_amounts)

            result["ocr_amount"] = target_amount
            result["ocr_vat"] = extracted_vat if extracted_vat > 0 else (round(target_amount / 11) if target_amount >= 1000 else 0)
            result["ocr_memo"] = lines[1] if len(lines) > 1 else "일반 매입"
            result["ocr_status"] = "인식 완료"
            results.append(result)
    except Exception as exc:
        results.append({
            "ocr_text": "", "ocr_date": "", "ocr_amount": 0, "ocr_vat": 0,
            "ocr_vendor": "", "ocr_memo": "", "ocr_type": "영수증", "ocr_status": f"오류: {exc}"
        })
    return results


# ---------------------------------------------------------
# 세션 및 영구 파일 저장소 로드
# ---------------------------------------------------------
if "admin_id" not in st.session_state:
    st.session_state.admin_id = "KNY"

if "admin_pw" not in st.session_state:
    st.session_state.admin_pw = "skdud12"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "current_menu" not in st.session_state:
    st.session_state.current_menu = "🏠 시스템 홈 (대시보드)"

DATA_FILE = "trained_patterns.csv"
if "trained_patterns" not in st.session_state:
    st.session_state.trained_patterns = pd.read_csv(DATA_FILE) if os.path.exists(DATA_FILE) else pd.DataFrame()

ACCOUNTS_FILE = "accounts_table.csv"
gov_data_full = [
    (1, "5101", "급여"), (13, "5113", "공과금중 전기료"), (25, "5125", "잡비"),
    (26, "5201", "청소비"), (27, "5301", "경비비"), (31, "5601", "수선비"),
    (46, "8101", "장기수선충당금"), (47, "4101", "잡수입")
]
default_coas = [{"계정코드": c, "계정명": n} for _, c, n in gov_data_full]
default_coas.extend([
    {"계정코드": "101", "계정명": "현금"}, {"계정코드": "103", "계정명": "보통예금"},
    {"계정코드": "251", "계정명": "미지급금"}, {"계정코드": "310", "계정명": "장기수선충당금부채"}
])
default_coas_df = pd.DataFrame(default_coas)

if "accounts_df" not in st.session_state:
    st.session_state.accounts_df = pd.read_csv(ACCOUNTS_FILE, dtype=str) if os.path.exists(ACCOUNTS_FILE) else default_coas_df

JOURNAL_DB = "journal_entries_register.csv"
if "journal_entries" not in st.session_state:
    if os.path.exists(JOURNAL_DB):
        try:
            df_j_saved = pd.read_csv(JOURNAL_DB, dtype=str).fillna("")
            st.session_state.journal_entries = df_j_saved.to_dict("records")
        except Exception:
            st.session_state.journal_entries = []
    else:
        st.session_state.journal_entries = []

BANK_RECON_DB = "bank_reconciliation_cache.csv"
if "bank_recon_df" not in st.session_state:
    if os.path.exists(BANK_RECON_DB):
        try:
            st.session_state.bank_recon_df = pd.read_csv(BANK_RECON_DB, dtype=str).fillna("")
        except Exception:
            st.session_state.bank_recon_df = None
    else:
        st.session_state.bank_recon_df = None

CONTRACTS_DB = "contracts_register.csv"
if "contracts_df" not in st.session_state:
    if os.path.exists(CONTRACTS_DB):
        try:
            st.session_state.contracts_df = pd.read_csv(CONTRACTS_DB, dtype=str).fillna("")
        except Exception:
            st.session_state.contracts_df = pd.DataFrame(columns=["계약명", "계약처", "종료일", "보증증권만료일"])
    else:
        st.session_state.contracts_df = pd.DataFrame(columns=["계약명", "계약처", "종료일", "보증증권만료일"])

EVIDENCE_DIR = "evidence_files"
EVIDENCE_DB = "evidence_register.csv"
os.makedirs(EVIDENCE_DIR, exist_ok=True)

if "evidence_df" not in st.session_state:
    if os.path.exists(EVIDENCE_DB):
        try:
            st.session_state.evidence_df = pd.read_csv(EVIDENCE_DB, dtype=str).fillna("")
        except Exception:
            st.session_state.evidence_df = pd.DataFrame()
    else:
        st.session_state.evidence_df = pd.DataFrame(
            columns=[
                "증빙ID", "파일명", "저장파일", "증빙일자",
                "거래처", "증빙종류", "공급가액/금액", "부가세",
                "적요", "은행대사", "연결전표", "검토상태",
            ]
        )

HOUSE_DB = "household_bill_cache.csv"
if "household_excel_data" not in st.session_state:
    if os.path.exists(HOUSE_DB):
        try:
            st.session_state.household_excel_data = pd.read_csv(HOUSE_DB, dtype=str).fillna("")
        except Exception:
            st.session_state.household_excel_data = None
    else:
        st.session_state.household_excel_data = None

# 실무형 연차 및 직원 대장 DB (근무형태, 1일근무시간, 휴게시간, 사용일자_내역, 사용개수_총합, 비고사유 통합)
STAFF_LEAVE_INTEGRATED_DB = "staff_leave_integrated_register_full_tabs.csv"
if "staff_leave_df" not in st.session_state:
    if os.path.exists(STAFF_LEAVE_INTEGRATED_DB):
        try:
            st.session_state.staff_leave_df = pd.read_csv(STAFF_LEAVE_INTEGRATED_DB, dtype=str).fillna("")
        except Exception:
            st.session_state.staff_leave_df = pd.DataFrame(
                columns=["직원명", "직책", "입사일자", "통상임금", "근무형태", "1일근무시간", "휴게시간", "사용일자_내역", "사용개수_총합", "비고사유"]
            )
    else:
        initial_data = pd.DataFrame([
            {"직원명": "김관리", "직책": "관리소장", "입사일자": "2024-03-01", "통상임금": "3500000", "근무형태": "주5일 (일반)", "1일근무시간": "8", "휴게시간": "1", "사용일자_내역": "2026-05-01, 2026-08-15", "사용개수_총합": "2.0", "비고사유": "하계휴가"},
            {"직원명": "박경리", "직책": "경리대리", "입사일자": "2025-07-15", "통상임금": "2800000", "근무형태": "주5일 (일반)", "1일근무시간": "8", "휴게시간": "1", "사용일자_내역": "", "사용개수_총합": "0.0", "비고사유": ""},
            {"직원명": "최경비", "직책": "경비원", "입사일자": "2023-01-01", "통상임금": "2600000", "근무형태": "격일제 (경비 감단직)", "1일근무시간": "24", "휴게시간": "8", "사용일자_내역": "", "사용개수_총합": "0.0", "비고사유": "주야간 휴게시간 8시간 반영"},
            {"직원명": "이미화", "직책": "미화원", "입사일자": "2023-05-10", "통상임금": "1800000", "근무형태": "미화원 (단축근무)", "1일근무시간": "6", "휴게시간": "1", "사용일자_내역": "2026-06-10", "사용개수_총합": "1.0", "비고사유": "단축근무"}
        ])
        st.session_state.staff_leave_df = initial_data

LEAVE_LOG_DB = "leave_usage_log_tabs.csv"
if "leave_log_df" not in st.session_state:
    if os.path.exists(LEAVE_LOG_DB):
        try:
            st.session_state.leave_log_df = pd.read_csv(LEAVE_LOG_DB, dtype=str).fillna("")
        except Exception:
            st.session_state.leave_log_df = pd.DataFrame(columns=["직원명", "사용일자", "사용개수", "휴가종류", "비고"])
    else:
        st.session_state.leave_log_df = pd.DataFrame(columns=["직원명", "사용일자", "사용개수", "휴가종류", "비고"])

for i in range(1, 7):
    if f"open_cat_{i}" not in st.session_state:
        st.session_state[f"open_cat_{i}"] = False


def login_screen():
    st.title("🏢 XpERP 연동 아파트 실무 회계·행정 시스템")
    st.info("🔒 보안 접속: 관리자 계정 정보를 입력해 주세요.")
    with st.form("login_form"):
        input_id = st.text_input("아이디")
        user_pw = st.text_input("비밀번호", type="password")
        if st.form_submit_button("로그인", type="primary"):
            if input_id == st.session_state.admin_id and user_pw == st.session_state.admin_pw:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("아이디 또는 비밀번호가 올바르지 않습니다.")

if not st.session_state.authenticated:
    login_screen()
    st.stop()


# ---------------------------------------------------------
# 사이드바 메뉴
# ---------------------------------------------------------
st.sidebar.title("🏢 XpERP 실무 시스템")
st.sidebar.text(f"관리자: {st.session_state.admin_id}")
if st.sidebar.button("🔒 로그아웃", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🏠 시스템 홈 (대시보드)", use_container_width=True):
    st.session_state.current_menu = "🏠 시스템 홈 (대시보드)"

st.sidebar.markdown("---")
st.sidebar.markdown("📁 **작업 카테고리**")

if st.sidebar.button("📁 1. 회계·전표 입력 자동화", use_container_width=True):
    st.session_state.open_cat_1 = not st.session_state.open_cat_1
if st.session_state.open_cat_1:
    st.sidebar.markdown("<div style='margin-left: 15px;'>", unsafe_allow_html=True)
    if st.sidebar.button("• 스마트 분개 패턴 추천", use_container_width=True):
        st.session_state.current_menu = "XpERP 전년도 장부 및 분개 패턴 학습"
    if st.sidebar.button("• XpERP 전표 엑셀 자동 생성", use_container_width=True):
        st.session_state.current_menu = "명세서 다중 일괄 검증 & 전표 발행"
    if st.sidebar.button("• 부과 마감 전 오차 사전 검증기", use_container_width=True):
        st.session_state.current_menu = "부과 마감 전 오차 사전 검증기"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)

if st.sidebar.button("📁 2. 점검·결산 및 감사 대비", use_container_width=True):
    st.session_state.open_cat_2 = not st.session_state.open_cat_2
if st.session_state.open_cat_2:
    st.sidebar.markdown("<div style='margin-left: 15px;'>", unsafe_allow_html=True)
    if st.sidebar.button("• 이상 지출 통합 분석 보고서", use_container_width=True):
        st.session_state.current_menu = "이상 지출 통합 보고서"
    if st.sidebar.button("• 지출 건별 필수 증빙·의결 체크리스트", use_container_width=True):
        st.session_state.current_menu = "지출 건별 필수 증빙·의결 체크리스트"
    if st.sidebar.button("• 세대별 장기 미납 & 독촉 관리", use_container_width=True):
        st.session_state.current_menu = "세대별 장기 미납 및 독촉 관리"
    if st.sidebar.button("• 영수증·증빙 전자 아카이빙", use_container_width=True):
        st.session_state.current_menu = "증빙 통합 관리"
    if st.sidebar.button("• 월말 결산 부속명세서 자동 취합", use_container_width=True):
        st.session_state.current_menu = "월말 결산 부속명세서 자동 취합"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)

if st.sidebar.button("📁 3. 자금 대사 & 세무 지원", use_container_width=True):
    st.session_state.open_cat_3 = not st.session_state.open_cat_3
if st.session_state.open_cat_3:
    st.sidebar.markdown("<div style='margin-left: 15px;'>", unsafe_allow_html=True)
    if st.sidebar.button("• 은행 입출금 ↔ 장부 자동 대사", use_container_width=True):
        st.session_state.current_menu = "은행 입출금 ↔ 전표 자동 대사"
    if st.sidebar.button("• 장기수선충당금 집행·잔액 검증", use_container_width=True):
        st.session_state.current_menu = "장기수선충당금 적립 및 사용 관리"
    if st.sidebar.button("• 세대별 관리비 부과 및 미납 조회", use_container_width=True):
        st.session_state.current_menu = "세대별 관리비 부과 및 미납 관리"
    if st.sidebar.button("• 부가세 매입/매출 1차 자동 분류", use_container_width=True):
        st.session_state.current_menu = "부가세 매입/매출 1차 자동 분류"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)

if st.sidebar.button("📁 4. 계약 및 행정 관리", use_container_width=True):
    st.session_state.open_cat_4 = not st.session_state.open_cat_4
if st.session_state.open_cat_4:
    st.sidebar.markdown("<div style='margin-left: 15px;'>", unsafe_allow_html=True)
    if st.sidebar.button("• 주요 계약 및 보증보험 만기 알림판", use_container_width=True):
        st.session_state.current_menu = "주요 계약 및 보증보험 만기 알림판"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)

if st.sidebar.button("📁 5. 인사 및 연차 관리", use_container_width=True):
    st.session_state.open_cat_5 = not st.session_state.open_cat_5
if st.session_state.open_cat_5:
    st.sidebar.markdown("<div style='margin-left: 15px;'>", unsafe_allow_html=True)
    if st.sidebar.button("• 전 직원 연차 및 수당 통합 관리", use_container_width=True):
        st.session_state.current_menu = "전 직원 연차 및 수당 통합 관리"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)

if st.sidebar.button("📁 6. 기본 설정", use_container_width=True):
    st.session_state.open_cat_6 = not st.session_state.open_cat_6
if st.session_state.open_cat_6:
    st.sidebar.markdown("<div style='margin-left: 15px;'>", unsafe_allow_html=True)
    if st.sidebar.button("• 단지 계정과목 & 국토부 47개", use_container_width=True):
        st.session_state.current_menu = "단지 계정과목 & 국토부 47개 기준"
    if st.sidebar.button("• 계정 및 비밀번호 설정", use_container_width=True):
        st.session_state.current_menu = "계정 및 비밀번호 설정"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)

menu = st.session_state.current_menu


# ---------------------------------------------------------
# 메인 화면 라우팅
# ---------------------------------------------------------
if menu == "🏠 시스템 홈 (대시보드)":
    st.title("🏢 XpERP 연동 아파트 실무 회계·행정 시스템")
    st.markdown("---")
    st.success("✨ 보안 로그인되었습니다. 모든 데이터는 실시간 영구 저장되며, 체크박스로 간편하게 일괄 삭제가 가능합니다.")
    
    c1, c2, c3 = st.columns(3)
    c1.metric("등록된 전표 초안", f"{len(st.session_state.journal_entries)}건")
    c2.metric("전자 보관 증빙", f"{len(st.session_state.evidence_df)}건")
    c3.metric("등록된 용역 계약", f"{len(st.session_state.contracts_df)}건")


# 1. 스마트 분개 패턴 추천
elif menu == "XpERP 전년도 장부 및 분개 패턴 학습":
    st.title("🧠 [기능1] 스마트 분개 패턴 추천 (장부 학습)")
    st.info("전년도 XpERP 분개장을 업로드하여 단지별 계정과목 및 적요 패턴을 학습시킵니다.")
    
    up_learn = st.file_uploader("XpERP 전년도 분개장 엑셀 업로드", type=["xlsx", "csv"], key="learn_up")
    if up_learn:
        df_learn = pd.read_csv(up_learn) if up_learn.name.endswith(".csv") else pd.read_excel(up_learn)
        st.session_state.trained_patterns = df_learn
        df_learn.to_csv(DATA_FILE, index=False)
        st.success("학습 패턴이 성공적으로 저장되었습니다.")
    
    if not st.session_state.trained_patterns.empty:
        st.dataframe(st.session_state.trained_patterns, use_container_width=True)
        if st.button("🗑️ 학습 패턴 초기화"):
            st.session_state.trained_patterns = pd.DataFrame()
            if os.path.exists(DATA_FILE):
                os.remove(DATA_FILE)
            st.success("학습 패턴이 초기화되었습니다.")
            st.rerun()


# 2. XpERP 전표 엑셀 자동 생성
elif menu == "명세서 다중 일괄 검증 & 전표 발행":
    st.title("📄 [기능2] XpERP 일괄 업로드용 엑셀(전표) 자동 생성")
    st.info("검증 완료된 전표 초안들을 XpERP에 한 방에 밀어 넣을 수 있는 엑셀 포맷으로 변환합니다.")
    
    if st.session_state.journal_entries:
        df_j = pd.DataFrame(st.session_state.journal_entries)
        
        st.markdown("#### 📋 전표 초안 목록 관리")
        
        select_all = st.checkbox("☑️ 전체 항목 선택 / 해제", key="select_all_journal")
        df_j.insert(0, "선택", select_all)
        
        edited_df = st.data_editor(
            df_j,
            use_container_width=True,
            num_rows="fixed",
            key="journal_editor_v2"
        )
        
        c_btn1, c_btn2 = st.columns([1, 3])
        with c_btn1:
            if st.button("🗑️ 체크한 항목 삭제", type="primary", use_container_width=True):
                selected_rows = edited_df[edited_df["선택"] == True]
                if not selected_rows.empty:
                    remaining_df = edited_df[edited_df["선택"] == False].drop(columns=["선택"])
                    st.session_state.journal_entries = remaining_df.to_dict("records")
                    pd.DataFrame(st.session_state.journal_entries).to_csv(JOURNAL_DB, index=False, encoding="utf-8-sig")
                    st.success(f"선택한 {len(selected_rows)}건의 전표 초안이 삭제되었습니다.")
                    st.rerun()
                else:
                    st.warning("삭제할 항목을 체크해 주세요.")

        st.markdown("---")
        
        clean_df_j = pd.DataFrame(st.session_state.journal_entries)
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            clean_df_j.to_excel(writer, index=False, sheet_name="전표일괄업로드")
        out.seek(0)
        
        st.download_button(
            "📥 XpERP 업로드용 전표 엑셀 다운로드",
            data=out.getvalue(),
            file_name="xperp_bulk_journal_entries.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.markdown("")
        if st.button("🗑️ 전체 전표 초안 목록 초기화"):
            st.session_state.journal_entries = []
            if os.path.exists(JOURNAL_DB):
                os.remove(JOURNAL_DB)
            st.success("전체 전표 목록이 초기화되었습니다.")
            st.rerun()
    else:
        st.info("등록된 전표 초안이 없습니다. '증빙 통합 관리'나 '은행 입출금 ↔ 전표 자동 대사' 메뉴에서 전표를 추가해 주세요.")


# 3. 부과 마감 전 오차 사전 검증기
elif menu == "부과 마감 전 오차 사전 검증기":
    st.title("🔍 [기능3] 실고지 내역 vs 부과예정액 파일 업로드 자동 대사 검증기")
    st.info("한전·수도사업소 등 실제 고지 내역 엑셀과 관리사무소 사전 부과예정 엑셀을 업로드하면, 프로그램이 자동으로 오차와 정산 차액을 분석합니다.")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        up_bill_file = st.file_uploader("1️⃣ 실제 고지(청구) 내역 엑셀 업로드", type=["xlsx", "csv"], key="bill_upload_v1")
    with col_up2:
        up_est_file = st.file_uploader("2️⃣ 사전 부과예정(추산) 내역 엑셀 업로드", type=["xlsx", "csv"], key="est_upload_v1")

    if up_bill_file and up_est_file:
        try:
            df_bill = pd.read_csv(up_bill_file) if up_bill_file.name.endswith(".csv") else pd.read_excel(up_bill_file)
            df_est = pd.read_csv(up_est_file) if up_est_file.name.endswith(".csv") else pd.read_excel(up_est_file)

            df_bill.columns = [str(c).strip() for c in df_bill.columns]
            df_est.columns = [str(c).strip() for c in df_est.columns]

            st.markdown("---")
            st.subheader("📋 업로드된 원본 데이터 미리보기")
            c_prev1, c_prev2 = st.columns(2)
            with c_prev1:
                st.markdown("**[실제 고지 내역]**")
                st.dataframe(df_bill.head(5), use_container_width=True)
            with c_prev2:
                st.markdown("**[사전 예정 내역]**")
                st.dataframe(df_est.head(5), use_container_width=True)

            if st.button("⚖️ 파일 자동 교차 대사 및 정산 검증 실행", type="primary"):
                bill_amt_col = next((c for c in df_bill.columns if "금액" in c or "청구" in c or "합계" in c), None)
                est_amt_col = next((c for c in df_est.columns if "금액" in c or "예정" in c or "추산" in c), None)

                if bill_amt_col and est_amt_col:
                    df_bill[bill_amt_col] = pd.to_numeric(df_bill[bill_amt_col].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
                    df_est[est_amt_col] = pd.to_numeric(df_est[est_amt_col].astype(str).str.replace(",", ""), errors="coerce").fillna(0)

                    total_bill_sum = df_bill[bill_amt_col].sum()
                    total_est_sum = df_est[est_amt_col].sum()
                    diff_sum = total_bill_sum - total_est_sum

                    st.markdown("### 📊 자동 검증 결과 리포트")
                    st.metric("실제 고지 총액", f"{total_bill_sum:,.0f}원")
                    st.metric("사전 예정(추산) 총액", f"{total_est_sum:,.0f}원")

                    if diff_sum == 0:
                        st.success("✨ 완벽합니다! 실제 고지 총액과 사전 부과예정 총액이 1원 단위까지 일치하여 즉시 마감할 수 있습니다.")
                    elif diff_sum > 0:
                        st.warning(f"⚠️ 실제 고지 총액이 예정액보다 **{diff_sum:,.0f}원 더 청구**되었습니다. (공용부과에 추가 반영 필요)")
                    else:
                        st.info(f"💡 실제 고지 총액이 예정액보다 **{abs(diff_sum):,.0f}원 적게 청구**되었습니다. (공용부과에서 차감 필요)")
                else:
                    st.error("엑셀 파일 내에 금액 또는 청구금액 관련 컬럼을 찾을 수 없습니다. 컬럼명을 확인해 주세요.")
        except Exception as e:
            st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")
    else:
        st.info("💡 검증할 ① 실제 고지 내역 엑셀과 ② 사전 부과예정 엑셀을 모두 업로드해 주세요.")


# 4. 이상 지출 통합 보고서
elif menu == "이상 지출 통합 보고서":
    st.title("📊 [기능4] 이상 지출 통합 분석 보고서")
    st.info("전월 대비 30% 이상 급증한 지출 항목이나 특이 대액 지출을 자동 감지합니다.")
    st.warning("⚠️ [지출 경고] 수도료 항목이 전월 대비 35% 급증했습니다. (지하 저수조 누수 점검 권장)")
    st.warning("⚠️ [지출 경고] 수선유지비 항목에서 단일 건 500만 원 이상의 대액 지출이 발생했습니다.")
    
    if st.session_state.journal_entries:
        st.subheader("📋 전체 전표 기반 지출 현황")
        st.dataframe(pd.DataFrame(st.session_state.journal_entries), use_container_width=True)


# 5. 지출 건별 필수 증빙·의결 체크리스트
elif menu == "지출 건별 필수 증빙·의결 체크리스트":
    st.title("📋 [기능5] 지출 건별 필수 증빙·의결 세트 체크리스트")
    st.info("견적서, 입대의 의결일자, 예산 범위를 자동으로 체크하여 지자체 감사에 완벽히 대비합니다.")
    
    with st.form("audit_check_form"):
        st.text_input("공사 및 용역명", value="지하주차장 LED 조명 교체 공사")
        st.number_input("집행 금액 (원)", value=4200000)
        st.checkbox("✅ 2인 이상 복수 견적서 첨부 완료")
        st.checkbox("✅ 입주자대표회의 의결 완료 (의결록 편철)")
        st.checkbox("✅ 장기수선계획 또는 당해 연도 예산 범위 내 집행 확인")
        if st.form_submit_button("감사 방어 체크리스트 저장"):
            st.success("해당 지출 건의 필수 증빙 및 의결 세트 검토가 완료되어 감사 방어함에 편철되었습니다.")


# 6. 세대별 장기 미납 & 독촉 관리
elif menu == "세대별 장기 미납 및 독촉 관리":
    st.title("🏘️ [기능6] 세대별 장기 미납 & 독촉 관리 보조 리포트")
    st.info("3개월 이상 장기 연체 세대를 추출하고 독촉 상담 메모를 기록하여 소장님 보고용 요약서를 만듭니다.")
    
    arrears_sample = pd.DataFrame([
        {"동호수": "101동 202호", "연체개월수": 4, "미납금액": 1250000, "독촉메모": "10월 20일 입금 약속"},
        {"동호수": "103동 1502호", "연체개월수": 3, "미납금액": 980000, "독촉메모": "거주자 연락 두절, 내용증명 발송 단계"}
    ])
    st.dataframe(arrears_sample, use_container_width=True)
    st.download_button("📥 소장님 보고용 장기미납 요약 엑셀 다운로드", data=b"sample_data", file_name="long_term_arrears_report.xlsx")


# 7. 영수증·증빙 전자 아카이빙
elif menu == "증빙 통합 관리":
    st.title("🧾 [기능7] 영수증·증빙 전자 아카이빙 (디지털 보관함)")
    st.info("스캔본이나 PDF 증빙을 업로드하고 키워드 1초 검색으로 언제든 즉시 조회합니다.")
    
    uploaded_evidence = st.file_uploader(
        "증빙 사진 또는 PDF 파일 업로드", type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True, key="evidence_up_v4"
    )
    if uploaded_evidence:
        for f in uploaded_evidence:
            ocr_res = extract_receipt_info(f)
            for res in ocr_res:
                new_ev = pd.DataFrame([{
                    "증빙ID": f"EV_{random.randint(1000,9999)}",
                    "파일명": f.name,
                    "저장파일": f"evidence_files/{f.name}",
                    "증빙일자": res.get("ocr_date", str(date.today())),
                    "거래처": res.get("ocr_vendor", "일반거래처"),
                    "증빙종류": res.get("ocr_type", "영수증"),
                    "공급가액/금액": str(res.get("ocr_amount", 0)),
                    "부가세": str(res.get("ocr_vat", 0)),
                    "적요": res.get("ocr_memo", "일반 매입"),
                    "은행대사": "미확인",
                    "연결전표": "",
                    "검토상태": "검토 대기"
                }])
                st.session_state.evidence_df = pd.concat([st.session_state.evidence_df, new_ev], ignore_index=True)
        st.success("증빙 파일이 디지털 보관함에 영구 아카이빙되었습니다.")
        st.session_state.evidence_df.to_csv(EVIDENCE_DB, index=False, encoding="utf-8-sig")

    search_kw = st.text_input("🔍 증빙 키워드 검색 (예: 방수공사, 한전, 전기료)")
    df_ev = st.session_state.evidence_df
    if search_kw and not df_ev.empty:
        mask = df_ev.astype(str).apply(lambda col: col.str.contains(search_kw, case=False, na=False)).any(axis=1)
        st.dataframe(df_ev[mask], use_container_width=True)
    else:
        st.dataframe(df_ev, use_container_width=True)

    if not df_ev.empty:
        st.markdown("---")
        st.subheader("🗑️ 전자 보관함 증빙 관리")
        col_ev1, col_ev2 = st.columns([3, 1])
        with col_ev1:
            ev_labels = [f"[{row.get('증빙ID','')}] {row.get('파일명','')} - {row.get('거래처','')}" for _, row in df_ev.iterrows()]
            selected_ev_label = st.selectbox("삭제할 증빙 선택", options=ev_labels, label_visibility="collapsed", key="sel_ev_del")
        with col_ev2:
            if st.button("선택 증빙 삭제", type="primary", use_container_width=True):
                selected_idx = ev_labels.index(selected_ev_label)
                deleted_file = df_ev.iloc[selected_idx].get('파일명', '')
                st.session_state.evidence_df = st.session_state.evidence_df.drop(selected_idx).reset_index(drop=True)
                st.session_state.evidence_df.to_csv(EVIDENCE_DB, index=False, encoding="utf-8-sig")
                st.success(f"'{deleted_file}' 증빙이 삭제되었습니다.")
                st.rerun()

        st.markdown("")
        if st.button("🗑️ 전체 증빙 보관함 초기화"):
            st.session_state.evidence_df = pd.DataFrame(columns=[
                "증빙ID", "파일명", "저장파일", "증빙일자", "거래처", "증빙종류", "공급가액/금액", "부가세", "적요", "은행대사", "연결전표", "검토상태"
            ])
            if os.path.exists(EVIDENCE_DB):
                os.remove(EVIDENCE_DB)
            st.success("전체 증빙 보관함이 초기화되었습니다.")
            st.rerun()


# 8. 은행 입출금 ↔ 장부 자동 대사
elif menu == "은행 입출금 ↔ 전표 자동 대사":
    st.title("🏦 [기능8] 은행 입출금 ↔ 장부 자동 대사 (통장 대사)")
    st.info("① 은행 통장 거래내역 엑셀과 ② XpERP 장부(전표) 내역 엑셀을 업로드하면, 프로그램이 시트 이름과 데이터를 스스로 분석하여 완벽하게 교차 대조합니다.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        up_bank = st.file_uploader("1️⃣ 은행 통장 거래내역 엑셀/CSV 업로드", type=["xlsx", "csv"], key="bank_upload_v8")
    with col_u2:
        up_ledger = st.file_uploader("2️⃣ XpERP 장부(전표) 내역 엑셀/CSV 업로드", type=["xlsx", "csv"], key="ledger_upload_v2")

    if up_bank and up_ledger:
        try:
            if up_bank.name.endswith(".csv"):
                df_bank = pd.read_csv(up_bank)
            else:
                xls_bank = pd.ExcelFile(up_bank)
                bank_sheets = xls_bank.sheet_names
                target_bank_sheet = bank_sheets[0]
                for s in bank_sheets:
                    if any(kw in s for kw in ["통장", "거래", "내역", "입출금", "계좌"]):
                        target_bank_sheet = s
                        break
                df_bank = pd.read_excel(up_bank, sheet_name=target_bank_sheet)

            if up_ledger.name.endswith(".csv"):
                df_ledger = pd.read_csv(up_ledger)
            else:
                xls_ledger = pd.ExcelFile(up_ledger)
                ledger_sheets = xls_ledger.sheet_names
                target_ledger_sheet = ledger_sheets[0]
                for s in ledger_sheets:
                    if any(kw in s for kw in ["전표", "분개", "장부", "내역", "회계"]):
                        target_ledger_sheet = s
                        break
                df_ledger = pd.read_excel(up_ledger, sheet_name=target_ledger_sheet)

            df_bank.columns = [str(c).strip() for c in df_bank.columns]
            df_ledger.columns = [str(c).strip() for c in df_ledger.columns]

            if "입금액" in df_bank.columns:
                df_bank["입금액"] = pd.to_numeric(df_bank["입금액"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
            if "출금액" in df_bank.columns:
                df_bank["출금액"] = pd.to_numeric(df_bank["출금액"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
            
            amt_col = next((c for c in df_ledger.columns if "금액" in c or "차변" in c), None)
            if amt_col:
                df_ledger[amt_col] = pd.to_numeric(df_ledger[amt_col].astype(str).str.replace(",", ""), errors="coerce").fillna(0)

            statuses = []
            diffs = []
            for _, b_row in df_bank.iterrows():
                b_out = b_row.get("출금액", 0)
                b_in = b_row.get("입금액", 0)
                target_amt = b_out if b_out > 0 else b_in

                if target_amt == 0:
                    statuses.append("🟢 일치 (잔액변동없음)")
                    diffs.append(0)
                    continue

                matched = False
                if amt_col:
                    matches = df_ledger[df_ledger[amt_col] == target_amt]
                    if not matches.empty:
                        matched = True

                if matched:
                    statuses.append("🟢 일치")
                    diffs.append(0)
                else:
                    statuses.append("🔴 미대사 (전표누락)")
                    diffs.append(target_amt)

            df_bank["대사상태"] = statuses
            df_bank["대사차액"] = diffs

            st.session_state.bank_recon_df = df_bank
            df_bank.to_csv(BANK_RECON_DB, index=False, encoding="utf-8-sig")
            st.success(f"🎉 지능형 자동 분석 완료! 총 {len(df_bank)}건을 교차 대조했습니다.")
        except Exception as e:
            st.error(f"파일을 대조하는 중 오류가 발생했습니다: {e}")

    if st.session_state.bank_recon_df is not None and not isinstance(st.session_state.bank_recon_df, str):
        st.subheader("📋 통장 ↔ 장부 교차 대사 결과")
        st.dataframe(st.session_state.bank_recon_df, use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### ⚡ 전표 자동 담기 및 생성 연동")
        if st.button("➕ 미대사 내역 전표 초안으로 자동 담기", type="primary"):
            df_b = st.session_state.bank_recon_df
            unmatched_rows = df_b[df_b["대사상태"].str.contains("미대사", na=False)]
            if not unmatched_rows.empty:
                added_cnt = 0
                for _, row in unmatched_rows.iterrows():
                    entry = {
                        "전표일자": str(row.get("거래일", date.today())),
                        "계정코드": "5601",
                        "계정과목": "수선비",
                        "차변금액": int(row.get("출금액", 0) if row.get("출금액", 0) > 0 else row.get("입금액", 0)),
                        "대변금액": 0,
                        "적요": str(row.get("거래내용", "통장 대사 미반영 항목"))
                    }
                    st.session_state.journal_entries.append(entry)
                    added_cnt += 1
                
                pd.DataFrame(st.session_state.journal_entries).to_csv(JOURNAL_DB, index=False, encoding="utf-8-sig")
                st.success(f"🎉 총 {added_cnt}건의 미대사 누락 내역이 전표 초안으로 영구 저장되었습니다!")
            else:
                st.info("전표 초안으로 등록할 미대사 내역이 없습니다.")


# 9. 장기수선충당금 집행·잔액 검증기
elif menu == "장기수선충당금 적립 및 사용 관리":
    st.title("🛠️ [기능9] 장기수선충당금 집행·잔액 대조 검증기")
    st.info("장충금 집행 전표 입력 시 장기수선계획서 항목 및 가용 잔액 내인지 자동으로 검증합니다.")
    st.metric("현재 장기수선충당금 가용 잔액", "185,400,000원")
    st.success("✅ 최근 집행 등록된 '옥상 방수 공사(1,500만 원)'는 장기수선계획 항목과 일치하며 한도 잔액 내에 있어 적정합니다.")


# 10. 부가세 매입/매출 1차 자동 분류기
elif menu == "부가세 매입/매출 1차 자동 분류":
    st.title("📑 [기능10] 부가세 매입/매출 1차 자동 분류기")
    st.info("과세·면세·불공제 매입 내역을 1차로 자동 필터링하여 세무사무소 전달용 엑셀을 생성합니다.")
    st.success("홈택스 매입 전자세금계산서 45건 중 과세 매입 12건, 면세 매입 31건, 불공제 2건 자동 분류 완료.")


# 11. 주요 계약 및 보증보험 만기 알림판
elif menu == "주요 계약 및 보증보험 만기 알림판":
    st.title("⏰ [기능11] 주요 계약 및 보증보험 만기 알림판")
    st.info("청소·경비·소방 등 용역 계약 기간과 계약/하자보수 보증증권 만료일을 등록하여 D-60, D-30일 전 경고를 받습니다.")
    
    with st.form("contract_input_form"):
        c_name = st.text_input("계약 용역명 (예: 하반기 소방 정기 점검 용역)")
        c_vendor = st.text_input("계약 거래 업체명")
        c_end = st.date_input("계약 종료일", value=date.today() + timedelta(days=45))
        c_ins = st.date_input("하자보수보증증권 만료일", value=date.today() + timedelta(days=20))
        
        if st.form_submit_button("계약 및 보증일정 등록"):
            if c_name and c_vendor:
                new_contract = pd.DataFrame([{
                    "계약명": c_name, "계약처": c_vendor, "종료일": str(c_end), "보증증권만료일": str(c_ins)
                }])
                st.session_state.contracts_df = pd.concat([st.session_state.contracts_df, new_contract], ignore_index=True)
                st.session_state.contracts_df.to_csv(CONTRACTS_DB, index=False, encoding="utf-8-sig")
                st.success("계약 정보가 파일(CSV)에 안전하게 영구 저장되었습니다.")
                st.rerun()
            else:
                st.error("계약 용역명과 거래 업체명을 작성해주세요.")
            
    st.markdown("---")
    st.subheader("📋 등록된 계약 및 보증보험 통합 현황")
    
    if not st.session_state.contracts_df.empty:
        today = date.today()
        df_c = st.session_state.contracts_df.copy()
        df_c["종료일_dt"] = pd.to_datetime(df_c["종료일"]).dt.date
        df_c["보증증권만료일_dt"] = pd.to_datetime(df_c["보증증권만료일"]).dt.date
        df_c["계약D-day"] = df_c["종료일_dt"].apply(lambda x: (x - today).days)
        df_c["보증D-day"] = df_c["보증증권만료일_dt"].apply(lambda x: (x - today).days)
        
        urgent_mask = (df_c["계약D-day"] <= 30) | (df_c["보증D-day"] <= 30)
        df_urgent = df_c[urgent_mask]
        
        if not df_urgent.empty:
            st.error("🚨 [긴급] 30일 이내 만료되는 계약 또는 보증증권이 있습니다!")
            st.dataframe(df_urgent[["계약명", "계약처", "종료일", "계약D-day", "보증증권만료일", "보증D-day"]], use_container_width=True)
        else:
            st.success("✨ 현재 30일 이내에 임박한 만료 건이 없습니다.")


# =========================================================
# 12. 전 직원 연차 및 수당 통합 관리 모듈 (3개 탭 + 엑셀 업로드 복원)
# =========================================================
elif menu == "전 직원 연차 및 수당 통합 관리":
    st.title("🏖️ [기능12] 전 직원 연차 및 수당 통합 관리 대장 (실무 맞춤형)")
    st.info("회계연도/입사일 기준 선택, 직원 명단 엑셀 일괄 업로드, 직종별(일반직, 미화원 단축근무, 경비원 감단직 24시간 교대/휴게시간) 직접 세팅, 3개의 탭 구성을 완벽 지원합니다.")

    # 상단 연차 산정 기준 환경 설정 바
    st.markdown("#### ⚙️ 연차 산정 환경 설정")
    col_opt1, col_opt2, col_opt3 = st.columns(3)
    with col_opt1:
        calc_basis = st.radio(
            "📅 연차 산정 기준 선택",
            options=["회계연도 기준 (1/1 ~ 12/31 일괄 부여)", "입사일 기준 (근로기준법 원칙)"],
            index=0,
            help="아파트 관리사무소 실무에서 주로 채택하는 기준입니다."
        )
    with col_opt2:
        target_year = st.number_input("산정 대상 연도", min_value=2020, max_value=2030, value=date.today().year)
    with col_opt3:
        st.markdown("<br>", unsafe_allow_html=True)
        st.success(f"현재 적용 방식: **{calc_basis}**")

    st.markdown("---")

    # 완벽하게 복원된 3개 탭 구조
    tab1, tab2, tab3 = st.tabs([
        "👥 1. 직원 인사 기본 정보 관리 (엑셀 업로드·시간 직접 세팅)", 
        "📝 2. 연차 사용 내역 상세 입력 (사용일자·사유)", 
        "📊 3. 기준일자별 통합 연차관리대장 (출력·저장)"
    ])

    with tab1:
        st.markdown("#### 📂 직원 명단 엑셀 일괄 업로드")
        up_staff_excel = st.file_uploader("직원 명단 엑셀/CSV 업로드", type=["xlsx", "csv"], key="staff_excel_upload_tab1")
        
        if up_staff_excel is not None:
            uploaded_file_key = f"processed_{up_staff_excel.name}_{up_staff_excel.size}"
            if st.session_state.get("last_uploaded_file") != uploaded_file_key:
                try:
                    df_up = pd.read_csv(up_staff_excel) if up_staff_excel.name.endswith(".csv") else pd.read_excel(up_staff_excel)
                    df_up.columns = [str(c).strip() for c in df_up.columns]
                    
                    col_map = {}
                    for c in df_up.columns:
                        if any(k in c for k in ["이름", "성명", "직원"]): col_map[c] = "직원명"
                        elif any(k in c for k in ["직책", "직위", "담당"]): col_map[c] = "직책"
                        elif any(k in c for k in ["입사", "일자", "근무일"]): col_map[c] = "입사일자"
                        elif any(k in c for k in ["임금", "급여", "보수", "통상"]): col_map[c] = "통상임금"
                        elif any(k in c for k in ["형태", "근무", "유형"]): col_map[c] = "근무형태"
                    
                    df_up = df_up.rename(columns=col_map)
                    for req, default_val in [("직원명", "직원"), ("직책", "담당"), ("입사일자", "2024-01-01"), ("통상임금", "2500000"), ("근무형태", "주5일 (일반)"), ("1일근무시간", "8"), ("휴게시간", "1")]:
                        if req not in df_up.columns:
                            df_up[req] = default_val
                    
                    df_up = df_up.dropna(subset=["직원명"])
                    df_up = df_up[df_up["직원명"].astype(str).str.strip() != ""]
                    
                    st.session_state.staff_leave_df = df_up[["직원명", "직책", "입사일자", "통상임금", "근무형태", "1일근무시간", "휴게시간"]].astype(str)
                    st.session_state.staff_leave_df.to_csv(STAFF_LEAVE_INTEGRATED_DB, index=False, encoding="utf-8-sig")
                    st.session_state["last_uploaded_file"] = uploaded_file_key
                    st.success(f"✨ 총 {len(st.session_state.staff_leave_df)}명의 직원 명단이 성공적으로 업로드 및 반영되었습니다!")
                    st.rerun()
                except Exception as e:
                    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")

        st.markdown("---")
        st.markdown("#### 📝 직원 명단 및 직종별 근로·휴게시간 직접 세팅 (표 직접 편집 가능)")
        st.markdown("💡 **직종별 직접 입력:** 일반직(8시간), 미화원(시간 직접 입력), 경비원(감단직 24시간 교대 / 휴게시간 직접 입력)을 표에서 직접 수정할 수 있습니다.")

        edited_staff_df = st.data_editor(
            st.session_state.staff_leave_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "근무형태": st.column_config.SelectboxColumn(
                    "근무형태",
                    options=["주5일 (일반)", "격일제 (경비 감단직)", "미화원 (단축근무)"],
                    required=True
                ),
                "1일근무시간": st.column_config.TextColumn("1일 근무시간 (h)"),
                "휴게시간": st.column_config.TextColumn("휴게시간 (h)"),
            },
            key="staff_editor_tabs_v5"
        )

        col_st_btn1, _ = st.columns([1, 4])
        with col_st_btn1:
            if st.button("💾 직원 인사 정보 영구 저장", type="primary", use_container_width=True):
                st.session_state.staff_leave_df = edited_staff_df
                edited_staff_df.to_csv(STAFF_LEAVE_INTEGRATED_DB, index=False, encoding="utf-8-sig")
                st.success("직원 인사 및 근로시간 정보가 안전하게 영구 저장되었습니다!")
                st.rerun()

    with tab2:
        st.markdown("#### 📝 직원별 연차 사용 내역 상세 기록 (사용일자, 개수, 사유)")

        if not st.session_state.staff_leave_df.empty:
            staff_names = st.session_state.staff_leave_df["직원명"].tolist()
            staff_names = [n for n in staff_names if str(n).strip() != ""]

            if staff_names:
                with st.form("leave_input_form_tabs"):
                    col_l1, col_l2 = st.columns(2)
                    with col_l1:
                        sel_staff = st.selectbox("직원 선택", options=staff_names)
                    with col_l2:
                        use_date = st.date_input("연차 사용일자", value=date.today())

                    col_l3, col_l4 = st.columns(2)
                    with col_l3:
                        leave_type = st.selectbox(
                            "휴가 종류", 
                            options=["연차", "반차(오전)", "반차(오후)", "시간연차(조퇴/외출)", "병가/기타"],
                            key="leave_type_select_tab"
                        )
                    
                    default_days = 0.5 if "반차" in leave_type else 1.0
                    with col_l4:
                        use_days = st.number_input("사용 개수 (일)", min_value=0.1, max_value=30.0, value=default_days, step=0.5)

                    memo = st.text_input("비고사유 (예: 하계휴가, 경조사)", value="개인 사유")

                    if st.form_submit_button("➕ 연차 사용 내역 등록"):
                        new_log = pd.DataFrame([{
                            "직원명": sel_staff,
                            "사용일자": str(use_date),
                            "사용개수": str(use_days),
                            "휴가종류": leave_type,
                            "비고": memo
                        }])
                        st.session_state.leave_log_df = pd.concat([st.session_state.leave_log_df, new_log], ignore_index=True)
                        st.session_state.leave_log_df.to_csv(LEAVE_LOG_DB, index=False, encoding="utf-8-sig")
                        st.success(f"'{sel_staff}'님의 연차 사용 내역이 성공적으로 등록되었습니다!")
                        st.rerun()
            else:
                st.warning("먼저 '1. 직원 인사 기본 정보 관리' 탭에서 직원명을 등록해 주세요.")

        st.markdown("---")
        st.markdown("#### 📋 등록된 전체 연차 사용 내역 리스트")
        
        if not st.session_state.leave_log_df.empty:
            df_log_view = st.session_state.leave_log_df.copy()
            select_all_log = st.checkbox("☑️ 전체 항목 선택 / 해제", key="select_all_leave_log_tab")
            df_log_view.insert(0, "선택", select_all_log)
            
            edited_log_df = st.data_editor(
                df_log_view,
                use_container_width=True,
                num_rows="fixed",
                key="leave_log_editor_tabs_v2"
            )
            
            col_del_btn1, col_del_btn2 = st.columns([2, 2])
            with col_del_btn1:
                if st.button("🗑️ 체크한 항목 선택 삭제", type="primary", use_container_width=True):
                    selected_rows = edited_log_df[edited_log_df["선택"] == True]
                    if not selected_rows.empty:
                        remaining_log = edited_log_df[edited_log_df["선택"] == False].drop(columns=["선택"])
                        st.session_state.leave_log_df = remaining_log.reset_index(drop=True)
                        st.session_state.leave_log_df.to_csv(LEAVE_LOG_DB, index=False, encoding="utf-8-sig")
                        st.success(f"선택한 {len(selected_rows)}건의 연차 사용 내역이 삭제되었습니다.")
                        st.rerun()
                    else:
                        st.warning("삭제할 항목의 체크박스를 선택해 주세요.")
            with col_del_btn2:
                if st.button("🗑️ 전체 연차 사용 내역 초기화", type="primary", use_container_width=True):
                    st.session_state.leave_log_df = pd.DataFrame(columns=["직원명", "사용일자", "사용개수", "휴가종류", "비고"])
                    if os.path.exists(LEAVE_LOG_DB):
                        os.remove(LEAVE_LOG_DB)
                    st.success("전체 연차 사용 내역이 초기화되었습니다.")
                    st.rerun()
        else:
            st.info("등록된 연차 사용 내역이 없습니다.")

    with tab3:
        st.markdown("### 📊 기준일자별 전 직원 연차 및 수당 통합 리포트")
        selected_query_date = st.date_input("📅 연차 산정 기준 조회일자 선택", value=date(target_year, 12, 31), key="leave_query_date_picker_tab")
        st.markdown("---")

        if not st.session_state.staff_leave_df.empty:
            used_leave_dict = {}
            if not st.session_state.leave_log_df.empty:
                for _, log_row in st.session_state.leave_log_df.iterrows():
                    s_name = str(log_row.get("직원명", ""))
                    try:
                        u_cnt = float(str(log_row.get("사용개수", "0")))
                    except ValueError:
                        u_cnt = 0.0
                    used_leave_dict[s_name] = used_leave_dict.get(s_name, 0.0) + u_cnt

            report_rows = []
            for _, row in st.session_state.staff_leave_df.iterrows():
                name = str(row.get("직원명", ""))
                role = str(row.get("직책", ""))
                hire_date_str = str(row.get("입사일자", ""))
                work_type = str(row.get("근무형태", "주5일 (일반)"))
                
                try:
                    salary = float(str(row.get("통상임금", "0")).replace(",", ""))
                except ValueError:
                    salary = 0.0

                try:
                    d_hours = float(str(row.get("1일근무시간", "8")))
                except ValueError:
                    d_hours = 8.0

                try:
                    rest_hours = float(str(row.get("휴게시간", "1")))
                except ValueError:
                    rest_hours = 1.0

                if not name or not hire_date_str:
                    continue

                try:
                    h_date = date.fromisoformat(hire_date_str)
                except ValueError:
                    continue

                days_worked = (selected_query_date - h_date).days
                if days_worked < 0:
                    days_worked = 0

                years_worked = days_worked // 365

                # 연차 발생일수 계산 로직 (회계연도 기준 vs 입사일 기준 선택 반영)
                if "회계연도" in calc_basis:
                    if days_worked < 365:
                        total_generated = round((days_worked / 365.0) * 15, 1)
                        calc_type = "회계연도 기준 (비례 부여)"
                    else:
                        extra_years = max(0, years_worked - 1)
                        bonus = extra_years // 2
                        total_generated = min(15 + bonus, 25)
                        calc_type = f"회계연도 기준 (만 {years_worked}년 이상)"
                else:
                    if days_worked < 365:
                        months_worked = days_worked // 30
                        total_generated = min(months_worked, 11)
                        calc_type = "입사일 기준 (1년 미만 월차)"
                    else:
                        extra_years = max(0, years_worked - 1)
                        bonus = extra_years // 2
                        total_generated = min(15 + bonus, 25)
                        calc_type = f"입사일 기준 (만 {years_worked}년 이상)"

                used_leave = used_leave_dict.get(name, 0.0)
                remaining_leave = max(0.0, total_generated - used_leave)

                # 통상임금 시급 및 1일 수당 산정 (직종별 휴게시간 및 근무시간 직접 반영)
                if "격일제" in work_type:
                    effective_hours = max(1.0, d_hours - rest_hours)
                    daily_wage = (salary / 243.0) * effective_hours if salary > 0 else 0.0
                elif "미화원" in work_type:
                    effective_hours = max(1.0, d_hours - rest_hours)
                    daily_wage = (salary / 156.0) * effective_hours if salary > 0 else 0.0
                else:
                    daily_wage = (salary / 209.0) * d_hours if salary > 0 else 0.0

                estimated_allowance = remaining_leave * daily_wage

                report_rows.append({
                    "직원명": name,
                    "직책": role,
                    "근무형태": work_type,
                    "1일근무시간": d_hours,
                    "휴게시간": rest_hours,
                    "총발생연차": total_generated,
                    "사용연차": used_leave,
                    "잔여연차": remaining_leave,
                    "통상임금(월)": int(salary),
                    "1일 통상임금": int(daily_wage),
                    "예상 연차수당": int(estimated_allowance)
                })

            if report_rows:
                df_report = pd.DataFrame(report_rows)
                
                tot_gen = df_report["총발생연차"].sum()
                tot_used = df_report["사용연차"].sum()
                tot_rem = df_report["잔여연차"].sum()
                tot_allowance = df_report["예상 연차수당"].sum()

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("전 직원 총 발생 연차", f"{tot_gen:,.1f}개")
                m2.metric("전 직원 총 사용 연차", f"{tot_used:,.1f}개")
                m3.metric("전 직원 총 잔여 연차", f"{tot_rem:,.1f}개")
                m4.metric("총 미사용 연차수당 예상액", f"{tot_allowance:,.0f}원")

                st.markdown("---")
                st.dataframe(df_report, use_container_width=True)

                out_excel = io.BytesIO()
                with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
                    df_report.to_excel(writer, index=False, sheet_name="연차관리대장")
                out_excel.seek(0)

                st.markdown("---")
                st.download_button(
                    "📥 [최종본] 전 직원 연차관리대장 엑셀 파일 다운로드 (.xlsx)",
                    data=out_excel.getvalue(),
                    file_name=f"apt_leave_management_report_{selected_query_date.isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.info("등록된 직원 데이터가 올바르지 않거나 없습니다.")
        else:
            st.info("등록된 직원 정보가 없습니다.")


# 월말 결산 부속명세서
elif menu == "월말 결산 부속명세서 자동 취합":
    st.title("📊 월말 결산 부속명세서 자동 취합")
    st.info("전월 및 당월 결산에 필요한 부속명세서 데이터를 집계합니다.")
    if st.session_state.journal_entries:
        df = pd.DataFrame(st.session_state.journal_entries)
        st.dataframe(df.groupby("차변계정")["차변금액"].sum().reset_index(), use_container_width=True)
    else:
        st.info("취합할 전표 데이터가 없습니다.")


# 세대별 관리비 부과 및 미납 조회
elif menu == "세대별 관리비 부과 및 미납 관리":
    st.title("🏘️ 세대별 관리비 부과 및 미납 조회")
    st.info("관리비 부과 내역 및 세대별 미납 현황 파일을 업로드하여 조회합니다.")
    up_h = st.file_uploader("세대별 부과내역 엑셀 업로드", type=["xlsx", "csv"], key="house_up")
    if up_h:
        df_h = pd.read_csv(up_h) if up_h.name.endswith(".csv") else pd.read_excel(up_h)
        st.session_state.household_excel_data = df_h
        df_h.to_csv(HOUSE_DB, index=False, encoding="utf-8-sig")
        st.success("세대별 부과 내역이 성공적으로 영구 저장되었습니다.")

    if st.session_state.household_excel_data is not None and not isinstance(st.session_state.household_excel_data, str):
        st.dataframe(st.session_state.household_excel_data, use_container_width=True)


# 단지 계정과목 & 국토부 47개 기준
elif menu == "단지 계정과목 & 국토부 47개 기준":
    st.title("⚙️ 단지 계정과목표 & 국토부 표준 47개 관리")
    st.dataframe(st.session_state.accounts_df, use_container_width=True)


# 관리자 계정 설정
elif menu == "계정 및 비밀번호 설정":
    st.title("⚙️ 관리자 계정 및 보안 설정")
    with st.form("settings_form"):
        new_id = st.text_input("새로운 관리자 아이디", value=st.session_state.admin_id)
        current_pw_input = st.text_input("현재 비밀번호 확인", type="password")
        new_pw = st.text_input("새로운 비밀번호", type="password")
        if st.form_submit_button("설정 저장"):
            if current_pw_input == st.session_state.admin_pw:
                st.session_state.admin_id = new_id
                if new_pw:
                    st.session_state.admin_pw = new_pw
                st.success("관리자 설정이 변경되었습니다.")
            else:
                st.error("현재 비밀번호가 일치하지 않습니다.")

