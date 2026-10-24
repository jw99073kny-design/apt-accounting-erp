import io
import os
import random
import re
import zipfile
import hashlib
import uuid
import tempfile
from datetime import date, timedelta
import cv2
import numpy as np
import openpyxl
import pandas as pd
import streamlit as st
from PIL import Image, ImageEnhance
import pytesseract
import fitz
import whisper

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


@st.cache_resource
def load_whisper_model():
    """음성 인식용 Whisper 모델 로드 (최초 실행 시 자동 다운로드)"""
    return whisper.load_model("base")


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
        return cropped_images
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


def sanitize_file_name(name):
    return re.sub(r'[\\/*?:"<>|]', "", name).replace(" ", "_")[:30]


# 💡 구어체, 잡음, 어색한 용어를 문맥에 맞게 완전히 다듬는 자연어 윤문 엔진
def smart_ai_refine_text(raw_text):
    if not raw_text or not raw_text.strip():
        return ""
    
    term_map = {
        r"생당기수?": "세대수",
        r"생당기": "세대",
        r"장충금": "장기수선충당금",
        r"기어소이기": "입주자 기여수입",
        r"공동기어소이기": "공동기여수입",
        r"비어한": "비영업",
        r"임대수의": "임대수입",
        r"수유조": "수익조정",
        r"발리비게자": "반환비용",
        r"스위칸|스위키": "수익창출",
        r"전패짝기": "전산회계",
        r"2사점|사점": "사업자",
        r"2자심센터": "자재센터",
        r"1th\s*1th": "",
        r"\b\d+\s*th\b": ""
    }
    
    cleaned = raw_text
    for pattern, repl in term_map.items():
        cleaned = re.sub(pattern, repl, cleaned, flags=re.IGNORECASE)
        
    cleaned = re.sub(r'(그렇죠\.?\s*)+', '확인되었습니다. ', cleaned)
    cleaned = re.sub(r'(이거\.?\s*)+', '이는 ', cleaned)
    cleaned = re.sub(r'(됐어요\.?\s*)+', '', cleaned)
    cleaned = re.sub(r'(자,?\s*)+', '', cleaned)
    cleaned = re.sub(r'아시겠어요\??', '', cleaned)
    cleaned = re.sub(r'넘어 가시고요\.?', '', cleaned)
    cleaned = re.sub(r'복사하겠습니다\.?', '', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()

    raw_sentences = [s.strip() for s in re.split(r'[.?!]\s+', cleaned) if len(s.strip()) > 3]
    refined_list = []
    
    for s in raw_sentences:
        s = re.sub(r'벌어 드렸으니까', '발생하였으므로', s)
        s = re.sub(r'순당기 수익으로 번 돈이고', '당기 순수익으로 계상되는 금액이며', s)
        s = re.sub(r'부가가치 세금입니다', '부가가치세 과세 대상 금액입니다', s)
        s = re.sub(r'국세청에 신고하시고', '국세청 관련 세무 신고를 진행해야 하고', s)
        s = re.sub(r'가져갈 수 밖에 없겠죠', '처리할 수밖에 없습니다', s)
        s = re.sub(r'보는 겁니다', '판단됩니다', s)
        s = re.sub(r'말하는 거예요', '의미합니다', s)
        
        if not s.endswith(('다.', '함.', '임.', '음.')):
            s += "."
        refined_list.append(s)

    polished_body = "\n".join([f"  - {line}" for line in refined_list[:15]])
    return f"[AI 실무 윤문 및 재번역 결과물]\n\n{polished_body}"


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
    if os.path.exists(DATA_FILE):
        try:
            st.session_state.trained_patterns = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
        except UnicodeDecodeError:
            st.session_state.trained_patterns = pd.read_csv(DATA_FILE, encoding="cp949")
    else:
        st.session_state.trained_patterns = pd.DataFrame()

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
    if os.path.exists(ACCOUNTS_FILE):
        try:
            st.session_state.accounts_df = pd.read_csv(ACCOUNTS_FILE, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            st.session_state.accounts_df = pd.read_csv(ACCOUNTS_FILE, dtype=str, encoding="cp949")
    else:
        st.session_state.accounts_df = default_coas_df

JOURNAL_DB = "journal_entries_register.csv"
if "journal_entries" not in st.session_state:
    if os.path.exists(JOURNAL_DB):
        try:
            df_j_saved = pd.read_csv(JOURNAL_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            df_j_saved = pd.read_csv(JOURNAL_DB, dtype=str, encoding="cp949").fillna("")
        except Exception:
            df_j_saved = pd.DataFrame()
            
        if not df_j_saved.empty:
            if "차변금액" in df_j_saved.columns:
                df_j_saved["차변금액"] = pd.to_numeric(df_j_saved["차변금액"].str.replace(",", ""), errors="coerce").fillna(0)
            if "대변금액" in df_j_saved.columns:
                df_j_saved["대변금액"] = pd.to_numeric(df_j_saved["대변금액"].str.replace(",", ""), errors="coerce").fillna(0)
            st.session_state.journal_entries = df_j_saved.to_dict("records")
        else:
            st.session_state.journal_entries = []
    else:
        st.session_state.journal_entries = []

BANK_RECON_DB = "bank_reconciliation_cache.csv"
if "bank_recon_df" not in st.session_state:
    if os.path.exists(BANK_RECON_DB):
        try:
            df_b_saved = pd.read_csv(BANK_RECON_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            df_b_saved = pd.read_csv(BANK_RECON_DB, dtype=str, encoding="cp949").fillna("")
        except Exception:
            df_b_saved = None
            
        if df_b_saved is not None and not df_b_saved.empty:
            for col in ["출금액", "입금액", "대사차액"]:
                if col in df_b_saved.columns:
                    df_b_saved[col] = pd.to_numeric(df_b_saved[col].str.replace(",", ""), errors="coerce").fillna(0)
            st.session_state.bank_recon_df = df_b_saved
        else:
            st.session_state.bank_recon_df = None
    else:
        st.session_state.bank_recon_df = None

CONTRACTS_DB = "contracts_register.csv"
if "contracts_df" not in st.session_state:
    if os.path.exists(CONTRACTS_DB):
        try:
            st.session_state.contracts_df = pd.read_csv(CONTRACTS_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            st.session_state.contracts_df = pd.read_csv(CONTRACTS_DB, dtype=str, encoding="cp949").fillna("")
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
            st.session_state.evidence_df = pd.read_csv(EVIDENCE_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            st.session_state.evidence_df = pd.read_csv(EVIDENCE_DB, dtype=str, encoding="cp949").fillna("")
        except Exception:
            st.session_state.evidence_df = pd.DataFrame()
    else:
        st.session_state.evidence_df = pd.DataFrame(
            columns=[
                "증빙ID", "파일명", "저장파일", "증빙일자",
                "거래처", "증빙종류", "공급가액/금액", "부가세",
                "적요", "은행대사", "연결전표", "검토상태", "파일해시"
            ]
        )

MEDIA_TRANSCRIPT_DB = "media_transcript_archive.csv"
if "media_transcripts_df" not in st.session_state:
    if os.path.exists(MEDIA_TRANSCRIPT_DB):
        try:
            st.session_state.media_transcripts_df = pd.read_csv(MEDIA_TRANSCRIPT_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            st.session_state.media_transcripts_df = pd.read_csv(MEDIA_TRANSCRIPT_DB, dtype=str, encoding="cp949").fillna("")
        except Exception:
            st.session_state.media_transcripts_df = pd.DataFrame(columns=["ID", "파일명", "추출일시", "전체텍스트", "윤문결과", "업무보고서"])
    else:
        st.session_state.media_transcripts_df = pd.DataFrame(columns=["ID", "파일명", "추출일시", "전체텍스트", "윤문결과", "업무보고서"])

# 구버전 CSV 호환 체크
if not st.session_state.media_transcripts_df.empty and "윤문결과" not in st.session_state.media_transcripts_df.columns:
    st.session_state.media_transcripts_df["윤문결과"] = st.session_state.media_transcripts_df["업무보고서"]

HOUSE_DB = "household_bill_cache.csv"
if "household_excel_data" not in st.session_state:
    if os.path.exists(HOUSE_DB):
        try:
            st.session_state.household_excel_data = pd.read_csv(HOUSE_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            st.session_state.household_excel_data = pd.read_csv(HOUSE_DB, dtype=str, encoding="cp949").fillna("")
        except Exception:
            st.session_state.household_excel_data = None
    else:
        st.session_state.household_excel_data = None

STAFF_LEAVE_INTEGRATED_DB = "staff_leave_integrated_register_full_tabs.csv"
if "staff_leave_df" not in st.session_state:
    if os.path.exists(STAFF_LEAVE_INTEGRATED_DB):
        try:
            st.session_state.staff_leave_df = pd.read_csv(STAFF_LEAVE_INTEGRATED_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            st.session_state.staff_leave_df = pd.read_csv(STAFF_LEAVE_INTEGRATED_DB, dtype=str, encoding="cp949").fillna("")
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
            st.session_state.leave_log_df = pd.read_csv(LEAVE_LOG_DB, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            st.session_state.leave_log_df = pd.read_csv(LEAVE_LOG_DB, dtype=str, encoding="cp949").fillna("")
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
    if st.sidebar.button("• 회의/전화 녹음·영상 텍스트 추출·보고서", use_container_width=True):
        st.session_state.current_menu = "회의 및 전화 녹음/영상 텍스트 추출·보고서"
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
    c1.metric("등록된 전표 초안 [실제 데이터]", f"{len(st.session_state.journal_entries)}건")
    c2.metric("전자 보관 증빙 [실제 데이터]", f"{len(st.session_state.evidence_df)}건")
    c3.metric("등록된 용역 계약 [실제 데이터]", f"{len(st.session_state.contracts_df)}건")


# 1. 스마트 분개 패턴 추천
elif menu == "XpERP 전년도 장부 및 분개 패턴 학습":
    st.title("🧠 [기능1] 스마트 분개 패턴 추천 (장부 학습)")
    st.info("전년도 XpERP 분개장을 업로드하여 단지별 계정과목 및 적요 패턴을 학습시킵니다.")
    
    up_learn = st.file_uploader("XpERP 전년도 분개장 엑셀 업로드", type=["xlsx", "csv"], key="learn_up")
    if up_learn:
        try:
            df_learn = pd.read_csv(up_learn, encoding="utf-8-sig") if up_learn.name.endswith(".csv") else pd.read_excel(up_learn)
        except UnicodeDecodeError:
            df_learn = pd.read_csv(up_learn, encoding="cp949")
        st.session_state.trained_patterns = df_learn
        df_learn.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
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
        
        st.markdown("#### 📋 전표 초안 목록 관리 [실제 사용자 입력 데이터]")
        
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
            if up_bill_file.name.endswith(".csv"):
                try:
                    df_bill = pd.read_csv(up_bill_file, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    df_bill = pd.read_csv(up_bill_file, encoding="cp949")
            else:
                df_bill = pd.read_excel(up_bill_file)

            if up_est_file.name.endswith(".csv"):
                try:
                    df_est = pd.read_csv(up_est_file, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    df_est = pd.read_csv(up_est_file, encoding="cp949")
            else:
                df_est = pd.read_excel(up_est_file)

            df_bill.columns = [str(c).strip() for c in df_bill.columns]
            df_est.columns = [str(c).strip() for c in df_est.columns]

            st.markdown("---")
            st.subheader("📋 업로드된 원본 데이터 미리보기 [실제 업로드 데이터]")
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
    
    st.info("💡 **[🔍 데모/예시 데이터 안내]** 아래 지출 경고 및 분석 리포트는 시스템 기능 시연을 위한 **하드코딩된 예시 지표**입니다.")
    st.warning("⚠️ [지출 경고 - 예시] 수도료 항목이 전월 대비 35% 급증했습니다. (지하 저수조 누수 점검 권장)")
    st.warning("⚠️ [지출 경고 - 예시] 수선유지비 항목에서 단일 건 500만 원 이상의 대액 지출이 발생했습니다.")
    
    if st.session_state.journal_entries:
        st.subheader("📋 전체 전표 기반 지출 현황 [실제 사용자 입력 데이터]")
        st.dataframe(pd.DataFrame(st.session_state.journal_entries), use_container_width=True)
    else:
        st.info("💡 등록된 실제 전표 초안 데이터가 없습니다. 전표를 등록하시면 실제 데이터 기반 분석이 함께 표시됩니다.")


# 5. 지출 건별 필수 증빙·의결 체크리스트
elif menu == "지출 건별 필수 증빙·의결 체크리스트":
    st.title("📋 [기능5] 지출 건별 필수 증빙·의결 세트 체크리스트")
    st.info("견적서, 입대의 의결일자, 예산 범위를 자동으로 체크하여 지자체 감사에 완벽히 대비합니다.")
    
    st.info("💡 **[🔍 데모/예시 데이터 안내]** 아래 입력폼 및 체크리스트 예시는 시스템 기능 시연을 위한 **데모 템플릿**입니다.")
    with st.form("audit_check_form"):
        st.text_input("공사 및 용역명 [데모 예시]", value="지하주차장 LED 조명 교체 공사")
        st.number_input("집행 금액 (원) [데모 예시]", value=4200000)
        st.checkbox("✅ 2인 이상 복수 견적서 첨부 완료")
        st.checkbox("✅ 입주자대표회의 의결 완료 (의결록 편철)")
        st.checkbox("✅ 장기수선계획 또는 당해 연도 예산 범위 내 집행 확인")
        if st.form_submit_button("감사 방어 체크리스트 저장"):
            st.success("해당 지출 건의 필수 증빙 및 의결 세트 검토가 완료되어 감사 방어함에 편철되었습니다.")


# 6. 세대별 장기 미납 & 독촉 관리
elif menu == "세대별 장기 미납 및 독촉 관리":
    st.title("🏘️ [기능6] 세대별 장기 미납 & 독촉 관리 보조 리포트")
    st.info("3개월 이상 장기 연체 세대를 추출하고 독촉 상담 메모를 기록하여 소장님 보고용 요약서를 만듭니다.")
    
    st.info("💡 **[🔍 데모/예시 데이터 안내]** 아래 미납 세대 리스트는 기능 설명을 위한 **하드코딩된 예시 데이터**입니다.")
    arrears_sample = pd.DataFrame([
        {"동호수": "101동 202호", "연체개월수": 4, "미납금액": 1250000, "독촉메모": "10월 20일 입금 약속 [예시]"},
        {"동호수": "103동 1502호", "연체개월수": 3, "미납금액": 980000, "독촉메모": "거주자 연락 두절, 내용증명 발송 단계 [예시]"}
    ])
    st.dataframe(arrears_sample, use_container_width=True)
    st.download_button("📥 소장님 보고용 장기미납 요약 엑셀 다운로드", data=b"sample_data", file_name="long_term_arrears_report.xlsx")


# 7. 영수
