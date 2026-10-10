import io
import os
import random
import re
import zipfile
import hashlib
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
from PIL import Image
import pytesseract
import fitz

st.set_page_config(
    page_title="XpERP 연동 AI 회계 시스템", page_icon="🏢", layout="wide"
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


# ---------------------------------------------------------
# 영수증·명세서 OCR 기능
# ---------------------------------------------------------
def extract_receipt_info(uploaded_file):
    """이미지/PDF에서 글자를 읽고 날짜·금액 후보를 추출합니다."""
    result = {
        "ocr_text": "",
        "ocr_date": "",
        "ocr_amount": 0,
        "ocr_status": "",
    }

    try:
        file_bytes = uploaded_file.getvalue()
        images = []

        if uploaded_file.name.lower().endswith(".pdf"):
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
            try:
                for page in pdf_doc:
                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(2, 2), alpha=False
                    )
                    images.append(
                        Image.open(
                            io.BytesIO(pix.tobytes("png"))
                        ).convert("RGB")
                    )
            finally:
                pdf_doc.close()
        else:
            images.append(
                Image.open(io.BytesIO(file_bytes)).convert("RGB")
            )

        text_parts = []
        for image_page in images:
            text_parts.append(
                pytesseract.image_to_string(
                    image_page, lang="kor+eng"
                )
            )

        text = "\n".join(text_parts).strip()
        result["ocr_text"] = text

        if not text:
            result["ocr_status"] = (
                "읽힌 글자가 없습니다. 날짜와 금액을 직접 입력해 주세요."
            )
            return result

        date_match = re.search(
            r"(20\d{2})\s*[./년-]\s*(\d{1,2})"
            r"\s*[./월-]\s*(\d{1,2})\s*일?",
            text,
        )

        if date_match:
            try:
                year, month, day = map(int, date_match.groups())
                result["ocr_date"] = date(
                    year, month, day
                ).isoformat()
            except ValueError:
                pass

        lines = [
            line.strip()
            for line in text.splitlines()
            if line.strip()
        ]

        priority_lines = [
            line for line in lines
            if re.search(
                r"합계|총액|결제금액|청구금액|받을금액|"
                r"총\s*금액|공급대가",
                line,
            )
        ]

        amount_candidates = []
        search_lines = priority_lines if priority_lines else lines

        for line in search_lines:
            numbers = re.findall(
                r"(?<!\d)[\d,]{2,}(?!\d)", line
            )
            for number in numbers:
                try:
                    amount = int(number.replace(",", ""))
                    if amount > 0:
                        amount_candidates.append(amount)
                except ValueError:
                    continue

        if amount_candidates:
            result["ocr_amount"] = amount_candidates[-1]
            result["ocr_status"] = (
                "인식 완료 — 날짜와 금액을 원본 증빙과 대조해 주세요."
            )
        else:
            result["ocr_status"] = (
                "글자는 읽었지만 금액을 찾지 못했습니다. "
                "직접 입력해 주세요."
            )

    except Exception as exc:
        result["ocr_status"] = (
            "OCR 실행 실패. Tesseract 프로그램과 한국어 언어 "
            f"데이터 설치를 확인하세요. 상세 오류: {exc}"
        )

    return result


# ---------------------------------------------------------
# 로그인 및 세션 관리
# ---------------------------------------------------------
if "admin_id" not in st.session_state:
    st.session_state.admin_id = "KNY"

if "admin_pw" not in st.session_state:
    st.session_state.admin_pw = "skdud12"

if "admin_hint" not in st.session_state:
    st.session_state.admin_hint = (
        "관리자 아이디 KNY / 초기 비밀번호는 지정하신 skdud12 입니다."
    )

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "current_menu" not in st.session_state:
    st.session_state.current_menu = "🏠 시스템 홈 (대시보드)"

DATA_FILE = "trained_patterns.csv"

if "trained_patterns" not in st.session_state:
    if os.path.exists(DATA_FILE):
        st.session_state.trained_patterns = pd.read_csv(DATA_FILE)
    else:
        st.session_state.trained_patterns = pd.DataFrame()

ACCOUNTS_FILE = "accounts_table.csv"

gov_data_full = [
    (1, "5101", "급여"),
    (2, "5102", "제수당"),
    (3, "5103", "상여금"),
    (4, "5104", "퇴직금"),
    (5, "5105", "산재보험료"),
    (6, "5106", "고용보험료"),
    (7, "5107", "국민연금"),
    (8, "5108", "국민건강보험료"),
    (9, "5109", "복리후생비(식대등)"),
    (10, "5110", "일반사무용품비"),
    (11, "5111", "도서인쇄비"),
    (12, "5112", "여비교통비"),
    (13, "5113", "공과금중 전기료"),
    (14, "5114", "통신료"),
    (15, "5115", "우편료"),
    (16, "5116", "제세공과금 등"),
    (17, "5117", "피복비"),
    (18, "5118", "교육훈련비"),
    (19, "5119", "연료비"),
    (20, "5120", "차량수리비"),
    (21, "5121", "차량보험료"),
    (22, "5122", "기타차량유지비"),
    (23, "5123", "관리용품구입비"),
    (24, "5124", "전문가자문비 등"),
    (25, "5125", "잡비"),
    (26, "5201", "청소비"),
    (27, "5301", "경비비"),
    (28, "5401", "소독비"),
    (29, "5501", "승강기유지비"),
    (30, "5502", "지능형홈네트워크설비유지비"),
    (31, "5601", "수선비"),
    (32, "5602", "시설유지비"),
    (33, "5701", "안전점검비"),
    (34, "5801", "재해예방비"),
    (35, "5901", "위탁관리수수료"),
    (36, "6101", "난방비"),
    (37, "6102", "급탕비"),
    (38, "6103", "가스사용료"),
    (39, "6104", "전기료"),
    (40, "6105", "수도료"),
    (41, "6106", "정화조오물수수료"),
    (42, "6107", "생활폐기물수수료"),
    (43, "7101", "입주자대표회의 운영비"),
    (44, "7102", "건물보험료"),
    (45, "7103", "선거관리위원회 운영비"),
    (46, "8101", "장기수선충당금"),
    (47, "4101", "잡수입"),
]

default_coas = [
    {"계정코드": code, "계정명": name}
    for num, code, name in gov_data_full
]

default_coas.extend([
    {"계정코드": "101", "계정명": "현금"},
    {"계정코드": "103", "계정명": "보통예금"},
    {"계정코드": "251", "계정명": "미지급금"},
    {"계정코드": "117", "계정명": "부가가치세대급금"},
    {"계정코드": "253", "계정명": "예수금(4대보험/소득세)"},
    {"계정코드": "310", "계정명": "장기수선충당금부채"},
])

default_coas_df = pd.DataFrame(default_coas)

if "accounts_df" not in st.session_state:
    if os.path.exists(ACCOUNTS_FILE):
        st.session_state.accounts_df = pd.read_csv(
            ACCOUNTS_FILE, dtype=str
        )
    else:
        st.session_state.accounts_df = default_coas_df

if "household_excel_data" not in st.session_state:
    st.session_state.household_excel_data = None

if "bank_reconciliation_df" not in st.session_state:
    st.session_state.bank_reconciliation_df = None

EVIDENCE_DIR = "evidence_files"
EVIDENCE_DB = "evidence_register.csv"
os.makedirs(EVIDENCE_DIR, exist_ok=True)

if "evidence_df" not in st.session_state:
    if os.path.exists(EVIDENCE_DB):
        try:
            st.session_state.evidence_df = (
                pd.read_csv(EVIDENCE_DB, dtype=str).fillna("")
            )
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

for category_number in range(1, 5):
    key = f"open_cat_{category_number}"
    if key not in st.session_state:
        st.session_state[key] = False


def login_screen():
    st.title("🏢 XpERP 연동 아파트 AI 회계 관리 시스템")
    st.info(
        "🔒 보안 접속: 등록된 관리자 계정과 비밀번호를 정확히 입력해 주세요."
    )

    tab_login, tab_find = st.tabs(["로그인", "🔑 비밀번호 힌트"])

    with tab_login:
        with st.form("login_form"):
            input_id = st.text_input("아이디", value="")
            user_pw = st.text_input("비밀번호", type="password")
            submit = st.form_submit_button("로그인", type="primary")

            if submit:
                if (
                    input_id == st.session_state.admin_id
                    and user_pw == st.session_state.admin_pw
                ):
                    st.session_state.authenticated = True
                    st.success("로그인 성공!")
                    st.rerun()
                else:
                    st.error("아이디 또는 비밀번호가 올바르지 않습니다.")

    with tab_find:
        st.info("💡 관리자 계정 정보 안내")
        st.write("🔐 **비밀번호 힌트**: 초기 설정된 계정 정보를 확인하세요.")


if not st.session_state.authenticated:
    login_screen()
    st.stop()

if "journal_entries" not in st.session_state:
    st.session_state.journal_entries = []

st.sidebar.title("🏢 XpERP 관리 시스템")
st.sidebar.text(f"관리자: {st.session_state.admin_id}")

if st.sidebar.button("🔒 로그아웃", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()

st.sidebar.markdown("---")

if st.sidebar.button("🏠 시스템 홈 (대시보드)", use_container_width=True):
    st.session_state.current_menu = "🏠 시스템 홈 (대시보드)"

st.sidebar.markdown("---")
st.sidebar.markdown("📁 **작업 카테고리**")


if st.sidebar.button("📁 회계·전표 검증", use_container_width=True):
    st.session_state.open_cat_1 = not st.session_state.open_cat_1

if st.session_state.open_cat_1:
    st.sidebar.markdown(
        "<div style='margin-left: 15px;'>", unsafe_allow_html=True
    )
    if st.sidebar.button(
        "• XpERP 장부 및 분개 패턴 학습", use_container_width=True
    ):
        st.session_state.current_menu = "XpERP 전년도 장부 및 분개 패턴 학습"
    if st.sidebar.button(
        "• 명세서 다중 일괄 검증 & 전표", use_container_width=True
    ):
        st.session_state.current_menu = "명세서 다중 일괄 검증 & 전표 발행"
    if st.sidebar.button(
        "• 단지 계정과목 & 국토부 47개", use_container_width=True
    ):
        st.session_state.current_menu = "단지 계정과목 & 국토부 47개 기준"
    if st.sidebar.button("• 이상 지출 통합 보고서", use_container_width=True):
        st.session_state.current_menu = "이상 지출 통합 보고서"
    if st.sidebar.button(
        "• 은행 입출금 ↔ 전표 자동 대사", use_container_width=True
    ):
        st.session_state.current_menu = "은행 입출금 ↔ 전표 자동 대사"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)


if st.sidebar.button("📁 결산·세무 관리", use_container_width=True):
    st.session_state.open_cat_2 = not st.session_state.open_cat_2

if st.session_state.open_cat_2:
    st.sidebar.markdown(
        "<div style='margin-left: 15px;'>", unsafe_allow_html=True
    )
    if st.sidebar.button(
        "• 부가가치세(VAT) 자동 안분 분개", use_container_width=True
    ):
        st.session_state.current_menu = "부가가치세(VAT) 자동 안분 분개"
    if st.sidebar.button(
        "• 월말 결산 부속명세서 자동 취합", use_container_width=True
    ):
        st.session_state.current_menu = "월말 결산 부속명세서 자동 취합"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)
if st.sidebar.button("📁 실무 통합 관리", use_container_width=True):
    st.session_state.open_cat_3 = not st.session_state.open_cat_3

if st.session_state.open_cat_3:
    st.sidebar.markdown(
        "<div style='margin-left: 15px;'>", unsafe_allow_html=True
    )
    if st.sidebar.button(
        "• 세대별 관리비 부과 및 미납", use_container_width=True
    ):
        st.session_state.current_menu = "세대별 관리비 부과 및 미납 관리"
    if st.sidebar.button(
        "• 장기수선충당금 적립 및 사용", use_container_width=True
    ):
        st.session_state.current_menu = "장기수선충당금 적립 및 사용 관리"
    if st.sidebar.button(
        "• 직원 급여 및 4대보험 자동 분개", use_container_width=True
    ):
        st.session_state.current_menu = "직원 급여 및 4대보험 자동 분개"
    if st.sidebar.button(
        "• 잡수입(재활용/주차) 통합 관리", use_container_width=True
    ):
        st.session_state.current_menu = "잡수입(재활용/주차) 통합 관리"
    if st.sidebar.button("• 영수증·증빙 통합 관리", use_container_width=True):
        st.session_state.current_menu = "증빙 통합 관리"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)


if st.sidebar.button("⚙️ 시스템 설정", use_container_width=True):
    st.session_state.open_cat_4 = not st.session_state.open_cat_4

if st.session_state.open_cat_4:
    st.sidebar.markdown(
        "<div style='margin-left: 15px;'>", unsafe_allow_html=True
    )
    if st.sidebar.button("• 계정 및 비밀번호 설정", use_container_width=True):
        st.session_state.current_menu = "계정 및 비밀번호 설정"
    st.sidebar.markdown("</div>", unsafe_allow_html=True)

menu = st.session_state.current_menu


# ---------------------------------------------------------
# 홈 화면
# ---------------------------------------------------------
if menu == "🏠 시스템 홈 (대시보드)":
    st.title("🏢 XpERP 연동 아파트 AI 회계 관리 시스템")
    st.markdown("---")
    st.success(
        "✨ 보안 로그인되었습니다. 좌측 사이드바에서 업무 메뉴를 선택하세요."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.info(
            "🧠 **AI 회계 패턴 학습**\n"
            "전년도 XpERP 장부를 등록해 단지별 회계 패턴을 관리합니다."
        )
    with col2:
        st.info(
            "📄 **명세서 OCR 검토**\n"
            "영수증을 읽고 날짜·금액 후보를 확인합니다."
        )
    with col3:
        st.info(
            "🧾 **실무 관리**\n"
            "증빙, 세대별 관리비, 장충금 및 잡수입을 관리합니다."
        )

    st.markdown("---")
    st.subheader("📊 현재 시스템 현황 요약")
    c1, c2, c3 = st.columns(3)
    c1.metric(
        "등록된 계정과목",
        f"{len(st.session_state.accounts_df)}개 (국토부 표준 포함)",
    )
    c2.metric("등록된 전표", f"{len(st.session_state.journal_entries)}건")
    c3.metric("학습된 패턴", f"{len(st.session_state.trained_patterns)}건")


# ---------------------------------------------------------
# 증빙 통합 관리
# ---------------------------------------------------------
elif menu == "증빙 통합 관리":
    st.title("🧾 영수증·세금계산서 증빙 통합 관리")
    st.info(
        "파일을 등록하면 OCR이 날짜와 금액 후보를 읽습니다. "
        "인식 결과는 원본 증빙과 대조한 뒤 등록하세요."
    )

    uploaded_evidence = st.file_uploader(
        "영수증 사진 또는 증빙 PDF 업로드 (여러 파일 가능)",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="evidence_upload_v1",
    )

    if uploaded_evidence:
        st.caption("날짜·금액 후보를 확인하고 필요하면 직접 수정하세요.")
        editor_rows = []

        for file_index, f in enumerate(uploaded_evidence):
            ocr = extract_receipt_info(f)

            with st.expander(f"🔎 OCR 결과 확인: {f.name}"):
                st.write(ocr["ocr_status"])
                if ocr["ocr_text"]:
                    st.text_area(
                        "인식된 원문",
                        value=ocr["ocr_text"],
                        height=140,
                        key=(
                            "ocr_text_"
                            + hashlib.sha256(
                                (f.name + str(file_index)).encode()
                            ).hexdigest()[:12]
                        ),
                    )

            editor_rows.append({
                "파일명": f.name,
                "증빙일자": ocr["ocr_date"] or str(date.today()),
                "거래처": "",
                "증빙종류": "영수증",
                "금액": int(ocr["ocr_amount"]),
                "부가세": 0,
                "적요": "",
            })

        edited_evidence = st.data_editor(
            pd.DataFrame(editor_rows),
            use_container_width=True,
            num_rows="fixed",
            key="evidence_metadata_editor",
            column_config={
                "파일명": st.column_config.TextColumn(
                    "파일명", disabled=True
                ),
                "증빙일자": st.column_config.TextColumn(
                    "증빙일자 (YYYY-MM-DD)"
                ),
                "증빙종류": st.column_config.SelectboxColumn(
                    "증빙종류",
                    options=[
                        "영수증", "세금계산서", "카드전표",
                        "거래명세서", "기타",
                    ],
                    required=True,
                ),
                "금액": st.column_config.NumberColumn(
                    "금액(원)", min_value=0, step=1000
                ),
                "부가세": st.column_config.NumberColumn(
                    "부가세(원)", min_value=0, step=100
                ),
            },
        )

        if st.button(
            "📥 증빙 등록 및 파일 저장",
            type="primary",
            key="save_evidence_v1",
        ):
            existing_df = st.session_state.evidence_df.copy()
            new_rows = []
            duplicate_names = []

            existing_hashes = set(
                existing_df["파일해시"].astype(str)
            ) if "파일해시" in existing_df.columns else set()

            for i, f in enumerate(uploaded_evidence):
                meta = edited_evidence.iloc[i].to_dict()
                file_bytes = f.getvalue()
                digest = hashlib.sha256(file_bytes).hexdigest()

                if digest in existing_hashes:
                    duplicate_names.append(f.name)
                    continue

                safe_name = re.sub(
                    r"[^A-Za-z0-9가-힣._-]", "_", f.name
                )
                stored_name = f"{digest[:12]}_{safe_name}"
                stored_path = os.path.join(EVIDENCE_DIR, stored_name)

                with open(stored_path, "wb") as out_file:
                    out_file.write(file_bytes)

                new_rows.append({
                    "증빙ID": digest[:16],
                    "파일명": f.name,
                    "저장파일": stored_path,
                    "파일해시": digest,
                    "증빙일자": str(meta.get("증빙일자", "")).strip(),
                    "거래처": str(meta.get("거래처", "")).strip(),
                    "증빙종류": str(meta.get("증빙종류", "영수증")).strip(),
                    "공급가액/금액": str(meta.get("금액", 0)),
                    "부가세": str(meta.get("부가세", 0)),
                    "적요": str(meta.get("적요", "")).strip(),
                    "은행대사": "미확인",
                    "연결전표": "",
                    "검토상태": "검토 대기",
                })

            if new_rows:
                updated = pd.concat(
                    [existing_df, pd.DataFrame(new_rows)],
                    ignore_index=True,
                )
                st.session_state.evidence_df = updated.fillna("")
                st.session_state.evidence_df.to_csv(
                    EVIDENCE_DB, index=False, encoding="utf-8-sig"
                )
                st.success(f"{len(new_rows)}개 증빙을 등록했습니다.")

            if duplicate_names:
                st.warning(
                    "동일한 파일은 중복 저장하지 않았습니다: "
                    + ", ".join(duplicate_names)
                )

            st.rerun()

    st.markdown("---")
    st.subheader("📋 등록 증빙 목록")
    evidence_df = st.session_state.evidence_df.copy()

    if evidence_df.empty:
        st.info("아직 등록된 증빙이 없습니다.")
    else:
        visible_cols = [
            col for col in [
                "증빙ID", "파일명", "증빙일자", "거래처",
                "증빙종류", "공급가액/금액", "부가세", "적요",
                "은행대사", "연결전표", "검토상태",
            ]
            if col in evidence_df.columns
        ]

        edited_register = st.data_editor(
            evidence_df[visible_cols],
            use_container_width=True,
            num_rows="fixed",
            key="evidence_register_editor",
            column_config={
                "은행대사": st.column_config.SelectboxColumn(
                    "은행대사",
                    options=[
                        "미확인", "연결 후보", "일치 확인",
                        "불일치", "해당 없음",
                    ],
                ),
                "검토상태": st.column_config.SelectboxColumn(
                    "검토상태",
                    options=["검토 대기", "검토 중", "승인", "보류", "반려"],
                ),
            },
        )

        if st.button(
            "💾 증빙 목록 변경사항 저장",
            key="save_evidence_edits",
        ):
            for col in edited_register.columns:
                st.session_state.evidence_df.loc[
                    edited_register.index, col
                ] = edited_register[col]

            st.session_state.evidence_df.to_csv(
                EVIDENCE_DB, index=False, encoding="utf-8-sig"
            )
            st.success("증빙 목록을 저장했습니다.")
            st.rerun()

        st.markdown("---")
        st.subheader("🔎 은행 거래 연결 후보 찾기")
        bank_df = st.session_state.bank_reconciliation_df

        if bank_df is None or bank_df.empty:
            st.info(
                "먼저 '은행 입출금 ↔ 전표 자동 대사' 메뉴에서 "
                "통장 내역을 처리하세요."
            )
        else:
            evidence_options = evidence_df["증빙ID"].astype(str).tolist()
            chosen_id = st.selectbox(
                "확인할 증빙 선택",
                evidence_options,
                key="evidence_match_id",
            )
            selected_evidence = evidence_df[
                evidence_df["증빙ID"].astype(str) == chosen_id
            ].iloc[0]

            try:
                evidence_amount = float(
                    str(
                        selected_evidence.get("공급가액/금액", "0")
                    ).replace(",", "")
                )
            except Exception:
                evidence_amount = 0.0

            bank_work = bank_df.copy()
            bank_work["_금액"] = (
                pd.to_numeric(
                    bank_work.get("입금액", 0), errors="coerce"
                ).fillna(0)
                + pd.to_numeric(
                    bank_work.get("출금액", 0), errors="coerce"
                ).fillna(0)
            )
            bank_work["_차액"] = (
                bank_work["_금액"] - evidence_amount
            ).abs()
            candidates = bank_work[bank_work["_차액"] == 0].copy()

            if candidates.empty:
                st.warning(
                    "금액이 정확히 같은 은행 거래가 없습니다. "
                    "날짜 차이·분할 결제 여부를 확인하세요."
                )
            else:
                st.caption(
                    "금액만으로 찾은 후보입니다. 자동으로 연결하지 않습니다."
                )
                st.dataframe(
                    candidates.drop(
                        columns=["_금액", "_차액"], errors="ignore"
                    ),
                    use_container_width=True,
                )
        st.markdown("---")
        st.subheader("🔗 전표 연결")
        journal_list = st.session_state.journal_entries

        if not journal_list:
            st.info("연결할 전표가 없습니다.")
        else:
            journal_labels = ["연결 안 함"] + [
                f"{i + 1}. {j.get('적요', j.get('파일명', '전표'))} / "
                f"{j.get('차변금액', j.get('대변금액', ''))}원"
                for i, j in enumerate(journal_list)
            ]

            link_id = st.selectbox(
                "연결할 증빙 선택",
                evidence_df["증빙ID"].astype(str).tolist(),
                key="evidence_link_id",
            )
            link_choice = st.selectbox(
                "연결할 전표 선택",
                journal_labels,
                key="evidence_journal_choice",
            )

            if st.button(
                "🔗 선택 증빙과 전표 연결 저장",
                key="link_evidence_journal",
            ):
                row_idx = st.session_state.evidence_df.index[
                    st.session_state.evidence_df["증빙ID"].astype(str)
                    == link_id
                ]

                if len(row_idx):
                    st.session_state.evidence_df.loc[
                        row_idx[0], "연결전표"
                    ] = (
                        "" if link_choice == "연결 안 함"
                        else link_choice
                    )
                    st.session_state.evidence_df.to_csv(
                        EVIDENCE_DB,
                        index=False,
                        encoding="utf-8-sig",
                    )
                    st.success("전표 연결 정보를 저장했습니다.")
                    st.rerun()

        st.markdown("---")
        st.subheader("📥 증빙 등록대장 엑셀 다운로드")
        export_df = st.session_state.evidence_df.drop(
            columns=["저장파일", "파일해시"], errors="ignore"
        )

        evidence_buffer = io.BytesIO()
        with pd.ExcelWriter(
            evidence_buffer, engine="openpyxl"
        ) as writer:
            export_df.to_excel(
                writer, index=False, sheet_name="증빙등록대장"
            )

        evidence_buffer.seek(0)
        st.download_button(
            "증빙 등록대장 다운로드",
            data=evidence_buffer.getvalue(),
            file_name="증빙_등록대장.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )


# ---------------------------------------------------------
# XpERP 전년도 장부 및 분개 패턴 학습
# ---------------------------------------------------------
elif menu == "XpERP 전년도 장부 및 분개 패턴 학습":
    st.title("🧠 XpERP 전년도 회계 장부 및 분개 패턴 AI 학습")
    st.info(
        "XpERP에서 내려받은 전년도 분개장이나 장부 파일을 올려 "
        "단지별 계정 처리 패턴을 관리합니다."
    )

    sample_learn_df = pd.DataFrame([
        {
            "적요내용": "한국전력 전기요금 납부",
            "차변계정코드": "5113",
            "차변계정명": "공과금중 전기료",
        },
        {
            "적요내용": "승강기 유지보수비 월정액",
            "차변계정코드": "5501",
            "차변계정명": "승강기유지비",
        },
        {
            "적요내용": "재활용품 매각 수수료 입금",
            "차변계정코드": "4101",
            "차변계정명": "잡수입",
        },
    ])

    buf_learn = io.BytesIO()
    sample_learn_df.to_excel(buf_learn, index=False)
    buf_learn.seek(0)

    st.download_button(
        "📥 [실무 양식] XpERP 학습용 분개장 샘플 엑셀 다운로드",
        data=buf_learn.getvalue(),
        file_name="sample_xperp_training_data.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    st.markdown("---")
    up_learn = st.file_uploader(
        "XpERP 전년도 분개장/장부 엑셀 파일 업로드",
        type=["xlsx", "csv"],
        key="up_learn_file",
    )

    if up_learn:
        df_learn_up = (
            pd.read_csv(up_learn)
            if up_learn.name.endswith(".csv")
            else pd.read_excel(up_learn)
        )

        st.success(
            f"총 {len(df_learn_up)}건의 XpERP 장부 데이터를 읽었습니다."
        )
        st.dataframe(df_learn_up, use_container_width=True)

        if st.button("🚀 단지 맞춤형 패턴 학습 실행", type="primary"):
            st.session_state.trained_patterns = df_learn_up
            df_learn_up.to_csv(DATA_FILE, index=False)
            st.success("학습 패턴을 저장했습니다.")

    if not st.session_state.trained_patterns.empty:
        st.markdown("---")
        col_title, col_reset = st.columns([3, 1])

        with col_title:
            st.subheader("✅ 현재 학습 패턴")
        with col_reset:
            if st.button("🗑️ 학습 패턴 초기화"):
                if os.path.exists(DATA_FILE):
                    os.remove(DATA_FILE)
                st.session_state.trained_patterns = pd.DataFrame()
                st.success("학습 데이터를 초기화했습니다.")
                st.rerun()

        st.dataframe(
            st.session_state.trained_patterns,
            use_container_width=True,
        )


# ---------------------------------------------------------
# 단지 계정과목표 관리
# ---------------------------------------------------------
elif menu == "단지 계정과목 & 국토부 47개 기준":
    st.title("⚙️ 단지별 계정과목표 관리")
    st.info("단지별 계정과목표를 업로드하거나 개별 계정을 관리하세요.")

    col_up, col_down = st.columns(2)

    with col_up:
        up_acc = st.file_uploader(
            "단지 계정과목표 엑셀/CSV 업로드 (덮어쓰기)",
            type=["xlsx", "csv"],
            key="up_accounts",
        )

        if up_acc:
            df_acc_up = (
                pd.read_csv(up_acc, dtype=str)
                if up_acc.name.endswith(".csv")
                else pd.read_excel(up_acc, dtype=str)
            )

            if (
                "계정코드" in df_acc_up.columns
                and "계정명" in df_acc_up.columns
            ):
                st.session_state.accounts_df = df_acc_up[
                    ["계정코드", "계정명"]
                ]
                st.session_state.accounts_df.to_csv(
                    ACCOUNTS_FILE, index=False
                )
                st.success("계정과목표를 저장했습니다.")
                st.rerun()
            else:
                st.error(
                    "파일에 '계정코드'와 '계정명' 열이 필요합니다."
                )

    with col_down:
        acc_buf = io.BytesIO()
        st.session_state.accounts_df.to_excel(acc_buf, index=False)
        acc_buf.seek(0)

        st.download_button(
            "📥 현재 계정과목표 엑셀 다운로드",
            data=acc_buf.getvalue(),
            file_name="단지_계정과목표.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    st.markdown("---")

    with st.form("add_account_form"):
        st.subheader("➕ 신규 계정과목 추가")
        c1, c2, c3 = st.columns([2, 3, 2])

        with c1:
            new_code = st.text_input("계정코드 (예: 5126)")
        with c2:
            new_name = st.text_input("계정명 (예: 단지 특화 수수료)")
        with c3:
            st.write("###")
            add_submitted = st.form_submit_button("계정 추가")

        if add_submitted:
            if new_code and new_name:
                if new_code in st.session_state.accounts_df["계정코드"].values:
                    st.error(f"이미 존재하는 계정코드({new_code})입니다.")
                else:
                    new_row = pd.DataFrame([
                        {
                            "계정코드": str(new_code),
                            "계정명": str(new_name),
                        }
                    ])
                    st.session_state.accounts_df = pd.concat(
                        [
                            st.session_state.accounts_df,
                            new_row,
                        ],
                        ignore_index=True,
                    )
                    st.session_state.accounts_df.to_csv(
                        ACCOUNTS_FILE, index=False
                    )
                    st.success("계정과목을 추가했습니다.")
                    st.rerun()
            else:
                st.error("계정코드와 계정명을 모두 입력하세요.")

    st.markdown("---")

    with st.form("del_account_form"):
        st.subheader("🗑️ 계정과목 삭제")
        account_options = st.session_state.accounts_df["계정코드"].tolist()

        code_to_del = st.selectbox(
            "삭제할 계정 선택",
            options=account_options,
            format_func=lambda x: (
                f"[{x}] "
                f"{st.session_state.accounts_df.loc[st.session_state.accounts_df['계정코드'] == x, '계정명'].values[0]}"
            ),
        )
        del_submitted = st.form_submit_button("선택 계정 삭제")

        if del_submitted and code_to_del:
            st.session_state.accounts_df = st.session_state.accounts_df[
                st.session_state.accounts_df["계정코드"] != code_to_del
            ]
            st.session_state.accounts_df.to_csv(
                ACCOUNTS_FILE, index=False
            )
            st.success("계정을 삭제했습니다.")
            st.rerun()

    st.markdown("---")
    st.subheader("📋 현재 등록된 계정과목")
    st.dataframe(
        st.session_state.accounts_df,
        use_container_width=True,
    )
# ---------------------------------------------------------
# 명세서 OCR 검토 및 전표 초안 생성
# ---------------------------------------------------------
elif menu == "명세서 다중 일괄 검증 & 전표 발행":
    st.title("📄 영수증 / 명세서 OCR 검토 및 전표 초안 생성")
    st.warning(
        "OCR 결과는 자동 인식 후보입니다. 원본 증빙과 대조한 뒤 "
        "담당자가 확인한 건만 전표 초안으로 추가하세요."
    )

    uploaded_invoices = st.file_uploader(
        "영수증 또는 명세서 파일 여러 개 업로드",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="invoice_ocr_upload_v2",
    )

    if uploaded_invoices:
        invoice_rows = []

        for file_index, f in enumerate(uploaded_invoices):
            ocr = extract_receipt_info(f)

            with st.expander(f"🔎 OCR 결과: {f.name}"):
                st.write(ocr["ocr_status"])
                if ocr["ocr_text"]:
                    st.text_area(
                        "인식된 원문",
                        value=ocr["ocr_text"],
                        height=120,
                        key=(
                            "invoice_ocr_text_"
                            + hashlib.sha256(
                                (f.name + str(file_index)).encode()
                            ).hexdigest()[:12]
                        ),
                    )

            invoice_rows.append({
                "파일명": f.name,
                "증빙일자": ocr["ocr_date"] or str(date.today()),
                "적요": "",
                "차변계정": "[5125] 잡비",
                "차변금액": int(ocr["ocr_amount"]),
                "대변계정": "[251] 미지급금",
                "대변금액": int(ocr["ocr_amount"]),
                "검토": False,
            })

        st.info(
            "날짜·적요·계정·금액을 확인하세요. 실제 증빙과 대조한 행에만 체크하세요."
        )

        edited_invoices = st.data_editor(
            pd.DataFrame(invoice_rows),
            use_container_width=True,
            num_rows="fixed",
            key="invoice_ocr_review_editor",
            column_config={
                "파일명": st.column_config.TextColumn(
                    "파일명", disabled=True
                ),
                "검토": st.column_config.CheckboxColumn(
                    "담당자 확인",
                    help="원본 증빙과 대조한 뒤 체크하세요.",
                ),
                "차변금액": st.column_config.NumberColumn(
                    "차변금액(원)", min_value=0, step=1000
                ),
                "대변금액": st.column_config.NumberColumn(
                    "대변금액(원)", min_value=0, step=1000
                ),
            },
        )

        if st.button(
            "✅ 담당자 확인 건만 전표 초안에 추가",
            type="primary",
            key="save_reviewed_invoices",
        ):
            reviewed_rows = edited_invoices[
                edited_invoices["검토"] == True
            ]

            if reviewed_rows.empty:
                st.warning(
                    "원본 증빙과 대조한 행의 '담당자 확인'을 체크하세요."
                )
            else:
                added = 0

                for _, row in reviewed_rows.iterrows():
                    debit_amount = int(row["차변금액"] or 0)
                    credit_amount = int(row["대변금액"] or 0)

                    if debit_amount <= 0 or credit_amount <= 0:
                        st.warning(
                            f"{row['파일명']}: 금액이 0이므로 추가하지 않았습니다."
                        )
                        continue

                    if debit_amount != credit_amount:
                        st.warning(
                            f"{row['파일명']}: 차변과 대변 금액이 달라 "
                            "추가하지 않았습니다."
                        )
                        continue

                    st.session_state.journal_entries.append({
                        "파일명": row["파일명"],
                        "증빙일자": row["증빙일자"],
                        "적요": sanitize_text(str(row["적요"])),
                        "차변계정": row["차변계정"],
                        "차변금액": debit_amount,
                        "대변계정": row["대변계정"],
                        "대변금액": credit_amount,
                        "검토상태": "담당자 확인 완료 - 전표 초안",
                    })
                    added += 1

                if added:
                    st.success(
                        f"전표 초안 {added}건을 추가했습니다. "
                        "계정과목은 다시 확인하세요."
                    )
                    st.rerun()
                else:
                    st.warning("추가할 전표가 없습니다. 금액을 확인하세요.")

    if st.session_state.journal_entries:
        st.subheader("📋 전표 목록")
        df_j = pd.DataFrame(st.session_state.journal_entries)
        st.dataframe(df_j, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_j.to_excel(writer, index=False, sheet_name="전표목록")

        output.seek(0)
        wb = openpyxl.load_workbook(output)
        ws = wb.active

        for col in ws.columns:
            max_len = 0
            col_letter = openpyxl.utils.get_column_letter(
                col[0].column
            )
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(
                max_len + 4, 12
            )

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.download_button(
            "📥 XpERP 연동 전표 장부 엑셀 다운로드",
            data=final_output.getvalue(),
            file_name="erp_journal_entries.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
        )


# ---------------------------------------------------------
# 이상 지출 통합 보고서
# ---------------------------------------------------------
elif menu == "이상 지출 통합 보고서":
    st.title("📊 이상 지출 통합 보고서")

    if st.session_state.journal_entries:
        st.dataframe(
            pd.DataFrame(st.session_state.journal_entries),
            use_container_width=True,
        )
    else:
        st.info("등록된 전표가 없습니다.")


# ---------------------------------------------------------
# 은행 입출금 ↔ 전표 자동 대사
# ---------------------------------------------------------
elif menu == "은행 입출금 ↔ 전표 자동 대사":
    st.title("🏦 은행 입출금 ↔ 전표 자동 대사")
    st.info(
        "통장 거래내역을 읽고 분개 초안을 제안합니다. "
        "추천 계정은 참고용이며 담당자 검토가 필요합니다."
    )

    sample_bank_df = pd.DataFrame([
        {
            "거래일자": "2026-06-01",
            "입금": 0,
            "출금": 3450000,
            "거래내용": "한국전력 전기요금 자동이체",
            "잔액": 45000000,
        },
        {
            "거래일자": "2026-06-03",
            "입금": 500000,
            "출금": 0,
            "거래내용": "입주자 잡수입 무통장입금",
            "잔액": 45500000,
        },
        {
            "거래일자": "2026-06-05",
            "입금": 0,
            "출금": 120000,
            "거래내용": "미확인 출금 내역",
            "잔액": 45380000,
        },
    ])

    buf_bank = io.BytesIO()
    sample_bank_df.to_excel(buf_bank, index=False)
    buf_bank.seek(0)

    st.download_button(
        "📥 은행 통장 거래내역 샘플 엑셀 다운로드",
        data=buf_bank.getvalue(),
        file_name="sample_bank_statement.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    st.markdown("---")
    st.subheader("📁 통장 거래내역 업로드")
    up_bank_file = st.file_uploader(
        "은행 통장 내역 엑셀(XLSX) 또는 CSV",
        type=["xlsx", "csv"],
        key="bank_step1_v2",
    )

    if up_bank_file:
        try:
            if up_bank_file.name.lower().endswith(".csv"):
                df_raw = pd.read_csv(up_bank_file)
                selected_sheet = "기본시트(CSV)"
            else:
                xls_file = pd.ExcelFile(up_bank_file)
                sheet_names = xls_file.sheet_names
                selected_sheet = st.selectbox(
                    "대사에 사용할 시트 선택",
                    sheet_names,
                )
                header_row = st.number_input(
                    "실제 헤더 행 번호 (0부터 시작)",
                    min_value=0,
                    value=0,
                    step=1,
                )
                df_raw = pd.read_excel(
                    up_bank_file,
                    sheet_name=selected_sheet,
                    header=header_row,
                )

            if len(df_raw.columns) == 0:
                st.error("파일에 읽을 수 있는 열이 없습니다.")
            else:
                st.write(f"선택한 시트: {selected_sheet}")
                st.dataframe(df_raw.head(), use_container_width=True)
                cols = list(df_raw.columns)

                col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)

                with col_m1:
                    c_date = st.selectbox(
                        "날짜 열", cols, key="m_date"
                    )
                with col_m2:
                    c_deposit = st.selectbox(
                        "입금액 열", cols, key="m_dep",
                        index=min(1, len(cols) - 1),
                    )
                with col_m3:
                    c_withdraw = st.selectbox(
                        "출금액 열", cols, key="m_wit",
                        index=min(2, len(cols) - 1),
                    )
                with col_m4:
                    c_desc = st.selectbox(
                        "거래내용 열", cols, key="m_desc",
                        index=min(3, len(cols) - 1),
                    )
                with col_m5:
                    c_bal = st.selectbox(
                        "잔액 열", cols, key="m_bal",
                        index=min(4, len(cols) - 1),
                    )

                if st.button(
                    "🚀 통장 검증 및 분개 초안 생성",
                    type="primary",
                ):
                    processed_rows = []

                    for _, row in df_raw.iterrows():
                        date_val = str(row[c_date]).strip()
                        desc_val = str(row[c_desc]).strip()

                        try:
                            dep_val = float(row[c_deposit]) if pd.notnull(
                                row[c_deposit]
                            ) else 0.0
                        except (ValueError, TypeError):
                            dep_val = 0.0

                        try:
                            wit_val = float(row[c_withdraw]) if pd.notnull(
                                row[c_withdraw]
                            ) else 0.0
                        except (ValueError, TypeError):
                            wit_val = 0.0

                        try:
                            bal_val = float(row[c_bal]) if pd.notnull(
                                row[c_bal]
                            ) else 0.0
                        except (ValueError, TypeError):
                            bal_val = 0.0

                        if (
                            not date_val
                            or date_val.lower() == "nan"
                            or (dep_val == 0 and wit_val == 0)
                        ):
                            continue

                        match_status = "미대사"
                        if "미확인" in desc_val:
                            match_status = "확인 필요"
                        elif any(
                            word in desc_val
                            for word in ["전기", "입금", "자동이체"]
                        ):
                            match_status = "일치 후보"

                        if "전기" in desc_val:
                            rec_account = "[5113] 공과금중 전기료"
                        elif dep_val > 0:
                            rec_account = "[4101] 잡수입"
                        else:
                            rec_account = "[5125] 잡비"

                        rec_amount = dep_val if dep_val > 0 else wit_val

                        processed_rows.append({
                            "선택": match_status == "일치 후보",
                            "거래일자": date_val,
                            "거래내용": sanitize_text(desc_val),
                            "입금액": dep_val,
                            "출금액": wit_val,
                            "잔액": bal_val,
                            "대사상태": match_status,
                            "추천계정": rec_account,
                            "분개금액": rec_amount,
                            "승인상태": "대기중",
                        })

                    st.session_state.bank_reconciliation_df = pd.DataFrame(
                        processed_rows
                    )
                    st.success("통장 거래 검토 자료를 만들었습니다.")
                    st.rerun()

        except Exception as exc:
            st.error(f"파일 처리 중 오류가 발생했습니다: {exc}")

    if (
        st.session_state.bank_reconciliation_df is not None
        and not st.session_state.bank_reconciliation_df.empty
    ):
        st.markdown("---")
        st.subheader("📊 거래 대사 현황 요약")

        df_rec = st.session_state.bank_reconciliation_df
        total_cnt = len(df_rec)
        done_cnt = len(df_rec[df_rec["승인상태"] == "승인완료"])
        check_cnt = len(df_rec[df_rec["대사상태"] == "확인 필요"])
        unmatched_cnt = len(df_rec[df_rec["대사상태"] == "미대사"])

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("전체 유효 거래", f"{total_cnt}건")
        k2.metric("대사 완료", f"{done_cnt}건")
        k3.metric("확인 필요", f"{check_cnt}건")
        k4.metric("미대사", f"{unmatched_cnt}건")

        st.subheader("📋 담당자 검토 및 승인")
        edited_df = st.data_editor(
            df_rec,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "선택": st.column_config.CheckboxColumn(
                    "엑셀 출력 대상 선택"
                ),
                "승인상태": st.column_config.SelectboxColumn(
                    "검토 상태",
                    options=["대기중", "승인완료", "보류", "반려"],
                    required=True,
                ),
            },
            key="bank_review_editor",
        )
        st.session_state.bank_reconciliation_df = edited_df

        c_exp1, c_exp2 = st.columns(2)

        with c_exp1:
            if st.button(
                "✅ 승인된 거래를 전표 초안에 추가",
                type="primary",
            ):
                approved_rows = edited_df[
                    (edited_df["승인상태"] == "승인완료")
                    & (edited_df["선택"] == True)
                ]

                if approved_rows.empty:
                    st.warning(
                        "먼저 내역을 확인하고 '선택'과 '승인완료'를 지정하세요."
                    )
                else:
                    for _, r in approved_rows.iterrows():
                        amount = float(r["분개금액"])
                        if amount <= 0:
                            continue

                        st.session_state.journal_entries.append({
                            "파일명": "은행대사_자동반영",
                            "적요": r["거래내용"],
                            "차변계정": r["추천계정"],
                            "차변금액": amount,
                            "대변계정": "[103] 보통예금",
                            "대변금액": amount,
                            "검토상태": "담당자 승인 - 전표 초안",
                        })

                    st.success(
                        f"{len(approved_rows)}건을 전표 초안에 추가했습니다."
                    )
                    st.rerun()

        with c_exp2:
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                edited_df.to_excel(
                    writer, index=False, sheet_name="은행대사보고서"
                )
            out_buf.seek(0)

            st.download_button(
                "📥 은행 대사 보고서 엑셀 다운로드",
                data=out_buf.getvalue(),
                file_name="monthly_bank_reconciliation_report.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )


# ---------------------------------------------------------
# 월말 결산 부속명세서
# ---------------------------------------------------------
elif menu == "월말 결산 부속명세서 자동 취합":
    st.title("📊 월말 결산 부속명세서 자동 생성")

    if st.session_state.journal_entries:
        df = pd.DataFrame(st.session_state.journal_entries)
        if "차변계정" in df.columns and "차변금액" in df.columns:
            st.dataframe(
                df.groupby("차변계정")["차변금액"]
                .sum()
                .reset_index(),
                use_container_width=True,
            )
        else:
            st.info("집계할 차변 계정 정보가 없습니다.")
    else:
        st.info("취합할 전표가 없습니다.")


# ---------------------------------------------------------
# 세대별 관리비 부과 및 미납 관리
# ---------------------------------------------------------
elif menu == "세대별 관리비 부과 및 미납 관리":
    st.title("🏘️ 세대별 관리비 부과 및 미납 관리")
    st.info(
        "세대별 부과·수납·미납 내역이 담긴 엑셀 파일을 업로드하면 "
        "시트별로 분리하여 조회할 수 있습니다."
    )

    sample_household_df = pd.DataFrame([
        {
            "동호수": "101동 101호",
            "당월부과액": 250000,
            "수납액": 250000,
            "미납액": 0,
            "납부상태": "완납",
        },
        {
            "동호수": "101동 102호",
            "당월부과액": 310000,
            "수납액": 0,
            "미납액": 310000,
            "납부상태": "미납",
        },
        {
            "동호수": "102동 501호",
            "당월부과액": 280000,
            "수납액": 280000,
            "미납액": 0,
            "납부상태": "완납",
        },
    ])

    buf_household = io.BytesIO()
    sample_household_df.to_excel(buf_household, index=False)
    buf_household.seek(0)

    st.download_button(
        "📥 세대별 부과내역 샘플 엑셀 다운로드",
        data=buf_household.getvalue(),
        file_name="sample_household_billing.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    up_household = st.file_uploader(
        "멀티 시트 엑셀 파일 업로드",
        type=["xlsx", "csv"],
        key="up_household_multi",
    )

    if up_household:
        try:
            if up_household.name.lower().endswith(".csv"):
                df_single = pd.read_csv(up_household)
                st.session_state.household_excel_data = {
                    "기본시트": df_single
                }
            else:
                xls = pd.ExcelFile(up_household)
                st.session_state.household_excel_data = pd.read_excel(
                    xls, sheet_name=None
                )

            st.success("세대별 관리비 파일을 읽었습니다.")
        except Exception as exc:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {exc}")

    if st.session_state.household_excel_data:
        sheet_names = list(st.session_state.household_excel_data.keys())
        tabs = st.tabs([f"📑 {name}" for name in sheet_names])

        for idx, tab in enumerate(tabs):
            sheet_name = sheet_names[idx]
            current_df = st.session_state.household_excel_data[sheet_name]

            with tab:
                st.write(
                    f"**시트명: {sheet_name}** (총 {len(current_df)}행)"
                )
                search_query = st.text_input(
                    "시트 내 검색", key=f"search_{sheet_name}"
                )

                if search_query:
                    mask = current_df.astype(str).apply(
                        lambda col: col.str.contains(
                            search_query, case=False, na=False
                        )
                    ).any(axis=1)
                    st.dataframe(
                        current_df[mask], use_container_width=True
                    )
                else:
                    st.dataframe(
                        current_df, use_container_width=True
                    )


# ---------------------------------------------------------
# 장기수선충당금 적립 및 사용 관리
# ---------------------------------------------------------
elif menu == "장기수선충당금 적립 및 사용 관리":
    st.title("🛠️ 장기수선충당금 적립 및 사용 관리")

    sample_repair_df = pd.DataFrame([
        {
            "거래일자": "2026-05-12",
            "구분": "공사집행",
            "내용": "옥상 방수 공사",
            "금액": 15000000,
        },
        {
            "거래일자": "2026-05-25",
            "구분": "월적립",
            "내용": "5월분 장기수선충당금 적립",
            "금액": 35000000,
        },
    ])

    buf_repair = io.BytesIO()
    sample_repair_df.to_excel(buf_repair, index=False)
    buf_repair.seek(0)

    st.download_button(
        "📥 장기수선충당금 내역 샘플 엑셀 다운로드",
        data=buf_repair.getvalue(),
        file_name="sample_longterm_repair.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    up_repair = st.file_uploader(
        "장기수선충당금 내역 엑셀 업로드",
        type=["xlsx", "csv"],
        key="up_repair",
    )

    if up_repair:
        df_repair = (
            pd.read_csv(up_repair)
            if up_repair.name.lower().endswith(".csv")
            else pd.read_excel(up_repair)
        )
        st.success(f"총 {len(df_repair)}건을 읽었습니다.")
        st.dataframe(df_repair, use_container_width=True)


# ---------------------------------------------------------
# 직원 급여 및 4대보험
# ---------------------------------------------------------
elif menu == "직원 급여 및 4대보험 자동 분개":
    st.title("👥 관리사무소 직원 급여 및 4대보험")

    with st.form("salary_form"):
        emp_name = st.text_input("직원 성명", value="")
        base_pay = st.number_input(
            "기본급 (원)", min_value=0, value=3500000, step=50000
        )
        insurances = st.number_input(
            "4대보험 회사부담금 합계 (원)",
            min_value=0,
            value=350000,
            step=10000,
        )

        if st.form_submit_button("급여 전표 초안 계산"):
            st.write(f"직원: {sanitize_text(emp_name)}")
            st.write(f"기본급: {base_pay:,.0f}원")
            st.write(f"회사부담 4대보험: {insurances:,.0f}원")
            st.warning(
                "급여·공제액을 실제 급여대장과 대조한 뒤 회계 처리하세요."
            )


# ---------------------------------------------------------
# 잡수입 통합 관리
# ---------------------------------------------------------
elif menu == "잡수입(재활용/주차) 통합 관리":
    st.title("💰 잡수입 통합 관리")

    sample_income_df = pd.DataFrame([
        {
            "발생일자": "2026-05-05",
            "항목": "알뜰시장 임대료",
            "금액": 1200000,
            "귀속처": "당월 관리비 차감",
        },
        {
            "발생일자": "2026-05-10",
            "항목": "재활용품 매각 대금",
            "금액": 450000,
            "귀속처": "잡수입 통장 적립",
        },
    ])

    buf_income = io.BytesIO()
    sample_income_df.to_excel(buf_income, index=False)
    buf_income.seek(0)

    st.download_button(
        "📥 잡수입 내역 샘플 엑셀 다운로드",
        data=buf_income.getvalue(),
        file_name="sample_misc_income.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    up_income = st.file_uploader(
        "잡수입 내역 엑셀/CSV 업로드",
        type=["xlsx", "csv"],
        key="up_income",
    )

    if up_income:
        df_income = (
            pd.read_csv(up_income)
            if up_income.name.lower().endswith(".csv")
            else pd.read_excel(up_income)
        )
        st.success(f"총 {len(df_income)}건을 읽었습니다.")
        st.dataframe(df_income, use_container_width=True)


# ---------------------------------------------------------
# 관리자 계정 설정
# ---------------------------------------------------------
elif menu == "계정 및 비밀번호 설정":
    st.title("⚙️ 관리자 계정 및 보안 설정")

    with st.form("settings_form"):
        new_id = st.text_input(
            "새로운 관리자 아이디",
            value=st.session_state.admin_id,
        )
        current_pw_input = st.text_input(
            "현재 비밀번호 확인", type="password"
        )
        new_pw = st.text_input(
            "새로운 비밀번호", type="password"
        )

        if st.form_submit_button("설정 저장"):
            if current_pw_input == st.session_state.admin_pw:
                st.session_state.admin_id = new_id
                if new_pw:
                    st.session_state.admin_pw = new_pw
                st.success(
                    "변경 완료. 이 설정은 현재 세션에 적용됩니다."
                )
            else:
                st.error("현재 비밀번호가 일치하지 않습니다.")

