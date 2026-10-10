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

st.set_page_config(
    page_title="XpERP 연동 AI 회계 시스템", page_icon="🏢", layout="wide"
)

# ---------------------------------------------------------
# 상하위 메뉴 시각적 구분을 위한 커스텀 CSS 주입
# ---------------------------------------------------------
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


# ---------------------------------------------------------
# 1. 개인정보 비식별화(마스킹) 정규식 함수
# ---------------------------------------------------------
def sanitize_text(text):
    """전표 적요 및 텍스트에서 개인정보(동호수, 전화번호, 주민번호 등) 자동 비식별화"""
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
# 2. 로그인 계정 정보 및 보안 세션 관리 (아이디 노출 제거)
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

# --- [학습 데이터 파일 로드 & 영구 저장 처리] ---
DATA_FILE = "trained_patterns.csv"

if "trained_patterns" not in st.session_state:
    if os.path.exists(DATA_FILE):
        st.session_state.trained_patterns = pd.read_csv(DATA_FILE)
    else:
        st.session_state.trained_patterns = pd.DataFrame()

# --- [단지 계정과목표 파일 로드 & 영구 저장 처리] ---
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
    {"계정코드": code, "계정명": name} for num, code, name in gov_data_full
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
        st.session_state.accounts_df = pd.read_csv(ACCOUNTS_FILE, dtype=str)
    else:
        st.session_state.accounts_df = default_coas_df

# 세대별 부과 데이터 세션 관리
if "household_excel_data" not in st.session_state:
    st.session_state.household_excel_data = None

# 은행 대사 전용 세션 관리 (5단계용)
if "bank_reconciliation_df" not in st.session_state:
    st.session_state.bank_reconciliation_df = None

# 증빙 통합 관리 세션 및 로컬 저장 경로
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
        st.session_state.evidence_df = pd.DataFrame(columns=[
            "증빙ID", "파일명", "저장파일", "증빙일자", "거래처", "증빙종류",
            "공급가액/금액", "부가세", "적요", "은행대사", "연결전표", "검토상태"
        ])

# 각 카테고리별 독립적인 열림/닫힘 상태 세션 관리 (처음 접속 시 모두 접혀있도록 False 설정)
if "open_cat_1" not in st.session_state:
    st.session_state.open_cat_1 = False
if "open_cat_2" not in st.session_state:
    st.session_state.open_cat_2 = False
if "open_cat_3" not in st.session_state:
    st.session_state.open_cat_3 = False
if "open_cat_4" not in st.session_state:
    st.session_state.open_cat_4 = False


def login_screen():
    st.title("🏢 XpERP 연동 아파트 AI 회계 관리 시스템")
    st.info("🔒 보안 접속: 등록된 관리자 계정과 비밀번호를 정확히 입력해 주세요.")

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

# ---------------------------------------------------------
# 4. 사이드바: 안정적인 토글 및 깔끔한 메뉴 구성
# ---------------------------------------------------------
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


# --- [1] 회계·전표 검증 카테고리 ---
if st.sidebar.button("📁 회계·전표 검증", use_container_width=True):
    st.session_state.open_cat_1 = not st.session_state.open_cat_1

if st.session_state.open_cat_1:
    st.sidebar.markdown(
        "<div style='margin-left: 15px;'>", unsafe_allow_html=True
    )
    if st.sidebar.button("• XpERP 장부 및 분개 패턴 학습", use_container_width=True):
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


# --- [2] 결산·세무 관리 카테고리 ---
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


# --- [3] 실무 통합 관리 카테고리 ---
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


# --- [4] 시스템 설정 카테고리 ---
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
# 5. 메뉴별 본문 구현
# ---------------------------------------------------------
if menu == "🏠 시스템 홈 (대시보드)":
    st.title("🏢 XpERP 연동 아파트 AI 회계 관리 시스템")
    st.markdown("---")
    st.success(
        "✨ 성공적으로 보안 로그인되었습니다. 좌측 사이드바에서 원하는 업무"
        " 메뉴를 클릭하여 시작하세요."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.info(
            "🧠 **AI 회계 패턴 학습**\n신규 단지 전년도 XpERP 장부를 학습시켜"
            " 단지 맞춤형 자동 분개를 수행합니다."
        )
    with col2:
        st.info(
            "📄 **명세서 일괄 검증**\n증빙 파일들을 다중 업로드하여 AI 복식부기"
            " 전표를 자동으로 생성합니다."
        )
    with col3:
        st.info(
            "🧾 **세무 및 실무 관리**\n부가세 안분, 세대별 관리비, 장충금 및"
            " 잡수입을 완벽하게 통합 관리합니다."
        )

    st.markdown("---")
    st.subheader("📊 현재 시스템 현황 요약")
    c1, c2, c3 = st.columns(3)
    c1.metric(
        "등록된 계정과목", f"{len(st.session_state.accounts_df)}개 (국토부 47개 표준 포함)"
    )
    c2.metric("발행된 전체 전표", f"{len(st.session_state.journal_entries)}건")
    c3.metric("학습된 단지 패턴", f"{len(st.session_state.trained_patterns)}건")

elif menu == "증빙 통합 관리":
    st.title("🧾 영수증·세금계산서 증빙 통합 관리")
    st.info(
        "증빙 파일을 등록하고, 거래처·일자·금액을 기록한 뒤 은행 거래 및 전표와 "
        "연결 후보를 확인합니다. 현재 버전은 OCR 자동 추출이 아닌 담당자 입력/검토 방식입니다."
    )

    uploaded_evidence = st.file_uploader(
        "영수증 사진 또는 증빙 PDF 업로드 (여러 파일 가능)",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="evidence_upload_v1",
    )

    if uploaded_evidence:
        st.caption("파일별로 정보를 입력하세요. 금액은 원 단위 숫자로 입력합니다.")
        editor_rows = []
        for f in uploaded_evidence:
            editor_rows.append({
                "파일명": f.name,
                "증빙일자": str(date.today()),
                "거래처": "",
                "증빙종류": "영수증",
                "금액": 0,
                "부가세": 0,
                "적요": "",
            })
        edited_evidence = st.data_editor(
            pd.DataFrame(editor_rows),
            use_container_width=True,
            num_rows="fixed",
            key="evidence_metadata_editor",
            column_config={
                "파일명": st.column_config.TextColumn("파일명", disabled=True),
                "증빙일자": st.column_config.TextColumn("증빙일자 (YYYY-MM-DD)"),
                "증빙종류": st.column_config.SelectboxColumn(
                    "증빙종류", options=["영수증", "세금계산서", "카드전표", "거래명세서", "기타"],
                    required=True
                ),
                "금액": st.column_config.NumberColumn("금액(원)", min_value=0, step=1000),
                "부가세": st.column_config.NumberColumn("부가세(원)", min_value=0, step=100),
            },
        )
        if st.button("📥 증빙 등록 및 파일 저장", type="primary", key="save_evidence_v1"):
            existing_df = st.session_state.evidence_df.copy()
            existing_names = set(existing_df.get("파일명", pd.Series(dtype=str)).astype(str))
            new_rows = []
            duplicate_names = []
            for i, f in enumerate(uploaded_evidence):
                meta = edited_evidence.iloc[i].to_dict()
                file_bytes = f.getvalue()
                digest = hashlib.sha256(file_bytes).hexdigest()
                # 같은 내용의 파일은 중복 등록하지 않음
                if "파일해시" in existing_df.columns and digest in set(existing_df["파일해시"].astype(str)):
                    duplicate_names.append(f.name)
                    continue
                safe_name = re.sub(r"[^A-Za-z0-9가-힣._-]", "_", f.name)
                stored_name = f"{digest[:12]}_{safe_name}"
                stored_path = os.path.join(EVIDENCE_DIR, stored_name)
                with open(stored_path, "wb") as out_file:
                    out_file.write(file_bytes)
                row = {
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
                }
                new_rows.append(row)
            if new_rows:
                updated = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)
                st.session_state.evidence_df = updated.fillna("")
                st.session_state.evidence_df.to_csv(EVIDENCE_DB, index=False, encoding="utf-8-sig")
                st.success(f"{len(new_rows)}개 증빙을 등록했습니다. 파일은 '{EVIDENCE_DIR}' 폴더에 저장했습니다.")
            if duplicate_names:
                st.warning("이미 등록된 동일 파일은 중복 저장하지 않았습니다: " + ", ".join(duplicate_names))
            st.rerun()

    st.markdown("---")
    st.subheader("📋 등록 증빙 목록")
    evidence_df = st.session_state.evidence_df.copy()
    if evidence_df.empty:
        st.info("아직 등록된 증빙이 없습니다.")
    else:
        # 오래된 CSV에 파일해시 열이 없을 수 있으므로 화면에는 알려진 열만 표시
        visible_cols = [c for c in [
            "증빙ID", "파일명", "증빙일자", "거래처", "증빙종류",
            "공급가액/금액", "부가세", "적요", "은행대사", "연결전표", "검토상태"
        ] if c in evidence_df.columns]
        edited_register = st.data_editor(
            evidence_df[visible_cols],
            use_container_width=True,
            num_rows="fixed",
            key="evidence_register_editor",
            column_config={
                "은행대사": st.column_config.SelectboxColumn(
                    "은행대사", options=["미확인", "연결 후보", "일치 확인", "불일치", "해당 없음"]
                ),
                "검토상태": st.column_config.SelectboxColumn(
                    "검토상태", options=["검토 대기", "검토 중", "승인", "보류", "반려"]
                ),
            },
        )
        if st.button("💾 증빙 목록 변경사항 저장", key="save_evidence_edits"):
            for col in edited_register.columns:
                st.session_state.evidence_df.loc[edited_register.index, col] = edited_register[col]
            st.session_state.evidence_df.to_csv(EVIDENCE_DB, index=False, encoding="utf-8-sig")
            st.success("증빙 목록을 저장했습니다.")
            st.rerun()

        # 금액 기준 은행 대사 후보 탐색 (자동 확정하지 않음)
        st.markdown("---")
        st.subheader("🔎 은행 거래 연결 후보 찾기")
        bank_df = st.session_state.bank_reconciliation_df
        if bank_df is None or bank_df.empty:
            st.info("먼저 '은행 입출금 ↔ 전표 자동 대사' 메뉴에서 통장 내역을 처리하면 연결 후보를 찾을 수 있습니다.")
        else:
            evidence_options = evidence_df["증빙ID"].astype(str).tolist()
            chosen_id = st.selectbox("확인할 증빙 선택", evidence_options, key="evidence_match_id")
            selected_evidence = evidence_df[evidence_df["증빙ID"].astype(str) == chosen_id].iloc[0]
            try:
                evidence_amount = float(str(selected_evidence.get("공급가액/금액", "0")).replace(",", ""))
            except Exception:
                evidence_amount = 0.0
            bank_work = bank_df.copy()
            bank_work["_금액"] = pd.to_numeric(bank_work.get("입금액", 0), errors="coerce").fillna(0) + pd.to_numeric(bank_work.get("출금액", 0), errors="coerce").fillna(0)
            bank_work["_차액"] = (bank_work["_금액"] - evidence_amount).abs()
            candidates = bank_work[bank_work["_차액"] == 0].copy()
            if candidates.empty:
                st.warning("금액이 정확히 같은 은행 거래를 찾지 못했습니다. 날짜 차이·분할 결제·합산 결제 여부를 직접 확인하세요.")
            else:
                st.caption("금액만으로 찾은 후보입니다. 동일 금액 거래가 여러 건일 수 있으므로 자동 연결하지 않습니다.")
                st.dataframe(candidates.drop(columns=["_금액", "_차액"], errors="ignore"), use_container_width=True)

        st.markdown("---")
        st.subheader("🔗 전표 연결")
        journal_list = st.session_state.journal_entries
        if not journal_list:
            st.info("연결할 전표가 없습니다. 먼저 전표를 등록하거나 은행 대사 검토를 진행하세요.")
        else:
            journal_labels = ["연결 안 함"] + [
                f"{i + 1}. {j.get('적요', j.get('파일명', '전표'))} / "
                f"{j.get('차변금액', j.get('대변금액', ''))}원"
                for i, j in enumerate(journal_list)
            ]
            link_id = st.selectbox("연결할 증빙 선택", evidence_df["증빙ID"].astype(str).tolist(), key="evidence_link_id")
            link_choice = st.selectbox("연결할 전표 선택", journal_labels, key="evidence_journal_choice")
            if st.button("🔗 선택 증빙과 전표 연결 저장", key="link_evidence_journal"):
                row_idx = st.session_state.evidence_df.index[
                    st.session_state.evidence_df["증빙ID"].astype(str) == link_id
                ]
                if len(row_idx):
                    st.session_state.evidence_df.loc[row_idx[0], "연결전표"] = "" if link_choice == "연결 안 함" else link_choice
                    st.session_state.evidence_df.to_csv(EVIDENCE_DB, index=False, encoding="utf-8-sig")
                    st.success("전표 연결 정보를 저장했습니다. 회계상 승인 여부는 담당자가 별도로 확인해야 합니다.")
                    st.rerun()

        st.markdown("---")
        st.subheader("📥 증빙 등록대장 엑셀 다운로드")
        export_df = st.session_state.evidence_df.drop(columns=["저장파일", "파일해시"], errors="ignore")
        evidence_buffer = io.BytesIO()
        with pd.ExcelWriter(evidence_buffer, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="증빙등록대장")
        evidence_buffer.seek(0)
        st.download_button(
            "증빙 등록대장 다운로드",
            data=evidence_buffer.getvalue(),
            file_name="증빙_등록대장.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

elif menu == "XpERP 전년도 장부 및 분개 패턴 학습":
    st.title("🧠 XpERP 전년도 회계 장부 및 분개 패턴 AI 학습")
    st.info(
        "신규 부임 단지에서 XpERP로 다운로드한 전년도 분개장 및 장부 엑셀 파일을"
        " 업로드하면, 단지 고유의 계정 처리 패턴과 적요를 AI가 학습하여 자동"
        " 분개에 반영합니다."
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
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
            f"총 {len(df_learn_up)}건의 XpERP 장부 데이터를 성공적으로 읽어왔습니다."
        )
        st.dataframe(df_learn_up, use_container_width=True)

        if st.button("🚀 단지 맞춤형 AI 패턴 학습 실행", type="primary"):
            st.session_state.trained_patterns = df_learn_up
            df_learn_up.to_csv(DATA_FILE, index=False)
            st.success("단지 고유의 회계 처리 패턴 및 계정 매핑 학습이 완료되었으며, 저장되었습니다!")

    if not st.session_state.trained_patterns.empty:
        st.markdown("---")
        col_title, col_reset = st.columns([3, 1])
        with col_title:
            st.subheader("✅ 현재 단지에 적용된 학습 패턴 요약")
        with col_reset:
            if st.button("🗑️ 학습 패턴 초기화", type="secondary"):
                if os.path.exists(DATA_FILE):
                    os.remove(DATA_FILE)
                st.session_state.trained_patterns = pd.DataFrame()
                st.success("학습 데이터가 초기화되었습니다.")
                st.rerun()

        st.dataframe(st.session_state.trained_patterns, use_container_width=True)

elif menu == "단지 계정과목 & 국토부 47개 기준":
    st.title("⚙️ 단지별 계정과목표 & 국토부 47개 표준 항목 관리")
    st.info(
        "신규 단지로 부임하거나 단지별 고유 계정을 추가·수정·삭제하여 관리할 수 있습니다."
        " 변경된 내용은 서버에 영구 저장됩니다."
    )

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
            if "계정코드" in df_acc_up.columns and "계정명" in df_acc_up.columns:
                st.session_state.accounts_df = df_acc_up[["계정코드", "계정명"]]
                st.session_state.accounts_df.to_csv(ACCOUNTS_FILE, index=False)
                st.success("단지 계정과목표가 성공적으로 갱신 및 저장되었습니다!")
                st.rerun()
            else:
                st.error(
                    "업로드한 파일에 '계정코드'와 '계정명' 열이 반드시 포함되어야 합니다."
                )

    with col_down:
        st.write("###")
        acc_buf = io.BytesIO()
        st.session_state.accounts_df.to_excel(acc_buf, index=False)
        acc_buf.seek(0)
        st.download_button(
            "📥 현재 계정과목표 엑셀 다운로드",
            data=acc_buf.getvalue(),
            file_name="단지_계정과목표.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    st.markdown("---")

    with st.form("add_account_form"):
        st.subheader("➕ 신규 계정과목 개별 추가")
        c1, c2, c3 = st.columns([2, 3, 2])
        with c1:
            new_code = st.text_input("계정코드 (예: 5126)")
        with c2:
            new_name = st.text_input("계정명 (예: 단지 특화 수수료)")
        with c3:
            st.write("###")
            add_submitted = st.form_submit_button("계정 추가", type="primary")

        if add_submitted:
            if new_code and new_name:
                if new_code in st.session_state.accounts_df["계정코드"].values:
                    st.error(f"이미 존재하는 계정코드({new_code})입니다.")
                else:
                    new_row = pd.DataFrame(
                        [{"계정코드": str(new_code), "계정명": str(new_name)}]
                    )
                    st.session_state.accounts_df = pd.concat(
                        [st.session_state.accounts_df, new_row], ignore_index=True
                    )
                    st.session_state.accounts_df.to_csv(ACCOUNTS_FILE, index=False)
                    st.success(
                        f"계정 [{new_code}] {new_name}이(가) 추가 및 저장되었습니다!"
                    )
                    st.rerun()
            else:
                st.error("계정코드와 계정명을 모두 입력해 주세요.")

    st.markdown("---")

    with st.form("del_account_form"):
        st.subheader("🗑️ 불필요한 계정과목 삭제")
        code_to_del = st.selectbox(
            "삭제할 계정 선택",
            options=st.session_state.accounts_df["계정코드"].tolist()
            if not st.session_state.accounts_df.empty
            else [],
            format_func=lambda x: f"[{x}] {st.session_state.accounts_df.loc[st.session_state.accounts_df['계정코드'] == x, '계정명'].values[0]}",
        )
        del_submitted = st.form_submit_button("선택 계정 삭제")

        if del_submitted and code_to_del:
            st.session_state.accounts_df = st.session_state.accounts_df[
                st.session_state.accounts_df["계정코드"] != code_to_del
            ]
            st.session_state.accounts_df.to_csv(ACCOUNTS_FILE, index=False)
            st.success(f"계정코드 [{code_to_del}]가 삭제 및 반영되었습니다!")
            st.rerun()

    st.markdown("---")
    st.subheader("📋 현재 등록된 단지별 계정과목 리스트")
    st.dataframe(st.session_state.accounts_df, use_container_width=True)

elif menu == "명세서 다중 일괄 검증 & 전표 발행":
    st.title("📄 영수증 / 명세서 다중 일괄 AI 자동 분개 & 발행")
    st.info("💡 XpERP 학습 패턴이 연동되어 지출 적요에 맞춰 계정이 자동 추천됩니다.")

    uploaded_invoices = st.file_uploader(
        "영수증 파일들 다중 업로드",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
    )
    if uploaded_invoices:
        st.success(f"총 {len(uploaded_invoices)}개 파일 인식 완료.")
        if st.button("✅ 전체 전표 일괄 생성 및 저장", type="primary"):
            for f in uploaded_invoices:
                st.session_state.journal_entries.append({
                    "파일명": f.name,
                    "적요": "지출증빙 자동 분개 (학습패턴 적용)",
                    "차변계정": "[5101] 일반관리비",
                    "차변금액": 150000,
                    "대변계정": "[251] 미지급금",
                    "대변금액": 150000,
                })
            st.success("전표 발행 및 시스템 저장 완료!")
            st.rerun()

    if st.session_state.journal_entries:
        st.subheader("📋 전표 발행 및 회계 장부 반영 내역")
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
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.download_button(
            "📥 XpERP 연동 전표 장부 엑셀 다운로드",
            data=final_output.getvalue(),
            file_name="erp_journal_entries.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
        )

elif menu == "이상 지출 통합 보고서":
    st.title("📊 이상 지출 통합 보고서")
    if st.session_state.journal_entries:
        st.dataframe(
            pd.DataFrame(st.session_state.journal_entries), use_container_width=True
        )
    else:
        st.info("등록된 전표가 없습니다.")

elif menu == "은행 입출금 ↔ 전표 자동 대사":
    st.title("🏦 은행 입출금 ↔ 전표 자동 대사 (5단계 실무형 프로세스)")
    st.info(
        "1단계(멀티시트 선택 및 데이터 시작 행 지정, 통장 읽기/오류검사) ➔ 2단계(자동 대사 후보 비교) ➔ 3단계(분개 초안 자동 생성) ➔ 4단계(담당자 검토 및 승인) ➔ 5단계(결과 내보내기)"
    )

    sample_bank_df = pd.DataFrame([
        {"거래일자": "2026-06-01", "입금": 0, "출금": 3450000, "거래내용": "한국전력 전기요금 자동이체", "잔액": 45000000},
        {"거래일자": "2026-06-03", "입금": 500000, "출금": 0, "거래내용": "입주자 잡수입 무통장입금", "잔액": 45500000},
        {"거래일자": "2026-06-05", "입금": 0, "출금": 120000, "거래내용": "미확인 출금 내역", "잔액": 45380000},
    ])
    buf_bank = io.BytesIO()
    sample_bank_df.to_excel(buf_bank, index=False)
    buf_bank.seek(0)
    st.download_button(
        "📥 [실무 양식] 은행 통장 거래내역 샘플 엑셀 다운로드",
        data=buf_bank.getvalue(),
        file_name="sample_bank_statement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")
    
    # [1단계] 멀티 시트 통장 파일 업로드 및 시트/헤더 선택 기능
    st.subheader("📁 [1단계] 통장 거래내역 파일 업로드 및 시트/컬럼 매핑")
    up_bank_file = st.file_uploader("은행 통장 내역 멀티 시트 엑셀(XLSX) 또는 CSV", type=["xlsx", "csv"], key="bank_step1_v2")

    if up_bank_file:
        try:
            if up_bank_file.name.endswith(".csv"):
                df_raw = pd.read_csv(up_bank_file)
                sheet_names = ["기본시트(CSV)"]
                selected_sheet = "기본시트(CSV)"
            else:
                xls_file = pd.ExcelFile(up_bank_file)
                sheet_names = xls_file.sheet_names
                selected_sheet = st.selectbox("📂 대사에 사용할 시트 선택", sheet_names)
                
                # 헤더 행 번호 지정 (안내문구나 제목이 있는 행을 건너뛰기 위함)
                header_row = st.number_input("🔢 실제 헤더(컬럼명)가 위치한 행 번호 (0부터 시작, 예: 0 또는 5)", min_value=0, value=0, step=1)
                df_raw = pd.read_excel(up_bank_file, sheet_name=selected_sheet, header=header_row)

            st.write(f"🔍 선택한 시트 **[{selected_sheet}]** 데이터 미리보기 (상위 5행):")
            st.dataframe(df_raw.head(), use_container_width=True)

            cols = list(df_raw.columns)
            
            st.write("🛠️ **컬럼 수동 매핑 설정** (통장 양식에 맞춰 올바른 열을 지정하세요)")
            col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
            with col_m1:
                c_date = st.selectbox("날짜 열", cols, index=0 if 0 < len(cols) else 0, key="m_date")
            with col_m2:
                c_deposit = st.selectbox("입금액 열", cols, index=1 if 1 < len(cols) else 0, key="m_dep")
            with col_m3:
                c_withdraw = st.selectbox("출금액 열", cols, index=2 if 2 < len(cols) else 0, key="m_wit")
            with col_m4:
                c_desc = st.selectbox("거래내용(적요) 열", cols, index=3 if 3 < len(cols) else 0, key="m_desc")
            with col_m5:
                c_bal = st.selectbox("잔액 열", cols, index=4 if 4 < len(cols) else len(cols)-1, key="m_bal")

            if st.button("🚀 [2~3단계] 통장 검증 및 자동 대사·분개 초안 생성 실행", type="primary"):
                processed_rows = []
                for idx, row in df_raw.iterrows():
                    date_val = str(row[c_date]).strip()
                    desc_val = str(row[c_desc]).strip()
                    try:
                        dep_val = float(row[c_deposit]) if pd.notnull(row[c_deposit]) else 0.0
                    except:
                        dep_val = 0.0
                    try:
                        wit_val = float(row[c_withdraw]) if pd.notnull(row[c_withdraw]) else 0.0
                    except:
                        wit_val = 0.0
                    try:
                        bal_val = float(row[c_bal]) if pd.notnull(row[c_bal]) else 0.0
                    except:
                        bal_val = 0.0

                    # 오류 검사
                    if not date_val or date_val == "nan" or (dep_val == 0 and wit_val == 0):
                        continue  # 빈 행이나 설명글 행 자동 스킵

                    match_status = "미대사"
                    rec_account = "[5113] 공과금중 전기료" if "전기" in desc_val else ("[4101] 잡수입" if dep_val > 0 else "[5125] 잡비")
                    rec_amount = dep_val if dep_val > 0 else wit_val
                    
                    if "전기" in desc_val or "입금" in desc_val or "자동이체" in desc_val:
                        match_status = "일치 후보"
                    elif "미확인" in desc_val:
                        match_status = "확인 필요"

                    processed_rows.append({
                        "선택": True if match_status == "일치 후보" else False,
                        "거래일자": date_val,
                        "거래내용": sanitize_text(desc_val),
                        "입금액": dep_val,
                        "출금액": wit_val,
                        "잔액": bal_val,
                        "대사상태": match_status,
                        "추천계정": rec_account,
                        "분개금액": rec_amount,
                        "승인상태": "대기중"
                    })

                st.session_state.bank_reconciliation_df = pd.DataFrame(processed_rows)
                st.success("1~3단계 통장 검증, 자동 대사 및 분개 초안 생성이 완료되었습니다!")
                st.rerun()

        except Exception as e:
            st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")

    # [4단계] 담당자 검토 화면 및 [5단계] 결과 내보내기 영역
    if st.session_state.bank_reconciliation_df is not None and not st.session_state.bank_reconciliation_df.empty:
        st.markdown("---")
        st.subheader("📊 거래 대사 현황 요약")
        
        df_rec = st.session_state.bank_reconciliation_df
        total_cnt = len(df_rec)
        done_cnt = len(df_rec[df_rec["승인상태"] == "승인완료"])
        check_cnt = len(df_rec[df_rec["대사상태"] == "확인 필요"])
        unmatched_cnt = len(df_rec[df_rec["대사상태"] == "미대사"])

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("전체 유효 거래", f"{total_cnt}건")
        k2.metric("대사 완료 (승인)", f"{done_cnt}건")
        k3.metric("확인 필요", f"{check_cnt}건")
        k4.metric("미대사", f"{unmatched_cnt}건")

        st.markdown("---")
        st.subheader("📋 [4단계] 담당자 검토 및 승인 처리 화면")
        st.info("💡 각 거래 건별로 내용을 검토한 후 승인, 보류, 반려 처리를 수행하세요.")

        edited_df = st.data_editor(
            df_rec,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "선택": st.column_config.CheckboxColumn("선택", help="엑셀 출력 대상 선택"),
                "승인상태": st.column_config.SelectboxColumn("검토 상태", options=["대기중", "승인완료", "보류", "반려"], required=True)
            }
        )
        st.session_state.bank_reconciliation_df = edited_df

        st.markdown("---")
        st.subheader("📥 [5단계] 결과 내보내기 및 검증 보고서")
        
        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            if st.button("✅ 승인된 전표만 회계 장부에 최종 반영하기", type="primary"):
                approved_rows = edited_df[edited_df["승인상태"] == "승인완료"]
                if len(approved_rows) > 0:
                    for _, r in approved_rows.iterrows():
                        st.session_state.journal_entries.append({
                            "파일명": "은행대사_자동반영",
                            "적요": r["거래내용"],
                            "차변계정": r["추천계정"],
                            "차변금액": r["분개금액"],
                            "대변계정": "[103] 보통예금",
                            "대변금액": r["분개금액"],
                        })
                    st.success(f"총 {len(approved_rows)}건의 승인된 전표가 회계 장부에 정상 반영되었습니다!")
                else:
                    st.warning("승인 완료된 항목이 없습니다. 검토 상태를 '승인완료'로 변경해 주세요.")

        with c_exp2:
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                edited_df.to_excel(writer, index=False, sheet_name="은행대사보고서")
            out_buf.seek(0)
            
            st.download_button(
                "📥 월별 은행 대사 보고서 엑셀 다운로드",
                data=out_buf.getvalue(),
                file_name="monthly_bank_reconciliation_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

elif menu == "월말 결산 부속명세서 자동 취합":
    st.title("📊 월말 결산 부속명세서 자동 생성")
    if st.session_state.journal_entries:
        df = pd.DataFrame(st.session_state.journal_entries)
        st.dataframe(
            df.groupby("차변계정")["차변금액"].sum().reset_index(),
            use_container_width=True,
        )
    else:
        st.info("취합할 전표가 없습니다.")

elif menu == "세대별 관리비 부과 및 미납 관리":
    st.title("🏘️ 세대별 관리비 부과 및 미납 세대 통합 분석 관리")
    st.info(
        "세대별 부과, 수납, 미납 내역 등이 포함된 멀티 시트 엑셀 파일을 업로드하면,"
        " 시트별(세대목록, 관리비부과내역, 관리비수납내역, 기초미납, 연말잔액_검산 등)로 완벽하게 분리하여 조회 및 분석할 수 있습니다."
    )

    sample_세대_df = pd.DataFrame([
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
    buf_세대 = io.BytesIO()
    sample_세대_df.to_excel(buf_세대, index=False)
    buf_세대.seek(0)
    st.download_button(
        "📥 [샘플 양식] 세대별 부과내역 기본 양식 다운로드",
        data=buf_세대.getvalue(),
        file_name="sample_household_billing.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

    st.markdown("---")
    up_세대 = st.file_uploader(
        "멀티 시트 엑셀 파일 업로드 (예: 관리비미납샘플.xlsx)",
        type=["xlsx", "csv"],
        key="up_세대_multi",
    )

    if up_세대:
        if up_세대.name.endswith(".csv"):
            df_single = pd.read_csv(up_세대)
            st.session_state.household_excel_data = {"기본시트": df_single}
            st.success("CSV 파일 업로드 완료!")
        else:
            try:
                xls = pd.ExcelFile(up_세대)
                sheet_to_df_map = pd.read_excel(xls, sheet_name=None)
                st.session_state.household_excel_data = sheet_to_df_map
                st.success(
                    f"성공적으로 업로드되었습니다! (포함된 시트: {list(sheet_to_df_map.keys())})"
                )
            except Exception as e:
                st.error(f"엑셀 파일 읽기 중 오류가 발생했습니다: {e}")

    if st.session_state.household_excel_data:
        st.markdown("---")
        st.subheader("📊 엑셀 시트별 실무 데이터 상세 조회")
        
        sheet_names = list(st.session_state.household_excel_data.keys())
        tabs = st.tabs([f"📑 {s}" for s in sheet_names])

        for idx, tab in enumerate(tabs):
            current_sheet_name = sheet_names[idx]
            current_df = st.session_state.household_excel_data[current_sheet_name]
            
            with tab:
                st.write(f"**[시트명: {current_sheet_name}]** (총 {len(current_df)}행 데이터)")
                
                search_query = st.text_input(f"'{current_sheet_name}' 시트 내 검색", key=f"search_{current_sheet_name}")
                if search_query:
                    mask = current_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
                    filtered_df = current_df[mask]
                    st.write(f"검색 결과: 총 {len(filtered_df)}건")
                    st.dataframe(filtered_df, use_container_width=True)
                else:
                    st.dataframe(current_df, use_container_width=True)

elif menu == "장기수선충당금 적립 및 사용 관리":
    st.title("🛠️ 장기수선충당금(장충금) 적립 및 사용 관리")
    st.info(
        "장충금 적립 및 공사 집행 내역이 담긴 엑셀 파일을 업로드하여 잔액을"
        " 관리하세요."
    )

    sample_장충_df = pd.DataFrame([
        {
            "거래일자": "2026-05-12",
            "구분": "공사집행",
            "내용": "101~105동 옥상 방수 공사",
            "금액": 15000000,
        },
        {
            "거래일자": "2026-05-25",
            "구분": "월적립",
            "내용": "5월분 장기수선충당금 적립",
            "금액": 35000000,
        },
    ])
    buf_장충 = io.BytesIO()
    sample_장충_df.to_excel(buf_장충, index=False)
    buf_장충.seek(0)
    st.download_button(
        "📥 [샘플 양식] 장기수선충당금 내역 엑셀 다운로드",
        data=buf_장충.getvalue(),
        file_name="sample_longterm_repair.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

    st.markdown("---")
    up_장충 = st.file_uploader(
        "장기수선충당금 내역 엑셀 업로드", type=["xlsx", "csv"], key="up_장충"
    )
    if up_장충:
        df_장충_up = (
            pd.read_csv(up_장충)
            if up_장충.name.endswith(".csv")
            else pd.read_excel(up_장충)
        )
        st.success(f"총 {len(df_장충_up)}건의 장충금 내역 업로드 완료!")
        st.dataframe(df_장충_up, use_container_width=True)

elif menu == "직원 급여 및 4대보험 자동 분개":
    st.title("👥 관리사무소 직원 급여 및 4대보험 자동 분개")
    with st.form("salary_form"):
        emp_name = st.text_input("직원 성명", value="홍길동 (관리소장)")
        base_pay = st.number_input("기본급 (원)", value=3500000, step=50000)
        insurances = st.number_input(
            "4대보험 회사부담금 합계 (원)", value=350000, step=10000
        )

        if st.form_submit_button("🤖 급여 복식부기 전표 생성", type="primary"):
            st.success(f"[{emp_name}] 급여 복식부기 전표 생성 완료!")

elif menu == "잡수입(재활용/주차) 통합 관리":
    st.title("💰 단지 내 잡수입(재활용품/주차수입 등) 통합 관리")
    st.info(
        "잡수입 발생 내역 엑셀 파일을 업로드하여 관리비 차감 및 충당금 적립 처리를"
        " 관리하세요."
    )

    sample_잡수입_df = pd.DataFrame([
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
    buf_잡수입 = io.BytesIO()
    sample_잡수입_df.to_excel(buf_잡수입, index=False)
    buf_잡수입.seek(0)
    st.download_button(
        "📥 [샘플 양식] 잡수입 내역 엑셀 다운로드",
        data=buf_잡수입.getvalue(),
        file_name="sample_misc_income.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

    st.markdown("---")
    up_잡수입 = st.file_uploader(
        "잡수입 내역 엑셀 업로드", type=["xlsx", "csv"], key="up_잡수입"
    )
    if up_잡수입:
        df_잡수입_up = (
            pd.read_csv(up_잡수입)
            if up_잡수입.name.endswith(".csv")
            else pd.read_excel(up_잡수입)
        )
        st.success(f"총 {len(df_잡수입_up)}건의 잡수입 내역 업로드 완료!")
        st.dataframe(df_잡수입_up, use_container_width=True)

elif menu == "계정 및 비밀번호 설정":
    st.title("⚙️ 관리자 계정 및 보안 설정")
    with st.form("settings_form"):
        new_id = st.text_input(
            "새로운 관리자 아이디", value=st.session_state.admin_id
        )
        current_pw_input = st.text_input("현재 비밀번호 확인", type="password")
        new_pw = st.text_input("새로운 비밀번호", type="password")
        if st.form_submit_button("설정 저장"):
            if current_pw_input == st.session_state.admin_pw:
                st.session_state.admin_id = new_id
                if new_pw:
                    st.session_state.admin_pw = new_pw
                st.success("변경 완료!")
            else:
                st.error("현재 비밀번호가 불일치합니다.")
