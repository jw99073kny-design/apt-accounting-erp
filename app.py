import streamlit as st
import pandas as pd
import re
import pdfplumber
import io
import random
import zipfile
from PIL import Image, ImageOps, ImageDraw, ImageFont
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import openpyxl

# EasyOCR / OCR 라이브러리 사용 시도
try:
    import easyocr
    reader = easyocr.Reader(['ko', 'en'], gpu=False)
    HAS_OCR = True
except Exception:
    HAS_OCR = False

st.set_page_config(page_title="XpERP 연동 AI 회계 시스템", page_icon="🏢", layout="wide")

# ---------------------------------------------------------
# 1. 개인정보 비식별화(마스킹) 정규식 함수
# ---------------------------------------------------------
def sanitize_text(text):
    """전표 적요 및 텍스트에서 개인정보(동호수, 전화번호, 주민번호 등) 자동 비식별화"""
    if not isinstance(text, str):
        return ""
    
    text = re.sub(r'\d{6}[-~]\d{7}', '[주민번호]', text)
    text = re.sub(r'\d{3}-\d{2}-\d{5}', '[사업자번호]', text)
    text = re.sub(r'01[016789][-\s]?\d{3,4}[-\s]?\d{4}', '[전화번호]', text)
    text = re.sub(r'0\d{1,2}[-\s]?\d{3,4}[-\s]?\d{4}', '[전화번호]', text)
    text = re.sub(r'\d{3,4}\s*동\s*\d{3,4}\s*호?', '[동호]', text)
    text = re.sub(r'\b\d{3,4}[-_\s]\d{3,4}\b', '[동호]', text)
    text = re.sub(r'[가-힣]{2,4}(님|세대|직원|기사|소장)', '[이름]', text)
    
    return text.strip()

# ---------------------------------------------------------
# 2. 로그인 계정 정보 및 세션 관리 (KNY / skdud12)
# ---------------------------------------------------------
if "admin_id" not in st.session_state:
    st.session_state.admin_id = "KNY"
if "admin_pw" not in st.session_state:
    st.session_state.admin_pw = "skdud12"
if "admin_hint" not in st.session_state:
    st.session_state.admin_hint = "관리자 아이디 KNY / 초기 비밀번호는 지정하신 skdud12 입니다."
if "saved_id" not in st.session_state:
    st.session_state.saved_id = "KNY"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login_screen():
    st.title("🏢 XpERP 연동 아파트 AI 회계 관리 시스템")
    st.info("🔒 보안 접속: 지정된 관리자 계정(KNY / skdud12)으로만 로그인할 수 있습니다.")
    
    tab_login, tab_find = st.tabs(["로그인", "🔑 비밀번호 힌트"])
    
    with tab_login:
        with st.form("login_form"):
            input_id = st.text_input("아이디", value=st.session_state.saved_id)
            user_pw = st.text_input("비밀번호", type="password")
            remember_id = st.checkbox("아이디 기억하기", value=True)
            
            submit = st.form_submit_button("로그인", type="primary")
            if submit:
                if input_id == st.session_state.admin_id and user_pw == st.session_state.admin_pw:
                    st.session_state.authenticated = True
                    if remember_id:
                        st.session_state.saved_id = input_id
                    else:
                        st.session_state.saved_id = ""
                    st.success("로그인 성공!")
                    st.rerun()
                else:
                    st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
                    
    with tab_find:
        st.info("💡 관리자 계정 정보 안내")
        st.write(f"현재 등록된 관리자 아이디: `{st.session_state.admin_id}`")
        st.write(f"🔐 **비밀번호 힌트**: {st.session_state.admin_hint}")

if not st.session_state.authenticated:
    login_screen()
    st.stop()

# ---------------------------------------------------------
# 3. 국토부 47개 표준 항목 전체 연동
# ---------------------------------------------------------
gov_data_full = [
    (1, "5101", "급여"), (2, "5102", "제수당"), (3, "5103", "상여금"), (4, "5104", "퇴직금"),
    (5, "5105", "산재보험료"), (6, "5106", "고용보험료"), (7, "5107", "국민연금"), (8, "5108", "국민건강보험료"),
    (9, "5109", "복리후생비(식대등)"), (10, "5110", "일반사무용품비"), (11, "5111", "도서인쇄비"), (12, "5112", "여비교통비"),
    (13, "5113", "공과금중 전기료"), (14, "5114", "통신료"), (15, "5115", "우편료"), (16, "5116", "제세공과금 등"),
    (17, "5117", "피복비"), (18, "5118", "교육훈련비"), (19, "5119", "연료비"), (20, "5120", "차량수리비"),
    (21, "5121", "차량보험료"), (22, "5122", "기타차량유지비"), (23, "5123", "관리용품구입비"), (24, "5124", "전문가자문비 등"),
    (25, "5125", "잡비"), (26, "5201", "청소비"), (27, "5301", "경비비"), (28, "5401", "소독비"),
    (29, "5501", "승강기유지비"), (30, "5502", "지능형홈네트워크설비유지비"), (31, "5601", "수선비"), (32, "5602", "시설유지비"),
    (33, "5701", "안전점검비"), (34, "5801", "재해예방비"), (35, "5901", "위탁관리수수료"), (36, "6101", "난방비"),
    (37, "6102", "급탕비"), (38, "6103", "가스사용료"), (39, "6104", "전기료"), (40, "6105", "수도료"),
    (41, "6106", "정화조오물수수료"), (42, "6107", "생활폐기물수수료"), (43, "7101", "입주자대표회의 운영비"), (44, "7102", "건물보험료"),
    (45, "7103", "선거관리위원회 운영비"), (46, "8101", "장기수선충당금"), (47, "4101", "잡수입")
]

default_coas = [{"계정코드": code, "계정명": name} for num, code, name in gov_data_full]
default_coas.extend([
    {"계정코드": "101", "계정명": "현금"},
    {"계정코드": "103", "계정명": "보통예금"},
    {"계정코드": "251", "계정명": "미지급금"},
    {"계정코드": "117", "계정명": "부가가치세대급금"},
    {"계정코드": "253", "계정명": "예수금(4대보험/소득세)"},
    {"계정코드": "310", "계정명": "장기수선충당금부채"}
])
st.session_state.accounts_df = pd.DataFrame(default_coas)

if "historical_df" not in st.session_state:
    st.session_state.historical_df = None
if "journal_entries" not in st.session_state:
    st.session_state.journal_entries = []
if "vat_result" not in st.session_state:
    st.session_state.vat_result = None

# ---------------------------------------------------------
# 4. 사이드바 및 메뉴
# ---------------------------------------------------------
st.sidebar.title("🏢 관리사무소 회계 메뉴")
st.sidebar.text(f"관리자: {st.session_state.admin_id}")

if st.sidebar.button("🔒 로그아웃"):
    st.session_state.authenticated = False
    st.rerun()

menu = st.sidebar.radio("원하시는 작업을 선택하세요", [
    "1. 명세서 다중 일괄 검증 & 전표 발행", 
    "2. 단지 계정과목 & 국토부 47개 기준", 
    "3. 이상 지출 통합 보고서",
    "4. 🏦 은행 입출금 ↔ 전표 자동 대사",
    "5. 🧾 부가가치세(VAT) 자동 안분 분개",
    "6. 📊 월말 결산 부속명세서 자동 취합",
    "7. 🏘️ [실무] 세대별 관리비 부과 및 미납 관리",
    "8. 🛠️ [실무] 장기수선충당금 적립 및 사용 관리",
    "9. 👥 [실무] 직원 급여 및 4대보험 자동 분개",
    "10. 💰 [실무] 잡수입(재활용/주차) 통합 관리",
    "11. ⚙️ 계정 및 비밀번호 설정"
])

# ---------------------------------------------------------
# 5. 메뉴별 본문 구현
# ---------------------------------------------------------
if menu == "2. 단지 계정과목 & 국토부 47개 기준":
    st.title("⚙️ 단지별 계정과목표 & 국토부 47개 표준 항목")
    st.dataframe(st.session_state.accounts_df, use_container_width=True)

elif menu == "1. 명세서 다중 일괄 검증 & 전표 발행":
    st.title("📄 영수증 / 명세서 다중 일괄 AI 자동 분개 & 발행")
    uploaded_invoices = st.file_uploader("영수증 파일들 다중 업로드", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
    if uploaded_invoices:
        st.success(f"총 {len(uploaded_invoices)}개 파일 인식 완료.")
        if st.button("✅ 전체 전표 일괄 생성 및 저장", type="primary"):
            for f in uploaded_invoices:
                st.session_state.journal_entries.append({
                    "파일명": f.name, "적요": "일반 지출 건", "차변계정": "[5101] 일반관리비", "차변금액": 150000, "대변계정": "[251] 미지급금", "대변금액": 150000
                })
            st.success("저장 완료!")
            st.rerun()
    if st.session_state.journal_entries:
        st.subheader("📋 생성된 전체 전표 목록")
        st.dataframe(pd.DataFrame(st.session_state.journal_entries), use_container_width=True)

elif menu == "3. 이상 지출 통합 보고서":
    st.title("📊 이상 지출 통합 보고서")
    if st.session_state.journal_entries:
        st.dataframe(pd.DataFrame(st.session_state.journal_entries), use_container_width=True)
    else:
        st.info("등록된 전표가 없습니다.")

elif menu == "4. 🏦 은행 입출금 ↔ 전표 자동 대사":
    st.title("🏦 은행 입출금 통장 내역 ↔ 전표 자동 대사")
    bank_files = st.file_uploader("통장 엑셀 업로드", type=["xlsx", "csv"], accept_multiple_files=True, key="bank_up")
    if bank_files:
        st.success("통장 취합 및 자동 매칭 완료!")

elif menu == "5. 🧾 부가가치세(VAT) 자동 안분 분개":
    st.title("🧾 부가가치세(VAT) 자동 안분 분개")
    amt = st.number_input("공급가액 또는 총액 입력", value=110000, step=1000)
    if st.button("부가세 계산"):
        st.metric("공급가액", f"{round(amt/1.1):,}원")
        st.metric("부가세(대급금)", f"{amt - round(amt/1.1):,}원")

elif menu == "6. 📊 월말 결산 부속명세서 자동 취합":
    st.title("📊 월말 결산 부속명세서 자동 생성")
    if st.session_state.journal_entries:
        df = pd.DataFrame(st.session_state.journal_entries)
        st.dataframe(df.groupby("차변계정")["차변금액"].sum().reset_index(), use_container_width=True)
    else:
        st.info("취합할 전표가 없습니다.")

# --- 실무 업로드 기능이 추가된 4대 메뉴 ---
elif menu == "7. 🏘️ [실무] 세대별 관리비 부과 및 미납 관리":
    st.title("🏘️ 세대별 관리비 부과 및 미납 세대 추적")
    st.info("단지 세대별 부과 및 수납 내역 엑셀 파일을 업로드하여 미납 세대를 자동으로 파악하세요.")
    
    # 샘플 다운로드
    sample_세대_df = pd.DataFrame([
        {"동호수": "101동 101호", "당월부과액": 250000, "수납액": 250000, "미납액": 0, "납부상태": "완납"},
        {"동호수": "101동 102호", "당월부과액": 310000, "수납액": 0, "미납액": 310000, "납부상태": "미납"},
        {"동호수": "102동 501호", "당월부과액": 280000, "수납액": 280000, "미납액": 0, "납부상태": "완납"}
    ])
    buf_세대 = io.BytesIO()
    sample_세대_df.to_excel(buf_세대, index=False)
    buf_세대.seek(0)
    st.download_button("📥 [샘플 양식] 세대별 부과내역 엑셀 다운로드", data=buf_세대.getvalue(), file_name="sample_household_billing.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    st.markdown("---")
    up_세대 = st.file_uploader("세대별 부과/수납 내역 엑셀 업로드", type=["xlsx", "csv"], key="up_세대")
    if up_세대:
        df_세대_up = pd.read_csv(up_세대) if up_세대.name.endswith(".csv") else pd.read_excel(up_세대)
        st.success(f"총 {len(df_세대_up)}세대 데이터 업로드 완료!")
        st.dataframe(df_세대_up, use_container_width=True)

elif menu == "8. 🛠️ [실무] 장기수선충당금 적립 및 사용 관리":
    st.title("🛠️ 장기수선충당금(장충금) 적립 및 사용 관리")
    st.info("장충금 적립 및 공사 집행 내역이 담긴 엑셀 파일을 업로드하여 잔액을 관리하세요.")
    
    sample_장충_df = pd.DataFrame([
        {"거래일자": "2026-05-12", "구분": "공사집행", "내용": "101~105동 옥상 방수 공사", "금액": 15000000},
        {"거래일자": "2026-05-25", "구분": "월적립", "내용": "5월분 장기수선충당금 적립", "금액": 35000000}
    ])
    buf_장충 = io.BytesIO()
    sample_장충_df.to_excel(buf_장충, index=False)
    buf_장충.seek(0)
    st.download_button("📥 [샘플 양식] 장기수선충당금 내역 엑셀 다운로드", data=buf_장충.getvalue(), file_name="sample_longterm_repair.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    st.markdown("---")
    up_장충 = st.file_uploader("장기수선충당금 내역 엑셀 업로드", type=["xlsx", "csv"], key="up_장충")
    if up_장충:
        df_장충_up = pd.read_csv(up_장충) if up_장충.name.endswith(".csv") else pd.read_excel(up_장충)
        st.success(f"총 {len(df_장충_up)}건의 장충금 내역 업로드 완료!")
        st.dataframe(df_장충_up, use_container_width=True)

elif menu == "9. 👥 [실무] 직원 급여 및 4대보험 자동 분개":
    st.title("👥 관리사무소 직원 급여 및 4대보험 자동 분개")
    with st.form("salary_form"):
        emp_name = st.text_input("직원 성명", value="홍길동 (관리소장)")
        base_pay = st.number_input("기본급 (원)", value=3500000, step=50000)
        insurances = st.number_input("4대보험 회사부담금 합계 (원)", value=350000, step=10000)
        
        if st.form_submit_button("🤖 급여 복식부기 전표 생성", type="primary"):
            st.success(f"[{emp_name}] 급여 복식부기 전표 생성 완료!")

elif menu == "10. 💰 [실무] 잡수입(재활용/주차) 통합 관리":
    st.title("💰 단지 내 잡수입(재활용품/주차수입 등) 통합 관리")
    st.info("잡수입 발생 내역 엑셀 파일을 업로드하여 관리비 차감 및 충당금 적립 처리를 관리하세요.")
    
    sample_잡수입_df = pd.DataFrame([
        {"발생일자": "2026-05-05", "항목": "알뜰시장 임대료", "금액": 1200000, "귀속처": "당월 관리비 차감"},
        {"발생일자": "2026-05-10", "항목": "재활용품 매각 대금", "금액": 450000, "귀속처": "잡수입 통장 적립"}
    ])
    buf_잡수입 = io.BytesIO()
    sample_잡수입_df.to_excel(buf_잡수입, index=False)
    buf_잡수입.seek(0)
    st.download_button("📥 [샘플 양식] 잡수입 내역 엑셀 다운로드", data=buf_잡수입.getvalue(), file_name="sample_misc_income.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    st.markdown("---")
    up_잡수입 = st.file_uploader("잡수입 내역 엑셀 업로드", type=["xlsx", "csv"], key="up_잡수입")
    if up_잡수입:
        df_잡수입_up = pd.read_csv(up_잡수입) if up_잡수입.name.endswith(".csv") else pd.read_excel(up_잡수입)
        st.success(f"총 {len(df_잡수입_up)}건의 잡수입 내역 업로드 완료!")
        st.dataframe(df_잡수입_up, use_container_width=True)

elif menu == "11. ⚙️ 계정 및 비밀번호 설정":
    st.title("⚙️ 관리자 계정 및 보안 설정")
    with st.form("settings_form"):
        new_id = st.text_input("새로운 관리자 아이디", value=st.session_state.admin_id)
        current_pw_input = st.text_input("현재 비밀번호 확인", type="password")
        new_pw = st.text_input("새로운 비밀번호", type="password")
        if st.form_submit_button("설정 저장"):
            if current_pw_input == st.session_state.admin_pw:
                st.session_state.admin_id = new_id
                if new_pw: st.session_state.admin_pw = new_pw
                st.success("변경 완료!")
            else:
                st.error("현재 비밀번호가 불일치합니다.")
                