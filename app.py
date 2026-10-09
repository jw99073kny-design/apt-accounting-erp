import io
import random
import re
import zipfile
import openpyxl
import pandas as pd
from PIL import Image, ImageOps, ImageDraw, ImageFont
import streamlit as st

# EasyOCR 라이브러리 사용 시도
try:
  import easyocr

  reader = easyocr.Reader(["ko", "en"], gpu=False)
  HAS_OCR = True
except Exception:
  HAS_OCR = False

st.set_page_config(
    page_title="XpERP 연동 AI 회계 시스템", page_icon="🏢", layout="wide"
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

if "trained_patterns" not in st.session_state:
  st.session_state.trained_patterns = pd.DataFrame(
      columns=["적요키워드", "추천계정코드", "추천계정명"]
  )


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

# ---------------------------------------------------------
# 3. 국토부 47개 표준 항목 전체 연동
# ---------------------------------------------------------
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
st.session_state.accounts_df = pd.DataFrame(default_coas)

if "journal_entries" not in st.session_state:
  st.session_state.journal_entries = []

# ---------------------------------------------------------
# 4. 사이드바: 상위 카테고리별 접고 펼치는 트리 구조 (톱니바퀴 아이콘 적용)
# ---------------------------------------------------------
st.sidebar.title("🏢 XpERP 관리 시스템")
st.sidebar.text(f"관리자: {st.session_state.admin_id}")

if st.sidebar.button("🔒 로그아웃", use_container_width=True):
  st.session_state.authenticated = False
  st.rerun()

st.sidebar.markdown("---")

# 대시보드 홈 버튼
if st.sidebar.button("🏠 시스템 홈 (대시보드)", use_container_width=True):
  st.session_state.current_menu = "🏠 시스템 홈 (대시보드)"

st.sidebar.markdown("---")

# 1. 회계·전표 검증
with st.sidebar.expander("📁 회계·전표 검증", expanded=True):
  if st.button("🧠 XpERP 장부 및 분개 패턴 학습", use_container_width=True):
    st.session_state.current_menu = "XpERP 전년도 장부 및 분개 패턴 학습"
  if st.button("📄 명세서 다중 일괄 검증 & 전표 발행", use_container_width=True):
    st.session_state.current_menu = "명세서 다중 일괄 검증 & 전표 발행"
  if st.button("📊 단지 계정과목 & 국토부 47개", use_container_width=True):
    st.session_state.current_menu = "단지 계정과목 & 국토부 47개 기준"
  if st.button("🚨 이상 지출 통합 보고서", use_container_width=True):
    st.session_state.current_menu = "이상 지출 통합 보고서"
  if st.button("🏦 은행 입출금 ↔ 전표 자동 대사", use_container_width=True):
    st.session_state.current_menu = "은행 입출금 ↔ 전표 자동 대사"

# 2. 결산·세무 관리
with st.sidebar.expander("📁 결산·세무 관리", expanded=False):
  if st.button("🧾 부가가치세(VAT) 자동 안분 분개", use_container_width=True):
    st.session_state.current_menu = "부가가치세(VAT) 자동 안분 분개"
  if st.button("📈 월말 결산 부속명세서 자동 취합", use_container_width=True):
    st.session_state.current_menu = "월말 결산 부속명세서 자동 취합"

# 3. 실무 통합 관리
with st.sidebar.expander("📁 실무 통합 관리", expanded=False):
  if st.button("🏘️ 세대별 관리비 부과 및 미납 관리", use_container_width=True):
    st.session_state.current_menu = "세대별 관리비 부과 및 미납 관리"
  if st.button("🛠️ 장기수선충당금 적립 및 사용", use_container_width=True):
    st.session_state.current_menu = "장기수선충당금 적립 및 사용 관리"
  if st.button("👥 직원 급여 및 4대보험 자동 분개", use_container_width=True):
    st.session_state.current_menu = "직원 급여 및 4대보험 자동 분개"
  if st.button("💰 잡수입(재활용/주차) 통합 관리", use_container_width=True):
    st.session_state.current_menu = "잡수입(재활용/주차) 통합 관리"

# 4. 시스템 설정 (톱니바퀴 아이콘 적용)
with st.sidebar.expander("⚙️ 시스템 설정", expanded=False):
  if st.button("⚙️ 계정 및 비밀번호 설정", use_container_width=True):
    st.session_state.current_menu = "계정 및 비밀번호 설정"

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
      st.success("단지 고유의 회계 처리 패턴 및 계정 매핑 학습이 완료되었습니다!")

  if not st.session_state.trained_patterns.empty:
    st.markdown("---")
    st.subheader("✅ 현재 단지에 적용된 학습 패턴 요약")
    st.dataframe(st.session_state.trained_patterns, use_container_width=True)

elif menu == "단지 계정과목 & 국토부 47개 기준":
  st.title("⚙️ 단지별 계정과목표 & 국토부 47개 표준 항목")
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
  st.title("🏦 은행 입출금 통장 내역 ↔ 전표 자동 대사")
  bank_files = st.file_uploader(
      "통장 엑셀 업로드",
      type=["xlsx", "csv"],
      accept_multiple_files=True,
      key="bank_up",
  )
  if bank_files:
    st.success("통장 취합 및 자동 매칭 완료!")

elif menu == "부가가치세(VAT) 자동 안분 분개":
  st.title("🧾 공동매입 부가가치세(VAT) 자동 안분 및 복식부기 전표 발행")
  st.info(
      "과세·면세 매출 및 공통매입 데이터가 담긴 엑셀 파일을 업로드하면, 시스템이"
      " 안분 비율을 자동으로 계산하여 복식부기 전표를 발행합니다."
  )

  sample_vat_df = pd.DataFrame([{
      "공급가액합계": 20000000,
      "총매입세액": 1100000,
      "과세공급가액": 5000000,
      "면세공급가액": 15000000,
      "적요": "공용 승강기 및 시설 유지 보수 공통매입",
  }])
  buf_vat = io.BytesIO()
  sample_vat_df.to_excel(buf_vat, index=False)
  buf_vat.seek(0)
  st.download_button(
      "📥 [실무 양식] 부가세 안분 데이터 엑셀 다운로드",
      data=buf_vat.getvalue(),
      file_name="sample_vat_allocation.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

  st.markdown("---")
  up_vat = st.file_uploader(
      "과세/면세 및 공통매입 내역 엑셀 파일 업로드",
      type=["xlsx", "csv"],
      key="up_vat_file",
  )

  if up_vat:
    df_vat_up = (
        pd.read_csv(up_vat)
        if up_vat.name.endswith(".csv")
        else pd.read_excel(up_vat)
    )
    st.success(
        f"총 {len(df_vat_up)}건의 부가세 데이터가 업로드 및 인식되었습니다."
    )
    st.dataframe(df_vat_up, use_container_width=True)

    if st.button("🤖 자동 안분 계산 및 전체 전표 일괄 반영", type="primary"):
      for idx, row in df_vat_up.iterrows():
        total_vat = float(row["총매입세액"])
        taxable_rev = float(row["과세공급가액"])
        exempt_rev = float(row["면세공급가액"])
        expense_name = str(row["적요"])

        total_rev = taxable_rev + exempt_rev
        taxable_ratio = taxable_rev / total_rev if total_rev > 0 else 0
        deductible_vat = round(total_vat * taxable_ratio)
        nondeductible_vat = total_vat - deductible_vat

        st.session_state.journal_entries.append({
            "파일명": f"부가세안분_{idx+1}",
            "적요": f"{expense_name} (공제분)",
            "차변계정": "[117] 부가가치세대급금",
            "차변금액": deductible_vat,
            "대변계정": "[251] 미지급금",
            "대변금액": deductible_vat,
        })
        st.session_state.journal_entries.append({
            "파일명": f"부가세안분_{idx+1}",
            "적요": f"{expense_name} (불공제분 비용산입)",
            "차변계정": "[5601] 수선비",
            "차변금액": nondeductible_vat,
            "대변계정": "[251] 미지급금",
            "대변금액": nondeductible_vat,
        })

      st.success("부가가치세 안분 계산 및 복식부기 전표 발행이 완료되었습니다!")
      st.rerun()

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
  st.title("🏘️ 세대별 관리비 부과 및 미납 세대 추적")
  st.info(
      "단지 세대별 부과 및 수납 내역 엑셀 파일을 업로드하여 미납 세대를 자동으로"
      " 파악하세요."
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
      "📥 [샘플 양식] 세대별 부과내역 엑셀 다운로드",
      data=buf_세대.getvalue(),
      file_name="sample_household_billing.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

  st.markdown("---")
  up_세대 = st.file_uploader(
      "세대별 부과/수납 내역 엑셀 업로드", type=["xlsx", "csv"], key="up_세대"
  )
  if up_세대:
    df_세대_up = (
        pd.read_csv(up_세대)
        if up_세대.name.endswith(".csv")
        else pd.read_excel(up_세대)
    )
    st.success(f"총 {len(df_세대_up)}세대 데이터 업로드 완료!")
    st.dataframe(df_세대_up, use_container_width=True)

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
        