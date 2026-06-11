import streamlit as st
import io, os, tempfile
from app import create_report_ppt

st.set_page_config(page_title="Hansol 사고보고서", page_icon="📋", layout="centered")

st.markdown("""
<style>
    .main { max-width: 680px; margin: 0 auto; }
    h1 { font-size: 20px !important; color: #00357A; }
    .stTextInput > label, .stTextArea > label, .stFileUploader > label,
    .stRadio > label, .stCheckbox > label { font-weight: 600; font-size: 13px; }
    .stButton > button { width: 100%; background: #00357A; color: white;
        font-size: 16px; font-weight: bold; padding: 12px; border-radius: 8px; border: none; }
    .stButton > button:hover { background: #002a62; }
    .stDownloadButton > button { width: 100%; background: #1a7f3c; color: white;
        font-size: 16px; font-weight: bold; padding: 12px; border-radius: 8px; border: none; }
</style>
""", unsafe_allow_html=True)

st.title("■ Hansol 공정사고 보고서")
st.caption("현장에서 입력 → PPT 자동 생성 (빈 칸 있어도 생성 가능)")

if st.button("🧪 더미데이터 자동입력 (테스트용)"):
    st.session_state.team = "생산2팀"
    st.session_state.date_str = "2026. 06. 11"
    st.session_state.title = "디지털 Safety Door Limit (Emergency_PB11) Fault 발생"
    st.session_state.datetime_val = "2026년 06월 11일 13시 25분"
    st.session_state.jijong = "TITAN 아트 118g"
    st.session_state.jungji = "13:25 ~ 00:05 (460분)"
    st.session_state.content = "가동중 Safety Door Limit Fault (Emergency_PB11) 발생"
    st.session_state.type_first = True
    st.session_state.situation = "13:25 가동중 Safety Door Limit (Emergency_PB11) Fault 발생으로 가동중지\n14:20 서비스반 Door 점검후 정상 가동, 20분후 동일 fault발생\n15:00 Safety Door Limit Fault 발생으로 추정으로 판단\n18:40 Steeter 전원 On/Off Reset 정상 가동, 20분후 동일 fault발생"
    st.session_state.cause = "PLC Hold (Turning Knife + Sheeter) 발생 추정\nDoor CH1 비정상 조건"
    st.session_state.action = "한국 알리스에 Safety Door Fault 관련 BWP측에 해제방법 요청\nDoor측 입출력값 및 Limit 동작확인 정상확인"
    st.session_state.product_handling = "전량 폐기"
    st.session_state.loss_cost = "적색"
    st.session_state.future = "동일 현상 발생시 PLC 판넬 동시 Reset (Turning Knife + Sheeter)\n가동중 Safety Door Limit Fault 발생 원인 및 해결 방법 문의\nSafety Door Limit Spare 구매"
    st.rerun()

st.divider()

# ── ① 기본 정보 ──────────────────────────────
st.subheader("① 기본 정보")
col1, col2 = st.columns(2)
with col1:
    team = st.text_input("팀명", placeholder="예) 생산2팀", key="team")
with col2:
    date_str = st.text_input("보고 날짜", placeholder="예) 2026. 06. 07", key="date_str")

title = st.text_input("사고 제목", placeholder="예) 디지털 Safety Door Limit (Emergency_PB11) Fault 발생", key="title")

st.divider()

# ── ② 개요 ───────────────────────────────────
st.subheader("② 개요")
datetime_val = st.text_input("일시", placeholder="예) 2026년 06월 07일 03시 20분", key="datetime_val")

col3, col4 = st.columns(2)
with col3:
    jijong = st.text_input("지종", placeholder="예) 뉴백상 플러스 80g", key="jijong")
with col4:
    jungji = st.text_input("중지시간", placeholder="예) 03:20 ~ 13:40 (620분)", key="jungji")

content = st.text_input("내용 요약", placeholder="예) Safety Door Limit 신호 이상 발생", key="content")

st.markdown("**발생유형**")
col5, col6 = st.columns(2)
with col5:
    type_first = st.checkbox("최초", key="type_first")
with col6:
    type_recur = st.checkbox("재발")

st.markdown("**결재란**")
col7, col8, col9 = st.columns(3)
with col7:
    approval_pl = st.text_input("P/L", placeholder="이름")
with col8:
    approval_team = st.text_input("팀장", placeholder="이름")
with col9:
    approval_factory = st.text_input("공장장", placeholder="이름")

st.divider()

# ── ③ 세부내용 ────────────────────────────────
st.subheader("③ 세부내용")
situation = st.text_area("발생 현황", height=120, key="situation",
    placeholder="시간대별 발생 내역을 입력하세요.\n예)\n03:20 제품측면 찍힘처 제거차 크레인으로 스폴 취외시...\n03:30 서비스반 설비점검 요청...")

action = st.text_area("점검 및 조치", height=100, key="action",
    placeholder="점검 및 조치 내용을 입력하세요.")

cause = st.text_area("발생 원인", height=100, key="cause",
    placeholder="발생 원인을 입력하세요.")

st.markdown("**발생 원인 사진 (최대 2장)**")
cause_images = st.file_uploader("사진 첨부", type=["jpg","jpeg","png"],
    accept_multiple_files=True, key="cause_imgs")
if cause_images and len(cause_images) > 2:
    st.warning("사진은 최대 2장까지 첨부 가능합니다.")
    cause_images = cause_images[:2]

col10, col11 = st.columns(2)
with col10:
    product_handling = st.text_input("기타 - 제품처리", placeholder="예) 전량 폐기", key="product_handling")
with col11:
    loss_cost = st.text_input("기타 - 손실비용", placeholder="예) 미정", key="loss_cost")

future = st.text_area("향후 대응방안", height=100, key="future",
    placeholder="향후 대응방안을 입력하세요.\n예)\n크레인 훅 체결방지 장치 타입변경 검토\n센서 예비품 확보")

st.divider()

# ── ④ 이메일 발송 ─────────────────────────────
st.subheader("④ 이메일 발송 (선택)")
email_to = st.text_input("수신 이메일", placeholder="받는 사람 이메일 주소")

st.divider()

# ── 생성 버튼 ─────────────────────────────────
if st.button("📊 PPT 생성"):
    with st.spinner("PPT 생성 중..."):
            # 이미지 임시 저장
            tmp_dir = tempfile.mkdtemp()
            sit_paths, cause_paths = [], []

            for f in []:
                p = os.path.join(tmp_dir, f.name)
                with open(p, 'wb') as out: out.write(f.read())
                sit_paths.append(p)

            for f in (cause_images or []):
                p = os.path.join(tmp_dir, f.name)
                with open(p, 'wb') as out: out.write(f.read())
                cause_paths.append(p)

            data = {
                'title': title, 'team': team, 'date_str': date_str,
                'datetime': datetime_val, 'jijong': jijong,
                'jungji': jungji, 'content': content,
                'type_first': type_first, 'type_recur': type_recur,
                'approval_pl': approval_pl, 'approval_team': approval_team,
                'approval_factory': approval_factory,
                'situation': situation, 'cause': cause, 'action': action,
                'product_handling': product_handling, 'loss_cost': loss_cost,
                'future': future,
            }

            ppt_bytes = create_report_ppt(data, sit_paths, cause_paths)

            # 임시 파일 정리
            for p in sit_paths + cause_paths:
                try: os.unlink(p)
                except: pass

            fname = f"공정사고보고서_{date_str.replace('.','').replace(' ','')}.pptx"

            st.success("✅ PPT 생성 완료!")
            st.download_button(
                label="⬇️ PPT 다운로드",
                data=ppt_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

            # 이메일 발송
            if email_to.strip():
                try:
                    import smtplib, os, io
                    from email.message import EmailMessage
                    mail_user = os.environ.get('MAIL_USER','')
                    mail_pass = os.environ.get('MAIL_PASS','')
                    msg = EmailMessage()
                    msg['From'] = mail_user
                    msg['To'] = email_to.strip()
                    msg['Subject'] = '[Hansol] Accident Report'
                    msg.set_payload('Hansol accident report attached.')
                    msg.add_attachment(ppt_bytes,
                                       maintype='application',
                                       subtype='octet-stream',
                                       filename='report.pptx')
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
                        s.login(mail_user, mail_pass)
                        s.send_message(msg)
                    st.success(f"메일 발송 완료!")
                except Exception as e:
                    import traceback
                    st.warning(f"메일 발송 실패: {traceback.format_exc()}")
