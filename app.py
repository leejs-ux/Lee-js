import streamlit as st
import pandas as pd
import ezdxf
import os
import tempfile
from io import BytesIO
import openpyxl
import json
import re
import time
import google.generativeai as genai
import gspread

# 💡 [서버 다운 방지] 웹 환경에서 이미지 그리기 도구를 안전하게 쓰기 위한 필수 설정
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from PIL import Image # AI에게 이미지를 전달할 정식 도구

st.set_page_config(page_title="2D DXF 하이브리드 자동 견적 시스템", page_icon="👁️", layout="wide")

# =========================================================================
# 💡 구글 시트 연결 (안전하게 정리)
# =========================================================================
SHEET_NAME = "견적프로그램_DB"

@st.cache_resource
def init_gspread():
    if "google_credentials" in st.secrets:
        try:
            cred_dict = json.loads(st.secrets["google_credentials"], strict=False)
            return gspread.service_account_from_dict(cred_dict)
        except Exception as e:
            st.error(f"⚠️ 구글 시트 인증 에러: {e}")
            return None
    return None

gc = init_gspread()

# =========================================================================
# 🔑 API 키 자동 저장/불러오기
# =========================================================================
API_KEY_FILE = "api_key.txt"
def load_api_key():
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, "r") as f:
            return f.read().strip()
    return ""

def save_api_key(key):
    with open(API_KEY_FILE, "w") as f:
        f.write(key.strip())

st.sidebar.header("🔑 AI 설정")
saved_key = load_api_key()
api_key = st.sidebar.text_input("Gemini API Key를 입력하세요", value=saved_key, type="password")
if st.sidebar.button("💾 API 키 저장하기"):
    if api_key:
        save_api_key(api_key)
        st.sidebar.success("✅ 키 저장 완료!")

st.title("👁️ AI 비전(Vision) 기반 하이브리드 견적 시스템")
st.markdown("---")

def safe_float(value):
    try:
        if isinstance(value, (int, float)): 
            return float(value)
        num_str = re.sub(r'[^0-9.]', '', str(value))
        return float(num_str) if num_str else 0.0
    except:
        return 0.0

# =========================================================================
# 1. 기준 단가표 관리 (가독성 개선)
# =========================================================================
default_material_db = pd.DataFrame({
    '재질': ['SS400', 'S45C', 'SPCC(레이져)', 'SM45C', 'AL2017', 'AL5052고베판', 'AL6061판재', 'AL7075', 'SUS304', 'BS(신주)', 'MC 나이론', '아세탈', '테프론', 'PC (국산)', 'PUR.'],
    'KG당 단가': [2400, 2400, 1500, 2400, 25000, 9300, 8000, 11500, 7650, 10000, 12000, 15000, 40000, 10000, 15000],
    '비중': [8.0, 8.0, 8.0, 8.0, 2.8, 2.8, 2.8, 2.8, 8.0, 8.0, 1.6, 1.41, 2.2, 1.2, 1.5]
})

default_post_db = pd.DataFrame({
    '표면처리': ['W-Anodizing', 'B-Anodizing', 'H-Anodizing', 'SOFT ANODIZING', '무전해니켈(ST)', '무전해니켈(AL)', '크롬도금', '전해연마', '아연도금', '흑색경질', 'POLISHING', 'PAINT'],
    'KG당 단가': [1500, 2500, 6000, 3000, 2500, 6000, 1500, 1500, 800, 7000, 2500, 600]
})

if 'material_db' not in st.session_state:
    if gc:
        try:
            ws_m = gc.open(SHEET_NAME).worksheet("material_db")
            data_m = ws_m.get_all_records()
            st.session_state.material_db = pd.DataFrame(data_m) if data_m else default_material_db
        except:
            st.session_state.material_db = default_material_db
    else:
        st.session_state.material_db = default_material_db

if 'post_db' not in st.session_state:
    if gc:
        try:
            ws_p = gc.open(SHEET_NAME).worksheet("post_db")
            data_p = ws_p.get_all_records()
            st.session_state.post_db = pd.DataFrame(data_p) if data_p else default_post_db
        except:
            st.session_state.post_db = default_post_db
    else:
        st.session_state.post_db = default_post_db

if 'parsed_df' not in st.session_state:
    st.session_state.parsed_df = pd.DataFrame()
if 'uploaded_file_names' not in st.session_state:
    st.session_state.uploaded_file_names = []

with st.expander("📊 1. 기준 단가표 관리 (클릭하여 펼치기)"):
    col1, col2 = st.columns(2)
    with col1:
        edited_material = st.data_editor(st.session_state.material_db, num_rows="dynamic", use_container_width=True)
    with col2:
        edited_post = st.data_editor(st.session_state.post_db, num_rows="dynamic", use_container_width=True)
    
    if st.button("💾 변경된 단가표 영구 저장하기 (구글 시트 연동)"):
        st.session_state.material_db = edited_material
        st.session_state.post_db = edited_post
        if gc:
            try:
                sh = gc.open(SHEET_NAME)
                ws_m = sh.worksheet("material_db")
                ws_m.clear()
                ws_m.update([edited_material.columns.values.tolist()] + edited_material.astype(str).values.tolist())
                
                ws_p = sh.worksheet("post_db")
                ws_p.clear()
                ws_p.update([edited_post.columns.values.tolist()] + edited_post.astype(str).values.tolist())
                st.success("✅ 구글 시트에 단가표가 동기화되었습니다!")
            except Exception as e:
                st.error(f"⚠️ 저장 실패: {e}")

st.markdown("---")

# =========================================================================
# 📸 2. DXF를 이미지(PIL Image)로 변환하는 함수
# =========================================================================
def dxf_to_image(doc):
    try:
        msp = doc.modelspace()
        fig = plt.figure(figsize=(12, 9), dpi=150) # 해상도를 높여 선명하게 캡처
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis('off') # 테두리 제거
        
        ctx = RenderContext(doc)
        out = MatplotlibBackend(ax)
        Frontend(ctx, out).draw_layout(msp, finalize=True)
        
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches='tight', pad_inches=0)
        plt.close(fig) # 메모리 누수 방지
        
        buf.seek(0)
        img = Image.open(buf) # AI가 가장 좋아하는 PIL 이미지 객체로 변환
        return img
    except Exception as e:
        st.error(f"⚠️ 도면 이미지 변환 실패: {e}")
        return None

# =========================================================================
# 🤖 진짜 AI (Gemini) 비전+텍스트 하이브리드 파싱 함수
# =========================================================================
def analyze_with_hybrid_gemini(filename, text_data, geometry_info, img_obj, api_key):
    genai.configure(api_key=api_key)
    
    # 이미지 분석에 가장 뛰어난 최신 1.5 플래시 모델 사용
    target_model_name = "gemini-1.5-flash"
    
    try:
        model = genai.GenerativeModel(target_model_name)
        
        prompt = f"""
        당신은 대한민국 최고 수준의 2D 가공 도면 해독 전문가입니다.
        이번엔 특별히 **도면 캡처 이미지**와 도면에서 추출한 **텍스트 데이터**를 동시에 제공합니다.

        [도면 해독 핵심 지침]
        1. 제공된 '도면 이미지'를 눈으로 직접 확인하여 표제란(부품표)의 위치와 형상의 가공 난이도를 파악하세요.
        2. 이미지 내의 한글이 네모(ㅁㅁㅁ)로 깨져 보일 수 있습니다. 당황하지 말고 제공된 [추출된 텍스트]에서 알맞은 글자를 찾아 짝을 맞추세요. (예: 표에 'ㅁㅁ'이 있고 텍스트에 'MC'가 있다면 재질은 'MC'입니다.)
        3. 표제란을 눈으로 찾아 재질, 규격(크기), 수량을 우선적으로 매핑하세요.
        4. SPEC이나 규격 칸에 적힌 "숫자X숫자X숫자" (예: 35X130X360) 패턴을 찾으면 가로, 세로, 두께로 완벽히 매핑하세요.
        5. 수량(Q'TY, 수량 등)을 정확히 찾아 숫자로 적으세요.

        [추출된 기하 정보]
        {geometry_info}
        
        [추출된 텍스트 (정확한 글자와 숫자)]
        {text_data}

        이 모든 시각 정보와 텍스트 문맥을 종합하여, 오직 아래 JSON 형식으로만 대답하세요.
        {{
            "도면번호": "문자열 (DWG.NO)",
            "품명": "문자열 (TITLE)",
            "재질": "문자열 (MC, SUS304 등)",
            "수량": 정수,
            "가로": 숫자,
            "세로": 숫자,
            "두께": 숫자,
            "후처리": "문자열 (없으면 '없음')",
            "비고": "도면 가공 특징 및 특이사항 요약 (이미지로 본 형상 난이도 필수 포함)"
        }}
        """
        
        # 💡 프롬프트와 이미지를 동시에 던지는 하이브리드 리스트 구성!
        contents = [prompt]
        if img_obj is not None:
            contents.append(img_obj)
            
        response = model.generate_content(contents)
        result_text = response.text.strip()
        
        if result_text.startswith("```json"):
            result_text = result_text[7:-3].strip()
        elif result_text.startswith("```"):
            result_text = result_text[3:-3].strip()
            
        return json.loads(result_text)
        
    except Exception as e:
        return {"도면번호": filename, "품명": "분석 실패", "재질": "미정", "수량": 1, "가로": 0, "세로": 0, "두께": 0, "후처리": "없음", "비고": f"AI 에러: {e}"}

# =========================================================================
# 2. DXF 업로드 및 실행 로직
# =========================================================================
st.subheader("2. DXF 도면 업로드 및 AI 비전 분석")
uploaded_files = st.file_uploader("📂 DXF 도면을 올려주세요. AI가 눈으로 도면을 직접 확인합니다.", type=['dxf'], accept_multiple_files=True)

if uploaded_files:
    if not api_key:
        st.warning("👈 왼쪽 사이드바에 Gemini API Key를 먼저 입력해 주세요!")
    else:
        current_file_names = [f.name for f in uploaded_files]
        
        if st.session_state.uploaded_file_names != current_file_names:
            parsed_results = []
            with st.spinner("📸 AI가 도면을 사진으로 찍어 텍스트와 함께 정밀 교차 검증 중입니다... (1장당 약 5~10초 소요)"):
                for idx, file in enumerate(uploaded_files):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp_file:
                        tmp_file.write(file.getvalue())
                        tmp_path = tmp_file.name
                    
                    try:
                        doc = ezdxf.readfile(tmp_path)
                        msp = doc.modelspace()
                        
                        # 1. 텍스트 추출 (100% 정확한 글자 및 숫자)
                        extracted_texts = [e.dxf.text for e in msp.query('TEXT MTEXT') if hasattr(e.dxf, 'text') and e.dxf.text]
                        clean_texts = " | ".join([t.strip() for t in extracted_texts if t.strip()])
                        
                        num_tols = sum(1 for t in extracted_texts if any(k in t for k in ['±', '%%p', '+', '-', 'H7', 'h7']))
                        num_holes = len(msp.query('CIRCLE'))
                        num_dims = len(msp.query('DIMENSION'))
                        geometry_info = f"원(구멍) 갯수: {num_holes}개, 치수기입 갯수: {num_dims}개, 공차추정: {num_tols}건"
                        
                        # 2. 도면을 이미지로 변환 (AI의 시각)
                        img_obj = dxf_to_image(doc)
                        
                        # 3. 하이브리드 분석 투척!
                        ai_result = analyze_with_hybrid_gemini(file.name, clean_texts, geometry_info, img_obj, api_key)
                        
                        # 데이터 후처리 및 정제
                        w = safe_float(ai_result.get("가로", 0))
                        h = safe_float(ai_result.get("세로", 0))
                        t = safe_float(ai_result.get("두께", 0))
                        qty = ai_result.get("수량", 1)
                        if not isinstance(qty, int): 
                            qty = 1
                        
                        ai_result["가로"], ai_result["세로"], ai_result["두께"], ai_result["수량"] = w, h, t, qty
                        
                        mat_name = str(ai_result.get("재질", "미정"))
                        post_name = str(ai_result.get("후처리", "없음"))
                        
                        # 💡 재질 단가표 스마트 매핑 ('MC'라고만 해도 'MC 나이론'을 찾아냄)
                        mat_info = pd.DataFrame()
                        if mat_name and mat_name != "미정":
                            mask = st.session_state.material_db['재질'].astype(str).str.lower().str.contains(mat_name.lower(), na=False)
                            matches = st.session_state.material_db[mask]
                            if not matches.empty:
                                mat_info = matches.iloc[[0]]
                            else:
                                mat_info = st.session_state.material_db[st.session_state.material_db['재질'] == mat_name]

                        if not mat_info.empty:
                            weight_ratio = mat_info['비중'].values[0]
                            mat_price_per_kg = mat_info['KG당 단가'].values[0]
                            weight = (w * h * t) * weight_ratio / 1000000 
                            ai_result["소재비"] = int(weight * mat_price_per_kg)
                        else: 
                            ai_result["소재비"] = 0
                        
                        post_info = st.session_state.post_db[st.session_state.post_db['표면처리'] == post_name]
                        if not post_info.empty:
                            post_price_per_kg = post_info['KG당 단가'].values[0]
                            ai_result["후처리비"] = int(weight * post_price_per_kg) if 'weight' in locals() else 0
                        else: 
                            ai_result["후처리비"] = 0
                        
                        ai_result["가공비(수동입력)"] = 0 
                        ai_result["최종합계"] = ai_result["소재비"] + ai_result["후처리비"]
                        
                        parsed_results.append(ai_result)
                    
                    except Exception as e:
                        st.error(f"{file.name} 처리 중 오류: {e}")
                    finally:
                        os.remove(tmp_path)
                    
                    if idx < len(uploaded_files) - 1: 
                        time.sleep(3) 
            
            st.session_state.parsed_df = pd.DataFrame(parsed_results)
            st.session_state.uploaded_file_names = current_file_names

        st.success("✅ AI 비전(눈)과 텍스트(두뇌)를 활용한 하이브리드 분석이 완료되었습니다!")

        # 💡 과거 이력 조회
        if gc and not st.session_state.parsed_df.empty:
            try:
                history_db = pd.DataFrame(gc.open(SHEET_NAME).worksheet("Quote_Database").get_all_records())
                if not history_db.empty:
                    for idx, row in st.session_state.parsed_df.iterrows():
                        drw_no = str(row.get('도면번호', ''))
                        if not drw_no: continue
                        matches = history_db[history_db['도면번호'].astype(str) == drw_no]
                        if not matches.empty:
                            last_quote = matches.iloc[-1]
                            st.warning(f"🕒 **과거 이력 발견!** [{drw_no}] 👉 가공비 {last_quote.get('가공비(수동입력)', 0):,}원 / 총액 {last_quote.get('최종합계', 0):,}원")
            except Exception as e: 
                pass 

        if not st.session_state.parsed_df.empty:
            st.markdown("---")
            st.subheader("3. 📝 최종 견적 검토 및 데이터 수정")
            
            edited_df = st.data_editor(st.session_state.parsed_df, disabled=["최종합계"], hide_index=True, use_container_width=True, key="quote_editor")
            
            final_df = edited_df.copy()
            final_df["최종합계"] = final_df["소재비"] + final_df["후처리비"] + final_df["가공비(수동입력)"]
            
            # 수량을 곱한 전체 프로젝트 총액 계산
            total_sum = sum(final_df["최종합계"] * final_df["수량"])
            st.markdown(f"### 💰 전체 프로젝트 총 견적액 (수량 반영): **{total_sum:,} 원**")

            st.markdown("---")
            st.subheader("4. 💾 견적 확정 및 엑셀 다운로드")
            
            if st.button("🚀 견적 확정 및 엑셀 폼 발행하기"):
                if gc:
                    try:
                        ws_q = gc.open(SHEET_NAME).worksheet("Quote_Database")
                        data_q = ws_q.get_all_values()
                        if not data_q: 
                            ws_q.update([final_df.columns.values.tolist()] + final_df.astype(str).values.tolist())
                        else: 
                            ws_q.append_rows(final_df.astype(str).values.tolist())
                        st.success(f"✅ 구글 시트 DB 누적 완료!")
                    except Exception as e: 
                        st.error(f"⚠️ 저장 실패: {e}")
                
                try:
                    wb = openpyxl.load_workbook("견적서.xlsx")
                    ws = wb["견적서(을지)"] if "견적서(을지)" in wb.sheetnames else wb.active
                    
                    start_row = 7
                    for index, row in final_df.iterrows():
                        current_row = start_row + index
                        qty = int(row['수량'])
                        
                        ws.cell(row=current_row, column=1).value = index + 1
                        ws.cell(row=current_row, column=2).value = row['도면번호']
                        ws.cell(row=current_row, column=3).value = row['품명']
                        ws.cell(row=current_row, column=4).value = f"{row['가로']} x {row['세로']} x {row['두께']}"
                        ws.cell(row=current_row, column=6).value = row['후처리']
                        ws.cell(row=current_row, column=7).value = qty
                        
                        # 엑셀 기입 시 개당 단가에 수량을 곱해서 입력
                        ws.cell(row=current_row, column=8).value = int(row['소재비']) * qty
                        ws.cell(row=current_row, column=9).value = int(row['가공비(수동입력)']) * qty
                        ws.cell(row=current_row, column=10).value = int(row['후처리비']) * qty
                        ws.cell(row=current_row, column=16).value = row['비고']
                        
                    output = BytesIO()
                    wb.save(output)
                    st.download_button(label="📊 회사 양식 최종 엑셀 다운로드 (.xlsx)", data=output.getvalue(), file_name="최종견적서_발행.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e: 
                    st.error(f"⚠️ 엑셀 템플릿 처리 중 오류: {e}")
