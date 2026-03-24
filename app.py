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

# [서버 다운 방지] 필수 설정
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from PIL import Image

st.set_page_config(page_title="2D DXF 하이브리드 자동 견적 시스템", page_icon="👁️", layout="wide")

# =========================================================================
# 💡 구글 시트 연결
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
# 1. 기준 단가표 관리
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
        except: st.session_state.material_db = default_material_db
    else: st.session_state.material_db = default_material_db

if 'post_db' not in st.session_state:
    if gc:
        try:
            ws_p = gc.open(SHEET_NAME).worksheet("post_db")
            data_p = ws_p.get_all_records()
            st.session_state.post_db = pd.DataFrame(data_p) if data_p else default_post_db
        except: st.session_state.post_db = default_post_db
    else: st.session_state.post_db = default_post_db

if 'parsed_df' not in st.session_state: st.session_state.parsed_df = pd.DataFrame()
if 'uploaded_file_names' not in st.session_state: st.session_state.uploaded_file_names = []

with st.expander("📊 1. 기준 단가표 관리 (클릭하여 펼치기)"):
    col1, col2 = st.columns(2)
    with col1: edited_material = st.data_editor(st.session_state.material_db, num_rows="dynamic", use_container_width=True)
    with col2: edited_post = st.data_editor(st.session_state.post_db, num_rows="dynamic", use_container_width=True)
    
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
            except Exception as e: st.error(f"⚠️ 저장 실패: {e}")

st.markdown("---")

# =========================================================================
# 📸 2. DXF를 이미지(PIL Image)로 변환하는 함수
# =========================================================================
def dxf_to_image(doc):
    try:
        msp = doc.modelspace()
        fig = plt.figure(figsize=(12, 9), dpi=150)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis('off')
        ctx = RenderContext(doc)
        out = MatplotlibBackend(ax)
        Frontend(ctx, out).draw_layout(msp, finalize=True)
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches='tight', pad_inches=0)
        plt.close(fig) 
        buf.seek(0)
        return Image.open(buf)
    except Exception as e:
        st.error(f"⚠️ 도면 이미지 변환 실패: {e}")
        return None

# =========================================================================
# 🤖 진짜 AI (Gemini) 비전+텍스트 하이브리드 파싱 함수
# =========================================================================
def analyze_with_hybrid_gemini(filename, text_data, geometry_info, img_obj, api_key):
    genai.configure(api_key=api_key)
    target_model_name = "gemini-1.5-flash"
    
    try:
        available_models = genai.list_models()
        for m in available_models:
            if 'generateContent' in m.supported_generation_methods and 'gemini' in m.name.lower() and 'flash' in m.name.lower():
                target_model_name = m.name
                break
    except: pass
    
    try:
        model = genai.GenerativeModel(target_model_name)
        prompt = f"""
        당신은 대한민국 최고 수준의 기계 가공 도면 해독 및 견적 산출 전문가입니다.
        제공된 **도면 캡처 이미지**와 **추출된 텍스트 데이터**를 종합하여 아래 임무를 완수하세요.

        [도면 해독 지침]
        1. 시각적 유추: 도면 형상을 보고 어떤 가공이 주를 이루는지 파악하세요. 
           - 원통형/회전체면 '선반', 각형태/포켓/구멍이 많으면 '밀링(MCT)', 얇은 판재를 자르는 모양이면 '레이저', 절곡 기호가 있으면 '판금', 용접 기호가 보이면 '용접'으로 추론하세요.
        2. 시간 추론: 형상의 복잡도, 구멍의 갯수, 공차 기호 유무를 파악하여 '예상 가공 시간'을 추론하세요. (예: "약 2시간", "약 30분" 등)
        3. 이미지 속 깨진 한글(ㅁㅁㅁ)은 [추출된 텍스트]에서 알맞은 글자를 찾아 매핑하세요. 표제란을 우선 탐색하여 재질, 규격(가로x세로x두께), 수량을 찾으세요.

        [추출된 기하 정보]
        {geometry_info}
        
        [추출된 텍스트]
        {text_data}

        이 모든 정보를 종합하여, 오직 아래 JSON 형식으로만 대답하세요. 절대 다른 설명은 붙이지 마세요.
        {{
            "도면번호": "문자열 (DWG.NO)",
            "품명": "문자열 (TITLE)",
            "재질": "문자열 (예: MC, SUS304 등)",
            "수량": 정수,
            "가로": 숫자,
            "세로": 숫자,
            "두께": 숫자,
            "후처리": "문자열 (없으면 '없음')",
            "가공방법": "문자열 (밀링, 선반, 레이저, 판금, 용접 중 형상에 맞는 것을 1개 이상 쉼표로 구분하여 적으세요)",
            "예상가공시간": "문자열 (도면 난이도를 바탕으로 추정한 예상 시간. 예: 약 1시간 30분)",
            "비고": "도면 가공 특징, 공차 유무, 특이사항 요약"
        }}
        """
        
        contents = [prompt]
        if img_obj is not None: 
            contents.append(img_obj)
            
        response = model.generate_content(contents)
        result_text = response.text.strip()
        
        # 💡 [핵심 수정 부분] 따옴표 에러 방지를 위해 작은따옴표(')로 변경하고 명확하게 분리했습니다!
        if result_text.startswith('```json'):
            result_text = result_text[7:-3].strip()
        elif result_text.startswith('```'):
            result_text = result_text[3:-3].strip()
            
        return json.loads(result_text)
        
    except Exception as e:
        return {"도면번호": filename, "품명": "분석 실패", "재질": "미정", "수량": 1, "가로": 0, "세로": 0, "두께": 0, "후처리": "없음", "가공방법": "알수없음", "예상가공시간": "알수없음", "비고": f"AI 에러: {e}"}

# =========================================================================
# 2. DXF 업로드 및 실행 로직
# =========================================================================
st.subheader("2. DXF 도면 업로드 및 AI 비전 분석")
uploaded_files = st.file_uploader("📂 DXF 도면을 올려주세요. AI가 눈으로 도면을 분석합니다.", type=['dxf'], accept_multiple_files=True)

if uploaded_files:
    if not api_key:
        st.warning("👈 왼쪽 사이드바에 Gemini API Key를 먼저 입력해 주세요!")
    else:
        current_file_names = [f.name for f in uploaded_files]
        if st.session_state.uploaded_file_names != current_file_names:
            parsed_results = []
            with st.spinner("📸 AI 전문가가 도면 형상을 확인하고 가공 방법과 소요 시간까지 추론 중입니다... (1장당 5~10초 소요)"):
                for idx, file in enumerate(uploaded_files):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp_file:
                        tmp_file.write(file.getvalue())
                        tmp_path = tmp_file.name
                    
                    try:
                        doc = ezdxf.readfile(tmp_path)
                        msp = doc.modelspace()
                        
                        extracted_texts = [e.dxf.text for e in msp.query('TEXT MTEXT') if hasattr(e.dxf, 'text') and e.dxf.text]
                        clean_texts = " | ".join([t.strip() for t in extracted_texts if t.strip()])
                        
                        num_tols = sum(1 for t in extracted_texts if any(k in t for k in ['±', '%%p', '+', '-', 'H7', 'h7']))
                        num_holes = len(msp.query('CIRCLE'))
                        num_dims = len(msp.query('DIMENSION'))
                        geometry_info = f"원(구멍) 갯수: {num_holes}개, 치수기입 갯수: {num_dims}개, 공차추정: {num_tols}건"
                        
                        img_obj = dxf_to_image(doc)
                        ai_result = analyze_with_hybrid_gemini(file.name, clean_texts, geometry_info, img_obj, api_key)
                        
                        w = safe_float(ai_result.get("가로", 0))
                        h = safe_float(ai_result.get("세로", 0))
                        t = safe_float(ai_result.get("두께", 0))
                        qty = ai_result.get("수량", 1)
                        if not isinstance(qty, int): 
                            qty = 1
                        
                        ai_result["가로"], ai_result["세로"], ai_result["두께"], ai_result["수량"] = w, h, t, qty
                        
                        ai_result["가공방법"] = str(ai_result.get("가공방법", "분석 불가"))
                        ai_result["예상가공시간"] = str(ai_result.get("예상가공시간", "분석 불가"))
                        
                        mat_name = str(ai_result.get("재질", "미정"))
                        post_name = str(ai_result.get("후처리", "없음"))
                        
                        # [스마트 매핑 로직]
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

        st.success("✅ 가공 방법 및 예상 소요 시간 분석이 완료되었습니다!")

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
                            st.warning(f"🕒 **과거 이력 발견!** [{drw_no}] 👉 기존 가공비 {last_quote.get('가공비(수동입력)', 0):,}원")
            except: 
                pass 

        if not st.session_state.parsed_df.empty:
            st.markdown("---")
            st.subheader("3. 📝 최종 견적 검토 및 데이터 수정")
            
            edited_df = st.data_editor(st.session_state.parsed_df, disabled=["최종합계", "가공방법", "예상가공시간"], hide_index=True, use_container_width=True, key="quote_editor")
            
            final_df = edited_df.copy()
            final_df["최종합계"] = final_df["소재비"] + final_df["후처리비"] + final_df["가공비(수동입력)"]
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
                        ws.cell(row=current_row, column=8).value = int(row['소재비']) * qty
                        ws.cell(row=current_row, column=9).value = int(row['가공비(수동입력)']) * qty
                        ws.cell(row=current_row, column=10).value = int(row['후처리비']) * qty
                        
                        combined_remarks = f"[{row['가공방법']} / {row['예상가공시간']}] {row['비고']}"
                        ws.cell(row=current_row, column=16).value = combined_remarks
                        
                    output = BytesIO()
                    wb.save(output)
                    st.download_button(label="📊 회사 양식 최종 엑셀 다운로드 (.xlsx)", data=output.getvalue(), file_name="최종견적서_발행.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e: 
                    st.error(f"⚠️ 엑셀 템플릿 처리 중 오류: {e}")
