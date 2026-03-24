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
        if isinstance(value, (int, float)): return float(value)
        num_str = re.sub(r'[^0-9.]', '', str(value))
        return float(num_str) if num_str else 0.0
    except: return 0.0

# =========================================================================
# 💡 [핵심] 엑셀처럼 금액을 실시간으로 재계산해 주는 마법의 함수!
# =========================================================================
def recalculate_costs(df, mat_db, post_db):
    for idx, row in df.iterrows():
        w = safe_float(row.get("가로", 0))
        h = safe_float(row.get("세로", 0))
        t = safe_float(row.get("두께", 0))
        mat_name = str(row.get("재질", "미정"))
        post_name = str(row.get("후처리", "없음"))
        
        # 재질 매핑
        mat_info = pd.DataFrame()
        if mat_name and mat_name != "미정":
            mask = mat_db['재질'].astype(str).str.lower().str.contains(mat_name.lower(), na=False)
            matches = mat_db[mask]
            if not matches.empty: 
                mat_info = matches.iloc[[0]]
            else: 
                mat_info = mat_db[mat_db['재질'] == mat_name]

        weight = 0
        if not mat_info.empty:
            weight_ratio = safe_float(mat_info['비중'].values[0])
            mat_price_per_kg = safe_float(mat_info['KG당 단가'].values[0])
            weight = (w * h * t) * weight_ratio / 1000000 
            df.at[idx, "소재비"] = int(weight * mat_price_per_kg)
        else: 
            df.at[idx, "소재비"] = 0
        
        # 후처리 매핑
        post_info = post_db[post_db['표면처리'] == post_name]
        if not post_info.empty:
            post_price_per_kg = safe_float(post_info['KG당 단가'].values[0])
            df.at[idx, "후처리비"] = int(weight * post_price_per_kg)
        else: 
            df.at[idx, "후처리비"] = 0
            
        # 최종 금액 합산
        manual_cost = int(safe_float(row.get("가공비(수동입력)", 0)))
        df.at[idx, "가공비(수동입력)"] = manual_cost
        df.at[idx, "최종합계"] = df.at[idx, "소재비"] + df.at[idx, "후처리비"] + manual_cost
    return df

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
                sh.worksheet("material_db").clear()
                sh.worksheet("material_db").update([edited_material.columns.values.tolist()] + edited_material.astype(str).values.tolist())
                sh.worksheet("post_db").clear()
                sh.worksheet("post_db").update([edited_post.columns.values.tolist()] + edited_post.astype(str).values.tolist())
                st.success("✅ 구글 시트에 단가표가 동기화되었습니다!")
            except Exception as e: st.error(f"⚠️ 저장 실패: {e}")

st.markdown("---")

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
        return None

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
        당신은 기계 가공 도면 해독 및 견적 산출 전문가입니다.
        제공된 **도면 캡처 이미지**와 **추출된 텍스트 데이터**를 종합하여 아래 임무를 완수하세요.

        1. 시각적 유추: 도면 형상을 보고 어떤 가공이 주를 이루는지 파악하세요 (밀링, 선반, 레이저, 판금, 용접 중 택).
        2. 시간 추론: 형상의 복잡도를 파악하여 '예상 가공 시간'을 추론하세요.
        3. 표제란을 우선 탐색하여 재질, 규격(숫자X숫자X숫자 패턴), 수량을 매핑하세요.

        [추출된 기하 정보]
        {geometry_info}
        
        [추출된 텍스트]
        {text_data}

        오직 아래 JSON 형식으로만 대답하세요. 절대 다른 설명은 붙이지 마세요.
        {{
            "도면번호": "문자열 (DWG.NO)",
            "품명": "문자열 (TITLE)",
            "재질": "문자열 (예: MC, SUS304 등)",
            "수량": 정수,
            "가로": 숫자,
            "세로": 숫자,
            "두께": 숫자,
            "후처리": "문자열",
            "가공방법": "문자열",
            "예상가공시간": "문자열",
            "비고": "문자열"
        }}
        """
        
        contents = [prompt]
        if img_obj is not None: 
            contents.append(img_obj)
            
        response = model.generate_content(contents)
        result_text = response.text.strip()
        
        if result_text.startswith('
http://googleusercontent.com/immersive_entry_chip/0
http://googleusercontent.com/immersive_entry_chip/1
