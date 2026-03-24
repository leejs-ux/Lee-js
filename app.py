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
import matplotlib.font_manager as fm

# 한글 폰트 깨짐 방지 설정 (Streamlit Cloud 환경 고려)
font_list = [f.name for f in fm.fontManager.ttflist]
if 'NanumGothic' in font_list:
    plt.rc('font', family='NanumGothic')
elif 'Malgun Gothic' in font_list:
    plt.rc('font', family='Malgun Gothic')
plt.rcParams['axes.unicode_minus'] = False 

from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from PIL import Image

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
        fig = plt.figure(figsize=(12, 9), dpi=200)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis('off') 
        
        ctx = RenderContext(doc)
        out = MatplotlibBackend(ax)
        Frontend(ctx, out).draw_layout(msp, finalize=True)
        
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches='tight', pad_inches=0)
        plt.close(fig) 
        
        buf.seek(0)
        img = Image.open(buf)
        return img
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
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        flash_models = [name for name in available_models if 'flash' in name.lower()]
        if flash_models:
            target_model_name = flash_models[0]
    except Exception as e:
        pass
    
    try:
        model = genai.GenerativeModel(target_model_name)
        
        prompt = f"""
        당신은 대한민국 최고 수준의 2D 가공 도면(DXF) 해독 수석 엔지니어입니다.
        제공된 **도면 이미지**와 추출된 **텍스트 데이터**를 조합하여 완벽한 견적 데이터를 산출하세요.

        [도면 해독 핵심 지침]
        1. **표제란(Title Block) 탐색:** 주로 우측 하단이나 상단에 위치한 표(Table)를 눈으로 먼저 찾으세요. DWG NO, TITLE, MAT'L(재질), Q'TY(수량)가 적혀있습니다.
        2. **텍스트 매핑 교차검증 (중요):** 이미지 내 한글이나 폰트가 네모(ㅁㅁ)로 깨져 보일 수 있습니다. 이 때는 무조건 제공된 [추출된 텍스트]에서 해당 위치에 들어갈 법한 단어를 유추하여 짝을 맞추세요.
        3. **규격(가로x세로x두께) 파악:** - 표제란의 SPEC이나 SIZE 칸에 "숫자X숫자X숫자" (예: 35X130X360) 패턴이 있다면 이를 우선시 하세요.
           - 만약 표제란에 규격이 없다면, 도면 형상에 기입된 가장 긴 치수들을 파악하여 외형 최대 사이즈(가로, 세로, 두께)를 직접 추론하세요.
        4. **가공 난이도 평가:** 제공된 [기하 정보] (구멍 개수, 공차, 치수 기입 개수)와 형상을 보고 "비고" 란에 가공 특이사항(예: 탭
