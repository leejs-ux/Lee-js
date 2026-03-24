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
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

st.set_page_config(page_title="2D DXF 하이브리드 자동 견적 시스템", page_icon="👁️", layout="wide")

# =========================================================================
# 💡 구글 시트 연결 및 API 설정
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
        with open(API_KEY_FILE, "r") as f: return f.read().strip()
    return ""
def save_api_key(key):
    with open(API_KEY_FILE, "w") as f: f.write(key.strip())

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
# 1. 단가표 관리 로직 (생략 없이 원본 유지)
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
        st.session_state.material_db, st.session_state.post_db = edited_material, edited_post
        if gc:
            try:
                sh = gc.open(SHEET_NAME)
                ws_m = sh.worksheet("material_db")
                ws_m.clear()
                ws_m.update([edited_material.columns.values.tolist()] + edited_material.astype(str).values.tolist())
                ws_p = sh.worksheet("post_db")
                ws_p.clear()
                ws_p.update([edited_post.columns.values.tolist()] + edited_post.astype(str).values.tolist())
                st.success("✅ 단가표 동기화 완료!")
            except Exception as e: st.error(f"⚠️ 저장 실패: {e}")

st.markdown("---")

# =========================================================================
# 📸 2. DXF를 이미지로 변환하는 마법의 함수
# =========================================================================
def dxf_to_image_bytes(doc):
    try:
        msp = doc.modelspace()
        fig = plt.figure(figsize=(10, 8), dpi=150) # 해상도 최적화
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis('off') # 테두리 제거
        ctx = RenderContext(doc)
        out = MatplotlibBackend(ax)
        Frontend(ctx, out).draw_layout(msp, finalize=True)
        
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches='tight', pad_inches=0)
        plt.close(fig)
        return buf.getvalue()
    except Exception as e:
        print(f"이미지 변환 실패: {e}")
        return None

# =========================================================================
# 🤖 진짜 AI (Gemini) 비전+텍스트 하이브리드 파싱 함수
# =========================================================================
def analyze_with_hybrid_gemini(filename, text_data, geometry_info, image_bytes, api_key):
    genai.configure(api_key=api_key)
    target_model_name = "gemini-1.5-flash" # 이미지 처리에 최적화된 최신 모델 고정
    
    try:
        model = genai.GenerativeModel(target_model_name)
        
        prompt = f"""
        당신은 대한민국 최고 수준의 기계 가공 도면 해독 전문가입니다.
        이번엔 특별히 **도면 이미지**와 추출된 **텍스트 데이터**를 함께 제공합니다.

        [도면 해독 핵심 지침]
        1. 첨부된 '도면 이미지'를 눈으로 보고 전체적인 맥락(표제란의 위치, 외곽 형상, 가공 난이도)을 파악하세요. 
           (주의: 이미지 내의 한글 폰트가 네모로 깨져 보일 수 있습니다. 깨진 글자는 무시하고 위치만 파악하세요.)
        2. 깨진 글자의 진짜 내용은 아래 제공된 [추출된 텍스트]에서 찾아내어 짝을 맞추세요.
        3. 표제란(부품표)을 눈으로 찾아 재질, 규격(크기), 수량을 우선적으로 매핑하세요.
        4. SPEC이나 규격에 적힌 "숫자X숫자X숫자" (예: 35X130X360) 패턴은 가로, 세로, 두께로 완벽히 매핑하세요.
        5. 수량(Q'TY, 수량 등)을 정확히 찾아 정수로 적으세요. (기본값은 1)

        [추출된 기하 정보]
        {geometry_info}
        
        [추출된 텍스트 (정확한 글자와 숫자)]
        {text_data}

        이 모든 시각 정보와 텍스트 문맥을 종합하여, 오직 아래 JSON 형식으로만 완벽하게 대답하세요. 다른 말은 절대 금지.
        {{
            "도면번호": "문자열 (DWG.NO)",
            "품명": "문자열 (TITLE)",
            "재질": "문자열 (MC, SUS304 등)",
            "수량": 정수,
            "가로": 숫자,
            "세로": 숫자,
            "두께": 숫자,
            "후처리": "문자열 (없으면 '없음')",
            "비고": "가공 특징 및 특이사항 요약 (이미지에서 본 난이도 포함)"
        }}
        """
        
        # 💡 하이브리드 투척! (프롬프트 + 이미지)
        contents = [prompt]
        if image_bytes:
            contents.append({"mime_type": "image/png", "data": image_bytes})
            
        response = model.generate_content(contents)
        result_text = response.text.strip()
        
        if result_text.startswith("
http://googleusercontent.com/immersive_entry_chip/0
http://googleusercontent.com/immersive_entry_chip/1

3. 초록색 **[Commit changes...]** 를 두 번 눌러 저장해 줍니다.
4. 스트림릿 서버에 가셔서 **[Reboot app]** 을 누르시면, 새로운 그리기 도구(`matplotlib`)를 설치하느라 평소보다 시간이 조금 더 걸릴 것입니다!

재부팅이 완료되고, 아까 폼이 달라서 오류가 났던 문제의 **'V-Block' 도면**을 올려보세요. 
과연 AI가 도면 사진을 눈으로 직접 보고, 흩어진 텍스트를 끼워 맞춰 완벽하게 재질과 숫자를 추론해 내는지 확인해 보실 차례입니다! 결과가 어떻게 나오는지 궁금하네요, 꼭 알려주세요!
