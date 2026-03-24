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
# 💡 실시간 재계산 함수
# =========================================================================
def recalculate_costs(df, mat_db, post_db):
    for idx, row in df.iterrows():
        w = safe_float(row.get("가로", 0))
        h = safe_float(row.get("세로", 0))
        t = safe_float(row.get("두께", 0))
        mat_name = str(row.get("재질", "미정")).strip()
        post_name = str(row.get("후처리", "없음")).strip()
        
        # 재질 매핑 (AI가 이미 잘 맞춰주겠지만, 혹시 모를 오차를 위해 이중 확인)
        mat_info = pd.DataFrame()
        if mat_name and mat_name != "미정":
            mask = mat_db['재질'].astype(str).str.lower().str.contains(mat_name.lower(), na=False)
            matches = mat_db[mask]
            if not matches.empty: 
                mat_info = matches.iloc[[0]]
            else: 
                mat_info = mat_db[mat_db['재질'].astype(str).str.strip() == mat_name]

        weight = 0
        if not mat_info.empty:
            weight_ratio = safe_float(mat_info['비중'].values[0])
            mat_price_per_kg = safe_float(mat_info['KG당 단가'].values[0])
            weight = (w * h * t) * weight_ratio / 1000000 
            df.at[idx, "소재비"] = int(weight * mat_price_per_kg)
        else: 
            df.at[idx, "소재비"] = 0
        
        # 후처리 매핑
        post_info = post_db[post_db['표면처리'].astype(str).str.strip() == post_name]
        if not post_info.empty:
            post_price_per_kg = safe_float(post_info['KG당 단가'].values[0])
            df.at[idx, "후처리비"] = int(weight * post_price_per_kg)
        else: 
            df.at[idx, "후처리비"] = 0
            
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

# 💡 [핵심 강화] DB 목록을 AI에게 전달하는 매개변수 추가
def analyze_with_hybrid_gemini(filename, text_data, geometry_info, img_obj, api_key, available_materials, available_posts):
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

        [우리 회사 DB 보유 목록]
        - 보유 재질: {available_materials}
        - 보유 후처리: {available_posts}

        [도면 해독 지침]
        1. 시각적 유추: 도면 형상을 보고 어떤 가공이 주를 이루는지 파악하세요 (밀링, 선반, 레이저, 판금, 용접 중 택).
        2. 시간 추론: 형상의 복잡도를 파악하여 '예상 가공 시간'을 추론하세요.
        3. [매우 중요] 재질 및 후처리 매핑:
           - 도면에 적힌 재질/후처리가 [우리 회사 DB 보유 목록]에 있는 항목의 동의어거나 같은 종류라면(예: 'MC흑색'->'MC 나이론'), 반드시 **DB에 등록된 정확한 명칭**으로 통일해서 적으세요.
           - 단, DB 목록에 전혀 없는 새로운 재질이거나 확신할 수 없다면, 억지로 DB 목록에 맞추지 말고 **도면에 적힌 원본 글자 그대로** 적으세요.
        4. 표제란을 우선 탐색하여 규격(숫자X숫자X숫자 패턴)과 수량을 매핑하세요.
        5. '비고' 란을 구체적으로 작성하세요. [추출된 기하 정보]를 바탕으로 홀/치수/공차 개수를 명시하고, 형상을 바탕으로 가공 특이사항과 주의할 점을 꼼꼼히 적으세요.

        [추출된 기하 정보]
        {geometry_info}
        
        [추출된 텍스트]
        {text_data}

        오직 아래 JSON 형식으로만 대답하세요. 절대 다른 설명은 붙이지 마세요.
        {{
            "도면번호": "문자열 (DWG.NO)",
            "품명": "문자열 (TITLE)",
            "재질": "문자열 (DB명칭 변환 혹은 도면 원본)",
            "수량": 정수,
            "가로": 숫자,
            "세로": 숫자,
            "두께": 숫자,
            "후처리": "문자열",
            "가공방법": "문자열",
            "예상가공시간": "문자열",
            "비고": "문자열 (예시: '▶특이사항: 탭 가공 및 깊은 포켓 존재 ▶주의사항: H7 끼워맞춤 공차 주의 ▶분석정보: 홀 5개, 공차 2건, 치수 15개')"
        }}
        """
        
        contents = [prompt]
        if img_obj is not None: 
            contents.append(img_obj)
            
        response = model.generate_content(contents)
        result_text = response.text.strip()
        
        bt = chr(96) * 3 
        if result_text.startswith(bt + "json"): 
            result_text = result_text[7:-3].strip()
        elif result_text.startswith(bt): 
            result_text = result_text[3:-3].strip()
            
        return json.loads(result_text)
    except Exception as e:
        return {"도면번호": filename, "품명": "분석 실패", "재질": "미정", "수량": 1, "가로": 0, "세로": 0, "두께": 0, "후처리": "없음", "가공방법": "알수없음", "예상가공시간": "알수없음", "비고": f"AI 에러: {e}"}

# =========================================================================
# 2. DXF 업로드 및 스마트 파일 처리 로직
# =========================================================================
st.subheader("2. DXF 도면 업로드 및 AI 비전 분석")
uploaded_files = st.file_uploader("📂 DXF 도면을 올려주세요.", type=['dxf'], accept_multiple_files=True)

if uploaded_files:
    if not api_key:
        st.warning("👈 왼쪽 사이드바에 Gemini API Key를 먼저 입력해 주세요!")
    else:
        current_file_names = [f.name for f in uploaded_files]
        old_file_names = st.session_state.uploaded_file_names
        
        if old_file_names != current_file_names:
            added_files = [f for f in uploaded_files if f.name not in old_file_names]
            removed_file_names = [name for name in old_file_names if name not in current_file_names]
            
            if removed_file_names:
                if not st.session_state.parsed_df.empty:
                    st.session_state.parsed_df = st.session_state.parsed_df[~st.session_state.parsed_df['도면번호'].isin(removed_file_names)]
                    st.toast("🗑️ 선택한 도면이 표에서 즉시 삭제되었습니다.")
            
            if added_files:
                new_parsed_results = []
                # 💡 [핵심] 현재 DB에 있는 재질/후처리 목록을 문자열로 예쁘게 뽑아서 AI에게 넘길 준비를 합니다.
                db_materials_str = ", ".join(st.session_state.material_db['재질'].astype(str).tolist())
                db_posts_str = ", ".join(st.session_state.post_db['표면처리'].astype(str).tolist())

                with st.spinner(f"📸 추가된 도면({len(added_files)}장)만 AI가 꼼꼼하게 분석 중입니다..."):
                    for idx, file in enumerate(added_files):
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
                            geometry_info = f"원 갯수: {num_holes}개, 치수 갯수: {num_dims}개, 공차: {num_tols}건"
                            
                            img_obj = dxf_to_image(doc)
                            # 💡 함수 호출 시 DB 메뉴판 2개(재질, 후처리)를 같이 던져줍니다!
                            ai_result = analyze_with_hybrid_gemini(file.name, clean_texts, geometry_info, img_obj, api_key, db_materials_str, db_posts_str)
                            
                            ai_result["가로"] = safe_float(ai_result.get("가로", 0))
                            ai_result["세로"] = safe_float(ai_result.get("세로", 0))
                            ai_result["두께"] = safe_float(ai_result.get("두께", 0))
                            qty = ai_result.get("수량", 1)
                            ai_result["수량"] = qty if isinstance(qty, int) else 1
                            ai_result["가공비(수동입력)"] = 0
                            
                            new_parsed_results.append(ai_result)
                        
                        except Exception as e: 
                            st.error(f"{file.name} 처리 중 오류: {e}")
                        finally: 
                            os.remove(tmp_path)
                            
                        if idx < len(added_files) - 1: 
                            time.sleep(3) 
                
                if new_parsed_results:
                    temp_df = pd.DataFrame(new_parsed_results)
                    temp_df = recalculate_costs(temp_df, st.session_state.material_db, st.session_state.post_db)
                    
                    col_order = ["도면번호", "품명", "재질", "수량", "가로", "세로", "두께", "후처리", 
                                 "소재비", "후처리비", "가공비(수동입력)", "최종합계", 
                                 "가공방법", "예상가공시간", "비고"]
                    
                    for c in col_order:
                        if c not in temp_df.columns:
                            temp_df[c] = 0 if "비" in c else ""
                    temp_df = temp_df[col_order]
                    
                    if st.session_state.parsed_df.empty:
                        st.session_state.parsed_df = temp_df
                    else:
                        st.session_state.parsed_df = pd.concat([st.session_state.parsed_df, temp_df], ignore_index=True)
                        
            st.session_state.uploaded_file_names = current_file_names
            st.rerun()

else:
    if st.session_state.uploaded_file_names:
        st.session_state.uploaded_file_names = []
        st.session_state.parsed_df = pd.DataFrame()
        st.rerun()

# =========================================================================
# 3. 데이터 검토 및 엑셀 다운로드 
# =========================================================================
if not st.session_state.parsed_df.empty:
    st.markdown("---")
    st.subheader("3. 📝 최종 견적 검토 및 데이터 수정 (실시간 자동 계산)")
    st.info("💡 '가로/세로/두께' 또는 '재질'을 직접 수정해 보세요! **소재비와 최종합계가 즉시 재계산됩니다!**")
    
    edited_df = st.data_editor(
        st.session_state.parsed_df, 
        disabled=["소재비", "후처리비", "최종합계", "가공방법", "예상가공시간"], 
        hide_index=True, 
        use_container_width=True, 
        key="quote_editor"
    )
    
    final_df = recalculate_costs(edited_df.copy(), st.session_state.material_db, st.session_state.post_db)
    
    if not final_df.equals(st.session_state.parsed_df):
        st.session_state.parsed_df = final_df
        st.rerun() 
    
    total_sum = sum(final_df["최종합계"] * final_df["수량"])
    st.markdown(f"### 💰 전체 프로젝트 총 견적액 (수량 반영): **{total_sum:,} 원**")

    st.markdown("---")
    st.subheader("4. 💾 견적 확정 및 엑셀 다운로드")
    
    st.session_state.final_df_to_save = final_df.copy()
    
    excel_data = None
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
            
            combined_remarks = f"[{row['가공방법']} / {row['예상가공시간']}]\n{row['비고']}"
            ws.cell(row=current_row, column=16).value = combined_remarks
            
        output = BytesIO()
        wb.save(output)
        excel_data = output.getvalue()
    except Exception as e: 
        st.error(f"⚠️ 엑셀 템플릿 처리 중 오류 (견적서.xlsx 파일 확인 필요): {e}")

    def save_to_db_on_download():
        if gc and 'final_df_to_save' in st.session_state:
            try:
                sh = gc.open(SHEET_NAME)
                try:
                    ws_q = sh.worksheet("Quote_Database")
                except:
                    ws_q = sh.add_worksheet(title="Quote_Database", rows="1000", cols="30")
                
                df_to_save = st.session_state.final_df_to_save
                data_q = ws_q.get_all_values()
                
                if not data_q: 
                    ws_q.append_rows([df_to_save.columns.values.tolist()] + df_to_save.astype(str).values.tolist())
                else: 
                    ws_q.append_rows(df_to_save.astype(str).values.tolist())
                    
                st.session_state.db_save_success = True
            except Exception as e: 
                st.session_state.db_save_error = str(e)

    if excel_data:
        st.download_button(
            label="🚀 견적 확정 및 엑셀 폼 발행하기 (.xlsx)",
            data=excel_data,
            file_name="최종견적서_발행.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=save_to_db_on_download
        )
        
        if st.session_state.get('db_save_success'):
            st.success("✅ 구글 시트(Quote_Database)에 견적 데이터가 무사히 누적 저장되었습니다!")
            st.session_state.db_save_success = False 
        elif st.session_state.get('db_save_error'):
            st.error(f"⚠️ 구글 시트 저장 실패 (에러 메시지를 확인하세요): {st.session_state.db_save_error}")
            st.session_state.db_save_error = None
