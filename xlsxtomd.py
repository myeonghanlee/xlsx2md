import streamlit as st
import pandas as pd
import io
import zipfile

st.title("엑셀 ➡ Markdown 변환기 (전처리 선택형)")

# --- 학점 배당표 전용 데이터 정제 함수 ---
def preprocess_curriculum_data(file, sheet_name):
    df = pd.read_excel(file, sheet_name=sheet_name, skiprows=2, header=[0, 1])
    
    new_columns = []
    for col in df.columns:
        col_name = "_".join([str(c) for c in col if 'Unnamed' not in str(c)])
        new_columns.append(col_name)
    df.columns = new_columns
    
    # 열 개수가 충분할 때만 ffill 및 dropna 적용 (에러 방지)
    if len(df.columns) >= 4:
        cols_to_ffill = df.columns[:3] 
        df[cols_to_ffill] = df[cols_to_ffill].ffill()
        
        subject_col = df.columns[3] 
        df = df.dropna(subset=[subject_col])
    
    df = df.fillna("-")
    return df
# ------------------------

uploaded_files = st.file_uploader(
    "엑셀 파일을 업로드하세요", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True
)

if uploaded_files:
    # ----------------------------------------------------
    # 1. 단일 파일 업로드
    # ----------------------------------------------------
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        xls = pd.ExcelFile(file)
        
        selected_sheet = st.selectbox(f"📄 '{file.name}'에서 변환할 시트를 선택하세요:", xls.sheet_names)
        
        # 사용자가 전처리 여부를 직접 선택할 수 있도록 토글 추가
        use_preprocess = st.toggle("⚙️ 이 시트에는 '학점 배당표 전처리' 적용하기", value=False)
        
        try:
            # 토글 상태에 따라 다르게 읽어오기
            if use_preprocess:
                df = preprocess_curriculum_data(file, selected_sheet)
            else:
                df = pd.read_excel(file, sheet_name=selected_sheet)
            
            md_text = df.to_markdown(index=False)
            
            st.subheader(f"미리보기: [{selected_sheet}] 시트")
            st.code(md_text, language='markdown')
            
            st.download_button(
                label="Markdown 다운로드",
                data=md_text,
                file_name=f"{file.name.split('.')[0]}_{selected_sheet}.md",
                mime="text/markdown"
            )
            
        except Exception as e:
            st.error(f"데이터를 읽는 중 오류가 발생했습니다. 전처리 옵션을 변경해 보시거나 표 형식을 확인해주세요. (상세 에러: {e})")

    # ----------------------------------------------------
    # 2. 다중 파일 업로드
    # ----------------------------------------------------
    else:
        st.subheader(f"📁 총 {len(uploaded_files)}개의 파일 설정")
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for file in uploaded_files:
                xls = pd.ExcelFile(file)
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    selected_sheet = st.selectbox(
                        f"'{file.name}'의 시트 선택:", 
                        xls.sheet_names, 
                        key=f"sheet_{file.name}"
                    )
                with col2:
                    # 다중 파일일 때도 개별적으로 전처리 여부 선택 가능
                    use_preprocess = st.checkbox(
                        "학점 배당표 전처리", 
                        key=f"prep_{file.name}"
                    )
                
                try:
                    if use_preprocess:
                        df = preprocess_curriculum_data(file, selected_sheet)
                    else:
                        df = pd.read_excel(file, sheet_name=selected_sheet)
                        
                    md_text = df.to_markdown(index=False)
                    md_filename = f"{file.name.split('.')[0]}_{selected_sheet}.md"
                    zip_file.writestr(md_filename, md_text)
                    
                except Exception as e:
                    st.error(f"'{file.name}' 처리 중 오류 발생: {e}")
        
        st.success("변환 준비가 완료되었습니다! (오류가 난 파일은 제외됩니다)")
        st.download_button(
            label="모든 변환 파일(ZIP) 다운로드",
            data=zip_buffer.getvalue(),
            file_name="converted_markdown_files.zip",
            mime="application/zip"
        )
