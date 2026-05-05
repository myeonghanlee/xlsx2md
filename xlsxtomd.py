import streamlit as st
import pandas as pd
import io
import zipfile

st.title("교육과정 엑셀 ➡ 데이터 분석용 Markdown 변환기")

# --- 데이터 정제 함수 ---
def preprocess_curriculum_data(file, sheet_name):
    """
    복잡한 교육과정 배당표 엑셀을 Markdown 변환에 최적화되게 정제합니다.
    """
    # 1. 상단 제목 2줄 스킵하고, 3~4번째 줄을 다중 헤더(2줄)로 읽어오기
    df = pd.read_excel(file, sheet_name=sheet_name, skiprows=2, header=[0, 1])
    
    # 2. 다중 헤더를 한 줄로 합치기 (예: '1학년' + '1학기' -> '1학년_1학기')
    new_columns = []
    for col in df.columns:
        # 'Unnamed'가 포함된 빈 상위 헤더는 무시하고 이름 조합
        col_name = "_".join([str(c) for c in col if 'Unnamed' not in str(c)])
        new_columns.append(col_name)
    df.columns = new_columns
    
    # 3. 병합된 셀(카테고리)의 빈칸 채우기 (Forward Fill)
    # 표의 첫 3열(구분, 교과(군), 과목 유형)은 병합되어 있으므로, 빈칸을 위에서부터 아래로 채움
    cols_to_ffill = df.columns[:3] 
    df[cols_to_ffill] = df[cols_to_ffill].ffill()
    
    # 4. 하단 '유의사항' 등 데이터가 아닌 행 제거 (과목명이 없는 행 삭제)
    # 과목명(보통 4번째 열)이 비어있거나 '유의사항' 같은 텍스트가 들어간 찌꺼기 행 제거
    subject_col = df.columns[3] 
    df = df.dropna(subset=[subject_col])
    
    # 5. 학점 빈칸(NaN)을 보기 좋게 하이픈(-) 또는 빈칸으로 변경
    df = df.fillna("-")
    
    return df
# ------------------------

uploaded_files = st.file_uploader(
    "교육과정 엑셀 파일을 업로드하세요", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        xls = pd.ExcelFile(file)
        
        selected_sheet = st.selectbox(f"📄 '{file.name}'에서 변환할 시트를 선택하세요:", xls.sheet_names)
        
        # 일반 read_excel 대신 새로 만든 전처리 함수 사용
        df = preprocess_curriculum_data(file, selected_sheet)
        md_text = df.to_markdown(index=False)
        
        st.subheader(f"미리보기: [{selected_sheet}] 시트")
        st.code(md_text, language='markdown')
        
        st.download_button(
            label="Markdown 다운로드",
            data=md_text,
            file_name=f"{file.name.split('.')[0]}_{selected_sheet}.md",
            mime="text/markdown"
        )

    else:
        st.subheader(f"📁 총 {len(uploaded_files)}개의 파일 설정")
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for file in uploaded_files:
                xls = pd.ExcelFile(file)
                selected_sheet = st.selectbox(
                    f"'{file.name}'의 시트 선택:", 
                    xls.sheet_names, 
                    key=file.name
                )
                
                # 전처리 함수 적용
                df = preprocess_curriculum_data(file, selected_sheet)
                md_text = df.to_markdown(index=False)
                
                md_filename = f"{file.name.split('.')[0]}_{selected_sheet}.md"
                zip_file.writestr(md_filename, md_text)
        
        st.success("모든 파일의 변환 준비가 완료되었습니다!")
        st.download_button(
            label="모든 변환 파일(ZIP) 다운로드",
            data=zip_buffer.getvalue(),
            file_name="converted_curriculum_md.zip",
            mime="application/zip"
        )
