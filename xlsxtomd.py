import streamlit as st
import pandas as pd
import io
import zipfile

st.title("엑셀 ➡ Markdown 변환기 (시트 선택 기능)")

# 파일 업로더 (여러 파일 허용)
uploaded_files = st.file_uploader(
    "엑셀 파일을 업로드하세요", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True
)

if uploaded_files:
    # ----------------------------------------------------
    # 1. 단일 파일 업로드 시 (미리보기 + 시트 선택)
    # ----------------------------------------------------
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        
        # 엑셀 파일의 시트 목록만 먼저 불러오기
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        
        # 사용자에게 시트 선택하게 하기
        selected_sheet = st.selectbox(f"📄 '{file.name}'에서 변환할 시트를 선택하세요:", sheet_names)
        
        # 선택한 시트의 데이터만 읽기
        df = pd.read_excel(file, sheet_name=selected_sheet)
        md_text = df.to_markdown(index=False)
        
        st.subheader(f"미리보기: [{selected_sheet}] 시트")
        st.code(md_text, language='markdown')
        
        # 파일명에 시트 이름을 포함하여 저장되도록 설정
        base_name = file.name.split('.')[0]
        
        st.download_button(
            label="Markdown 다운로드",
            data=md_text,
            file_name=f"{base_name}_{selected_sheet}.md",
            mime="text/markdown"
        )

    # ----------------------------------------------------
    # 2. 다중 파일 업로드 시 (개별 시트 선택 + ZIP 다운로드)
    # ----------------------------------------------------
    else:
        st.subheader(f"📁 총 {len(uploaded_files)}개의 파일 설정")
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            # 업로드된 각 파일마다 시트 선택 UI를 제공
            for file in uploaded_files:
                xls = pd.ExcelFile(file)
                sheet_names = xls.sheet_names
                
                # key 속성에 파일명을 넣어 각 selectbox가 독립적으로 동작하게 함
                selected_sheet = st.selectbox(
                    f"'{file.name}'의 시트 선택:", 
                    sheet_names, 
                    key=file.name
                )
                
                df = pd.read_excel(file, sheet_name=selected_sheet)
                md_text = df.to_markdown(index=False)
                
                # 압축 파일 내부에 들어갈 파일명 (예: data_Sheet1.md)
                base_name = file.name.split('.')[0]
                md_filename = f"{base_name}_{selected_sheet}.md"
                zip_file.writestr(md_filename, md_text)
        
        st.success("모든 파일의 변환 준비가 완료되었습니다!")
        
        st.download_button(
            label="모든 변환 파일(ZIP) 다운로드",
            data=zip_buffer.getvalue(),
            file_name="converted_markdown_files.zip",
            mime="application/zip"
        )
