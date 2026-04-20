import streamlit as st
import pandas as pd
import io
import zipfile

st.title("엑셀 ➡ Markdown 변환기")

# 파일 업로더 (여러 파일 허용)
uploaded_files = st.file_uploader(
    "엑셀 파일을 업로드하세요", 
    type=['xlsx'], 
    accept_multiple_files=True
)

if uploaded_files:
    # 1개만 올렸을 때: 미리보기 제공
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        df = pd.read_excel(file)
        md_text = df.to_markdown(index=False)
        
        st.subheader(f"📄 {file.name} 미리보기")
        st.code(md_text, language='markdown') # 미리보기 창
        
        st.download_button(
            label="Markdown 다운로드",
            data=md_text,
            file_name=f"{file.name.split('.')[0]}.md",
            mime="text/markdown"
        )

    # 여러 개 올렸을 때: ZIP 압축 파일 제공
    else:
        st.subheader(f"📁 총 {len(uploaded_files)}개의 파일 처리 중...")
        
        # 압축 파일을 담을 메모리 버퍼
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for file in uploaded_files:
                df = pd.read_excel(file)
                md_text = df.to_markdown(index=False)
                
                # md 파일 이름을 생성하여 압축 파일 내에 저장
                md_filename = f"{file.name.split('.')[0]}.md"
                zip_file.writestr(md_filename, md_text)
        
        st.success("모든 파일이 변환되었습니다.")
        
        st.download_button(
            label="모든 변환 파일(ZIP) 다운로드",
            data=zip_buffer.getvalue(),
            file_name="converted_markdown_files.zip",
            mime="application/zip"
        )
