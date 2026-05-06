import streamlit as st
import pandas as pd
import io
import zipfile

st.title("엑셀 ➡️ 마크다운(MD) 변환기")
st.write("여러 개의 엑셀 파일을 일괄 변환하고 미리보기를 확인할 수 있습니다.")

# 다중 파일 업로드 기능
uploaded_files = st.file_uploader("엑셀 파일(.xlsx)을 여러 개 업로드해주세요.", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    # 첫 번째 파일을 기준으로 시트 목록 추출 (모든 파일의 시트 구성이 동일하다는 조건)
    first_file = uploaded_files[0]
    xls = pd.ExcelFile(first_file)
    sheet_names = xls.sheet_names

    st.markdown("---")
    # 조건 1 & 2: 변환할 시트 선택 (한 번만 선택)
    selected_sheet = st.selectbox("변환할 시트를 선택하세요 (모든 파일에 동일 적용):", sheet_names)

    st.markdown("---")
    # 조건 5 & 6: 파일 개수에 따른 미리보기 UI 분기
    if len(uploaded_files) == 1:
        preview_file = uploaded_files[0]
        st.subheader(f"📄 미리보기: {preview_file.name}")
    else:
        preview_file_name = st.selectbox("미리보기를 확인할 파일을 선택하세요:", [f.name for f in uploaded_files])
        preview_file = next(f for f in uploaded_files if f.name == preview_file_name)

    # 조건 3: 병합된 셀 처리
    # 마크다운 표 자체는 셀 병합 기능을 지원하지 않습니다.
    # 따라서 pandas가 병합된 셀을 읽을 때 생기는 빈 공간(NaN)을 빈 문자열("")로 치환하여 표 형태가 깨지지 않도록 구분합니다.
    df_preview = pd.read_excel(preview_file, sheet_name=selected_sheet)
    df_preview = df_preview.fillna("") 

    # 화면에 데이터프레임 및 마크다운 코드 미리보기 출력
    st.dataframe(df_preview)
    st.markdown("##### 마크다운 변환 결과")
    st.code(df_preview.to_markdown(index=False), language="markdown")

    st.markdown("---")
    # 일괄 변환 및 다운로드 기능 (ZIP 압축)
    if st.button("모든 파일 마크다운으로 변환하여 다운로드"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file in uploaded_files:
                df = pd.read_excel(file, sheet_name=selected_sheet)
                df = df.fillna("")
                md_content = df.to_markdown(index=False)
                
                # 기존 엑셀 파일명에서 확장자만 .md로 변경
                new_filename = file.name.rsplit(".", 1)[0] + ".md"
                zip_file.writestr(new_filename, md_content)
        
        st.download_button(
            label="📦 전체 파일(.zip) 다운로드",
            data=zip_buffer.getvalue(),
            file_name="markdown_files.zip",
            mime="application/zip"
        )
