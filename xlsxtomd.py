import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def handle_merged_cells(file, sheet_name):
    """셀 병합을 해제하고 병합된 값을 모든 셀에 채워넣는 함수"""
    wb = openpyxl.load_workbook(file, data_only=True)
    sheet = wb[sheet_name]
    
    # 병합된 셀 정보 파악 및 값 복사
    merged_cells = list(sheet.merged_cells.ranges)
    for merged_range in merged_cells:
        # 병합된 영역의 첫 번째 셀 값 가져오기
        min_col, min_row, max_col, max_row = merged_range.bounds
        top_left_value = sheet.cell(row=min_row, column=min_col).value
        
        # 병합 해제 후 모든 셀에 값 채우기
        sheet.unmerge_cells(str(merged_range))
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                sheet.cell(row=row, column=col).value = top_left_value
                
    # 처리된 데이터를 DataFrame으로 변환
    data = sheet.values
    cols = next(data)
    df = pd.DataFrame(data, columns=cols)
    return df

st.set_page_config(page_title="Excel to MD Converter", layout="wide")
st.title("📊 엑셀 → AI 분석용 MD 변환기")
st.info("여러 개의 엑셀 파일을 업로드하면 동일한 시트를 찾아 MD 파일로 변환해 드립니다.")

# 1. 파일 업로드
uploaded_files = st.file_uploader("엑셀 파일들을 선택하세요 (여러 개 가능)", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    # 첫 번째 파일에서 시트 목록 가져오기
    first_file = uploaded_files[0]
    temp_wb = openpyxl.load_workbook(first_file)
    sheet_names = temp_wb.sheetnames
    
    # 2. 시트 선택 (모든 파일에 공통 적용)
    selected_sheet = st.selectbox("변환할 시트를 선택하세요", sheet_names)
    
    if st.button("변환 시작"):
        for uploaded_file in uploaded_files:
            try:
                st.write(f"📄 {uploaded_file.name} 처리 중...")
                
                # 파일 객체를 BytesIO로 다시 읽기 (Openpyxl 사용을 위해)
                file_content = BytesIO(uploaded_file.getvalue())
                
                # 셀 병합 처리하여 DataFrame 생성
                df = handle_merged_cells(file_content, selected_sheet)
                
                # 결측치 처리 (AI 분석 시 깨끗한 데이터를 위해)
                df = df.fillna("")
                
                # 3. MD 변환 및 미리보기
                md_output = df.to_markdown(index=False)
                
                # 파일명 생성
                md_filename = uploaded_file.name.replace(".xlsx", f"_{selected_sheet}.md")
                
                # 다운로드 버튼 생성
                st.download_button(
                    label=f"📥 {md_filename} 다운로드",
                    data=md_output,
                    file_name=md_filename,
                    mime="text/markdown"
                )
                
                with st.expander(f"{uploaded_file.name} 미리보기"):
                    st.code(md_output, language='markdown')
                    
            except Exception as e:
                st.error(f"{uploaded_file.name} 처리 중 오류 발생: {e}")
