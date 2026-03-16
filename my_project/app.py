import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    # 데이터 연산
    df = pd.read_excel(uploaded_file)
    result = df.iloc[:, 0].sum()
    
    template_path = os.path.join("templates", "template_B.xlsx")
    
    if os.path.exists(template_path):
        wb = load_workbook(template_path)
        ws = wb.active
        ws['B2'] = result
        
        output_path = "result_B.xlsx"
        wb.save(output_path)
        
        # 파일명 동적 생성 (파일명 + _Report)
        base_name = os.path.splitext(uploaded_file.name)[0]
        download_name = f"{base_name}_Report.xlsx"
        
        with open(output_path, "rb") as f:
            st.download_button(
                label="결과 파일 다운로드",
                data=f,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
