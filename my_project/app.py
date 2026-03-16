import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

# 1. 파일 업로드
uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    # 2. 로딩 시각화 (spinner)
    with st.spinner('파일을 연산하고 보고서를 생성 중입니다...'):
        try:
            # 데이터 연산
            df = pd.read_excel(uploaded_file)
            result = df.iloc[:, 0].sum()
            
            # 템플릿 로드
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active

                
                # 메모리에서 엑셀 파일 생성
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                # 다운로드 버튼
                base_name = os.path.splitext(uploaded_file.name)[0]
                download_name = f"{base_name}_Report.xlsx"
                
                st.download_button(
                    label="결과 파일 다운로드",
                    data=output,
                    file_name=download_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(f"템플릿 파일을 찾을 수 없습니다.")
                
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
