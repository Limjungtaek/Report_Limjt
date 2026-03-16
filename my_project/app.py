import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

# 1. 파일 업로드
uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    try:
        # 2. 데이터 연산
        df = pd.read_excel(uploaded_file)
        result = df.iloc[:, 0].sum()
        st.write(f"연산된 합계 값: {result}")

        # 3. 템플릿 로드 (절대 경로 강제 지정)
        # app.py가 위치한 폴더의 경로를 가져와서 templates 폴더를 찾습니다.
        base_path = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_path, "templates", "template_B.xlsx")


        if os.path.exists(template_path):
            wb = load_workbook(template_path)
            ws = wb.active
            ws['B2'] = result
            
            # 4. 메모리에서 엑셀 파일 생성
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            # 5. 다운로드 버튼
            base_name = os.path.splitext(uploaded_file.name)[0]
            download_name = f"{base_name}_Report.xlsx"
            
            st.download_button(
                label="결과 파일 다운로드",
                data=output,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error(f"템플릿 파일을 찾을 수 없습니다: {template_path}")
            
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
