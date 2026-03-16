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
        result = df.iloc[:, 0].sum()  # 첫 번째 열의 합계
        st.write(f"연산된 합계 값: {result}")

        # 3. 템플릿 로드 (상대 경로 사용)
        template_path = os.path.join("templates", "template_B.xlsx")
        
        if os.path.exists(template_path):
            wb = load_workbook(template_path)
            ws = wb.active
            
            # 특정 셀(B2)에 값 기록
            ws['B2'] = result
            
            # 4. 메모리에서 엑셀 파일 생성 (서버 파일 생성 방지)
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            # 5. 다운로드 파일 이름 설정 (파일명_Report.xlsx)
            base_name = os.path.splitext(uploaded_file.name)[0]
            download_name = f"{base_name}_Report.xlsx"
            
            # 다운로드 버튼
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
