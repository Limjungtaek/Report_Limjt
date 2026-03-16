import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('파일을 연산 중입니다...'):
        try:
            # 1. 데이터 읽기 (header=None으로 하여 첫 행부터 데이터로 인식)
            df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            
            # 2. D열(인덱스 3)의 6행(인덱스 5)부터 10000행까지
            target_range = df.iloc[5:10000, 3]
            
            # [디버깅] 실제 데이터 확인용 (화면에 출력)
            st.write("D열 6행부터의 실제 데이터 샘플:", target_range.head(10).tolist())
            
            # 3. 데이터 카운팅
            # 데이터 타입을 문자열로 통일하고 공백 제거
            clean_data = target_range.astype(str).str.strip()
            
            count_normal = (clean_data == '정상').sum()
            count_closed = (clean_data == '폐업').sum()
            
            st.write(f"카운트 결과 - 정상: {count_normal}, 폐업: {count_closed}")

            # 4. 템플릿 로드 및 저장
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                ws['B7'] = count_normal
                ws['F7'] = count_closed
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                download_name = f"{os.path.splitext(uploaded_file.name)[0]}_Report.xlsx"
                st.download_button("결과 파일 다운로드", output, download_name)
            else:
                st.error("템플릿 파일을 찾을 수 없습니다.")
                
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
