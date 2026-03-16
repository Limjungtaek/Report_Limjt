import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('데이터를 처리 중입니다...'):
        try:
            # 1. 데이터 읽기
            # header=None으로 D열(인덱스 3)에 직접 접근
            df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            
            # 2. D열(인덱스 3)의 6행(인덱스 5)부터 10000행까지 데이터 추출
            target_range = df.iloc[5:10000, 3]
            
            # 3. 데이터 연산
            # B7: 비어있지 않은 모든 셀 개수 (COUNTA)
            count_total = target_range.count()
            
            # 문자열로 변환 후 공백 제거하여 '정상'/'폐업' 판별
            clean_data = target_range.astype(str).str.strip()
            # D7: '정상' 개수
            count_normal = len(clean_data[clean_data == '정상'])
            # F7: '폐업' 개수
            count_closed = len(clean_data[clean_data == '폐업'])
            
            # 4. 템플릿 처리
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # B7(전체 개수), D7(정상), F7(폐업)에 기입
                ws['B7'] = count_total
                ws['D7'] = count_normal
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
